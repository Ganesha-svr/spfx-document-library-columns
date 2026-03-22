// DocumentLibraryService.ts  (v10 — List View Threshold Fix)
//
// SharePoint blocks queries on lists >5000 items unless:
//   1. The WHERE clause only uses INDEXED columns, OR
//   2. No WHERE clause at all (full scan with RowLimit)
//
// Strategy:
//   - Fetch pages of 100 items ordered by Modified (indexed by default)
//   - Apply text search CLIENT-SIDE after fetching each page
//   - Stop fetching when enough results found OR no more pages
//   - Brand filter uses Eq on an indexed field (if BrandNo is indexed)
//   - Max pages fetched = 20 (2000 items) to prevent timeout

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IDocumentItem, ILibraryInfo } from "../components/IDocumentItem";

const POST_OPTIONS: ISPHttpClientOptions = {
  headers: {
    "Accept": "application/json",
    "Content-Type": "application/json",
    "Prefer": "HonorNonIndexedQueriesWarningMayFailOnLargeList"
  }
};

const PAGE_SIZE    = 100;   // items per page — safe for threshold
const MAX_RESULTS  = 200;   // stop after this many matches found
const MAX_PAGES    = 20;    // never fetch more than 20 pages (2000 items)

const BUILTIN_FIELDS = [
  "ID", "FileLeafRef", "FileRef", "Title",
  "FileDirRef", "FSObjType", "Created", "Modified",
  "Editor", "File_x0020_Size", "ContentType", "Author",
  "HTML_x0020_File_x0020_Type"
];

const CUSTOM_FIELD_CANDIDATES: { [key: string]: string[] } = {
  productNo:   ["ProductNo", "Product_x0020_No", "ProductNumber", "Product_No"],
  brandNo:     ["BrandNo",   "Brand_x0020_No",   "BrandNumber",   "Brand_No"],
  description: ["Description0", "Description", "Description1", "Desc"]
};

export class DocumentLibraryService {
  private context: WebPartContext;
  private siteUrl: string;
  private siteOrigin: string;
  private fieldCache: { [lib: string]: { [key: string]: string } } = {};

  constructor(context: WebPartContext, siteUrl?: string) {
    this.context = context;
    this.siteUrl = (siteUrl || context.pageContext.web.absoluteUrl).replace(/\/$/, "");
    this.siteOrigin = this.siteUrl.match(/^https?:\/\/[^/]+/)?.[0] || "";
  }

  // ─── GET helper — no custom headers ───────────────────────────────────────

  private async _get(url: string): Promise<any> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url, SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`HTTP ${response.status}: ${text.substring(0, 200)}`);
    }
    return response.json();
  }

  // ─── POST helper ──────────────────────────────────────────────────────────

  private async _post(url: string, body: string): Promise<any> {
    const response = await this.context.spHttpClient.post(
      url, SPHttpClient.configurations.v1,
      { ...POST_OPTIONS, body }
    );
    if (!response.ok) {
      const err = await response.text();
      throw new Error(`HTTP ${response.status}: ${err.substring(0, 300)}`);
    }
    return response.json();
  }

  // ─── Field discovery ──────────────────────────────────────────────────────

  private async _discoverFields(libraryTitle: string): Promise<{ [key: string]: string }> {
    if (this.fieldCache[libraryTitle]) return this.fieldCache[libraryTitle];
    const result: { [key: string]: string } = {};
    try {
      const url  =
        `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(libraryTitle)}')` +
        `/fields?$select=InternalName&$filter=Hidden eq false`;
      const data = await this._get(url);
      const names: string[] = (data?.value || data?.d?.results || [])
        .map((f: any) => f.InternalName as string);
      for (const key of Object.keys(CUSTOM_FIELD_CANDIDATES)) {
        const found = CUSTOM_FIELD_CANDIDATES[key].find(c =>
          names.some(n => n.toLowerCase() === c.toLowerCase())
        );
        if (found) {
          result[key] = names.find(n => n.toLowerCase() === found.toLowerCase()) || found;
        }
      }
    } catch (e) {
      console.warn("[Service] Field discovery failed:", e);
    }
    this.fieldCache[libraryTitle] = result;
    return result;
  }

  // ─── Client-side text match ────────────────────────────────────────────────

  private _matchesSearch(item: IDocumentItem, searchText: string): boolean {
    if (!searchText || !searchText.trim()) return true;
    const q = searchText.trim().toLowerCase();
    return (
      item.fileName.toLowerCase().indexOf(q)    !== -1 ||
      item.title.toLowerCase().indexOf(q)       !== -1 ||
      item.productNo.toLowerCase().indexOf(q)   !== -1 ||
      item.brandNo.toLowerCase().indexOf(q)     !== -1 ||
      item.description.toLowerCase().indexOf(q) !== -1 ||
      item.folderPath.toLowerCase().indexOf(q)  !== -1
    );
  }

  // ─── Thumbnail / Folder helpers ───────────────────────────────────────────

  public getThumbnailUrl(serverRelativeUrl: string, fileExt: string): string {
    if (!serverRelativeUrl) return "";
    const exts = ["png","jpg","jpeg","gif","bmp","webp","pdf","doc","docx","xls","xlsx","ppt","pptx"];
    if (exts.indexOf(fileExt.toLowerCase()) !== -1) {
      return `${this.siteOrigin}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(serverRelativeUrl)}&resolution=0`;
    }
    return "";
  }

  public getFolderUrl(fileDirRef: string): string {
    return fileDirRef ? `${this.siteOrigin}${fileDirRef}` : "";
  }

  // ─── List all Document Libraries ──────────────────────────────────────────

  public async getAllDocumentLibraries(): Promise<ILibraryInfo[]> {
    const url =
      `${this.siteUrl}/_api/web/lists` +
      `?$filter=BaseType eq 1 and Hidden eq false and IsCatalog eq false` +
      `&$select=Id,Title,ItemCount,LastItemModifiedDate,DefaultViewUrl,Description` +
      `&$orderby=Title asc`;
    const data  = await this._get(url);
    const lists = data?.value || data?.d?.results || [];
    return lists.map((lib: any): ILibraryInfo => ({
      id:             lib.Id,
      title:          lib.Title,
      itemCount:      lib.ItemCount || 0,
      lastModified:   lib.LastItemModifiedDate,
      defaultViewUrl: lib.DefaultViewUrl || "",
      description:    lib.Description   || ""
    }));
  }

  // ─── Search Items — paged fetch + client-side filter ─────────────────────
  // Avoids threshold by: no Contains in CAML, uses indexed Modified orderby,
  // fetches PAGE_SIZE items per request, filters client-side

  public async searchItems(
    libraryTitle: string,
    searchText: string,
    brandFilter: string
  ): Promise<IDocumentItem[]> {
    const fields    = await this._discoverFields(libraryTitle);
    const allFields = [
      ...BUILTIN_FIELDS,
      ...(fields.productNo   ? [fields.productNo]   : []),
      ...(fields.brandNo     ? [fields.brandNo]     : []),
      ...(fields.description ? [fields.description] : [])
    ];
    const viewFields = allFields.map(f => `<FieldRef Name="${f}"/>`).join("");
    const apiUrl     = `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(libraryTitle)}')/getItems`;

    const results:  IDocumentItem[] = [];
    let   position: string | null   = null;  // ListItemCollectionPosition token
    let   page      = 0;

    while (page < MAX_PAGES && results.length < MAX_RESULTS) {
      page++;

      // Build CAML — NO WHERE clause (avoids threshold on non-indexed columns)
      // OrderBy Modified which is indexed by default
      const camlQuery = `
        <View Scope="RecursiveAll">
          <ViewFields>${viewFields}</ViewFields>
          <Query>
            <OrderBy><FieldRef Name="Modified" Ascending="FALSE"/></OrderBy>
          </Query>
          <RowLimit Paged="TRUE">${PAGE_SIZE}</RowLimit>
        </View>`;

      const queryObj: any = { ViewXml: camlQuery };
      if (position) queryObj.ListItemCollectionPosition = position;

      const data     = await this._post(apiUrl, JSON.stringify({ query: queryObj }));
      const rawItems = data?.value || data?.d?.results || [];

      // Map all raw items
      const mapped = rawItems
        .map((item: any) => this._mapItem(item, libraryTitle, fields))
        .filter((item: IDocumentItem) => item.fileName !== "" || item.isDocumentSet);

      // Client-side filter: search text + brand filter
      const filtered = mapped.filter((item: IDocumentItem) => {
        const matchesText  = this._matchesSearch(item, searchText);
        const matchesBrand = !brandFilter || brandFilter === "All"
          ? true
          : item.brandNo.toLowerCase() === brandFilter.toLowerCase();
        return matchesText && matchesBrand;
      });

      results.push(...filtered);

      // Get next page token
      const nextToken = data?.__next || data?.d?.__next || null;
      if (!nextToken || rawItems.length < PAGE_SIZE) break;  // no more pages
      position = nextToken;
    }

    return results.slice(0, MAX_RESULTS);
  }

  // ─── Document Set children ────────────────────────────────────────────────

  public async getDocumentSetChildren(
    libraryTitle: string,
    documentSetServerRelativeUrl: string
  ): Promise<IDocumentItem[]> {
    try {
      const fields     = await this._discoverFields(libraryTitle);
      const allFields  = [
        ...BUILTIN_FIELDS,
        ...(fields.productNo   ? [fields.productNo]   : []),
        ...(fields.brandNo     ? [fields.brandNo]     : []),
        ...(fields.description ? [fields.description] : [])
      ];
      const viewFields = allFields.map(f => `<FieldRef Name="${f}"/>`).join("");
      const camlQuery  = `
        <View Scope="RecursiveAll">
          <ViewFields>${viewFields}</ViewFields>
          <Query><OrderBy><FieldRef Name="FileLeafRef" Ascending="TRUE"/></OrderBy></Query>
          <RowLimit>200</RowLimit>
        </View>`;

      const apiUrl = `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(libraryTitle)}')/getItems`;
      const data   = await this._post(apiUrl, JSON.stringify({
        query: {
          ViewXml: camlQuery,
          FolderServerRelativeUrl: documentSetServerRelativeUrl
        }
      }));

      return (data?.value || data?.d?.results || [])
        .map((item: any) => this._mapItem(item, libraryTitle, fields))
        .filter((item: IDocumentItem) => !item.isDocumentSet && item.fileName !== "");
    } catch (e) {
      console.error("[Service] getDocumentSetChildren:", e);
      return [];
    }
  }

  // ─── Brand options ────────────────────────────────────────────────────────

  public async getBrandOptions(libraryTitle: string): Promise<string[]> {
    const fields = await this._discoverFields(libraryTitle);
    if (!fields.brandNo) return [];
    try {
      const url  =
        `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(libraryTitle)}')` +
        `/fields/getByInternalNameOrTitle('${fields.brandNo}')?$select=Choices,FieldTypeKind`;
      const data    = await this._get(url);
      const choices = data?.Choices?.results || data?.Choices || data?.d?.Choices?.results || data?.d?.Choices || [];
      if (choices.length > 0) return choices;
    } catch { /* fall through */ }
    try {
      const url  =
        `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(libraryTitle)}')` +
        `/items?$select=${fields.brandNo}&$top=200`;
      const data  = await this._get(url);
      const items = data?.value || data?.d?.results || [];
      const all   = items.map((i: any) => i[fields.brandNo]).filter((b: any) => b && b.toString().trim());
      return all.reduce((acc: string[], val: string) => {
        if (acc.indexOf(val) === -1) acc.push(val);
        return acc;
      }, []).sort();
    } catch { return []; }
  }

  // ─── Map raw item ─────────────────────────────────────────────────────────

  private _mapItem(item: any, libraryTitle: string, fields: { [key: string]: string }): IDocumentItem {
    const fileLeafRef   = item["FileLeafRef"] || "";
    const fileDirRef    = item["FileDirRef"]  || "";
    const serverRelUrl  = item["FileRef"]     || "";
    const htmlFileType  = (item["HTML_x0020_File_x0020_Type"] || "").toLowerCase();
    const contentType   = item["ContentType"] || "";
    const isFolder      = item["FSObjType"] === 1 || item["FSObjType"] === "1";
    const isDocumentSet = isFolder && (
      contentType.toLowerCase().indexOf("document set") !== -1 ||
      htmlFileType.indexOf("sharepoint.documentset") !== -1
    );
    const folderPath    = this._extractFolderPath(fileDirRef, libraryTitle);
    const folderName    = fileDirRef.split("/").pop() || "";
    const folderUrl     = this.getFolderUrl(fileDirRef);
    const fileExt       = fileLeafRef.split(".").pop()?.toLowerCase() || "";
    const thumbnailUrl  = isFolder ? "" : this.getThumbnailUrl(serverRelUrl, fileExt);
    const documentSetUrl = isDocumentSet ? `${this.siteOrigin}${serverRelUrl}/Forms/default.aspx` : "";

    const getTitle = (f: any): string => {
      if (!f) return "";
      if (typeof f === "string") return f;
      if (f.Title) return f.Title;
      if (f.results && f.results[0]) return f.results[0].Title || "";
      return "";
    };

    return {
      id:                item["ID"]  || item["Id"] || 0,
      title:             item["Title"]             || "",
      name:              item["Title"]             || fileLeafRef || "(No Name)",
      productNo:         fields.productNo   ? (item[fields.productNo]   || "") : "",
      brandNo:           fields.brandNo     ? (item[fields.brandNo]     || "") : "",
      description:       fields.description ? (item[fields.description] || "") : "",
      fileName:          fileLeafRef,
      folderName,        folderPath,         folderUrl,
      fileRef:           serverRelUrl ? `${this.siteOrigin}${serverRelUrl}` : "",
      serverRelativeUrl: serverRelUrl,
      libraryTitle,
      created:           item["Created"]           || "",
      modified:          item["Modified"]          || "",
      modifiedBy:        getTitle(item["Editor"]),
      fileSize:          item["File_x0020_Size"]   || 0,
      isFolder,          isDocumentSet,             documentSetUrl,
      contentType,       thumbnailUrl
    };
  }

  // ─── Folder path ──────────────────────────────────────────────────────────

  private _extractFolderPath(fileDirRef: string, libraryTitle: string): string {
    if (!fileDirRef) return "";
    const parts    = fileDirRef.split("/");
    const libIndex = parts.findIndex(p =>
      p.toLowerCase() === libraryTitle.toLowerCase() ||
      decodeURIComponent(p).toLowerCase() === libraryTitle.toLowerCase()
    );
    if (libIndex === -1) return parts[parts.length - 1] || "";
    const sub = parts.slice(libIndex + 1).filter(Boolean);
    return sub.length > 0 ? sub.join(" / ") : "(Root)";
  }

  private _escapeXml(str: string): string {
    return str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
              .replace(/"/g,"&quot;").replace(/'/g,"&apos;");
  }
}
