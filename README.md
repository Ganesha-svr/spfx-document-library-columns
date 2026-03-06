# SPFx Document Library Search — v4
### Product Document Library Web Part

A SharePoint Framework (SPFx) web part that searches across **all nested folders** of any document library on your site, tailored for a product document library with the following columns:

| Column | SharePoint Internal Name | Type |
|--------|--------------------------|------|
| **Title** | `Title` | Single line of text |
| **Name** | `FileLeafRef` | File name (built-in) |
| **Product No** | `ProductNo` | Single line of text |
| **Brand No** | `BrandNo` | Text or Choice |
| **Description** | `Description0` *(see note below)* | Multi-line text |
| **File Name** | `FileLeafRef` | Built-in file name |

> ⚠️ **Description column note:** SharePoint sometimes appends `0` to the internal name of a custom Description column to avoid a clash with its built-in field. Go to **Library Settings → click the Description column → check the URL for `Field=`** to confirm whether your internal name is `Description` or `Description0`, and update `CAML_FIELDS` in `DocumentLibraryService.ts` accordingly.

---

## Features

| Feature | Details |
|---------|---------|
| 🔍 **Smart Search** | Searches across Title, Name, Product No, Brand No, Description simultaneously |
| 🔁 **Recursive Folder Search** | Finds documents in ANY nested subfolder at any depth using CAML `RecursiveAll` |
| 📚 **Library Picker** | Auto-discovers all document libraries on the site — no hardcoding needed |
| 🏷️ **Brand No Filter** | Filter chips auto-loaded from your library's BrandNo field values |
| 🔗 **Direct File Links** | Every result card links directly to open the file |
| 📁 **Full Folder Path** | Shows complete nested path e.g. `Products › Brand A › 2024 › Q1` |
| ✕ **Clear Button** | One-click clear resets search text, filters, and results |
| 🔎 **Highlight Matches** | Search terms highlighted in yellow across all result fields |
| 📋 **Expandable Details** | Click any card to reveal all column values + file URL |
| ⚡ **Debounced Search** | Auto-searches 400ms after you stop typing (min 2 characters) |

---

## Project Structure

```
spfx-document-library-columns/
├── config/
│   ├── config.json                              # Bundle entry point
│   └── package-solution.json                   # Solution packaging config
├── src/
│   ├── typings/
│   │   └── scss.d.ts                           # Global SCSS module type declaration
│   └── webparts/
│       └── documentLibraryColumns/
│           ├── components/
│           │   ├── DocumentLibraryColumns.tsx          # Main React component (UI)
│           │   ├── DocumentLibraryColumns.module.scss  # Styles
│           │   ├── DocumentLibraryColumns.module.scss.ts  # SCSS type declarations (fixes TS error)
│           │   └── IDocumentItem.ts                    # TypeScript interfaces
│           ├── services/
│           │   └── DocumentLibraryService.ts           # SharePoint REST + CAML API service
│           ├── DocumentLibraryColumnsWebPart.ts        # SPFx WebPart entry point
│           └── DocumentLibraryColumnsWebPart.manifest.json
├── gulpfile.js
├── package.json
├── tsconfig.json
└── README.md
```

---

## Prerequisites

- **Node.js** v16 LTS or v18 LTS
- **SPFx** 1.18.2 compatible environment
- **SharePoint Online** tenant
- Yeoman & Gulp CLI:

```bash
npm install -g yo gulp-cli @microsoft/generator-sharepoint
```

---

## Installation

```bash
# 1. Navigate into the project folder
cd spfx-document-library-columns

# 2. Install all dependencies
npm install

# 3. Start local dev server (workbench)
gulp serve
```

Then open your SharePoint workbench:
```
https://<your-tenant>.sharepoint.com/sites/<yoursite>/_layouts/15/workbench.aspx
```

---

## Build & Deploy

```bash
# Production build
gulp bundle --ship

# Package the solution
gulp package-solution --ship
```

This produces:
```
sharepoint/solution/spfx-document-library-columns.sppkg
```

**Deploy to SharePoint:**
1. Go to your **App Catalog**: `https://<tenant>.sharepoint.com/sites/appcatalog/AppCatalog`
2. Upload the `.sppkg` file
3. Check **"Make this solution available to all sites"** for tenant-wide deployment
4. Add the **"Document Library Columns"** web part to any modern SharePoint page

---

## Configuration (Property Pane)

Click the ✏️ edit pencil on the web part to access:

| Setting | Description | Default |
|---------|-------------|---------|
| **Site URL** | Full URL for a different site's libraries. Leave blank for current site. | *(current site)* |

---

## Customising Column Internal Names

If your SharePoint column internal names differ from the defaults, open:

```
src/webparts/documentLibraryColumns/services/DocumentLibraryService.ts
```

Update the `CAML_FIELDS` array:

```typescript
private readonly CAML_FIELDS = [
  "ID",
  "FileLeafRef",
  "FileRef",
  "Title",
  "ProductNo",       // ← change to your actual internal name
  "BrandNo",         // ← change to your actual internal name
  "Description0",    // ← change to "Description" if needed
  "FileDirRef",
  "FSObjType",
  "Created",
  "Modified",
  "Editor",
  "File_x0020_Size",
  "ContentType",
  "Author"
];
```

Also update the mapping in `_mapToDocumentItem()`:

```typescript
productNo:   item["ProductNo"]    || "",   // ← match your internal name
brandNo:     item["BrandNo"]      || "",   // ← match your internal name
description: item["Description0"] || item["Description"] || "",
```

### How to find a column's internal name

1. Go to **Library Settings** in SharePoint
2. Click the column name
3. Look at the browser URL — the `Field=` parameter is the internal name

Or use the REST API directly:
```
https://<site>/_api/web/lists/getByTitle('<LibraryName>')/fields?$select=Title,InternalName&$filter=Hidden eq false
```

---

## How Recursive Search Works

The key to searching nested folders is the **CAML query with `Scope="RecursiveAll"`**:

```xml
<View Scope="RecursiveAll">
  <Query>
    <Where>
      <Or>
        <Contains><FieldRef Name="FileLeafRef"/><Value Type="Text">invoice</Value></Contains>
        <Contains><FieldRef Name="ProductNo"/><Value Type="Text">invoice</Value></Contains>
      </Or>
    </Where>
  </Query>
</View>
```

This is sent as a **POST** to `/_api/web/lists/getByTitle('Library')/getItems` with the header:
```
Prefer: HonorNonIndexedQueriesWarningMayFailOnLargeList
```

| Approach | Root only | Nested folders |
|----------|-----------|---------------|
| Old: `GET /items?$filter=...` | ✅ | ❌ |
| New: `POST /getItems` CAML RecursiveAll | ✅ | ✅ |

---

## Fixing "Cannot find module *.module.scss" Error

This solution includes the fix already. Every SCSS module needs two files:

```
DocumentLibraryColumns.module.scss       ← styles
DocumentLibraryColumns.module.scss.ts    ← TypeScript type declarations  ✅ included
src/typings/scss.d.ts                    ← global wildcard fallback       ✅ included
```

If you add a **new component** with its own SCSS module, create a matching `.module.scss.ts` file listing all class names as `readonly string`.

---

## Search Behaviour

| User types… | Searches in… |
|-------------|-------------|
| `"3M"` | Title, Name, ProductNo, BrandNo, Description |
| `"PRD-001"` | Title, Name, **ProductNo**, BrandNo, Description |
| `"Scotch"` | Title, Name, ProductNo, **BrandNo**, Description |
| `"adhesive tape"` | Title, Name, ProductNo, BrandNo, **Description** |

All searches are **case-insensitive** and use `<Contains>` (substring match), so partial terms work.

---

## Version History

| Version | Changes |
|---------|---------|
| **v4** | Updated for product library: Title, Name, Product No, Brand No, Description, FileName columns |
| **v3** | CAML RecursiveAll for nested folder search; full folder path breadcrumb |
| **v2** | Auto-discover libraries; enhanced search UI; expandable result cards; clear button |
| **v1** | Initial release with basic column display |

---

## License

MIT
