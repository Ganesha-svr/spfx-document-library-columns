// DocumentLibraryColumnsWebPart.ts  v5 — with diagnostics
import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import DocumentLibraryColumns from "./components/DocumentLibraryColumns";

export interface IWebPartProps { siteUrl: string; }

export default class DocumentLibraryColumnsWebPart extends BaseClientSideWebPart<IWebPartProps> {

  // Run a quick connectivity check before rendering the component
  // This surfaces the exact HTTP error code so we know what is failing
  private async _runDiagnostics(): Promise<string | null> {
    try {
      const siteUrl = (this.properties.siteUrl || this.context.pageContext.web.absoluteUrl).replace(/\/$/, "");
      const testUrl = `${siteUrl}/_api/web?$select=Title`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        testUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
          }
        }
      );

      if (!response.ok) {
        const body = await response.text();
        return `HTTP ${response.status} — ${response.statusText}\n\nDetails: ${body.substring(0, 300)}`;
      }

      return null; // no error — all good
    } catch (e) {
      return `Connection error: ${e.message}`;
    }
  }

  public async render(): Promise<void> {
    const diagError = await this._runDiagnostics();

    if (diagError) {
      // Show a clear diagnostic panel instead of a broken web part
      this.domElement.innerHTML = `
        <div style="
          font-family: 'Segoe UI', sans-serif;
          padding: 20px;
          background: #fff;
          border-radius: 8px;
          border-left: 5px solid #d13438;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          max-width: 700px;
        ">
          <div style="font-size:20px; font-weight:700; color:#d13438; margin-bottom:12px;">
            ⚠️ Document Library Web Part — Connection Error
          </div>
          <div style="font-size:13px; color:#333; margin-bottom:16px;">
            The web part could not connect to SharePoint. See the error below:
          </div>
          <pre style="
            background:#fef0f0;
            border:1px solid #f4c0c0;
            border-radius:6px;
            padding:12px;
            font-size:12px;
            color:#a4262c;
            white-space:pre-wrap;
            word-break:break-all;
          ">${diagError}</pre>
          <div style="margin-top:16px; font-size:13px; color:#555;">
            <strong>Common causes:</strong><br/>
            &bull; <strong>401</strong> — Not logged in or session expired. Refresh the page.<br/>
            &bull; <strong>403</strong> — No permission to access this site. Check site permissions.<br/>
            &bull; <strong>406</strong> — Accept header rejected. Update GET_HEADERS in DocumentLibraryService.ts.<br/>
            &bull; <strong>404</strong> — Site URL is wrong. Check the Site URL in web part properties.<br/>
            &bull; <strong>Connection error</strong> — Network issue or CORS blocked.<br/>
          </div>
          <div style="margin-top:16px; padding:12px; background:#f0f8ff; border-radius:6px; font-size:12px; color:#333;">
            <strong>Current Site URL being used:</strong><br/>
            <code>${this.properties.siteUrl || this.context.pageContext.web.absoluteUrl}</code><br/><br/>
            <strong>Logged in user:</strong><br/>
            <code>${this.context.pageContext.user.displayName} (${this.context.pageContext.user.loginName})</code>
          </div>
        </div>`;
      return;
    }

    // Diagnostics passed — render the normal component
    const element = React.createElement(DocumentLibraryColumns, {
      context: this.context,
      siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }

  protected get dataVersion(): Version { return Version.parse("2.0"); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: "Configure Document Library Search" },
        groups: [{
          groupName: "Settings",
          groupFields: [
            PropertyPaneTextField("siteUrl", {
              label: "Site URL (optional)",
              description: "Leave blank to use the current site. Enter full URL for a different site.",
              placeholder: "https://tenant.sharepoint.com/sites/mysite"
            })
          ]
        }]
      }]
    };
  }
}
