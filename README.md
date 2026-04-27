# Merge Documents — SPFx ListViewCommandSet Extension

![version](https://img.shields.io/badge/SPFx-1.22.2-green.svg)

A SharePoint Framework (SPFx) **ListViewCommandSet** extension that lets users select multiple documents in a Document Library, reorder them, and bundle them into a single merged document via a Power Automate flow.

---

## Features

- **Multi-select awareness** — the "Merge Documents" button only appears in the command bar when 2 or more documents are selected.
- **Drag-to-reorder dialog** — a modal lets users drag documents into the desired merge order before sending.
- **Blocking loading screen** — while the Power Automate flow runs, a full-screen spinner prevents accidental interaction.
- **Result dialog with link** — on success, a dialog displays the merged document URL with an "Open Document" button that opens it in a new tab.
- **Authenticated flow trigger** — the flow is called using `AadHttpClientFactory` (user's own identity, not a shared connection). See [this blog post](https://cognicoast.com/blogs/Execute%20Power%20Automate%20workflow%20from%20SPFx%20NOT%20as%20Anyone.html) for background.

---

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | 22.x |
| SPFx | 1.22.2 |
| React | 17.0.1 |
| FluentUI React | 8.x |

---

## Configuration

Before building, open [src/extensions/mergeWebPart/MergeWebPartCommandSet.ts](src/extensions/mergeWebPart/MergeWebPartCommandSet.ts) and update the two constants near the top of the file:

```typescript
// Power Automate HTTP trigger URL — replace with your flow's HTTP POST URL
const FLOW_URL: string = 'https://<YOUR-FLOW-TRIGGER-URL>';

// Full URL of the Document Library where merged documents are saved
// e.g. https://<tenant>.sharepoint.com/sites/<site>/Shared%20Documents
const MERGED_DOCS_LIBRARY_URL: string = 'https://<YOUR-TENANT>.sharepoint.com/sites/<YOUR-SITE>/<YOUR-LIBRARY>';
```

| Constant | Description |
|---|---|
| `FLOW_URL` | The HTTP POST trigger URL from your Power Automate flow. Found in the **"When an HTTP request is received"** trigger of the flow. |
| `MERGED_DOCS_LIBRARY_URL` | Full URL of the Document Library where the merged file is saved. Used as a fallback base URL when the flow returns only a filename. |

---

## Power Automate Flow — Expected Response

See [this blog Easily Combine Multiple SharePoint Files into One PDF Power Automate No Code Solution](https://clavinfernandes.wordpress.com/2025/11/16/execute-power-automate-workflow-from-spfx-not-as-anyone/) 

The flow must return a JSON body from a **"Respond to a PowerApp or flow"** action. The extension recognises any of the following output property names (checked in order):

| Property name | Example value |
|---|---|
| `mergedFileUrl` | `https://contoso.sharepoint.com/sites/HR/Docs/Merged.pdf` |
| `fileUrl` | `https://contoso.sharepoint.com/sites/HR/Docs/Merged.pdf` |
| `fileName` | `Merged.pdf` |
| `name` | `Merged.pdf` |
| `url` | `https://contoso.sharepoint.com/sites/HR/Docs/Merged.pdf` |

**Recommendation:** use `mergedFileUrl` with the full absolute URL — Power Automate's SharePoint "Create file" action exposes this directly as **"Link to item"**.

If the returned value is a relative path or filename only, the extension automatically prepends `MERGED_DOCS_LIBRARY_URL` to construct the full link.

---

## Payload Sent to the Flow

The extension POSTs the following JSON to `FLOW_URL`:

```json
{
  "fileUrls": [
    "/sites/HR/Shared Documents/DocumentA.docx",
    "/sites/HR/Shared Documents/DocumentB.docx"
  ]
}
```

`fileUrls` is an ordered array of **server-relative URLs** in the order the user set in the drag-to-reorder dialog.

---

## Local Development

1. Install dependencies:
   ```bash
   npm install
   ```

2. Update `config/serve.json` — set `pageUrl` to a Document Library `AllItems.aspx` page on your tenant:
   ```json
   "pageUrl": "https://<YOUR-TENANT>.sharepoint.com/sites/<YOUR-SITE>/Shared%20Documents/Forms/AllItems.aspx"
   ```

3. Start the dev server:
   ```bash
   npm start
   ```

4. When the browser opens, click **Load debug scripts** to activate the extension.

5. Select **2 or more documents** — the **Merge Documents** button will appear in the command bar.

---

## Deployment

1. Update the two constants in `MergeWebPartCommandSet.ts` (see [Configuration](#configuration) above).

2. Build and package:
   ```bash
   npm run build
   ```
   This produces `sharepoint/solution/merge-web-part.sppkg`.

3. Upload `merge-web-part.sppkg` to the **SharePoint App Catalog** (Tenant or Site Collection).

4. When prompted, click **Deploy** to make it available tenant-wide.

5. **First-time only:** approve the API permission for `Microsoft Flow Service / Flows.Read.All` in **SharePoint Admin Centre → Advanced → API access**.

> **Note:** The extension targets **Document Libraries only** (`ListTemplateId="101"`). It will not appear on generic lists.

---

## Updating the Deployment

After any code or configuration change, increment the `version` in `config/package-solution.json` before repackaging:

```json
"version": "1.0.1.0"
```

This ensures SharePoint detects the update and upgrades the deployed feature.

---

## File Structure

```
src/extensions/mergeWebPart/
├── MergeWebPartCommandSet.ts     # Main command set — visibility logic & flow trigger
├── MergeDocumentsDialog.tsx      # Modal with drag-to-reorder document list
├── MergeLoadingDialog.tsx        # Blocking spinner shown while flow is running
├── MergeResultDialog.tsx         # Result modal with merged document link
└── loc/
    ├── en-us.js                  # English strings
    └── myStrings.d.ts            # Localisation type definitions

sharepoint/assets/
├── elements.xml                  # Custom action registration (Document Library — template 101)
└── ClientSideInstance.xml        # Tenant-wide auto-deployment configuration

config/
├── package-solution.json         # Solution metadata and API permission requests
└── serve.json                    # Local debug server configuration
```

---

## Solution

| Solution | Author |
|---|---|
| Merge Documents SPFx Extension | Clavin Fernandes |

## Version History

| Version | Date | Comments |
|---|---|---|
| 1.0.0 | April 2026 | Initial release |

---

## References

- [Using Custom Dialogs with SPFx](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/guidance/using-custom-dialogs-with-spfx)
- [Execute Power Automate from SPFx (not as anyone)](https://clavinfernandes.wordpress.com/2025/11/16/execute-powerautomate-workflow-from-spfx-not-as-anyone/)
- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Heft Documentation](https://heft.rushstack.io/)
