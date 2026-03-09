# SharePoint Slideshow

A self-contained, single-file web app that authenticates with Microsoft 365 and displays images from a SharePoint list as a full-screen slideshow. Built for kiosk or presentation use cases where images are managed in SharePoint.

## Development
This project was built with the assistance of [Claude](https://claude.ai) by Anthropic.

## How It Works

1. The user opens the page and signs in with their Microsoft 365 account.
2. The app fetches all unique category values from the configured SharePoint list.
3. The user selects a category (or "All images") to begin the slideshow.
4. Images are displayed full-screen with a blurred background fill, progress bar, and optional captions.
5. Tapping/clicking the screen reveals playback controls and a settings panel.

## Features

- Microsoft 365 / Azure AD authentication (MSAL)
- Reads images from any SharePoint list via Microsoft Graph API
- Category filter screen populated dynamically from list data
- Sorted or randomized playback order (persisted in localStorage)
- Configurable slide interval with quick-select preset chips (long-press a chip to customize it)
- Fine-tune interval with +/− buttons (adjusts in configurable steps)
- HUD overlay (tap to show/hide) with previous, next, pause, and settings controls
- Settings panel to change interval, order, category, refresh images, or sign out
- Smooth fade transitions and blurred ambient background

## Prerequisites

- An **Azure AD app registration** with the `Sites.Read.All` delegated permission granted
- The redirect URI of the app registration must match where the file is hosted (e.g. `https://yourdomain.com/slideshow.html`)
- A **SharePoint list** containing an image column and (optionally) category and title columns

## Configuration

All configuration is in the `CONFIG` block near the top of `index.html`:

```js
const CONFIG = {
  // Azure AD app registration
  tenantId:   "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  clientId:   "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy",
  redirectUri: window.location.origin + window.location.pathname,

  // SharePoint
  siteUrl:        "https://yourorg.sharepoint.com/sites/YourSite",
  listName:       "Your List Name",
  imageColumn:    "Thumbnail",
  categoryColumn: "Category",
  titleColumn:    "",

  // UX defaults
  defaultIntervalSec: 60,
  minIntervalSec:      5,
  maxIntervalSec:   3600,
  intervalStepSec:     5,

  // Sort columns
  sortColumn1: "Date",
  sortColumn2: "SlideNumber",
};
```

### Configuration Reference

| Key | Description |
|-----|-------------|
| `tenantId` | Azure AD tenant ID (GUID) from your app registration |
| `clientId` | Application (client) ID (GUID) from your app registration |
| `redirectUri` | OAuth redirect URI — defaults to the current page URL; must match the app registration |
| `siteUrl` | Full URL to the SharePoint site (not the list, just the site root) |
| `listName` | Internal or display name of the SharePoint list containing the images |
| `imageColumn` | Internal column name of the image field (e.g. `Thumbnail`, `Image0`) |
| `categoryColumn` | Internal column name used to populate the category filter screen |
| `titleColumn` | Internal column name to use as a slide caption; set to `""` to disable captions |
| `defaultIntervalSec` | Default seconds per slide on first load |
| `minIntervalSec` | Minimum allowed slide interval in seconds |
| `maxIntervalSec` | Maximum allowed slide interval in seconds |
| `intervalStepSec` | How many seconds the +/− buttons adjust the interval by |
| `sortColumn1` | Primary sort field (internal SharePoint name); sorted descending |
| `sortColumn2` | Secondary sort field (internal SharePoint name); sorted ascending |

### Finding Internal Column Names

Internal column names differ from display names. To find them:

1. Go to your SharePoint list → **Settings** → **List settings**
2. Click on a column under "Columns"
3. The internal name appears in the URL as the `Field` parameter

Alternatively, use the Graph API Explorer to browse `GET /sites/{siteId}/lists/{listId}/items?$expand=fields` and inspect the `fields` object keys.

## Deployment

The app is a single `index.html` file with no build step. Host it anywhere the browser can reach:

- SharePoint itself (as a page or in a document library)
- GitHub Pages
- Azure Static Web Apps
- Any web server or CDN

Make sure the hosting URL is added as a **redirect URI** in the Azure AD app registration under **Authentication**.

## Runtime Settings (persisted in localStorage)

| Setting | Storage key | Default |
|---------|-------------|---------|
| Slide order (sorted/random) | `slideshow_order` | `sorted` |
| Custom preset intervals | `slideshow_presets` | `[5, 10, 60, 300, 3600]` |

# Setup Guide

## Prerequisites
- A Microsoft 365 account with access to SharePoint
- A GitHub account
- Your SharePoint list already created with the required columns (see [SharePoint List Requirements](#sharepoint-list-requirements))

---

## 1. Azure App Registration

The slideshow authenticates users via Microsoft 365 using OAuth. You need to register an application in Microsoft Entra ID (formerly Azure Active Directory) to enable this.

### Create the App Registration

1. Go to [portal.azure.com](https://portal.azure.com) and sign in
2. Click **Microsoft Entra ID** on the home screen (or search for it in the top bar)
3. In the left sidebar, click **App registrations**
4. Click **+ New registration** at the top
5. Fill in the form:
   - **Name:** something descriptive, e.g. `SharePoint Slideshow`
   - **Supported account types:** select **Accounts in this organizational directory only**
   - **Redirect URI:** leave blank for now — you will add it in the next step
6. Click **Register**

### Add the Redirect URI

1. In your new app registration, click **Authentication** in the left sidebar
2. Click **+ Add a platform**
3. Choose **Single-page application (SPA)** — this is critical; do not choose "Web"
4. Enter your redirect URI:
   - For GitHub Pages: `https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/`
   - For local development: `http://localhost:5500/index.html`
5. Click **Configure**
6. You can add multiple redirect URIs to the same SPA platform entry — add both the GitHub Pages URL and your localhost URL so both work

### Add API Permission

1. Click **API permissions** in the left sidebar
2. Click **+ Add a permission**
3. Choose **Microsoft Graph**
4. Choose **Delegated permissions**
5. Search for `Sites.Read.All` and check it
6. Click **Add permissions**
7. If you are a Microsoft 365 admin, click **Grant admin consent** to avoid users seeing a consent prompt on first login. If you are not an admin, users will see a one-time consent screen when they first sign in.

### Note Your IDs

From the app registration **Overview** page, copy:
- **Application (client) ID** — you will put this in the `CONFIG` block as `clientId`
- **Directory (tenant) ID** — you will put this in the `CONFIG` block as `tenantId`

Neither of these values is a secret. It is safe to hardcode them in a public repository.

---

## 2. Create a copy of this Repo and Enable GitHub Pages

1. Click **Fork** at the top right of this repository on GitHub
2. GitHub creates a copy under your own account

### Configure the App

Open `index.html` and fill in the `CONFIG` block near the top of the file:

```javascript
const CONFIG = {
  tenantId:         "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",  // from Entra ID Overview
  clientId:         "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy",  // from App Registration Overview
  redirectUri:      "https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/",

  siteUrl:          "https://YOUR-TENANT.sharepoint.com/sites/YOUR-SITE",
  listName:         "YOUR-LIST-INTERNAL-NAME",
  imageColumn:      "YOUR-IMAGE-COLUMN-INTERNAL-NAME",
  categoryColumn:   "YOUR-CATEGORY-COLUMN-INTERNAL-NAME",
  titleColumn:      "Title",

  sortColumn1:      "Date",
  sortColumn2:      "Slide_x0023_",
  ...
};
```

> **Finding internal column names:** In your SharePoint list, go to **List Settings** (gear icon → List Settings), then click on the column name. The internal name appears in the browser URL after `&Field=`.

Commit and push this change to your repository.

### Enable GitHub Pages

1. Go to your repository on GitHub
2. Click **Settings**
3. Click **Pages** in the left sidebar
4. Under **Source**, select **Deploy from a branch**
5. Set the branch to `main` and the folder to `/ (root)`
6. Click **Save**
7. GitHub will display your Pages URL — it will be in the format:
   ```
   https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/
   ```
8. Wait 1–3 minutes for the first build. Check the **Actions** tab in your repo — a green checkmark on the `pages build and deployment` workflow confirms it is live.

> **Note:** GitHub Pages requires the repository to be **public** on a free GitHub account. The `clientId`, `tenantId`, and SharePoint URLs in your config are safe to expose publicly — they are not secrets. Never commit a client secret or password.

---

## 3. SharePoint List Requirements

### Required Columns

| Display Name | Internal Name | Type | Purpose |
|---|---|---|---|
| Image | *(check your list)* | Image | The slide image |
| Category | *(check your list)* | Choice | Used to filter slides by category |

### Optional Columns

| Display Name | Internal Name | Type | Purpose |
|---|---|---|---|
| Title | `Title` | Single line of text | Displayed as the slide caption |
| Date | `Date` | Date only | Primary sort — newest dates shown first |
| Slide # | `Slide_x0023_` | Number (integer) | Secondary sort — ordering within a date |

### Adding an Index to a Column

SharePoint limits filtering and sorting via the API to **indexed columns** when a list contains more than 5,000 items. Even below that threshold, indexes improve query performance and prevent errors as your list grows. You should index any column used in a `$filter` or `$orderby` query.

**Columns that need an index in this app:**
- `Category` — used for `$filter`
- `Date` — used for `$orderby`
- `Slide #` (`Slide_x0023_`) — used for `$orderby`

**How to add an index:**

1. Go to your SharePoint list
2. Click the gear icon → **List Settings**
3. Scroll down to the **Columns** section and click **Indexed columns**
4. Click **Create a new index**
5. Under **Primary column**, select the column you want to index (e.g. `Date`)
6. Leave **Secondary column** as `(none)` — a simple single-column index is sufficient
7. Click **Create**
8. Repeat for each column that needs an index

> You can also add a **compound index** on `Date` + `Slide #` together, which can improve performance for the combined sort query. To do this, set `Date` as the primary column and `Slide_x0023_` as the secondary column when creating the index.

---

## 4. Local Development

To test changes without waiting for GitHub Pages to rebuild:

1. Install the **Live Server** extension in VS Code (by Ritwick Dey)
2. In VS Code Settings, search for **Live Server Host** and change the value from `127.0.0.1` to `localhost`
3. Right-click `index.html` → **Open with Live Server**
4. The app opens at `http://localhost:5500/index.html`
5. Make sure this URL is added as a redirect URI in your Azure app registration (under the SPA platform)

Changes you save in VS Code are reflected in the browser immediately — no commit or push required until you are ready to deploy.