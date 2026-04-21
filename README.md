# SCORM Generator — PowerPoint Add-in

A PowerPoint add-in that converts the currently-open presentation into a SCORM 1.2 or 2004 package. Runs fully offline — no slide content ever leaves your machine.

---

## What you get

- **Ribbon button** in PowerPoint's Home tab labeled *"Export as SCORM"*
- **Task pane** with full configuration: course title, branding colors, layout (scroll/paginated), videos, content protection, completion tracking
- **One-click import** of the current presentation's slides (parses titles, text, bullets, images, embedded videos, card layouts, per-shape geometry)
- **Downloadable SCORM zip** ready to upload to Moodle, Canvas, Cornerstone, or any SCORM-compliant LMS

---

## Quick setup (for small-team sideload)

You need three things: (1) host the files on HTTPS, (2) update the manifest URLs, (3) sideload in PowerPoint. Takes about 10 minutes.

### Step 1 — Host the files on GitHub Pages (free, HTTPS built in)

1. Create a new public GitHub repository, e.g. `scorm-addin`
2. Copy these files into the repo root:
   - `taskpane.html`
   - `icons/icon-16.png`
   - `icons/icon-32.png`
   - `icons/icon-64.png`
   - `icons/icon-80.png`
3. Go to the repo's **Settings → Pages**, under "Source" select "Deploy from a branch", pick `main` branch and `/` (root), click Save.
4. After ~60 seconds GitHub Pages will publish at `https://YOUR-USERNAME.github.io/scorm-addin/`. Verify by opening `https://YOUR-USERNAME.github.io/scorm-addin/taskpane.html` — it should load (with a warning that Office.js can't connect outside PowerPoint, which is expected).

### Step 2 — Update the manifest with your URL

Open `manifest.xml` in a text editor. Find and replace every occurrence of:

```
https://USER.github.io/ADDIN/
```

with your actual GitHub Pages URL:

```
https://YOUR-USERNAME.github.io/scorm-addin/
```

There are 8 URLs to replace. Save the file.

Also replace the `<Id>` value at the top with a fresh GUID. Generate one at https://www.guidgenerator.com/ — paste the new value between `<Id>` and `</Id>`.

### Step 3 — Sideload the manifest in PowerPoint

**On Windows:**

1. Open PowerPoint (any version 2016 or newer, including Microsoft 365)
2. Open any `.pptx` file
3. Click **Insert → My Add-ins**
4. In the dialog, click **Upload My Add-in** (bottom-left link)
5. Browse to and select your `manifest.xml`
6. PowerPoint validates it and adds the ribbon button

**On Mac:**

1. Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (create the `wef` folder if missing)
2. Restart PowerPoint
3. Click **Insert → My Add-ins → Shared Folder** to see and activate your add-in

**For an entire team (shared folder):**

1. Put `manifest.xml` on a network share folder
2. Each user adds the shared folder as a trusted location under **File → Options → Trust Center → Trusted Add-in Catalogs**
3. Restart PowerPoint. The add-in appears under **Insert → My Add-ins → Shared Folder**

---

## Using the add-in

1. Open any `.pptx` in PowerPoint
2. Click **Home → Export as SCORM** (your new ribbon button)
3. Task pane opens on the right
4. Click **Import from current presentation** — slides are read into the SCO manager
5. Configure course title, branding, videos, content protection etc. in the task pane
6. Click **Build & Download** → SCORM zip downloads
7. Upload the zip to your LMS

---

## Updating the add-in later

When you make changes, just push new files to your GitHub repo. The next time someone opens PowerPoint and uses the add-in, they get the latest version automatically (cached for up to 24 hours — force refresh with `Ctrl+F5` in the task pane).

No need for users to reinstall.

---

## Troubleshooting

**"This add-in could not be started"**
- Double-check that `taskpane.html` is accessible at the URL in your manifest. Open it directly in a browser.
- Make sure the URL is **HTTPS**, not HTTP. PowerPoint rejects non-HTTPS add-ins.
- Check that your manifest XML is valid — run it through Microsoft's [Office Add-in Validator](https://aka.ms/check-manifest) if unsure.

**Ribbon button doesn't appear**
- Confirm the manifest specifies `<Host Name="Presentation" />` (not `Document` or `Workbook`)
- On Mac, make sure `manifest.xml` is in the exact path above and PowerPoint is restarted
- On Windows, try **Insert → My Add-ins → Manage My Add-ins** to see if it was registered

**"Office.js failed to load"**
- The user is probably opening `taskpane.html` in a regular browser. It only works inside PowerPoint.

**Import reads 0 slides**
- The .pptx file must be saved first (not just "unsaved new presentation")
- Legacy binary `.ppt` files aren't supported — save as `.pptx` first

---

## Files in this package

```
ppt_addin/
├── manifest.xml         ← upload this to PowerPoint to install
├── taskpane.html        ← host this at an HTTPS URL (GitHub Pages / Azure / etc.)
├── icons/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
└── README.md            ← this file
```

---

## Privacy

All slide parsing and SCORM packaging happens 100% in the task pane (your machine's browser engine). The only external resource loaded is Microsoft's own `office.js` from `appsforoffice.microsoft.com` (required for add-ins to work). No slide data, no text, no images are ever sent anywhere.
