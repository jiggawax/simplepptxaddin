# Slide+ PowerPoint Add-in

A minimal PowerPoint add-in with a single button to add a new blank slide.

---

## Files

| File | Purpose |
|------|---------|
| `taskpane.html` | The add-in UI (self-contained, no build step needed) |
| `manifest.xml` | Office Add-in manifest — tells PowerPoint about the plugin |

---

## Setup (3 steps)

### 1. Host `taskpane.html`

The file must be served over **HTTPS**. Easiest options:

- **GitHub Pages** — push the file to a repo, enable Pages, use the generated URL
- **Netlify / Vercel** — drag-and-drop deploy
- **localhost with HTTPS** — use `npx office-addin-dev-certs install` then `npx http-server -S -C cert.pem`

### 2. Update `manifest.xml`

Replace both occurrences of `https://YOUR-SERVER/taskpane.html` with your actual URL.

### 3. Sideload the add-in into PowerPoint

**Windows (Desktop):**
1. Copy `manifest.xml` to a shared network folder
2. In PowerPoint → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs
3. Add the folder path → OK
4. Insert → My Add-ins → Shared Folder → pick "Slide+"

**Mac (Desktop):**
1. Copy `manifest.xml` to: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
2. Restart PowerPoint → Insert → My Add-ins

**PowerPoint Online:**
1. Insert → Add-ins → Upload My Add-in → select `manifest.xml`

---

## Usage

Once loaded, a **"Slide+"** panel opens. Click the big **ADD SLIDE** button to insert a blank slide after the last slide. The status bar confirms success and shows the current slide count.

---

## Local Dev (no server needed)

```bash
npm install -g office-addin-dev-certs http-server
office-addin-dev-certs install
http-server . -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```

Then set the manifest URL to `https://localhost:3000/taskpane.html`.
