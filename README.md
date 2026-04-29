# Word Auto-Save

A Microsoft Word add-in that saves your local `.docx` files automatically on a configurable interval.

Word's built-in AutoSave only works with cloud storage (OneDrive / SharePoint). This add-in fills that gap for files stored locally on your Mac — no server, no Node.js, no background process. The static files are hosted on Netlify; Word fetches the JS once when a document opens and runs the timer entirely inside its own WebView.

## How it works

- On every document open, Word fires `OnDocumentOpened` and starts a save timer in a **shared runtime** (so the timer stays alive for the whole session).
- Every N minutes (default: 5), the add-in calls `context.document.save()` — identical to pressing ⌘S.
- New unsaved documents (no file path yet) are skipped automatically — no Save As dialog is triggered.
- A ribbon button in the **Home** tab opens a small settings panel where you can change the interval or save immediately.

## Installation (Mac)

One-time setup per Mac. Requires an internet connection the first time a document is opened (Word fetches the JS from Netlify); subsequent saves in that session work offline.

```bash
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef

curl -o ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml \
  https://word-auto-save.netlify.app/manifest.xml
```

Then **restart Word**. That's it.

The manifest persists across reboots — no repeat steps needed unless you uninstall Word.

## Uninstall

```bash
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

Restart Word.

## Changing the save interval

1. Open any `.docx` file.
2. Click **Auto-Save Settings** in the Home ribbon.
3. Enter the interval in minutes and click **Apply**.

The setting is stored in the browser's `localStorage` and survives restarts.

## Deploying to Netlify

The repo is deployed as a static site — no build step needed.

1. Push this repo to GitHub.
2. In Netlify: **Add new site → Import an existing project** → connect the repo.
3. Leave build command and publish directory blank (Netlify serves the repo root as-is).
4. After deploy, confirm your site URL (e.g. `https://word-auto-save.netlify.app`).
5. If the URL differs, update the three `https://word-auto-save.netlify.app` references in `manifest.xml` and redeploy.

## File structure

```
word-auto-save/
├── manifest.xml          # Sideloaded into Word's wef folder on each Mac
├── assets/
│   ├── icon-32.png
│   └── icon-80.png
└── src/
    ├── commands.js       # OnDocumentOpened handler + save timer
    ├── taskpane.html     # Settings UI
    ├── taskpane.js       # Taskpane logic
    └── taskpane.css
```

## Verification

| Test | Expected result |
|------|----------------|
| Open a `.docx`, wait N minutes | File modification date updates in Finder |
| Click Save Now | File modification date updates immediately |
| Open a new unsaved document | No crash, no Save As dialog |
| Reboot Mac, open Word | Add-in still loads automatically |
| Disconnect internet after first open | Saves continue for that session |
