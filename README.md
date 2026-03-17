# ✉️ AI Reply Assistant — Desktop Outlook (Windows)

An Office Add-in that puts an AI reply drafting panel inside **Desktop Outlook for Windows**.  
Powered by Claude. **Nothing is ever sent automatically.**

---

## What you'll need

- Desktop Outlook for Windows (Microsoft 365 / Outlook 2019+)
- An Anthropic API key — free at https://console.anthropic.com

---

## Two ways to install

### Option A — GitHub Pages (recommended, free, permanent)

This hosts your add-in files online so Outlook can always reach them.

1. **Create a free GitHub account** at https://github.com if you don't have one

2. **Create a new repository** — click the `+` icon → New repository  
   Name it anything, e.g. `outlook-ai-reply`. Set it to **Public**.

3. **Upload all files** from this folder into the repository  
   (drag and drop them onto the GitHub page, or use the "uploading an existing file" link)

4. **Enable GitHub Pages**  
   Go to your repo → Settings → Pages → Source: select `main` branch → Save  
   Your URL will be: `https://YOUR-USERNAME.github.io/outlook-ai-reply`

5. **Edit `manifest.xml`** — open it in Notepad and replace every instance of:
   ```
   YOUR_GITHUB_PAGES_URL
   ```
   with your actual URL, e.g.:
   ```
   https://myname.github.io/outlook-ai-reply
   ```
   Save the file.

6. **Run `install.ps1`** — right-click it → "Run with PowerShell"  
   It copies the manifest into Outlook's trusted add-in folder automatically.

7. Restart Outlook — the **✦ AI Reply Draft** button appears in the Home ribbon.

---

### Option B — Local server (for testing, no GitHub needed)

1. Install **Node.js** from https://nodejs.org
2. Double-click **`start-local-server.bat`** — this starts a local HTTPS server and patches the manifest automatically
3. Run **`install.ps1`** to sideload the manifest
4. Restart Outlook
5. Keep the server window open while using the add-in

---

## First-time setup in Outlook

1. Open any email
2. Click **✦ AI Reply Draft** in the **Home** ribbon
3. The AI panel opens on the right — enter your Anthropic API key and click **Save**

---

## How to use

1. Open an email in Outlook
2. Click **✦ AI Reply Draft** in the ribbon
3. Pick a **tone** — Professional, Friendly, Concise, Formal, Empathetic, or Assertive
4. Click **Generate Draft Reply**
5. Read the draft in the panel
6. Click **↳ Insert into Reply** — Outlook opens a reply window with the draft pre-filled
7. Edit as you like, then **send manually when ready**

Use the **ON/OFF toggle** at the top to pause the add-in anytime.

---

## Uninstall

Delete the file at:
```
%APPDATA%\Microsoft\Outlook\Wef\AIReplyAssistant.xml
```
Then restart Outlook.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Button doesn't appear | Restart Outlook; check the manifest was copied to `%APPDATA%\Microsoft\Outlook\Wef\` |
| "Add-in couldn't be loaded" | Make sure the hosting URL in `manifest.xml` is correct and reachable |
| API key error | Confirm the key starts with `sk-ant-` and is active at console.anthropic.com |
| PowerShell security error | Right-click `install.ps1` → Properties → Unblock → OK, then try again |
