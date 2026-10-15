
## 1. Start the local server
Run these in `c:\Development\Templates\Add-in\addin`:
```powershell
npm install
npm run dev-server
```
This serves the add-in at `https://localhost:3000`. If the browser or Outlook rejects the HTTPS certificate, install the dev certificate first:
```powershell
npx office-addin-dev-certs install
```
Leave the server running while you use the add-in.

## 2. Sideload it (pick one)

**Option A: the automated way.** This starts the dev server and registers the add-in for you, using the desktop Outlook it finds:
```powershell
npm start          # uses manifest.json
npm run start:xml  # uses manifest.xml
npm run stop       # unregisters it when you're done
```

**Option B: Outlook on the web or new Outlook (manual).**
1. Open https://aka.ms/olksideload. Alternatively, in Outlook go to **Get Add-ins** → **My add-ins** → **Custom Addins** → **Add a custom add-in** → **Add from file**.
2. Choose `manifest.xml`. Use the XML file here, since the manual upload dialog is most reliable with it.
3. Accept the warning.

**Option C: classic desktop Outlook (Windows).**
1. Open a mail item, then choose **Get Add-ins**. You can also use Home → **Get Add-ins**.
2. Go to **My add-ins** → **Custom Addins** → **Add a custom add-in** → **Add from file**, and select `manifest.xml`.
3. Accept the warning.

## 3. Use it
Open or compose a new message. The **Email Templates** button should appear on the ribbon. In a compose window, it may be under **…** (More).

## Troubleshooting
- **Blank task pane:** check that `https://localhost:3000/taskpane.html` loads in a browser without certificate errors.
- **Add-in doesn't appear:** restart Outlook. Changes to the manifest can take a few minutes to show up.
- **Cached old version:** remove the add-in and sideload it again.
- **Sideloading is blocked:** your tenant admin may have disabled custom add-ins. In that case an admin has to deploy it from the Microsoft 365 admin center under **Integrated apps**.
- **Before sharing it:** replace the `localhost:3000` URLs in the manifests with your real hosting URL. The package description mentions GitHub Pages.