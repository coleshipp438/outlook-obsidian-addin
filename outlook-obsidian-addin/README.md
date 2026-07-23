# Outlook → Obsidian Add-in

This package creates an Outlook for Mac button that sends the open email to:

- Vault: `Cole Command Center`
- File: `02 Tasks/Task Inbox.md`
- Heading: `Work`

The task is scheduled using the date selected in the task pane.

## Files

- `manifest.xml`: file you sideload into Outlook
- `taskpane.html`, `taskpane.js`, `taskpane.css`: hosted add-in
- `config.js`: Obsidian destination settings
- icon files: Outlook button artwork

## Deploy with GitHub Pages

1. Create a new public GitHub repository named `outlook-obsidian-addin`.
2. Upload every file in this folder to the repository root.
3. In GitHub, open **Settings → Pages**.
4. Under **Build and deployment**, select **Deploy from a branch**.
5. Select branch `main`, folder `/ (root)`, then save.
6. Wait for GitHub Pages to publish the site.
7. Your site will normally be:
   `https://YOUR-GITHUB-USERNAME.github.io/outlook-obsidian-addin/`

## Update the manifest

Open `manifest.xml` in a plain-text editor.

Replace every occurrence of:

`YOUR-GITHUB-USERNAME`

with your actual GitHub username.

Save the file after replacing all occurrences.

## Sideload into Outlook

1. Open Outlook and open the Add-ins dialog.
2. Choose **Add a custom add-in** or **Add from File**.
3. Select the edited `manifest.xml`.
4. Accept the prompts.
5. Open an email.
6. Click **Add to Obsidian**.

Microsoft no longer supports Add from URL for manual sideloading, so use the XML file.

## First test

1. Open an email.
2. Open the add-in.
3. Confirm the proposed task and date.
4. Click **Add task to Obsidian**.
5. Approve opening Obsidian if macOS asks.
6. Check `02 Tasks/Task Inbox.md` beneath `## Work`.

## If Outlook blocks the Obsidian link

Click **Copy Obsidian link**, paste it into Safari or Chrome, and approve opening Obsidian.

## Change the destination later

Edit `config.js`:

```js
window.ADDIN_CONFIG = {
  vaultName: "Cole Command Center",
  taskFile: "02 Tasks/Task Inbox.md",
  taskHeading: "Work"
};
```

## Privacy

The add-in runs from the static files you host. It does not send email data to a server created by this package. It only reads the currently opened message through Office.js and constructs an Obsidian URI on your Mac.
