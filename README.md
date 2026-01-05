# Outlook Inbox Analyzer Add-in

An Outlook Add-in that automatically analyzes your mailbox and shows email volume breakdown by sender. No CSV import needed - it pulls data directly from Outlook.

## Features

- 📊 **Automatic Analysis** - Scans emails directly from your Outlook mailbox
- 📬 **Sender Breakdown** - See who sends you the most email
- 📈 **Visual Stats** - Summary stats and bar charts for quick insights
- 📥 **CSV Export** - Export results for further analysis
- 📁 **Multiple Folders** - Analyze Inbox, Sent, Drafts, Deleted, or Archive

## Quick Start

### Prerequisites

- Node.js 16+ installed
- Outlook (Microsoft 365, Outlook.com, or Exchange)
- HTTPS for local development (the add-in requires secure connections)

### Installation

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Install dev certificates** (required for HTTPS):
   ```bash
   npx office-addin-dev-certs install
   ```

3. **Start the development server:**
   ```bash
   npm run dev
   ```
   This starts a local server at `https://localhost:3000`

### Sideload the Add-in

#### Option A: Outlook on the Web
1. Go to https://outlook.office.com
2. Open any email
3. Click the "..." menu → "Get Add-ins"
4. Click "My add-ins" → "Add a custom add-in" → "Add from file"
5. Upload the `manifest.xml` file

#### Option B: Outlook Desktop (Windows)
1. Open Outlook
2. Go to File → Manage Add-ins (opens in browser)
3. Click "Add a custom add-in" → "Add from file"
4. Upload the `manifest.xml` file

#### Option C: Using Office Add-in tools
```bash
npm run sideload
```

### Usage

1. Open any email in Outlook
2. Look for the "Analyze Inbox" button in the ribbon
3. Click it to open the analyzer panel
4. Select a folder and how many emails to scan
5. Click "Analyze" to see results

## Project Structure

```
outlook-inbox-analyzer/
├── manifest.xml          # Add-in manifest (config for Outlook)
├── package.json          # npm dependencies
├── src/
│   ├── taskpane.html     # Main UI
│   ├── taskpane.js       # Analysis logic
│   └── functions.html    # Required function file
└── assets/               # Icons (you'll need to add these)
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

## How It Works

The add-in uses Office.js to access the mailbox:

1. **EWS (Exchange Web Services)** - Primary method, works with Exchange/Microsoft 365
2. **REST API Fallback** - For newer Outlook versions that support it

It fetches emails from the selected folder, extracts sender information, and aggregates the counts without storing any email content.

## Configuration

### Changing the Default Scan Limit

Edit `src/taskpane.html` and modify the `<select id="limit">` options.

### Adding More Folders

Edit the `<select id="folder">` in `taskpane.html`. Available folder IDs:
- `inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`
- `junkemail`, `outbox`, `calendar`, `contacts`, `tasks`

## Deployment

For production deployment:

1. Host the files on a web server with HTTPS
2. Update all URLs in `manifest.xml` to point to your server
3. Deploy through Microsoft 365 Admin Center or AppSource

## Troubleshooting

**"Failed to fetch emails" error:**
- Make sure you're using an Exchange or Microsoft 365 mailbox
- Check browser console for detailed errors
- Try a smaller scan limit

**Add-in not appearing:**
- Verify the dev server is running on port 3000
- Check that HTTPS certificates are installed
- Clear Outlook's add-in cache

**CORS errors:**
- The dev server should have CORS enabled
- Make sure manifest URLs match your server

## License

MIT
