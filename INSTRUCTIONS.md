# Document Redaction Add-in - Setup and Usage Instructions

This add-in redacts sensitive information (emails, phone numbers, and SSNs) from Word documents, adds a confidentiality header, and tracks all changes.

## Prerequisites

- Node.js and npm installed on your machine
- A Microsoft account (for Word on the web) or Microsoft 365 subscription
- The test document: `Document-To-Be-Redacted.docx`

## Step-by-Step Setup Instructions

### 1. Install Dependencies

Open a terminal in the project directory and run:

```bash
npm install
```

This will install all required dependencies including TypeScript, Office.js types, and development tools.

### 2. Generate SSL Certificates (First Time Only)

The add-in requires HTTPS certificates for local development. Run:

```bash
npx office-addin-dev-certs install
```

Follow the prompts to install and trust the certificates. This only needs to be done once.

### 3. Start the Development Server

You need to run two commands in separate terminal windows:

**Terminal 1 - TypeScript Compiler (Watch Mode):**
```bash
npm run watch
```

This will continuously compile your TypeScript code as you make changes.

**Terminal 2 - Development Server:**
```bash
npm run dev-server
```

This starts the HTTPS server on port 3000. You should see:
```
Available on:
  https://127.0.0.1:3000
  https://localhost:3000
```

Keep both terminals running while developing.

### 4. Prepare Your Document

1. Go to [Office.com](https://office.com) and sign in with your Microsoft account
2. Navigate to **OneDrive**
3. Upload the `Document-To-Be-Redacted.docx` file to your OneDrive
4. Open the document in **Word on the web** (click on the file in OneDrive)

### 5. Enable Track Changes (IMPORTANT - Word Online Only)

**⚠️ Important Note:** This solution was developed and tested using **Word on the web** (Word Online) because the desktop version of Word was not available. Word Online has limited support for the Track Changes API, so you need to manually enable it:

1. In the Word document, click on the **Review** tab in the ribbon
2. Click the **Track Changes** button to turn it ON (it should be highlighted/active)
3. This ensures all modifications made by the add-in will be tracked

**Why is this necessary?** Word Online doesn't fully support the Track Changes API that would allow the add-in to automatically enable tracking. If you're using Word Desktop, the add-in will attempt to enable tracking automatically, but for Word Online, manual activation is required.

### 6. Sideload the Add-in

1. In Word on the web, click **Home** → **Add-ins** → **More Settings** (or **More add-ins**)
2. In the **Office Add-ins** dialog, click **Upload My Add-in**
3. Click **Browse** and navigate to your project folder
4. Select the `manifest.xml` file
5. Click **Upload**

The add-in task pane should now appear on the right side of your Word document.

### 7. Run the Redaction

1. Make sure **Track Changes is ON** (see Step 5)
2. In the add-in task pane, click the **Run Redaction** button
3. Watch the status messages as the add-in:
   - Analyzes the document
   - Enables track changes (if supported)
   - Adds the confidentiality header
   - Scans for sensitive information
   - Redacts found patterns

### 8. Verify Results

After the redaction completes, you should see:

- **Status message** showing how many items were redacted (e.g., "Redaction complete. Replaced 18 items (6 emails, 8 phones, 4 SSNs)")
- **Confidentiality header** at the top of the document (if it wasn't already there)
- **Redacted content** - sensitive information replaced with markers like `[REDACTED EMAIL]`, `[REDACTED PHONE]`, `[REDACTED SSN]`
- **Tracked changes** - all modifications should appear as tracked revisions (if Track Changes was enabled)

## Troubleshooting

### Add-in doesn't load
- Make sure both `npm run watch` and `npm run dev-server` are running
- Check that the server is accessible at `https://localhost:3000`
- Try re-uploading the `manifest.xml` file
- Clear your browser cache and reload

### "ApiNotFound" or Track Changes errors
- This is expected in Word Online - manually enable Track Changes in the Review tab before running redaction
- The add-in will continue working even if it can't enable tracking automatically

### No sensitive patterns found
- Check that your document actually contains emails, phone numbers, or SSNs
- The patterns look for:
  - Emails: `user@domain.com` format
  - Phones: Various formats like `(555) 123-4567`, `555-123-4567`, `555 123 4567`
  - SSNs: `XXX-XX-XXXX` format

### Certificate errors
- Run `npx office-addin-dev-certs install` again
- Make sure you accepted the certificate installation prompts
- On Mac, you may need to add the certificate to Keychain Access manually

### Changes not being tracked
- Verify Track Changes is ON in the Review tab
- In Word Online, the add-in cannot automatically enable tracking - you must do it manually
- If using Word Desktop, the add-in should enable it automatically

## Development Notes

- The TypeScript source is in `taskpane.ts`
- Compiled JavaScript output is `taskpane.js`
- The UI is in `index.html`
- Configuration is in `manifest.xml`
- After making code changes, the watch process will automatically recompile
- Reload the add-in task pane to see changes (close and reopen, or re-upload manifest)

## Testing

Use the provided `Document-To-Be-Redacted.docx` file to test the add-in. It contains:
- Multiple email addresses
- Various phone number formats
- Social security numbers

The add-in should successfully redact all of these and add the confidentiality header.

## Stopping the Development Server

When you're done:
- Press `Ctrl+C` in both terminal windows to stop the servers
- Or run `npm stop` to stop the add-in debugging process

---

**Note:** This solution was developed for Word on the web due to unavailability of Word Desktop. The manual Track Changes step is required for Word Online but may not be necessary if testing on Word Desktop.

