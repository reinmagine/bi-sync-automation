# Bidirectional Sync Automation

This folder contains the Apps Script files for automating bidirectional data synchronization between Google Sheets.

## What it does

- Adds a custom Google Sheets menu: `DATA SYNC AUTOMATION`
- Runs `Sync Data` from the menu
- Overwrites rows `2..end` in the destination sheet
- Uses a configurable key column as the unique identifier
- Preserves unmapped destination values when a matching key already exists

## Local files

- `Code.gs` - main Apps Script
- `appsscript.json` - Apps Script manifest
- `sync_config.json` - local configuration snapshot
- `package.json` - optional local test runner and clasp helper
- `test-mapping.js` - local mapping check

## Recommended setup

1. Install Node.js.
2. Install clasp.
3. Create or bind the Apps Script project to the destination spreadsheet.
4. Push the local files.

```powershell
npm install -g @google/clasp
clasp login
cd "C:\Users\ludrein.salvador_glo\Downloads\bi-sync-automation"
clasp create --title "Bidirectional Sync Automation" --type sheets --parentId YOUR_DEST_SPREADSHEET_ID
clasp push
```

## How to use

1. Open the destination spreadsheet.
2. Refresh the page so Apps Script runs `onOpen()`.
3. Click `DATA SYNC AUTOMATION`.
4. Click `Sync Data`.

## Notes

- This version is manual only.
- It does not create an error sheet or trigger.
- A configurable key column (by default `GPO #`) serves as the record identifier.
- Any column not mapped from the source keeps the existing destination value when the same key exists.
