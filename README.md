# Monstera TNR Lenggong Tracker

A mobile-friendly web app for tracking stray cat colonies and TNR (Trap-Neuter-Return) status around Monstera Lenggong, Perak, Malaysia.

## Quick Start

1. **Open the app:**
   - Double-click `index.html` to open in browser
   - Or host on GitHub Pages for sharing

2. **Report a cat:**
   - Click "Report" tab
   - Fill in Name, Location, Type, Status
   - Click "Submit Report"

3. **Update status:**
   - Tap any entry in the list
   - Change status in popup
   - Click "Save Changes"

## Cloud Sync (Optional)

To sync data across multiple users:

1. Create a Google Sheet at https://sheets.new
2. Go to Extensions → Apps Script
3. Paste code from `Code.gs`
4. Run `setupSheets()` once
5. Deploy as Web App (Execute as: Me, Access: Anyone)
6. Open app → Click ⚙️ → Paste URL → Save

## Files
- `index.html` - Main app (single file, no dependencies)
- `Code.gs` - Google Apps Script backend
- `README.md` - This file

## Demo Data
The app comes with 4 demo entries. Click "Reset" to restore them.

## License
MIT License - Free to use and modify.
