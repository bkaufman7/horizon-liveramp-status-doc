# HMI LiveRamp Alerts Tracker

Google Apps Script automation for syncing LiveRamp alerts and managing Horizon team responses.

## Overview

This system mirrors a LiveRamp-owned Google Sheet into Horizon's internal tracking sheet, allows structured updates, and pushes responses back to LiveRamp.

## Features

- **Automated Daily Sync**: Pulls alerts from LiveRamp sheet
- **Status Tracking**: Identifies New/Updated/Ongoing/Resolved alerts
- **Template System**: User-editable response templates
- **Automated Push**: Appends Horizon updates to LiveRamp (column F only)
- **Daily Email Reports**: Color-coded HTML emails with status tables
- **Scheduled Automation**: Daily sync (pre-email), email (7:30 AM ET), auto-push (11:00 PM ET)

## Files

- `Code.js` - Main Apps Script with all automation logic
- `init.js` - One-time initialization script for sheet setup
- `appsscript.json` - Apps Script manifest
- `.clasp.json` - Clasp configuration (gitignored)

## Setup

1. Create a new Google Sheet for Horizon internal tracking
2. Open Apps Script editor (Extensions > Apps Script)
3. Copy contents of `Code.js` and `init.js` into the editor
4. Run `initHmiLiveRampUi()` once to create all tabs
5. Fill in Config tab with LiveRamp sheet URL and settings
6. Add recipients in Recipients tab
7. Run `setup()` from menu to install triggers

## Usage

### Daily Workflow

1. Review Working_Alerts tab (auto-synced daily)
2. Select template and add notes for each alert
3. Check HMI_Push_Ready box for items to send
4. Use menu: HMI LiveRamp > Push Updates to LiveRamp Now

### Menu Actions

- Refresh from LiveRamp
- Build Email Preview
- Send Email Now
- Push Updates to LiveRamp Now
- Refresh Template Dropdowns

### Templates

Edit the Templates tab to add/modify dropdown options. Run "Refresh Template Dropdowns" from menu to apply changes.

## Configuration

All settings in Config tab:
- LiveRamp Sheet URL
- Daily email time (default: 7:30 AM ET)
- Auto push time (default: 11:00 PM ET)
- Email subject prefix
- Timezone (America/New_York)

## Error Notifications

All errors are automatically emailed to: bkaufman@horizonmedia.com

## Technical Details

### Status Logic
- **New**: Thread key not seen before
- **Updated**: Thread key exists but content changed
- **Ongoing**: No changes since last sync
- **Resolved**: Marked "READY TO BE RESOLVED"

### Thread Key
Stable hash of Product + Workflow + Issue + Request (excludes Date) to track same issue across multiple days.

### Append Logic
Never overwrites LiveRamp column F. Format:
```
MM/DD/YYYY h:mm AM/PM ET
Horizon: <message>
```

## Version

2.0 - February 2026

## Support

For questions: bkaufman@horizonmedia.com
