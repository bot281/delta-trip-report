# Delta Company - Trip Report System

## Overview
This project replaces Delta Company's Google Form with a mobile-friendly web app for driver trip reporting (arrival/departure).

## Key Features
- **Dynamic dropdowns** (vehicles, drivers, locations, etc.) loaded from Google Sheets
- **Two-step workflow**: Arrive → Depart (updates same row)
- **Pending trips** view for drivers
- **Admin-editable config** via Google Sheets
- **Offline support** with localStorage

## Architecture
```plaintext
┌─────────────┐    ┌─────────────┐    ┌─────────────────────┐
│   Web App   │───▶│ Apps Script │───▶│ Google Sheets       │
│ (GitHub Pages)│    │ (Backend)    │    │ - Câu trả lời biểu mẫu 1 │
└─────────────┘    └─────────────┘    │ - Cấu hình           │
                                      └─────────────────────┘
```

## URLs
- **Live App**: https://bot281.github.io/delta-trip-report/
- **GitHub Repo**: https://github.com/bot281/delta-trip-report
- **Backend**: https://script.google.com/macros/s/AKfycbzDyjhAhngx-D1N_P-WWMR1Hx2XEjHB78SEGJPXWkLoPD3rSMu2PCoe5XbEJ4t_fTd_/exec
- **Google Sheet**: https://docs.google.com/spreadsheets/d/1yWzRFhSScCtNdMB8Dk82sY5v9RHYzQFYSerVQA3KKzE/edit

## Project Structure
```plaintext
/
├── apps-script/       # Google Apps Script code
├── web/               # Frontend (GitHub Pages)
├── automation/        # Scripts for automation tasks
├── docs/               # Documentation
├── assets/            # Static assets (logos, etc.)
├── reports/           # Generated reports
├── backups/           # Data backups
├── CONTEXT.md         # Project context
└── README.md          # This file
```

## Setup
1. **Google Sheets**: Ensure the sheet has tabs:
   - `Câu trả lời biểu mẫu 1` (responses)
   - `Cấu hình` (config)
2. **Apps Script**: Deploy as web app (execute as `bot@delta.com.vn`, access "Anyone")
3. **GitHub Pages**: Push to `main` branch

## Usage
- Drivers: Use the web app to report arrivals/departures
- Admins: Edit dropdown options in the "Cấu hình" sheet

## Notes
- Uses GET requests instead of POST to bypass Google Workspace CORS restrictions
- All fields required except notes
- Mobile-first design with large touch targets

*Last updated: 2026-03-04*