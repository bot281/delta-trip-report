# Delta Company - Client Project

## About
- **Company:** Công ty Cổ phần Quốc tế Delta (Delta International JSC)
- **Website:** https://delta.com.vn
- **Industry:** Logistics & Transportation (Vietnam)
- **Services:** Domestic transport (~200 vehicles), international freight, cold chain logistics, warehouse rental (Thu Duc & Hanoi), customs clearance, project cargo

## Project
- **Type:** Client work — automation, workflows, task execution
- **Started:** 2026-03-04
- **Status:** Active
- **Topic ID:** 4911

## Deliverable: Trip Report Web App

### Live URLs
- **Web App:** https://bot281.github.io/delta-trip-report/
- **GitHub Repo:** https://github.com/bot281/delta-trip-report
- **Apps Script:** https://script.google.com/macros/s/AKfycbzDyjhAhngx-D1N_P-WWMR1Hx2XEjHB78SEGJPXWkLoPD3rSMu2PCoe5XbEJ4t_fTd_/exec
- **Apps Script Project:** https://script.google.com/u/0/home/projects/16ysKXxiM5IL65cr4szDuh0JTMzzFQWDQEVY25vNMrIXzGZuXip-D-gQa/edit
- **Google Sheet:** https://docs.google.com/spreadsheets/d/1yWzRFhSScCtNdMB8Dk82sY5v9RHYzQFYSerVQA3KKzE/edit

### Architecture
- **Frontend:** Single HTML file (vanilla JS, mobile-first) hosted on GitHub Pages
- **Backend:** Google Apps Script (v3) deployed as web app under bot@delta.com.vn
- **Data:** Google Sheets — "Câu trả lời biểu mẫu 1" (responses) + "Cấu hình" (config)
- **Auth:** Execute as bot@delta.com.vn, access "Anyone"
- **CORS fix:** Uses GET with `?data=` URL param instead of POST (Google Workspace blocks external POST)

### Sheet Columns (A→T)
| Col | Header | Section |
|-----|--------|---------|
| A | Thời gian đến nơi | Arrival |
| B | Đến nơi / Rời đi | Arrival |
| C | Mã số xe / Biển số xe | Arrival |
| D | Biển số romooc | Arrival |
| E | Tên lái xe | Arrival |
| F | Địa điểm | Arrival |
| G | Số lô | Arrival |
| H | Số km trên đồng hồ | Arrival |
| I | % pin trên đồng hồ | Arrival |
| J | Nghiệp vụ | Arrival |
| K | Ghi chú | Arrival |
| L | Dấu thời gian | Departure |
| M | Đến nơi / Rời đi (Chỉnh sửa) | Departure |
| N | Biển số romooc (Chỉnh sửa) | Departure |
| O | Số km trên đồng hồ (Chỉnh sửa) | Departure |
| P | % pin trên đồng hồ (Chỉnh sửa) | Departure |
| Q | Ghi chú (Chỉnh sửa) | Departure |
| R | Số lô (khi rời đi) | Departure |
| S | Nghiệp vụ (khi rời đi) | Departure |
| T | Cột 1 | (unused) |

### Config Sheet ("Cấu hình")
7 columns — admin editable, web app loads dynamically on each visit:
- Biển số xe | Biển số romooc | Tên lái xe | Địa điểm | Số lô | Nghiệp vụ (đến) | Nghiệp vụ (rời)

### Key Features
- **Driver profile saved in localStorage** — vehicle, trailer, name auto-fill
- **2-step submit:** Arrive (creates row A-K) → Depart (updates same row L-S)
- **Pending trips tab** — shows arrivals without departure
- **Dynamic dropdowns** — loaded from "Cấu hình" sheet, admin can edit anytime

### API Endpoints (all via GET with `?data=JSON`)
- `getConfig` — returns dropdown options from config sheet
- `arrive` — appends row, returns rowId
- `depart` — updates existing row by rowId
- `getPending` — finds rows without departure for a driver

### Browser Profile
- **Profile:** delta-company (port 18810, DodgerBlue)
- **Logged in as:** bot@delta.com.vn (BOT DELTA)

### Lessons Learned
- Google Workspace accounts block external POST to Apps Script — use GET with URL params
- Apps Script deploy dialog crashes from non-owner accounts — must use the Workspace account
- Monaco editor in Apps Script accessible via `window.monaco.editor.getEditors()[0]`

## Notes
*(Track decisions, deliverables, and learnings here)*

---

*Last updated: 2026-03-04*
