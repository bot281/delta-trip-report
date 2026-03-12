# Trip Report Web App — v2 Spec

## Changes from v1

### 1. Đến nơi (Arrive) — Add date field
- **Current:** Time only (`HH:mm`)
- **New:** Date + Time (`DD/MM/YYYY HH:mm`)
- Default: current date & time, editable
- Stored in column A as `DD/MM/YYYY HH:mm:ss`

### 2. Trip Code — Enable trip handoff
- On successful arrival, server generates a **4-character alphanumeric code** (uppercase, e.g. `A3K9`)
- Code is **unique per active trip** (trips without departure)
- Code stored in **new column U** ("Mã chuyến")
- Confirmation screen shows the code prominently so driver can copy or screenshot it
- Code is used to look up the trip for departure

### 3. Rời đi (Depart) — New fields + code search

#### Finding a trip (two options):
- **Tab: "Chuyến của tôi"** — shows pending trips for current driver (from localStorage, like v1)
- **Tab: "Nhập mã chuyến"** — input 4-char code to find ANY pending trip (enables Driver B handoff)

#### New departure fields:
- **Số xe** (vehicle) — dropdown, pre-filled from profile but editable
- **Tên tài xế** (driver name) — dropdown, pre-filled from profile but editable
- Existing: Trailer, KM, Battery, Lot, Task, Note

#### Data behavior:
- Departure data is written to columns L–S (existing) + driver/vehicle in new columns or existing departure columns
- **Arrival data (columns A–K) is NEVER modified** — preserves Driver A's history
- A trip records who arrived and who departed — they can be different people/vehicles

### 4. Sheet Column Updates

| Col | Header | Section |
|-----|--------|---------|
| A | Thời gian đến nơi | Arrival (now `DD/MM/YYYY HH:mm:ss`) |
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
| L | Dấu thời gian | Departure (timestamp) |
| M | Đến nơi / Rời đi (Chỉnh sửa) | Departure |
| N | Biển số romooc (Chỉnh sửa) | Departure |
| O | Số km trên đồng hồ (Chỉnh sửa) | Departure |
| P | % pin trên đồng hồ (Chỉnh sửa) | Departure |
| Q | Ghi chú (Chỉnh sửa) | Departure |
| R | Số lô (khi rời đi) | Departure |
| S | Nghiệp vụ (khi rời đi) | Departure |
| T | Cột 1 | (unused) |
| **U** | **Mã chuyến** | **Trip code (NEW)** |
| **V** | **Tên lái xe (rời đi)** | **Departure driver (NEW)** |
| **W** | **Biển số xe (rời đi)** | **Departure vehicle (NEW)** |

### 5. API Changes

#### `arrive` response — now includes trip code:
```json
{
  "success": true,
  "rowId": 42,
  "tripCode": "A3K9",
  "message": "Đã ghi nhận đến nơi"
}
```

#### New action: `findByCode` — look up trip by code:
```json
// Request
{ "action": "findByCode", "code": "A3K9" }

// Response
{
  "success": true,
  "trip": {
    "rowId": 42,
    "time": "07/03/2026 09:30:00",
    "vehicle": "50E-196.22",
    "trailer": "15R-091.48",
    "driver": "Lê Anh Toàn",
    "location": "Cảng Bình Dương",
    "lot": "MEV49",
    "km": "1500",
    "battery": "85%",
    "task": "Lấy rỗng"
  }
}
```

#### `depart` — now includes driver & vehicle:
```json
{
  "action": "depart",
  "rowId": 42,
  "driver": "Lê Thanh Tú",
  "vehicle": "50H-180.03",
  "trailer": "15R-129.37",
  "km": "1520",
  "battery": "80%",
  "lot": "MEV50",
  "task": "Hạ cont đầy",
  "note": ""
}
```

### 6. UX Flow

```
ARRIVE
  ┌─────────────────────────────┐
  │ 👤 Thông tin lái xe          │
  │   Tên / Xe / Romooc         │
  ├─────────────────────────────┤
  │ 📋 Chi tiết chuyến           │
  │   📅 Ngày + ⏰ Giờ           │  ← NEW: date picker
  │   Địa điểm / Lô / KM / Pin  │
  │   Nghiệp vụ / Ghi chú       │
  ├─────────────────────────────┤
  │  [📍 Ghi nhận ĐẾN NƠI]      │
  └─────────────────────────────┘
           │
           ▼
  ┌─────────────────────────────┐
  │ ✅ Đã ghi nhận!              │
  │                              │
  │   Mã chuyến: A3K9            │  ← NEW: trip code
  │   [📋 Sao chép]              │
  │                              │
  │ Gửi mã này cho tài xế tiếp  │
  │ theo nếu cần chuyển chuyến.  │
  └─────────────────────────────┘

DEPART
  ┌──────────────────────────────┐
  │ [Chuyến của tôi] [Nhập mã]  │  ← NEW: two ways
  ├──────────────────────────────┤
  │                               │
  │  Nhập mã chuyến: [____]      │  ← NEW: 4-char input
  │  [🔍 Tìm chuyến]             │
  │                               │
  ├──────────────────────────────┤
  │ 📍 Thông tin đến nơi (read)  │
  │   Driver A / Vehicle A / ... │
  ├──────────────────────────────┤
  │ 🚀 Thông tin rời đi           │
  │   👤 Tên lái xe  [dropdown]   │  ← NEW
  │   🚛 Biển số xe  [dropdown]   │  ← NEW
  │   Romooc / KM / Pin           │
  │   Lô / Nghiệp vụ / Ghi chú   │
  ├──────────────────────────────┤
  │  [🚀 Ghi nhận RỜI ĐI]        │
  └──────────────────────────────┘
```

### 7. Trip Code Generation (server-side)
- Characters: `ABCDEFGHJKLMNPQRSTUVWXYZ23456789` (no I/O/0/1 to avoid confusion)
- 4 characters = 28^4 = 614,656 combinations
- Check uniqueness against active trips (no departure timestamp)
- If collision, regenerate (max 5 attempts)

---

*Approved by: P O — 2026-03-07*
