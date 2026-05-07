# CSKVTT — Quản Lý Đèn Tắt Hệ Thống Chiếu Sáng Công Cộng

**Author**: Mai Vũ Lâm  
**Version**: V1.2  
**Description**: Web app quản lý, theo dõi và báo cáo sự cố đèn chiếu sáng công cộng bị tắt/hư tại TP.HCM. Hiển thị vị trí đèn trên bản đồ Leaflet với dot icon màu sắc theo trạng thái, pulse animation cho đèn hư. Có trang đăng nhập xác thực qua Google Sheet. Hỗ trợ nhập/xuất Excel, thêm/sửa marker, tìm kiếm, định vị GPS, reverse geocode, chỉ đường, xuất CAD, báo cáo hàng ngày, lọc theo trạng thái. Ghi dữ liệu qua Google Apps Script → Google Sheet.

**URL triển khai:** `https://neo-era.github.io/cskvtt/Den%20tat/dentat.html`  
**Google Apps Script URL:** `https://script.google.com/macros/s/AKfycbygHZRMQa6cC6-7mXIMgzHaRFg4b2t_QjM0kMooVXt9GqX7chK-knxwwiYjhBVzwKPM7Q/exec`  
**Google Sheet CSV URL:** `https://docs.google.com/spreadsheets/d/e/2PACX-1vQC6mnGNSNDjUVzs5C4Se9Q9JQCGF9_YQRTRXewhsJhg0QDAcp6NqtxNsFl-fs8g1yyYBQEUqPhgwBv/pub?output=csv`

---

## Cấu trúc thư mục

```
cskvtt/Den tat/
├── dentat.html          # File chính — bản đồ quản lý đèn tắt
├── gas.js               # Mã nguồn Google Apps Script (reference, không chạy ở browser)
├── CLAUDE.md            # Tài liệu dự án này
├── manifest.json        # PWA manifest
├── sw.js                # Service Worker
├── data/
│   ├── Danhsachdentat.xlsx      # File Excel dữ liệu đèn (backup local)
│   ├── bangron.xlsx             # Dữ liệu bảng rộng (dự phòng)
│   ├── Camera.xlsx              # Dữ liệu camera (dự phòng)
│   └── thietbitichhop.xlsx      # Dữ liệu thiết bị tích hợp (dự phòng)
└── images/
    ├── icon-192.png     # PWA icon 192×192
    ├── icon-512.png     # PWA icon 512×512
    └── *.jpeg           # Ảnh chụp hiện trường đèn sự cố
```

---

## Trạng thái đèn — STATUS_CONFIG

| Mã | Màu | Ý nghĩa | Pulse |
|---|---|---|---|
| `1` | `#ef4444` (đỏ) | Đèn LED đang hư | ✓ |
| `7` | `#f97316` (cam) | Đèn HPS đang hư | ✓ |
| `5` | `#7c3aed` (tím) | Đèn hư quá 10 ngày | ✓ |
| `3` | `#10b981` (xanh lá) | Đèn LED sáng bình thường | ✗ |
| `4` | `#0ea5e9` (xanh dương) | Đèn HPS sáng bình thường | ✗ |

Marker dùng `L.divIcon` với CSS class `den-dot` — dot tròn màu theo trạng thái, có `den-pulse` animation cho đèn hư.

---

## Thư viện chính (CDN)

| Thư viện | Phiên bản | Mục đích |
|---|---|---|
| Leaflet | 1.9.4 | Bản đồ tương tác |
| Leaflet MarkerCluster | 1.5.3 | Gom nhóm marker |
| XLSX (SheetJS) | 0.18.5 | Đọc file Excel, parse CSV từ Google Sheet |
| Google Fonts (Inter) | — | Font chữ giao diện |
| Nominatim OSM | — | Reverse geocode (tên đường/phường từ tọa độ) |
| OSRM | — | Routing (chỉ đường lái xe) |

---

## Đăng nhập (Authentication)

### Luồng xác thực
1. Trang tải → `window.onload` → `checkAuth()` → kiểm tra `localStorage('dt_user')`
2. Nếu có phiên hợp lệ → ẩn login overlay → `loadDataFromSheet()`
3. Nếu không có → hiện login overlay, ẩn loading spinner
4. Người dùng nhập username/password → `doLogin()` → POST GAS `action: 'login'`
5. GAS tra cứu sheet **TaiKhoan** → trả về `{status: 'ok', user: {...}}` hoặc `{status: 'error', message: '...'}`
6. Thành công → lưu `currentUser` vào `localStorage('dt_user')` → ẩn overlay → tải dữ liệu
7. Đăng xuất → `logout()` → xóa `localStorage('dt_user')` → hiện overlay lại

### Sheet TaiKhoan (cần tạo thủ công)
Tab mới trong cùng Google Sheet, tên tab: `TaiKhoan`, cấu trúc:

| Cột A | Cột B | Cột C | Cột D |
|---|---|---|---|
| `tenDangNhap` | `matKhau` | `hoTen` | `vaiTro` |
| nguyenvana | 123456 | Nguyễn Văn A | user |
| admin | adminpass | Quản trị viên | admin |

- `vaiTro`: `admin` hoặc `quanly` → hiện icon 👑; khác → 👷
- Mật khẩu lưu plaintext (đủ cho công cụ nội bộ)

### State người dùng
```js
let currentUser = null; // { username, displayName, role }
// Lưu vào localStorage key: 'dt_user'
```

### Hàm auth
- `checkAuth()` — kiểm tra localStorage; nếu có phiên → show app + tải dữ liệu; nếu không → show login
- `doLogin()` — gọi GAS với `action: 'login'`; dùng `Content-Type: text/plain` để tránh CORS preflight
- `logout()` — xóa localStorage, clear markers/data, hiện login overlay
- `updateSidebarUser()` — cập nhật tên + vai trò hiển thị trong sidebar

---

## Dữ liệu

Nguồn dữ liệu chính: **Google Sheet** (tải về qua URL CSV `pub?output=csv`), không dùng file Excel local.  
Ghi dữ liệu: qua **Google Apps Script** (`gas.js` deploy thành Web App).

### Schema cột FIELDS (index trong code)

| Index | Key | Tên cột Sheet | Ghi chú |
|---|---|---|---|
| 0 | `id` | `ID` | Mã định danh |
| 1 | `soTru` | `Số trụ` | Tên chính marker |
| 2 | `tenTu` | `Tên tủ` | Tủ điều khiển quản lý đèn |
| 3 | `latitude` | `latitude` | WGS84 |
| 4 | `longitude` | `lontitude` | WGS84 — **tên cột Sheet là `lontitude` (sai chính tả, giữ nguyên)** |
| 5 | `loaiDen` | `Loại đèn` | `LED` hoặc `HPS` |
| 6 | `congSuat` | `Công suất` | Đơn vị W |
| 7 | `trangThai` | `Trang thai` | Mã: 1/3/4/5/7 |
| 8 | `duong` | `Đường` | Tự động từ reverse geocode |
| 9 | `phuong` | `Phường` | Tự động từ reverse geocode |
| 10 | `ngayPhatHien` | `Ngày phát hiện` | Định dạng `dd/mm/yyyy` |
| 11 | `nguoiPhatHien` | `Người phát hiện` | Tự động điền từ `currentUser.displayName` khi thêm mới |
| 12 | `ngaySua` | `Ngày sửa` | Định dạng `dd/mm/yyyy` |
| 13 | `nguoiSua` | `Người sửa` | |
| 14 | `vatTuSua` | `Vật tư sửa` | Danh sách linh kiện |
| 15 | `hinhAnh` | `HÌnh ảnh` | Đường dẫn trong `Den tat/images/` (**tên cột có lỗi 'Ì' thay 'ì'**) |
| 16 | `ghiChu` | `Ghi chú` | |
| 17 | `vn2000x` | `VN2000-X` | Tọa độ VN2000 Easting |
| 18 | `vn2000y` | `VN2000-Y` | Tọa độ VN2000 Northing |

---

## Kiến trúc JavaScript (inline trong dentat.html)

### Khởi tạo
- `initializeMap()` — khởi tạo Leaflet map, tile layers (OSM, Google Satellite), layer control, đăng ký sự kiện click bản đồ
- `loadDataFromSheet()` — fetch CSV từ Google Sheet, parse bằng SheetJS, gọi `addMarkersToMap()`
- `window.onload` — gọi `initializeMap()`, `buildFilterChips()`, `buildLegend()`, `checkAuth()`; **không gọi `loadDataFromSheet()` trực tiếp** — checkAuth quyết định khi nào tải dữ liệu

### Xác thực
- `checkAuth()` — kiểm tra localStorage, quyết định show login hay load app
- `doLogin()` — POST GAS `action:'login'`, xử lý response, lưu session
- `logout()` — xóa session, reset state, hiện login
- `updateSidebarUser()` — render tên + vai trò trong chip user sidebar

### Xử lý dữ liệu
- `processFile(event)` — đọc file Excel local bằng SheetJS (import thủ công)
- `parseRow(row)` — parse hàng dữ liệu (array hoặc object) thành object chuẩn; xử lý ngày qua `parseDate()`
- `parseMarkerRow(row)` — alias của `parseRow()`
- `getVal(row, keys, fb)` — trích xuất giá trị từ hàng theo nhiều tên cột khác nhau
- `applyFieldsToRow(row, updates)` — merge updates vào row gốc, giữ nguyên key tên cột Sheet gốc
- `rowToArray(row)` — chuyển object marker → array 19 phần tử để xuất Excel

### Ngày tháng
- `excelSerialToViDate(serial)` — chuyển Excel serial number → chuỗi `dd/mm/yyyy`
- `parseDate(val)` — chuẩn hóa giá trị ngày: serial number → `dd/mm/yyyy`, hoặc giữ nguyên nếu là string
- `parseViDate(dateStr)` — parse `dd/mm/yyyy` hoặc `yyyy-mm-dd` → Date object (không lệch timezone)
- `daysBetween(dateStr)` — số ngày từ ngày đến hôm nay
- `todayStr()` — trả về hôm nay dạng `dd/mm/yyyy`
- `toInputDate(vn)` — `dd/mm/yyyy` → `yyyy-mm-dd` cho `<input type="date">`
- `fromInputDate(iso)` — `yyyy-mm-dd` → `dd/mm/yyyy` để lưu vào Sheet

### Marker
- `addMarkerRow(row, idx)` — tạo Leaflet marker (dot icon) + label, thêm vào cluster; lazy popup callback
- `addMarkersToMap(data)` — clear và load lại toàn bộ marker, gọi `applyFilter()`, `updateStats()`, v.v.
- `createPopupContent(row, rowIndex)` — HTML popup: tính số ngày hư, nút "Đã sửa" (quickFix), nút chỉ đường
- `createStatusIcon(trangThai)` — `L.divIcon` dot tròn màu theo trạng thái, có pulse animation

### Lọc & thống kê
- `toggleFilter(status)` — ẩn/hiện marker theo trạng thái
- `applyFilter()` — áp dụng `hiddenFilters` lên cluster và labelLayerGroup
- `updateStats()` — cập nhật đếm số đèn theo từng trạng thái
- `buildFilterChips()` — render filter chips trong sidebar
- `buildLegend()` — render chú thích màu trong sidebar

### Status List Panel
- `openStatusList(status)` — mở panel trượt từ dưới, liệt kê đèn theo trạng thái
- `closeStatusList()` — đóng panel
- `flyToEntry(index)` — nhảy bản đồ đến marker theo index, mở popup

### Tìm kiếm
- `searchMarkers()` — tìm theo số trụ, đường, phường, tên tủ; 1 kết quả → fly to, nhiều → dropdown
- `normalizeText(t)` / `normalizeKey(v)` — bỏ dấu tiếng Việt, lowercase

### Sidebar
- `openSidebar()` / `closeSidebar()` — mở/đóng sidebar drawer

### Form thêm/sửa
- `showFormAt(lat, lon)` — mở form thêm mới tại tọa độ, gọi `reverseGeocode()`, tự điền `fNguoiPhatHien = currentUser.displayName`
- `openEditPopup(index)` — mở form chỉnh sửa marker theo index
- `fillForm(row, lat, lon)` — điền dữ liệu vào form; nếu `row = null` → thêm mới (đặt ngày hôm nay, trạng thái `1`, tự điền người phát hiện từ `currentUser`)
- `hideForm()` / `cancelMarkerPopup()` — đóng form
- `saveMarkerPopup()` — đọc form → payload → POST GAS; merge với `applyFieldsToRow()` khi edit

### Chế độ thêm marker
- `startAddMarker()` — toggle; hỏi GPS hay click bản đồ
- `enableAddMode()` / `disableAddMode()` — bật/tắt `pendingNewMarker`, cập nhật nút

### Định vị & Reverse Geocode
- `setMarkerToCurrentLocation()` — GPS (`enableHighAccuracy:true`, `maximumAge:0`, `timeout:15000`); spinner; điền lat/lon vào form; gọi `reverseGeocode()`
- `centerOnUserLocation()` — GPS + setView + marker GPS xanh
- `reverseGeocode(lat, lon)` — Nominatim API lấy tên đường/phường, điền vào form

### Chỉ đường
- `routeToMarker(index)` — lấy GPS hiện tại, gọi `fetchRoute()`
- `fetchRoute(origin, dest)` — OSRM API, vẽ polyline route, hiện khoảng cách + thời gian

### Quick Fix
- `quickFix(index)` — đánh dấu đèn đã sửa: đổi trạng thái (LED→3, HPS→4), `ngaySua = hôm nay`, gửi GAS

### Báo cáo
- `openReportModal()` / `closeReportModal()`
- `getReportRows()` — lọc theo khoảng ngày phát hiện
- `updateReportPreview()` — preview tóm tắt theo trạng thái
- `exportReport()` — xuất Excel báo cáo chi tiết (có cột "Số ngày hư")

### Lưu & Xuất
- `saveMarkerData()` — xuất `loadedData[]` ra file Excel (⚠️ xem bug bên dưới)
- `exportMarkersToCad()` — xuất DXF (AutoCAD) với tọa độ VN2000
- `convertLatLonToVn2000(lat, lon)` — WGS84 → VN2000 (UTM zone tự động, GRS80)

### Ảnh
- `triggerImageInput()` — kích hoạt input file (có `capture="environment"` → mở camera)
- `handleImageFile(e)` — đọc file ảnh → `resizeImage()` → preview
- `resizeImage(dataUrl, cb)` — resize max 800px, giảm quality rồi kích thước cho đến khi ≤ `MAX_IMG_LEN`
- `normalizeImagePath(img)` — chuẩn hóa đường dẫn ảnh

### Google Apps Script (gas.js)
- `doPost(e)` — nhận POST; nếu `action: 'login'` → `handleLogin()`; nếu `action: 'full_update'` → `findRowNum()` → `updateRow()` hoặc `appendRow()`; ngược lại → GPS-only update
- `handleLogin(username, password)` — tra cứu sheet **TaiKhoan**, so sánh plaintext, trả `{status, user}`
- `findRowNum()` — tìm hàng theo ID (ưu tiên) hoặc Số trụ, so sánh normalize
- `updateRow()` / `updateRowFields()` — ghi từng ô theo `hIdx`
- `appendRow()` — thêm hàng mới theo đúng thứ tự cột
- `buildFieldValues(data)` — chuyển payload camelCase → tên cột Sheet qua `FIELD_MAP`
- `doGet(e)` — health check (trả `{status: 'ok', message: 'Den tat GAS v3 — login ready'}`)

---

## Trạng thái toàn cục (global state)

```js
const GOOGLE_SCRIPT_URL = '...';
const GOOGLE_SHEET_CSV_URL = '...';
const MAX_IMG_LEN = 32767;

let map, markersCluster, labelLayerGroup;
let markers = [];               // [{lat, lon, name, marker, row, rowIndex, labelMarker}]
let loadedData = [];
let editingIndex = null;
let addMarkerMode = false;
let pendingNewMarker = false;
let markerImageDataUrl = '', markerImagePath = '';
let currentLocationMarker = null;
let routeLine = null;
let hiddenFilters = new Set();
let toastTimer;
let currentUser = null;         // { username, displayName, role } — lưu trong localStorage 'dt_user'
```

---

## Tile layers

```js
'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png'
'https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'  // subdomains: mt0–mt3
```

---

## Google Apps Script (gas.js) — FIELD_MAP

| Key JS | Tên cột Sheet |
|---|---|
| `soTru` | `Số trụ` |
| `lat` / `latitude` | `latitude` |
| `lon` / `longitude` / `lontitude` | `lontitude` |
| `trangThai` | `Trang thai` |
| `hinhAnh` | `HÌnh ảnh` |
| `vn2000x` / `vn2000y` | `VN2000-X` / `VN2000-Y` |

Tìm hàng: ưu tiên theo `ID`, fallback theo `Số trụ` (cả hai đều normalize trước khi so sánh).

---

## PWA

- `manifest.json`: name "Quản Lý Đèn Tắt CSCC", short_name "Đèn Tắt", start_url `./dentat.html`, icons 192/512, theme `#0ea5e9`
- `sw.js`: 2 cache: `dentat-static-v1` (assets), `dentat-tiles-v1` (map tiles, giới hạn 200)
- **Không cache**: Google Sheet CSV, Apps Script URL, Nominatim, OSRM
- Meta: `theme-color: #0f172a` (navy trong header)

---

## Hướng dẫn thiết lập tài khoản

1. Mở Google Sheet của dự án
2. Tạo tab mới, đặt tên chính xác: `TaiKhoan` (phân biệt hoa thường)
3. Hàng 1 (header): `tenDangNhap` | `matKhau` | `hoTen` | `vaiTro`
4. Thêm tài khoản từ hàng 2 trở đi
5. **Không publish tab này** — chỉ GAS đọc nội bộ (Execute as: Me)
6. Deploy lại GAS sau khi copy gas.js mới (version `v3`)

---

## Lưu ý kỹ thuật

- **Login dùng `text/plain` không dùng `no-cors`**: Khác với các POST ghi dữ liệu đèn (`mode:'no-cors'`), login cần đọc response → dùng `Content-Type: text/plain;charset=utf-8` để tránh CORS preflight, GAS xử lý và trả về JSON.
- **Phiên đăng nhập lưu localStorage**: Key `'dt_user'`, không có expiry — phiên tồn tại đến khi đăng xuất hoặc xóa cache.
- **Auto-fill người phát hiện**: Khi mở form thêm mới, `fNguoiPhatHien` tự điền `currentUser.displayName` (hoặc username nếu không có displayName).
- **Cột `lontitude`**: Tên cột Sheet sai chính tả, giữ nguyên để tương thích GAS.
- **Tên cột `HÌnh ảnh`**: Có lỗi chữ 'Ì' hoa — đã map trong gas.js và FIELDS.
- **⚠️ Bug xuất Excel**: `saveMarkerData()` header 17 cột nhưng `rowToArray()` trả về 19 → thiếu `Người phát hiện` và `Người sửa` trong header → dữ liệu lệch cột từ cột 11.
- **Mật khẩu plaintext**: Lưu plaintext trong Sheet vì đây là công cụ nội bộ, không phải hệ thống công khai. Sheet `TaiKhoan` không được publish CSV.
- **mode: 'no-cors' cho ghi đèn**: Response opaque, không đọc được — fire-and-forget; lỗi GAS chỉ thấy trong Apps Script logs.
- **Lazy popup**: callback `() => createPopupContent(...)` — HTML chỉ render khi mở popup.
- **GPS**: `enableHighAccuracy:true`, `maximumAge:0`, `timeout:15000`; hiện spinner khi chờ.

---

## Kế hoạch phát triển

### Sửa bug xuất Excel
- Thêm `Người phát hiện` và `Người sửa` vào header array trong `saveMarkerData()`

### Bảo mật login nâng cao
- Hash mật khẩu phía GAS (SHA-256) thay vì plaintext
- Hoặc dùng Google Workspace OAuth nếu tổ chức có Google Workspace

### Session expiry
- Lưu `loggedAt` trong `currentUser`, kiểm tra khi `checkAuth()`: nếu > 24h → yêu cầu đăng nhập lại

### Giới hạn chỉnh sửa theo vai trò
- `admin/quanly`: full access (thêm/sửa/xóa)
- `user`: chỉ báo cáo đèn mới, không sửa đèn của người khác

### Tối ưu mobile
- Font size input 16px để tránh auto-zoom iOS
- Test trên Android Chrome + iOS Safari

---

## Chạy local

Mở bằng Live Server (VS Code extension) trên port 5501, trỏ vào `Den tat/dentat.html`.  
Không cần build step — toàn bộ là HTML/CSS/JS thuần.
