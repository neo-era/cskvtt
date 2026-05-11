# CSKVTT — Bản Đồ Khảo Sát Hệ Thống Chiếu Sáng Công Cộng

**Author**: Mai Vũ Lâm  
**Version**: V1.2  
**Description**: Web app quản lý và hiển thị vị trí hạ tầng chiếu sáng công cộng (TP.HCM) trên bản đồ Leaflet. Hỗ trợ đăng nhập, nhập/xuất Excel, thêm/sửa marker, tìm kiếm, định vị GPS, chỉ đường, xuất CAD, đồng bộ GitHub và PWA offline. Backend dữ liệu là Google Sheets, ghi qua Google Apps Script proxy.

**Google Apps Script (khaosat):** `https://script.google.com/macros/s/AKfycbxPFn7lnJZYdjUg5WlVq1P1JGI5O3rvcry_IVE3bhF5foDmbgT6KwBU4xFxHNwKruHRvQ/exec`  
**Google Apps Script (github_write_file):** `https://script.google.com/macros/s/AKfycbyNtO-Q8KPuRRgx0VacIEDF3VXeMfvAUsg-t2nJEqfVxzSlRGCyxjB4_kLMP7AqP9mu/exec`  
**Google Sheet CSV:** `https://docs.google.com/spreadsheets/d/e/2PACX-1vRjBdN5epr38kL2jLituPUC4wJO_Ohr_AsFpzaZh6Azxxs_gkzGT5s3JiZ1uFkNQFP5epR2U_wYT6wT/pub?output=csv`

---

## Cấu trúc thư mục

```
cskvtt/
├── index.html          # Trang chính — khảo sát chiếu sáng (tất cả quận)
├── gas-khaosat.js      # Mã nguồn Google Apps Script (reference, không chạy ở browser)
├── camera-map.html     # Bản đồ camera giám sát
├── quan1.html          # Bản đồ riêng Quận 1
├── test.html           # Bản sao index.html dùng để test
├── backup-index.html   # Backup index.html
├── css/                # CSS files (Bootstrap 5, Font Awesome, plugins)
├── js/                 # JS files (jQuery 3.3.1, Bootstrap bundle, plugins)
├── images/             # Icons UI, logo, favicon, PWA icons + ảnh khảo sát
├── data/               # File Excel (.xlsx) — dữ liệu marker local (fallback)
├── Backup/             # Các phiên bản HTML cũ
├── manifest.json       # PWA manifest
├── sw.js               # Service Worker (cache-first cho assets)
└── .vscode/            # Live Server port 5501
```

---

## Thư viện chính (CDN)

| Thư viện | Phiên bản | Mục đích |
|---|---|---|
| Leaflet | 1.9.4 | Bản đồ tương tác |
| Leaflet MarkerCluster | 1.5.3 | Gom nhóm marker |
| XLSX (SheetJS) | 0.18.5 | Đọc/ghi file Excel, parse CSV |
| ExcelJS | 4.3.1 | Xử lý workbook (lưu có ảnh) |
| Bootstrap | 5 (bundle) | UI responsive |
| jQuery | 3.3.1 (`js/jquery.min.js`) | DOM & modal |
| Font Awesome | 4.0.3 | Icons |

---

## Dữ liệu

### Nguồn dữ liệu
- **Chính**: Google Sheet → `KHAOSAT_CSV_URL` (CSV published to web), tải sau khi đăng nhập
- **Fallback**: `data/khaosat.xlsx` (local), tải khi chưa đăng nhập

### Schema (dùng chỉ số cột, cả Excel lẫn CSV Google Sheet)

| Index | Tên trường | Google Sheet header | Ghi chú |
|---|---|---|---|
| 0 | ID | `ID` | Có thể để trống |
| 1 | tên trụ | `Tên trụ` | Tên chính marker |
| 2 | latitude | `Lat` | WGS84 |
| 3 | longitude | `Lon` | WGS84 |
| 4 | ghi chú | `Ghi chú` | |
| 5 | người KS | `Người KS` | Tự điền từ `currentUser.displayName` khi thêm mới |
| 6 | loại | `Loại` | Số nguyên — xem ICON_CONFIG |
| 7 | tủ điều khiển | `Tủ điều khiển` | |
| 8 | loại trụ | `Loại trụ` | Ví dụ: STK, BTLT... |
| 9 | loại cần | `Loại cần` | |
| 10 | loại đèn | `Loại đèn` | |
| 11 | hình ảnh | `Ảnh` | `images/xxx.jpg` (GitHub) hoặc base64 |
| 12 | thời gian cập nhật | `Thời gian cập nhật` | |
| 13 | marker gốc | `Marker gốc` | Tên marker tham chiếu |
| 14 | khoảng cách (m) | `Khoảng cách (m)` | Khoảng cách đến marker gốc |

### ICON_CONFIG (loại trụ → màu dot)

| Mã | Label | Màu |
|---|---|---|
| 0 | Trụ STK + Đèn LED | `#0ea5e9` xanh dương |
| 1 | Trụ STK + Đèn HPS | `#f59e0b` vàng |
| 2 | Trụ trang trí | `#ec4899` hồng |
| 3 | Trụ BTLT + Đèn LED | `#10b981` xanh lá |
| 4 | Trụ BTLT + Đèn HPS | `#f97316` cam |
| 5 | Tủ lắp nổi | icon `images/9.png` |
| 6 | Tủ lắp ngầm | icon `images/9.png` (nhỏ hơn) |
| 7 | Trụ BTLT điện lực ql | `#64748b` xám |
| 8 | Trụ BTLT CS ql | `#7c3aed` tím |
| 9 | Trụ BTLT đề xuất trồng mới | `#22c55e` xanh lá, pulse |

---

## Kiến trúc JavaScript (trong HTML)

Toàn bộ logic viết inline trong `index.html`. Các hàm chính:

### Đăng nhập (Auth)
- `checkAuth()` — kiểm tra `localStorage('ks_user')`; nếu có session → `loadDataFromSheet()`; nếu không → hiện login overlay + `loadDefaultFile()` fallback
- `doLogin()` — POST GAS `action:'login'`; dùng `Content-Type: text/plain` tránh CORS preflight; lưu session vào `localStorage('ks_user')`
- `logout()` — xóa session, clear marker, hiện login overlay, load lại `loadDefaultFile()`
- `updateSidebarUser()` — render tên + vai trò (👑 admin/quanly, 👷 user) lên topbar và sidebar

### Khởi tạo
- `initializeMap()` — khởi tạo Leaflet map, tile layers (OSM, Google Satellite, Bing, WMS), control layers, `createMapControls()`; đăng ký `zoomend` listener **1 lần duy nhất** tại đây
- `loadDefaultFile()` — fetch `data/khaosat.xlsx` → `processFile()` (dùng khi chưa login)
- `loadDataFromSheet()` — fetch `KHAOSAT_CSV_URL` → XLSX parse → `addMarkersToMap()` (dùng sau login)
- `window.onload` — `initializeMap()` → `checkAuth()`

### Xử lý dữ liệu
- `processFile(file)` — đọc file Excel bằng SheetJS, lưu vào `loadedData[]`, gọi `addMarkersToMap()`
- `triggerExcelImport()` — kích hoạt input file import Excel
- `handleExcelImport(event)` — đọc file Excel import → `addMarkersToMap()`
- `downloadTemplateExcel()` — tạo và tải file Excel mẫu (có header + sheet hướng dẫn) bằng SheetJS

### Google Sheet / GAS
- `loadDataFromSheet()` — tải CSV từ Google Sheet, parse bằng XLSX
- `pushMarkerToSheet(row)` — upload ảnh (nếu base64) rồi ghi row lên Google Sheet qua `KHAOSAT_GAS_URL`
- GAS actions: `login`, `upload_image` (ảnh lên GitHub `images/`), `full_update` (ghi dữ liệu vào `DanhSachTru`)
- `GAS_URL` — GAS thứ 2 (Den Tat GAS), dùng riêng cho `github_write_file` (sync Excel lên GitHub)

### Marker
- `addMarkerRowToMap(row)` — tạo Leaflet marker (dot icon, draggable) + label; popup lazy (`bindPopup('')`); `markers.push({lat, lon, name, note, marker, row})`
- `addMarkersToMap(data)` — clear và load lại toàn bộ marker; áp dụng trạng thái label sau khi load xong
- `createMarkerPopupContent(row, markerIdx)` — HTML popup card (`.pc-*` classes): topbar màu, avatar ảnh, thông tin, nút "🧭 Đi đến" + "✏️ Chỉnh sửa"
- `openEditMarker(markerIdx)` — lấy `markers[idx].row`, điền vào form, set `editingMarkerIndex`
- `saveMarkerPopup()` — nếu `editingMarkerIndex != null` → cập nhật row in-place + `pushMarkerToSheet`; nếu không → thêm mới
- `updateMarkerCoordinatesInData(markerName, newLat, newLon)` — cập nhật tọa độ khi kéo marker
- `enableAddMarkerMode()` / `disableAddMarkerMode()` / `startAddMarker()` — toggle chế độ click-to-add
- `showMarkerPopupAt(lat, lon)` — mở form thêm marker mới tại tọa độ
- `cancelMarkerPopup()` / `hideMarkerPopup()` — đóng form, reset `editingMarkerIndex`
- `makeElementDraggable(handle, target)` — Pointer Events API để drag form popup

### Tên & khoảng cách trụ
- `getNextMarkerName(previousName)` — tự động tăng số cuối tên trụ (VD: "A01" → "A02")
- `normalizeMarkerBaseName(name)` — tách phần gốc tên (bỏ số cuối)
- `getSameNameMarkers(name)` — lấy danh sách marker cùng tên gốc
- `getDistanceMeters(lat1, lon1, lat2, lon2)` — tính khoảng cách Haversine (m)

### Tìm kiếm & điều hướng
- `searchMarkers()` — tìm theo tên (normalize không dấu), nhảy đến marker
- `goToMarker(index)` — fly đến marker theo index
- `normalizeText(text)` — bỏ dấu tiếng Việt, lowercase
- `clearSearchResults()` — xóa kết quả tìm kiếm
- `navigateToPage()` — chuyển trang theo dropdown `#pages`
- `centerOnUserLocation()` — dùng Geolocation API, setView

### Định vị & chỉ đường
- `startTrackingCurrentLocation()` — theo dõi GPS liên tục (`watchPosition`, `enableHighAccuracy:true`, `maximumAge:5000`)
- `stopTrackingCurrentLocation()` — dừng watchPosition
- `updateCurrentLocationMarker(lat, lon)` — cập nhật marker GPS; chỉ pan khi di chuyển > 15m (`lastPanLocation`)
- `setMarkerToCurrentLocation()` — GPS 1 lần (`maximumAge:0`, `timeout:15s`); spinner; hiện `±Xm`
- `createMapControls()` — tạo control panel tùy chỉnh trên bản đồ
- `routeToMapCenter()` — chỉ đường đến tâm bản đồ
- `routeToMarker(index)` — chỉ đường đến marker cụ thể (xe máy)
- `fetchMotorbikeRoute(origin, destination)` — gọi OSRM API lấy route xe máy
- `drawRoute(route)` — vẽ polyline route lên bản đồ
- `clearRoute()` — xóa route

### Lưu & xuất
- `saveMarkerData()` — xuất `loadedData[]` ra file Excel (dùng `createExcelBufferWithImages`)
- `createExcelBufferWithImages(rows)` — tạo workbook ExcelJS có nhúng ảnh (async)
- `exportMarkersToCad()` — xuất DXF (AutoCAD) với tọa độ VN2000
- `convertLatLonToVn2000(lat, lon)` — chuyển WGS84 → VN2000 (UTM zone tự động, GRS80)
- `downloadTextFile(filename, content)` — tải file text
- `sanitizeFileName(name)` — chuẩn hóa tên file

### GitHub Integration
- `updateGitHubExcel()` — sync file Excel + ảnh lên GitHub, gọi `GAS_URL` action `github_write_file`
- `uploadGitHubImageFile(...)` — upload ảnh đơn qua `GAS_URL` (`github_write_file`)
- `getGitHubFileContent(owner, repo, path, branch, token)` — lấy nội dung + SHA file từ GitHub API
- `checkGitHubBranchExists(owner, repo, branch, token)` — kiểm tra branch tồn tại

### Ảnh
- `resolveImageUrl(val)` — chuẩn hóa đường dẫn ảnh: `images/xxx.jpg` → `GITHUB_RAW_BASE + val`; `http...` và `data:` giữ nguyên
- `handleMarkerImageFile(event)` — đọc file ảnh → `resizeImageDataUrl` → base64
- `resizeImageDataUrl(dataUrl, callback)` — resize ảnh max **800px**, giữ < 32767 chars
- `getImageExtensionFromDataUrl(dataUrl)` — lấy đuôi file từ data URI
- `getBase64FromDataUrl(dataUrl)` — tách phần base64 từ data URI
- `arrayBufferToBase64(buffer)` — chuyển ArrayBuffer → base64

### Tiện ích
- `setMarkerToPickOnMap()` — bật chế độ chọn vị trí bằng cách click map
- `displayError(message)` — hiển thị thông báo lỗi/trạng thái
- `closeControlsModal()` — đóng modal controls

---

## Trạng thái toàn cục (global state)

```js
let map;                          // Leaflet map instance
let markersCluster;               // L.markerClusterGroup()
let labelLayerGroup;              // L.layerGroup() cho nhãn tên
let markers = [];                 // [{lat, lon, name, note, marker, row}]
let loadedData = [];              // Mảng dữ liệu từ Excel/CSV (array rows)
let newMarkerRows = [];           // Marker thêm mới trong session
let currentUser = null;           // { username, displayName, role } — từ localStorage('ks_user')
let editingMarkerIndex = null;    // Index trong markers[] đang chỉnh sửa (null = thêm mới)
let currentLocation = null;       // {lat, lon} vị trí GPS hiện tại
let currentLocationMarker = null; // Marker GPS hiện tại
let tempMarker = null;            // Marker tạm khi chọn vị trí
let routeLine = null;             // Polyline route đang hiển thị
let pendingNewMarker = false;     // Đang chờ thêm marker
let pendingMarkerLocation = null; // [lat, lon] vị trí chờ xác nhận
let markerPopupLocation = null;   // [lat, lon] của popup đang mở
let markerImageDataUrl = null;    // base64 ảnh đang chọn
let addMarkerMode = false;        // boolean — chế độ click-to-add
let pickMarkerMode = false;       // Chế độ click-to-pick vị trí marker
let lastMarkerName = '';          // Tên marker vừa thêm (auto-increment)
let lastMarkerNote = '';
let lastMarkerCabinet = '';
let lastMarkerPoleType = '';
let lastMarkerLampType = '';
let currentLocationWatchId = null; // ID từ watchPosition
let lastPanLocation = null;       // [lat, lon] lần pan cuối (throttle setView)
let deferredPrompt;               // PWA install prompt event

// Constants
const KHAOSAT_GAS_URL = '...';   // GAS: login, upload_image, full_update
const GAS_URL = '...';            // GAS: github_write_file (upload Excel/ảnh lên GitHub)
const KHAOSAT_CSV_URL = '...';   // Google Sheet CSV URL
const GITHUB_RAW_BASE = 'https://raw.githubusercontent.com/neo-era/cskvtt/sub1/';
const MAX_EXCEL_TEXT_LENGTH = 32767;
```

---

## Tile layers

```js
'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png'          // OSM (default)
'https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'         // Google Satellite, subdomains: mt0-mt3
'https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'         // Google Satellite (alt)
'https://ecn.t{s}.tiles.virtualearth.net/tiles/a{q}.jpeg?g=1' // Bing Maps
'https://ows.mundialis.de/services/service?'                   // HCMC Admin WMS (overlay)
```

---

## Google Apps Script (`gas-khaosat.js`)

File reference (không chạy trực tiếp ở browser). Deploy trên Google Apps Script, liên kết với Google Sheet có 2 tab:

### Tab `DanhSachTru`
Header: `ID | Tên trụ | Lat | Lon | Ghi chú | Người KS | Loại | Tủ điều khiển | Loại trụ | Loại cần | Loại đèn | Ảnh | Thời gian cập nhật | Marker gốc | Khoảng cách (m)`

### Tab `TaiKhoan`
Header: `tenDangNhap | matKhau | hoTen | vaiTro`  
(`vaiTro`: `admin` hoặc `quanly` → 👑; khác → 👷)

### Actions xử lý trong `doPost`
| Action | Hàm xử lý | Ghi chú |
|---|---|---|
| `login` | `handleLogin` | Tra cứu TaiKhoan, trả `{status, user}` |
| `upload_image` | `handleImageUpload` | Upload ảnh lên GitHub `images/` trong `neo-era/cskvtt` branch `sub1` |
| `full_update` | `findRowNum` → `updateRow`/`appendRow` | Ghi dữ liệu marker vào DanhSachTru |

### Hướng dẫn deploy
1. Mở Google Sheet → Extensions → Apps Script → dán nội dung `gas-khaosat.js`
2. Deploy → New deployment → Web App (Execute as: Me, Anyone)
3. Copy URL → dán vào `KHAOSAT_GAS_URL` trong `index.html`
4. Publish Sheet tab `DanhSachTru` dạng CSV → copy URL → `KHAOSAT_CSV_URL`
5. Script Properties: `GITHUB_TOKEN` = PAT với scope `contents:write`

---

## PWA

- `manifest.json`: name "Lighting System V1.2", start_url `index.html`, icons 192/512
- `sw.js` cache `lighting-system-v1.3`:
  - Install: `Promise.allSettled()` thay vì `addAll()` → không fail khi 1 asset thiếu; `skipWaiting()` để kích hoạt ngay
  - Activate: xóa cache version cũ, `clients.claim()`
  - Static assets: cache-first
  - Map tiles (OSM + Google): cache-first, giới hạn 200 tile, tự trim
  - Google Sheet CSV + GAS URL: luôn network (không cache)
- Meta: `apple-mobile-web-app-capable`, `theme-color`

---

## Lưu ý kỹ thuật

- **Auth**: Login qua GAS → `localStorage('ks_user')` → không có expiry; logout xóa key
- **Data flow**: Chưa login → `loadDefaultFile()` (`data/khaosat.xlsx`); Sau login → `loadDataFromSheet()` (Google Sheet CSV)
- **Ảnh**: `resolveImageUrl()` chuyển `images/xxx.jpg` → `GITHUB_RAW_BASE + val` để hiển thị từ GitHub raw; upload ảnh qua GAS proxy (không cần token ở client)
- **Resize ảnh**: max **800px**, giữ dưới 32767 chars (giới hạn Excel cell)
- **Edit marker**: `editingMarkerIndex` lưu index trong `markers[]`; `markers[].row` là tham chiếu trực tiếp vào `loadedData[]` → cập nhật in-place
- **Popup card**: CSS `.pc-*` — topbar màu theo loại trụ, avatar ảnh, grid 2 nút (Đi đến + Chỉnh sửa)
- **zoomend listener**: đăng ký **1 lần** trong `initializeMap()` — nếu đặt trong `addMarkerRowToMap()` sẽ tích lũy N listener = freeze
- **Chỉ đường**: OSRM API (xe máy), vẽ polyline; hỗ trợ chỉ đường đến tâm bản đồ hoặc marker
- **VN2000**: UTM zone tự động từ kinh độ, ellipsoid GRS80
- **GPS watchPosition**: `maximumAge:5000`; pan bản đồ chỉ khi di chuyển > 15m (`lastPanLocation`)
- **jQuery**: dùng `js/jquery.min.js` (v3.3.1) — **không dùng** `js/jquery-3.0.0.min.js` (đó là jQuery Migrate)
- **CSS**: `css/style.css` có `@import` một số file (`swiper.min.css`, v.v.) không tồn tại — chỉ là cảnh báo, không ảnh hưởng chức năng
- **GitHub token**: lưu trong GAS Script Properties, client không bao giờ nhận token
- **`GAS_URL` vs `KHAOSAT_GAS_URL`**: 2 GAS deployment khác nhau — `KHAOSAT_GAS_URL` xử lý data marker; `GAS_URL` xử lý `github_write_file` (hiện dùng nhờ GAS Den Tat)
- **Live Server**: port 5501 (`.vscode/settings.json`)

---

## Bảo mật

- **Login**: mật khẩu plaintext trong Sheet (công cụ nội bộ); tab `TaiKhoan` không publish CSV
- **GitHub token**: lưu trong GAS `Script Properties`, không bao giờ về browser; scope tối thiểu `contents:write`
- **Session**: `localStorage('ks_user')` — không có expiry tự động
- **Không commit token**: không hardcode URL GAS token hay PAT trong source

---

## Kế hoạch phát triển

### Hợp nhất 2 GAS thành 1
- Thêm `github_write_file` handler vào `gas-khaosat.js` → `index.html` chỉ cần 1 GAS URL duy nhất
- Bỏ dependency vào GAS Den Tat cho tính năng GitHub sync

### Tối ưu giao diện mobile
- Test popup form trên màn hình nhỏ (không bị chồng lên bản đồ)
- Nút control panel đủ lớn, không che tile layer switcher
- Test Android Chrome + iOS Safari (PWA)

### Session expiry
- Lưu `loggedAt` trong `currentUser`, kiểm tra khi `checkAuth()`: nếu > 24h → yêu cầu đăng nhập lại

### Phân quyền theo vai trò
- `admin/quanly`: full access (thêm/sửa/xóa)
- `user`: chỉ xem và thêm mới, không sửa marker của người khác

---

## Chạy local

Mở bằng Live Server (VS Code extension) trên port 5501, hoặc bất kỳ static file server nào.  
Không cần build step — toàn bộ là HTML/CSS/JS thuần.  
Lần đầu chạy cần xóa Service Worker cũ: DevTools → Application → Service Workers → Unregister → hard refresh.
