# CSKVTT — Quản Lý Đèn Tắt Hệ Thống Chiếu Sáng Công Cộng

**Author**: Mai Vũ Lâm  
**Version**: V1.0  
**Description**: Web app quản lý, theo dõi và báo cáo sự cố đèn chiếu sáng công cộng bị tắt/hư tại TP.HCM. Hiển thị vị trí đèn trên bản đồ Leaflet với icon màu sắc theo trạng thái. Hỗ trợ nhập/xuất Excel, thêm/sửa marker, tìm kiếm, định vị GPS, tracking vị trí, chỉ đường, đồng bộ Google Sheet qua Apps Script. Dùng cho công tác quản lý và báo cáo đèn tắt hệ thống chiếu sáng công cộng.

**URL triển khai:** `https://neo-era.github.io/cskvtt/Den%20tat/dentat.html`  
**Google Apps Script URL:** `https://script.google.com/macros/s/AKfycbz_mVOEHikgVeB9dGvGZrfXSaSHJ89yrJOoEccSjG9FnfKKsFuSwj93VfrPT6HHuzC2JQ/exec`  
**Google Sheet CSV URL:** `https://docs.google.com/spreadsheets/d/e/2PACX-1vSWWGTRtkbcls__sB1VlAcSJZTNvDigroNMJgqmce-1Ug3j181zS8VwyvZZb6jCs_W1tbFSq17YNK3B/pub?output=csv`

---

## Cấu trúc thư mục

```
cskvtt/Den tat/
├── dentat.html          # File chính — bản đồ quản lý đèn tắt
├── CLAUDE.md            # Tài liệu dự án này
├── data/
│   ├── Danhsachdentat.xlsx      # File Excel dữ liệu đèn (backup local)
│   ├── Camera.xlsx              # Dữ liệu camera (dự phòng)
│   └── thietbitichhop.xlsx      # Dữ liệu thiết bị tích hợp (dự phòng)
└── images/
    ├── 1.png            # Icon: Đèn LED đang hư
    ├── 3.png            # Icon: Đèn LED sáng bình thường
    ├── 4.png            # Icon: Đèn HPS sáng bình thường
    ├── 5.png            # Icon: Đèn hư quá 10 ngày
    ├── 7.png            # Icon: Đèn HPS đang hư
    ├── 8.png            # Icon: Marker mặc định (vị trí hiện tại)
    ├── icon-192.png     # PWA icon 192×192
    ├── icon-512.png     # PWA icon 512×512
    └── ...              # Ảnh chụp đèn sự cố (tên tự động theo marker)
```

---

## Trạng thái đèn — Mã icon

| Mã số | File icon | Ý nghĩa |
|---|---|---|
| `1` | `images/1.png` | Đèn LED đang hư |
| `3` | `images/3.png` | Đèn LED sáng bình thường |
| `4` | `images/4.png` | Đèn HPS sáng bình thường |
| `5` | `images/5.png` | Đèn hư quá 10 ngày |
| `7` | `images/7.png` | Đèn HPS đang hư |

Icon được chọn tự động dựa theo giá trị trường `trạng thái` của mỗi marker.

---

## Thư viện chính (CDN)

| Thư viện | Phiên bản | Mục đích |
|---|---|---|
| Leaflet | 1.9.4 | Bản đồ tương tác |
| Leaflet MarkerCluster | 1.5.3 | Gom nhóm marker |
| XLSX (SheetJS) | 0.18.5 | Đọc/ghi file Excel, parse CSV từ Google Sheet |
| Bootstrap | 5 (bundle) | UI responsive |
| jQuery | 3.0.0 | DOM & modal |
| Font Awesome | 4.0.3 | Icons UI |
| Google Fonts (Inter) | — | Font chữ giao diện |
| Google Apps Script | — | Backend proxy: đọc/ghi Google Sheet, cập nhật GitHub |

---

## Dữ liệu

Nguồn dữ liệu chính: **Google Sheet** (tải về qua URL CSV `pub?output=csv`), không dùng file Excel local.  
File Excel `Danhsachdentat.xlsx` chỉ dùng để backup hoặc import thủ công.

### Cấu trúc cột (rowFieldKeys — hiện tại)

> **Lưu ý:** Code hiện tại kế thừa schema từ module K76.A11. Cần cập nhật `rowFieldKeys` về schema đèn tắt bên dưới.

| Index | Tên field (code) | Tên cột Sheet | Ghi chú |
|---|---|---|---|
| 0 | `id` | ID | Mã định danh đèn/trụ |
| 1 | `name` | name / Tên | Số trụ — tên chính marker |
| 2 | `latitude` | latitude | WGS84 |
| 3 | `longitude` | longitude / longtitude | WGS84 (giữ lỗi chính tả để tương thích) |
| 4 | `north` | N | Tọa độ VN2000 X (Northing) |
| 5 | `east` | E | Tọa độ VN2000 Y (Easting) |
| 6 | `status` | trạng thái | Mã trạng thái đèn: 1/3/4/5/7 |
| 7 | `homeAddress` | địa chỉ nhà riêng | → dùng làm trường **Đường** |
| 8 | `officeAddress` | địa chỉ cơ quan | → dùng làm trường **Phường** |
| 9 | `phone1` | điện thoại 1 | → dùng làm **Ngày phát hiện** |
| 10 | `phone2` | điện thoại 2 | → dùng làm **Ngày sửa** |
| 11 | `facebook` | facebook | → dùng làm **Vật tư sửa** |
| 12 | `company` | cơ quan công tác | → dùng làm **Loại đèn** |
| 13 | `title` | chức vụ | → dùng làm **Công suất** |
| 14 | `image` | hình ảnh | Đường dẫn ảnh trong `Den tat/images/` |

### Schema đèn tắt mục tiêu (cần cập nhật rowFieldKeys)

| Index | Tên field | Tên cột Sheet | Ghi chú |
|---|---|---|---|
| 0 | `id` | ID | Mã định danh |
| 1 | `soTru` | Số trụ | Tên chính marker |
| 2 | `tenTu` | Tên tủ | Tủ điều khiển quản lý đèn |
| 3 | `latitude` | latitude | WGS84 |
| 4 | `longitude` | longitude | WGS84 |
| 5 | `loaiDen` | Loại đèn | `LED` hoặc `HPS` |
| 6 | `congSuat` | Công suất | Đơn vị W |
| 7 | `trangThai` | Trạng thái | Mã số: 1/3/4/5/7 |
| 8 | `duong` | Đường | Tên đường — tự động từ reverse geocode |
| 9 | `phuong` | Phường | Tên phường — tự động từ reverse geocode |
| 10 | `ngayPhatHien` | Ngày phát hiện | Định dạng `dd/mm/yyyy` |
| 11 | `ngaySua` | Ngày sửa | Định dạng `dd/mm/yyyy` |
| 12 | `vatTuSua` | Vật tư sửa | Danh sách linh kiện thay thế |
| 13 | `hinhAnh` | Hình ảnh | Đường dẫn trong `Den tat/images/` |
| 14 | `ghiChu` | Ghi chú | Ghi chú tự do |
| 15 | `vn2000x` | VN2000X | Tọa độ VN2000 Easting |
| 16 | `vn2000y` | VN2000Y | Tọa độ VN2000 Northing |

---

## Kiến trúc JavaScript (trong HTML)

Toàn bộ logic viết inline trong `dentat.html` (không có file JS riêng).

### Khởi tạo
- `initializeMap()` — khởi tạo Leaflet map, tile layers (OSM, Google Satellite), layer control
- `loadDataFromSheet()` — fetch CSV từ Google Sheet, parse bằng XLSX, gọi `addMarkersToMap()`
- `window.onload` — gọi `initializeMap()`, `loadDataFromSheet()`, setup draggable popup, đóng GitHub modal khi click ngoài

### Xử lý dữ liệu
- `processFile(file)` — đọc file Excel local bằng SheetJS (import thủ công), gọi `applyCoordOverrides()`
- `parseMarkerRow(row)` — parse hàng dữ liệu (array hoặc object) thành object chuẩn
- `getRowValue(row, keys, fallback)` — trích xuất giá trị từ hàng theo nhiều tên cột khác nhau
- `rowToArray(row)` — chuyển object marker → array để xuất Excel

### Marker
- `addMarkerRowToMap(row, rowIndex)` — tạo Leaflet marker + label, thêm vào cluster; icon chọn theo `status`
- `addMarkersToMap(data)` — clear và load lại toàn bộ marker
- `createMarkerPopupContent(row, rowIndex)` — tạo HTML popup hiển thị thông tin đèn
- `createMarkerIcon(parsed)` — tạo icon từ đường dẫn ảnh hoặc dùng icon mặc định
- `enableAddMarkerMode()` / `disableAddMarkerMode()` — bật/tắt chế độ click-to-add
- `startAddMarker()` — toggle chế độ thêm marker (GPS hoặc click bản đồ)
- `showMarkerPopupAt(lat, lon)` — mở form nhập marker mới tại tọa độ
- `openMarkerEditPopup(index)` — mở form chỉnh sửa marker theo index
- `fillMarkerPopupForm(row, lat, lon)` — điền dữ liệu vào form popup
- `saveMarkerPopup()` — lưu form → gửi `POST` đến GAS (action: `full_update`), cập nhật local map
- `cancelMarkerPopup()` / `hideMarkerPopup()` — đóng form popup
- `updateAddMarkerButton()` — cập nhật trạng thái nút thêm marker

### Tìm kiếm & điều hướng
- `searchMarkers()` — tìm theo tên (normalize không dấu) hoặc nhảy đến tọa độ
- `normalizeText(text)` — bỏ dấu tiếng Việt, lowercase
- `normalizeMarkerBaseName(name)` — tách phần gốc tên (bỏ số cuối)
- `getDistanceMeters(lat1, lon1, lat2, lon2)` — tính khoảng cách Haversine (m)

### Định vị & chỉ đường
- `centerOnUserLocation()` — lấy GPS một lần, setView đến vị trí
- `setMarkerToCurrentLocation()` — lấy GPS một lần (`maximumAge:0`, `timeout:15000`), điền vào form
- `makeElementDraggable(handle, target)` — drag popup bằng Pointer Events API

### Lưu & xuất
- `saveMarkerData()` — xuất `loadedData[]` ra file Excel bằng SheetJS
- `convertLatLonToVn2000(lat, lon)` — chuyển WGS84 → VN2000 (UTM zone tự động, GRS80)
- `exportMarkersToCad()` — xuất DXF (AutoCAD) với tọa độ VN2000
- `downloadFileFromDataUrl(dataUrl, filename)` — tải file từ data URI

### Google Apps Script (backend)
- `saveMarkerPopup()` — gửi `POST` với `action: "full_update"` → GAS ghi vào Sheet + cập nhật GitHub
- `submitGitHubUpdate()` — gửi `POST` chỉ với `{name, lat, lon}` → GAS chỉ cập nhật tọa độ trên Sheet
- `handleFullUpdate(index)` — gửi đầy đủ thông tin marker lên GAS (dùng khi cần sync thủ công)
- `applyCoordOverrides()` — tải `data/coords-override.json` từ GitHub, áp dụng tọa độ ghi đè

### GitHub modal
- `openGitHubModal()` / `closeGitHubModal()` — mở/đóng modal cập nhật tọa độ GPS
- `getGPSForGitHub()` — lấy GPS cho modal cập nhật
- `setGhStatus(msg, isErr)` — hiển thị trạng thái trong modal
- `openGHSettings()` / `saveGHToken()` — cài đặt và lưu GitHub token vào localStorage

### Ảnh
- `handleMarkerImageFile(event)` — đọc file ảnh → resize → base64
- `triggerMarkerImageInput()` — kích hoạt input file ảnh
- `resizeImageDataUrl(dataUrl, callback)` — resize ảnh max 400px, giữ < 32767 chars
- `normalizeImagePath(image)` — chuẩn hóa đường dẫn ảnh (xử lý `images/`, `data:`, URL)

---

## Trạng thái toàn cục (global state)

```js
const GH_OWNER = 'neo-era';
const GH_REPO  = 'cskvtt';
const GH_BRANCH = 'sub1';
const GH_COORDS_PATH = 'data/coords-override.json';
const MAX_EXCEL_TEXT_LENGTH = 32767;
const GOOGLE_SCRIPT_URL = '...';       // URL Google Apps Script proxy
const GOOGLE_SHEET_CSV_URL = '...';    // URL Google Sheet CSV public

let map;                        // Leaflet map instance
let markersCluster;             // L.markerClusterGroup({ maxClusterRadius: 40 })
let labelLayerGroup;            // L.layerGroup() cho nhãn tên
let markers = [];               // [{lat, lon, name, marker, row, rowIndex, labelMarker}]
let loadedData = [];            // Mảng dữ liệu từ Sheet/Excel
let newMarkerRows = [];         // Marker thêm mới trong session
let currentLocation = null;     // [lat, lon] vị trí GPS hiện tại
let currentLocationMarker = null;
let routeLine = null;           // Polyline route đang hiển thị
let pendingNewMarker = false;
let pendingMarkerLocation = null;
let markerPopupLocation = null;
let markerImageDataUrl = null;  // base64 ảnh đang chọn
let markerImageFileName = '';
let markerImagePath = '';
let editingMarkerIndex = null;  // Index marker đang chỉnh sửa (null = thêm mới)
let addMarkerMode = false;
let lastMarkerName = '';
let ghPendingLat = null;        // Tọa độ GPS chờ gửi lên GitHub
let ghPendingLon = null;
```

---

## Tile layers

```js
// OpenStreetMap (mặc định)
'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png'

// Google Satellite
'https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'  // subdomains: mt0–mt3
```

---

## PWA

- `manifest.json`: name "Quản Lý Đèn Tắt CSCC", short_name "Đèn Tắt", start_url `./dentat.html`, icons 192/512, theme `#0ea5e9`
- `sw.js`: 2 cache riêng biệt:
  - `dentat-static-v1` — static assets (HTML, manifest, Leaflet CSS/JS, XLSX, Google Fonts)
  - `dentat-tiles-v1` — map tiles (OSM + Google), cache-first, giới hạn 200 tile, tự trim khi vượt
  - `activate` tự xóa các cache version cũ
- **Không cache**: Google Sheet CSV và Apps Script URL (luôn fetch mới để có dữ liệu realtime)
- Meta tags: `apple-mobile-web-app-capable`, `theme-color`, `apple-touch-icon`
- Service Worker đăng ký trong `window.addEventListener('load', ...)` (trước `window.onload`)

---

## Lưu ý kỹ thuật

- **Nguồn dữ liệu chính là Google Sheet**: `loadDataFromSheet()` fetch CSV, parse bằng SheetJS `XLSX.read(csvText, {type:'string'})`
- **Không dùng file data GitHub**: File Excel local chỉ dùng để import thủ công qua input file
- **Backend proxy GAS**: Mọi thao tác ghi (thêm/sửa marker) đều đi qua Google Apps Script — không gọi GitHub API trực tiếp từ browser
- **`action: "full_update"`**: Khi lưu form chỉnh sửa, payload phải có `action: "full_update"` để GAS phân biệt với request chỉ cập nhật tọa độ GPS
- **Icon theo trạng thái**: `createMarkerIcon()` đọc trường `status` → chọn file icon `images/{status}.png`
- **VN2000**: tính theo UTM zone tự động từ kinh độ, dùng ellipsoid GRS80
- **Giới hạn ảnh**: 32767 ký tự/ô Excel — `resizeImageDataUrl()` tự scale xuống
- **MarkerCluster + Label**: label dùng `L.divIcon` thêm vào `labelLayerGroup` riêng (không cluster)
- **Lazy popup**: `bindPopup('')` khi tạo marker — HTML popup chỉ render khi mở, giảm tải CPU lúc load
- **GPS getCurrentPosition**: `maximumAge:0` + `timeout:15000` — luôn lấy vị trí mới, hiện spinner chờ
- **GitHub token**: lưu trong `localStorage` key `'gh_token_k76a11'`, người dùng nhập mỗi phiên
- **`coords-override.json`**: file JSON trên GitHub ghi đè tọa độ từ Sheet (dùng khi học viên/kỹ thuật viên cập nhật GPS từ điện thoại)
- **Live Server**: port 5501 (`.vscode/settings.json`)

---

## Bảo mật

- **GitHub token**: lưu tạm trong `localStorage`, không hardcode vào source. GAS proxy sẽ dùng `PropertiesService.getScriptProperties()` khi nâng cấp
- **Scope token tối thiểu**: PAT chỉ cần `contents:write` cho repo `neo-era/cskvtt`
- **Không commit token**: `.gitignore` hoặc quy trình CI không được để token lọt vào lịch sử commit
- **GAS là proxy an toàn**: client gửi dữ liệu đến GAS, GAS tự thêm credentials khi gọi GitHub API

---

## Kế hoạch phát triển

### Cập nhật schema dữ liệu đèn tắt
- Đổi `rowFieldKeys` từ schema K76.A11 (tên/điện thoại/facebook) sang schema đèn tắt (số trụ, loại đèn, trạng thái, ngày phát hiện, ngày sửa, vật tư)
- Cập nhật `parseMarkerRow()`, `fillMarkerPopupForm()`, `createMarkerPopupContent()` theo schema mới
- Cập nhật `rowToArray()` và header Excel khi xuất

### Reverse geocode tự động
- Khi thêm marker mới, tự động gọi Nominatim API lấy tên đường và phường từ tọa độ GPS
- Điền vào trường `duong` và `phuong` thay vì nhập tay

### Xuất báo cáo hàng ngày
- Thêm chức năng lọc marker theo ngày phát hiện
- Xuất Excel báo cáo các đèn hư trong ngày với đầy đủ thông tin sự cố

### Bộ lọc theo trạng thái
- Thêm UI lọc marker: chỉ hiện đèn hư / đèn đã sửa / đèn hư quá 10 ngày
- Đếm số lượng theo từng trạng thái trên UI

### Nâng cấp PWA offline
- ✅ Đã có `manifest.json` và `sw.js`
- Có thể pre-cache tile vùng TP.HCM khi install để dùng hoàn toàn offline ngoài thực địa

### Tối ưu giao diện mobile
- Nút thao tác đủ lớn cho màn hình nhỏ (ngón tay)
- Test trên Android Chrome + iOS Safari

---

## Chạy local

Mở bằng Live Server (VS Code extension) trên port 5501, trỏ vào `Den tat/dentat.html`.  
Không cần build step — toàn bộ là HTML/CSS/JS thuần.
