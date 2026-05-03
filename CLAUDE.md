# CSKVTT — Bản Đồ Khảo Sát Hệ Thống Chiếu Sáng Công Cộng

**Author**: Mai Vũ Lâm  
**Version**: V1.1  
**Description**: Web app quản lý và hiển thị vị trí hạ tầng chiếu sáng công cộng (TP.HCM) trên bản đồ Leaflet. Hỗ trợ nhập/xuất Excel, thêm/sửa marker, tìm kiếm, định vị GPS, chỉ đường, xuất CAD, đồng bộ GitHub và PWA offline.

---

## Cấu trúc thư mục

```
cskvtt/
├── index.html          # Trang chính — khảo sát chiếu sáng (tất cả quận)
├── camera-map.html     # Bản đồ camera giám sát
├── quan1.html          # Bản đồ riêng Quận 1
├── test.html           # Bản sao index.html dùng để test
├── backup-index.html   # Backup index.html
├── css/                # CSS files (Bootstrap 5, Font Awesome, plugins)
├── js/                 # JS files (jQuery 3, Bootstrap bundle, plugins)
├── images/             # Icons UI, logo, favicon, PWA icons
├── images1/            # Ảnh marker + bản sao icons UI
├── data/               # File Excel (.xlsx) — dữ liệu marker từng quận
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
| XLSX (SheetJS) | 0.18.5 | Đọc/ghi file Excel |
| ExcelJS | 4.3.1 | Xử lý workbook (lưu có ảnh) |
| Bootstrap | 5 (bundle) | UI responsive |
| jQuery | 3.0.0 | DOM & modal |
| Font Awesome | 4.0.3 | Icons |

---

## Dữ liệu (`/data`)

File mặc định index.html load: `data/khaosat.xlsx`

Mỗi hàng Excel (dùng chỉ số cột, không có header):

| Index | Tên trường | Ghi chú |
|---|---|---|
| 0 | ID | Có thể để trống |
| 1 | tên trụ | Tên chính marker |
| 2 | latitude | WGS84 |
| 3 | longitude | WGS84 (tên cột cũ là `longtitude` — lỗi chính tả, giữ nguyên) |
| 4 | ghi chú | |
| 5 | (trống) | |
| 6 | loại | Số nguyên: 1=trụ thường, 2=STK, 6=có tủ điều khiển |
| 7 | tủ điều khiển | |
| 8 | loại trụ | Ví dụ: STK, BTLT... |
| 9 | loại cần | |
| 10 | loại đèn | |
| 11 | hình ảnh | base64 data URI hoặc đường dẫn file |
| 12 | thời gian cập nhật | `toLocaleString('vi-VN')` |
| 13 | marker gốc | Tên marker tham chiếu |
| 14 | khoảng cách (m) | Khoảng cách đến marker gốc |

---

## Kiến trúc JavaScript (trong HTML)

Toàn bộ logic viết inline trong mỗi file HTML (không có file JS riêng). Các hàm chính:

### CAD Export Modal
- `showExportCadModal()` — mở modal chọn loại xuất DXF
- `closeExportCadModal()` — đóng modal
- `updateExportCadCountInfo()` — cập nhật thông tin số lượng marker sẽ xuất
- `confirmExportCad()` — tạo và tải file DXF

### Khởi tạo
- `initializeMap()` — khởi tạo Leaflet map, tile layers (OSM, Google Satellite, Bing, WMS), control layers, createMapControls
- `loadDefaultFile()` — fetch `data/khaosat.xlsx` và gọi `processFile()`
- `window.onload` — gọi `initializeMap()`, `loadDefaultFile()`, setup draggable popup

### Xử lý dữ liệu
- `processFile(file)` — đọc file Excel bằng SheetJS, lưu vào `loadedData[]`, gọi `addMarkersToMap()`

### Marker
- `addMarkerRowToMap(row)` — tạo Leaflet marker (draggable) + label, thêm vào cluster, bind popup
- `addMarkersToMap(data)` — clear và load lại toàn bộ marker
- `createMarkerPopupContent(row)` — tạo HTML nội dung popup (hiển thị trụ cùng tên, khoảng cách)
- `updateMarkerCoordinatesInData(markerName, newLat, newLon)` — cập nhật tọa độ trong `loadedData` khi kéo marker
- `enableAddMarkerMode()` / `disableAddMarkerMode()` — bật/tắt chế độ click-to-add
- `updateAddMarkerButton()` — cập nhật trạng thái nút thêm marker
- `startAddMarker()` — toggle chế độ thêm marker
- `showMarkerPopupAt(lat, lon)` — mở form nhập marker mới tại tọa độ
- `saveMarkerPopup()` — lưu form vào `loadedData[]` + `newMarkerRows[]`, cập nhật map
- `cancelMarkerPopup()` / `hideMarkerPopup()` — đóng form
- `makeElementDraggable(handle, target)` — dùng Pointer Events API để drag popup

### Tên & khoảng cách trụ
- `getNextMarkerName(previousName)` — tự động tăng số cuối tên trụ (VD: "A01" → "A02")
- `normalizeMarkerBaseName(name)` — tách phần gốc tên (bỏ số cuối)
- `getMarkerCountByName(name)` — đếm marker cùng tên gốc
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
- `startTrackingCurrentLocation()` — theo dõi GPS liên tục (watchPosition)
- `stopTrackingCurrentLocation()` — dừng watchPosition
- `updateCurrentLocationMarker(lat, lon)` — cập nhật marker vị trí hiện tại
- `createMapControls()` — tạo control panel tùy chỉnh trên bản đồ
- `routeToMapCenter()` — chỉ đường đến tâm bản đồ
- `routeToMarker(index)` — chỉ đường đến marker cụ thể (xe máy)
- `fetchMotorbikeRoute(origin, destination)` — gọi OSRM API lấy route xe máy
- `drawRoute(route)` — vẽ polyline route lên bản đồ
- `clearRoute()` — xóa route

### Lưu & xuất
- `saveMarkerData()` — xuất `loadedData[]` ra file Excel (dùng `createExcelBufferWithImages`)
- `createExcelBufferWithImages(rows)` — tạo workbook ExcelJS có nhúng ảnh (async)
- `attachImagesToWorksheet(worksheet, rows, startRow)` — chèn ảnh vào worksheet
- `prepareImagePathsForSave(rows)` — lưu ảnh base64 ra file, cập nhật đường dẫn vào row
- `exportMarkersToCad()` — xuất DXF (AutoCAD) với tọa độ VN2000
- `createDxfForMarkers(rows)` — tạo nội dung file DXF
- `convertLatLonToVn2000(lat, lon)` — chuyển WGS84 → VN2000 (UTM zone tự động, ellipsoid GRS80)
- `downloadTextFile(filename, content)` — tải file text
- `downloadDataUrlFile(dataUrl, suggestedPath)` — tải file từ data URI
- `sanitizeFileName(name)` — chuẩn hóa tên file

### GitHub Integration
- `updateGitHubExcel()` — đồng bộ file Excel + ảnh lên GitHub (async)
- `uploadGitHubImageFile(...)` — upload ảnh đơn lên GitHub qua API (async)
- `prepareImagePathsForGitHub(rows, owner, repo, branch, token)` — upload tất cả ảnh, cập nhật đường dẫn (async)
- `getGitHubFileContent(owner, repo, path, branch, token)` — lấy nội dung file từ GitHub (async)
- `updateGitHubFile(owner, repo, path, branch, token, contentBase64, sha)` — cập nhật file trên GitHub (async)
- `getGitHubRepoContents(owner, repo, path, branch, token)` — liệt kê thư mục GitHub (async)
- `renderGitHubContents(contents, currentPath)` — render danh sách file/thư mục
- `browseGitHubPath(path)` — điều hướng thư mục GitHub (async)
- `selectGitHubFilePath(filePath)` — chọn file từ trình duyệt GitHub
- `checkGitHubBranchExists(owner, repo, branch, token)` — kiểm tra branch tồn tại (async)

### Ảnh
- `handleMarkerImageFile(event)` — đọc file ảnh → resize → base64
- `triggerMarkerImageInput()` — kích hoạt input file ảnh
- `resizeImageDataUrl(dataUrl, callback)` — resize ảnh max 400px, giữ < 32767 chars (giới hạn Excel cell)
- `getImageExtensionFromDataUrl(dataUrl)` — lấy đuôi file từ data URI
- `getImagePathForMarker(name, dataUrl)` — tạo đường dẫn file ảnh từ tên marker
- `getBase64FromDataUrl(dataUrl)` — tách phần base64 từ data URI
- `arrayBufferToBase64(buffer)` — chuyển ArrayBuffer → base64

### Tiện ích
- `setMarkerToCurrentLocation()` — đặt vị trí marker về GPS hiện tại
- `setMarkerToPickOnMap()` — bật chế độ chọn vị trí bằng cách click map
- `displayError(message)` — hiển thị thông báo lỗi/trạng thái
- `closeControlsModal()` — đóng modal controls

---

## Trạng thái toàn cục (global state)

```js
let map;                        // Leaflet map instance
let markersCluster;             // L.markerClusterGroup()
let labelLayerGroup;            // L.layerGroup() cho nhãn tên
let markers = [];               // [{lat, lon, name, marker, row, labelMarker}]
let loadedData = [];            // Mảng dữ liệu từ Excel (array rows)
let newMarkerRows = [];         // Marker thêm mới trong session
let currentLocation = null;     // {lat, lon} vị trí GPS hiện tại
let currentLocationMarker = null; // Marker GPS hiện tại
let tempMarker = null;          // Marker tạm khi chọn vị trí
let routeLine = null;           // Polyline route đang hiển thị
let pendingNewMarker = false;   // Đang chờ thêm marker
let pendingMarkerLocation = null; // [lat, lon] vị trí chờ xác nhận
let markerPopupLocation = null; // [lat, lon] của popup đang mở
let markerImageDataUrl = null;  // base64 ảnh đang chọn
let addMarkerMode = false;      // boolean — chế độ click-to-add
let lastMarkerName = '';        // Tên marker vừa thêm (để auto-increment)
let lastMarkerNote = '';
let lastMarkerCabinet = '';
let lastMarkerPoleType = '';
let lastMarkerLampType = '';
let githubOwner = '';           // GitHub integration settings
let githubRepo = '';
let githubBranch = '';
let githubToken = '';
let currentLocationWatchId = null; // ID từ watchPosition
let deferredPrompt;             // PWA install prompt event
let pickMarkerMode = false;     // Chế độ click-to-pick vị trí marker
const MAX_EXCEL_TEXT_LENGTH = 32767; // Giới hạn ký tự/ô Excel
```

---

## Tile layers

```js
// OpenStreetMap (default)
'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png'

// Google Satellite (2 variant)
'https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'  // subdomains: mt0-mt3
'https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'

// Bing Maps
'https://ecn.t{s}.tiles.virtualearth.net/tiles/a{q}.jpeg?g=1'

// HCMC Admin WMS (overlay)
'https://ows.mundialis.de/services/service?' (OSM-Overlay-WMS)
```

---

## PWA

- `manifest.json`: name "Lighting System V1.1", start_url `index.html`, icons 192/512
- `sw.js`: cache-first strategy, cache key `lighting-map-v1`
- Meta tags: `apple-mobile-web-app-capable`, `theme-color`
- Install prompt: `deferredPrompt` (BeforeInstallPromptEvent)

---

## Lưu ý kỹ thuật

- **Giới hạn ảnh Excel**: 32767 ký tự/ô — `resizeImageDataUrl` tự scale ảnh xuống
- **VN2000**: tính theo UTM zone tự động từ kinh độ, dùng ellipsoid GRS80
- **Tọa độ**: WGS84 (lat/lon), cột Excel tên cũ là `longtitude` (sai chính tả, giữ nguyên để tương thích)
- **Dữ liệu dùng chỉ số**: `row[0..14]` (không dùng object key như phiên bản cũ)
- **Marker draggable**: kéo marker cập nhật tọa độ trong `loadedData` qua `updateMarkerCoordinatesInData`
- **Chỉ đường**: OSRM API (xe máy), vẽ polyline, hỗ trợ chỉ đường đến tâm bản đồ hoặc marker cụ thể
- **GitHub sync**: upload Excel + ảnh base64 lên repo qua GitHub REST API (cần token)
- **Popup form draggable**: dùng Pointer Events API (`pointerdown/move/up`)
- **MarkerCluster + Label**: label dùng `L.divIcon` thêm vào `labelLayerGroup` riêng (không cluster)
- **Live Server**: port 5501 (`.vscode/settings.json`)

---

## Chạy local

Mở bằng Live Server (VS Code extension) trên port 5501, hoặc bất kỳ static file server nào.  
Không cần build step — toàn bộ là HTML/CSS/JS thuần.
