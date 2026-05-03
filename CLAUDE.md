# CSKVTT — Bản Đồ Khảo Sát Hệ Thống Chiếu Sáng Công Cộng

**Author**: Mai Vũ Lâm  
**Version**: V1.1  
**Description**: Web app quản lý và hiển thị vị trí hạ tầng chiếu sáng công cộng (TP.HCM) trên bản đồ Leaflet. Hỗ trợ nhập/xuất Excel, thêm/sửa marker, tìm kiếm, định vị GPS, xuất CAD và PWA offline.

---

## Cấu trúc thư mục

```
cskvtt/
├── index.html          # Trang chính — khảo sát chiếu sáng (tất cả quận)
├── k76a11.html         # Trang lớp K76.A11 — danh sách học viên
├── camera-map.html     # Bản đồ camera giám sát
├── quan1.html          # Bản đồ riêng Quận 1
├── test.html           # Bản sao index.html dùng để test
├── backup-index.html   # Backup index.html
├── css/                # 26 CSS files (Bootstrap 5, Font Awesome, plugins)
├── js/                 # 17 JS files (jQuery 3, Bootstrap bundle, plugins)
├── images/             # Icons UI, logo, favicon, PWA icons
├── images1/            # Ảnh học viên K76.A11 + bản sao icons UI
├── data/               # 23 file Excel (.xlsx) — dữ liệu marker từng quận
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
| ExcelJS | 4.3.1 | Xử lý workbook |
| Bootstrap | 5 (bundle) | UI responsive |
| jQuery | 3.0.0 | DOM & modal |
| Font Awesome | 4.0.3 | Icons |

---

## Dữ liệu (`/data`)

Mỗi file Excel có cột:  
`ID | name | latitude | longtitude | N | E | trạng thái | địa chỉ nhà riêng | địa chỉ cơ quan | điện thoại 1 | điện thoại 2 | facebook | cơ quan công tác | chức vụ | hình ảnh`

File mặc định mỗi trang load: `data/danhsach.xlsx` (k76a11.html).

---

## Kiến trúc JavaScript (trong HTML)

Toàn bộ logic ứng dụng được viết inline trong mỗi file HTML (không có file JS riêng cho business logic). Các hàm chính:

### Khởi tạo
- `initializeMap()` — khởi tạo Leaflet map, tile layers, event listeners
- `loadDefaultFile()` — fetch `data/danhsach.xlsx` và gọi `processFile()`
- `window.onload` — gọi `initializeMap()`, `loadDefaultFile()`, setup draggable popup

### Xử lý dữ liệu
- `parseMarkerRow(row)` — chuẩn hóa 1 hàng Excel → object `{id, name, lat, lon, north, east, status, ...}`
- `rowFieldKeys` — map tên cột (đa ngôn ngữ, chỉ số) → field name
- `getRowValue(row, keys, fallback)` — đọc giá trị từ row theo nhiều key khả năng
- `rowToArray(row)` — chuyển row → mảng để xuất Excel

### Marker
- `addMarkerRowToMap(row, rowIndex)` — tạo Leaflet marker + label, thêm vào cluster
- `addMarkersToMap(data)` — clear và load lại toàn bộ marker
- `createMarkerIcon(parsed)` — tạo icon từ ảnh hoặc dùng `customIcons.default`
- `createMarkerPopupContent(row, rowIndex)` — tạo HTML nội dung popup
- `openMarkerEditPopup(index)` — mở form chỉnh sửa marker đã có
- `showMarkerPopupAt(lat, lon)` — mở form thêm marker mới tại tọa độ
- `saveMarkerPopup()` — lưu form vào `loadedData[]`, cập nhật map
- `cancelMarkerPopup()` / `hideMarkerPopup()` — đóng form

### Tìm kiếm & điều hướng
- `searchMarkers()` — tìm theo tên (normalize không dấu), nhảy đến marker
- `normalizeText(text)` — bỏ dấu tiếng Việt, lowercase
- `centerOnUserLocation()` — dùng Geolocation API, setView

### Lưu & xuất
- `saveMarkerData()` — xuất `loadedData[]` ra file `danhsach-updated.xlsx`
- `exportMarkersToCad()` — xuất DXF (AutoCAD) với tọa độ VN2000
- `convertLatLonToVn2000(lat, lon)` — chuyển WGS84 → VN2000 (UTM zone tự động)

### Ảnh
- `handleMarkerImageFile(event)` — đọc file ảnh → resize → base64
- `resizeImageDataUrl(dataUrl, callback)` — resize ảnh max 400px, giữ < 32767 chars (giới hạn Excel cell)
- `normalizeImagePath(image)` — chuẩn hóa đường dẫn ảnh (hỗ trợ data URI, HTTP, path tương đối)

---

## Trạng thái toàn cục (global state)

```js
let map;                    // Leaflet map instance
let markersCluster;         // L.markerClusterGroup()
let labelLayerGroup;        // L.layerGroup() cho nhãn tên
let markers = [];           // [{lat, lon, name, marker, row, rowIndex, labelMarker}]
let loadedData = [];        // Mảng dữ liệu từ Excel (object rows)
let newMarkerRows = [];     // Marker thêm mới trong session
let editingMarkerIndex;     // null hoặc index đang sửa
let addMarkerMode;          // boolean — chế độ click-to-add
let pendingNewMarker;       // boolean — đang chờ click trên map
let markerPopupLocation;    // [lat, lon] của popup đang mở
let markerImageDataUrl;     // base64 ảnh đang chọn
let markerImagePath;        // đường dẫn lưu ảnh (images1/...)
```

---

## Tile layers

```js
// OpenStreetMap (default)
'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png'

// Google Satellite
'https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}'
// subdomains: ['mt0','mt1','mt2','mt3']
```

---

## PWA

- `manifest.json`: name "Lighting System V1.1", start_url `index.html`, icons 192/512
- `sw.js`: cache-first strategy, cache key `lighting-map-v1`
- Meta tags: `apple-mobile-web-app-capable`, `theme-color`

---

## Lưu ý kỹ thuật

- **Giới hạn ảnh Excel**: 32767 ký tự/ô — `resizeImageDataUrl` tự scale ảnh xuống
- **VN2000**: tính theo UTM zone tự động từ kinh độ, dùng ellipsoid GRS80
- **Tọa độ**: WGS84 (lat/lon), cột Excel tên là `longtitude` (sai chính tả, giữ nguyên để tương thích)
- **Modal left**: CSS custom slide-in từ trái, width 400px desktop / 100vw mobile
- **Popup form draggable**: dùng Pointer Events API (`pointerdown/move/up`)
- **MarkerCluster + Label**: label dùng `L.divIcon` thêm vào `labelLayerGroup` riêng (không cluster)
- **Live Server**: port 5501 (`.vscode/settings.json`)

---

## Chạy local

Mở bằng Live Server (VS Code extension) trên port 5501, hoặc bất kỳ static file server nào.  
Không cần build step — toàn bộ là HTML/CSS/JS thuần.
