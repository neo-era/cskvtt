# CLAUDE.md — Lighting System V1.1

> File này là context và spec cho AI (Claude / v0 / Cursor / Copilot) khi làm việc trên dự án **Lighting System V1.1** — web app bản đồ khảo sát hệ thống chiếu sáng công cộng.

---

## 1. Tổng quan dự án

- **Tên:** Lighting System V1.1
- **Mục đích:** Ứng dụng bản đồ khảo sát hệ thống chiếu sáng công cộng, cho phép kỹ sư cắm marker tại vị trí trụ đèn, nhập thông số, chụp ảnh, đồng bộ dữ liệu lên Google Sheets, hình ảnh lên GitHub, và xuất file CAD (DXF).
- **Bối cảnh sử dụng:**
  - Người dùng chính: kỹ sư khảo sát, làm việc ngoài trời (ánh sáng mạnh, cầm điện thoại 1 tay, có thể đeo găng).
  - Thiết bị: ~70% mobile (Android/iOS, 5–6.7 inch), ~30% desktop để xem báo cáo & export.
  - Tác vụ chính: cắm marker → nhập thông số → chụp ảnh → đồng bộ Google Sheets + GitHub.

---

## 2. Tech stack (BẮT BUỘC GIỮ)

### Frontend
- Vanilla HTML / CSS / JS
- Bootstrap 4
- Leaflet 1.9.4 + `leaflet.markercluster`
- jQuery 3 + FontAwesome 4
- **KHÔNG dùng framework** (React/Vue) — phải chạy được như file HTML đơn

### Lưu trữ & đồng bộ
- **Dữ liệu (marker, thông số trụ):** Google Sheets (qua Google Apps Script Web App — xem mục 3).
- **Hình ảnh:** GitHub repo (qua GitHub Contents API, đường dẫn được ghi ngược lại Google Sheets).
- **Export local:** XLSX + ExcelJS (chỉ dùng cho nút "Tải Excel" và "Export to CAD/DXF", không còn là nguồn dữ liệu chính).

---

## 3. Data Architecture

### 3.1. Sơ đồ luồng dữ liệu

```
┌────────────┐    POST marker     ┌──────────────────────┐    append    ┌──────────────┐
│  Web App   │ ─────────────────► │ Apps Script Web App  │ ───────────► │ Google Sheet │
│ (browser)  │ ◄───────────────── │   (doGet / doPost)   │ ◄─── read ── │              │
└─────┬──────┘    JSON response   └──────────────────────┘              └──────────────┘
      │
      │ PUT image (base64)
      ▼
┌──────────────┐
│  GitHub API  │ ──► raw.githubusercontent.com/.../images/<ten-tru>-<timestamp>.jpg
└──────────────┘
      ▲
      │ URL ảnh
      └──────────► được ghi vào cột "Ảnh" của Google Sheet
```

### 3.2. Google Sheets — vai trò "database"

**Cách kết nối: Apps Script Web App** (recommended)

Lý do chọn Apps Script thay vì gọi thẳng Google Sheets API v4:
- Không lộ API key / OAuth token trong frontend.
- Tránh phải share sheet công khai.
- Có thể thêm validation, dedupe, format ở server-side.
- Miễn phí, không cần backend riêng.
- CORS dễ xử lý (Apps Script tự bật).

**Cấu trúc sheet:**

| Sheet name        | Vai trò                                                       |
| ----------------- | ------------------------------------------------------------- |
| `markers`         | Bảng chính, mỗi row = 1 trụ                                   |
| `cabinets`        | Bảng tủ điều khiển (tham chiếu)                               |
| `audit_log`       | Log mọi thay đổi (created_at, updated_at, user, action)       |
| `config`          | Cấu hình runtime (GitHub repo path, branch, options)          |

**Schema sheet `markers`** (cột A → O, giữ tương thích với array hiện tại `row[0..14]`):

| Cột | Index | Tên                | Kiểu        | Ghi chú                                  |
| --- | ----- | ------------------ | ----------- | ---------------------------------------- |
| A   | 0     | `id`               | string      | UUID hoặc auto-increment                 |
| B   | 1     | `ten_tru`          | string      | Tên hiển thị (vd `T1-001`)               |
| C   | 2     | `lat`              | number      | Vĩ độ                                    |
| D   | 3     | `lon`              | number      | Kinh độ                                  |
| E   | 4     | `ghi_chu`          | string      |                                          |
| F   | 5     | `nguoi_khaosat`    | string      | Lấy tên đăng nhập                 |
| G   | 6     | `loai_marker`      | number      | 1=trụ thường, 2=STK, 6=tủ điều khiển     |
| H   | 7     | `tu_dieu_khien`    | string      |                                          |
| I   | 8     | `loai_tru`         | string      |                                          |
| J   | 9     | `loai_can`         | string      |                                          |
| K   | 10    | `loai_den`         | string      |                                          |
| L   | 11    | `anh_url`          | string      | URL GitHub raw (KHÔNG lưu base64)        |
| M   | 12    | `updated_at`       | datetime    | ISO 8601                                 |
| N   | 13    | `marker_goc`       | string      | Tên marker tham chiếu để tính khoảng cách |
| O   | 14    | `khoang_cach_m`    | number      | Tới marker gốc                           |

### 3.3. Endpoints Apps Script Web App

Deploy 1 file `Code.gs` thành Web App (Execute as: Me, Access: Anyone). Tất cả request qua `POST` với body JSON, có field `action`:

| Action            | Mô tả                                                       | Body                                  |
| ----------------- | ----------------------------------------------------------- | ------------------------------------- |
| `list_markers`    | Lấy toàn bộ marker (có filter optional theo `district`)     | `{ action, district? }`               |
| `create_marker`   | Thêm 1 marker mới                                           | `{ action, marker: {...} }`           |
| `update_marker`   | Cập nhật theo `id`                                          | `{ action, id, patch: {...} }`        |
| `delete_marker`   | Xóa mềm (set flag, không xóa row)                           | `{ action, id }`                      |
| `bulk_upsert`     | Đồng bộ batch (cho chế độ offline)                          | `{ action, markers: [...] }`          |
| `get_config`      | Lấy cấu hình GitHub (owner, repo, branch, path)             | `{ action }`                          |

**Response shape chuẩn:**
```json
{ "ok": true, "data": [...], "error": null }
```

### 3.4. GitHub — chỉ để lưu ảnh

- Repo: `neo-era/cskvtt` (đã hardcode trong code hiện tại — chuyển vào Apps Script `config`).
- Branch: `sub1` (chuyển vào config).
- Đường dẫn: `images/<ten-tru-sanitized>-<timestamp>.<ext>`
- URL truy cập:
  - `https://raw.githubusercontent.com/<owner>/<repo>/<branch>/images/...` (trực tiếp, không cache)
  - hoặc `https://cdn.jsdelivr.net/gh/<owner>/<repo>@<branch>/images/...` (qua CDN, nhanh hơn nhiều cho production)
- **GitHub token KHÔNG được hardcode trong frontend.** Có 2 lựa chọn:
  - **(a) Đi qua Apps Script:** Apps Script giữ token, frontend POST ảnh base64 → Apps Script upload lên GitHub → trả URL về. *Khuyến nghị.*
  - **(b) GitHub App với fine-grained token chỉ ghi vào folder `images/`:** vẫn rủi ro nếu lộ.

### 3.5. Chế độ offline

Khảo sát ngoài trời có thể mất 4G. App phải:
- Cache mọi marker mới vào `localStorage` (key: `pending_markers`) khi chưa đồng bộ được.
- Hiển thị badge "⏳ N marker chưa đồng bộ" ở top bar.
- Có nút "Đồng bộ ngay" trong panel điều khiển.
- Khi có mạng lại: gọi `bulk_upsert` rồi xóa cache.
- Ảnh: cache base64 trong IndexedDB (vì có thể >5MB tổng), upload sau.

---

## 4. Quy tắc khi redesign UI

Giữ nguyên **toàn bộ logic JavaScript** hiện tại, **chỉ thay đổi** HTML structure và CSS.

### 4.1. Cấu trúc UI

#### Map full-screen
- `#map` chiếm `100vw × 100vh` làm nền.

#### Top bar nổi (thay FAB "Điều khiển" hiện tại)
- Mảnh, dạng "command bar" lơ lửng phía trên (giống Google Maps mobile):
  - Trái: logo / icon ứng dụng
  - Giữa: ô tìm kiếm trụ + tọa độ (luôn nhìn thấy, không cần mở modal)
  - Phải: avatar / nút menu (mở panel điều khiển)
- Trên mobile: chỉ giữ search bar + nút menu hamburger.

#### Bottom action bar (mobile-first, thumb-reachable)
Hàng nút lớn ở đáy màn hình, mỗi nút cao ≥ 56px:
- `[+]` Thêm marker (primary, to nhất, nổi bật)
- `[📍]` Vị trí của tôi
- `[🧭]` Tracking on/off
- `[📤]` Export / Lưu (mở bottom sheet với các option)

#### Side panel "Điều khiển" (thay modal slide-left hiện tại)
- **Desktop:** panel cố định bên trái, rộng 320px, có thể collapse.
- **Mobile:** bottom sheet kéo lên (drag handle ở trên), **KHÔNG phải full-screen modal**.
- Nội dung: dropdown chọn quận, đường dẫn GitHub, các nút Export to CAD, Cập nhật GitHub, Lưu dữ liệu.

#### Marker form (cải thiện popup hiện tại)
Hiện tại form đè giữa màn hình, che bản đồ → khó căn vị trí. Đổi thành:
- **Mobile:** bottom sheet kéo từ dưới lên, có 2 mức cao (50% / 90% màn hình), để người dùng vẫn nhìn được pin trên bản đồ khi nhập.
- **Desktop:** panel phải hoặc dialog có thể kéo.
- Các trường: Tên trụ, Marker gốc (cùng basename), Tủ điều khiển, Loại trụ, Loại đèn, Ghi chú, Ảnh, Tọa độ.
- 3 nút chính ở footer luôn fixed visible: `[Hủy] [Lưu marker] [📷 Ảnh]`.
- Hiển thị khoảng cách tới marker gốc realtime khi nhập (logic đã có).

#### Floating controls cụm phải-dưới
Giữ nhưng làm sang hơn: Tracking, Locate, Route, Clear route — dạng cluster button bo tròn, có shadow, kèm tooltip.

#### Toast notification (thay `#error` hiện tại)
Hiện tại đang dùng 1 div đỏ làm cả error và success → confusing. Đổi thành toast nổi góc trên/dưới, tự ẩn sau 3s, có 3 màu:
- `success` (xanh), `warning` (vàng), `error` (đỏ).

### 4.2. Phong cách

- **Tham chiếu:** Google Maps mobile (layout) + Mapbox Studio (control style) + Linear (typography sạch).
- **Tone:** chuyên nghiệp kỹ thuật, không lòe loẹt, ưu tiên function.
- **Color palette:**
  - Primary: `#2563eb` (xanh dương, gần màu LED ban ngày)
  - Secondary: `#f59e0b` (vàng amber — gợi đèn đường)
  - Success: `#10b981`, Warning: `#f59e0b`, Danger: `#ef4444`
  - Surface: `#ffffff` với shadow nhẹ, border `#e5e7eb`
  - Text: `#111827` (chính), `#6b7280` (phụ)
- **Typography:** System font stack (`-apple-system, Segoe UI, Roboto`) — **KHÔNG load Google Fonts** (chậm khi 4G yếu ngoài hiện trường).
- **Border radius:** 12px cho card, 8px cho button, 999px cho FAB.
- **Shadow:** nhiều lớp mềm (giống Material 3 elevation).
- **Icon:** giữ FontAwesome 4 (đã có sẵn).

### 4.3. Chi tiết quan trọng

- **Outdoor readability:** tăng contrast text, dùng nền trắng đặc thay vì `rgba` trong suốt (ngoài nắng nhìn không rõ).
- **Touch targets:** tất cả button tương tác ≥ 44×44px (Apple HIG).
- **Một tay sử dụng:** action quan trọng ở 1/3 dưới màn hình.
- **Mất mạng:** hiển thị indicator offline ở top bar khi mất GPS / network.
- **Safe area:** dùng `env(safe-area-inset-bottom)` cho bottom bar trên iPhone có notch.
- **Dark mode:** thêm `@media (prefers-color-scheme: dark)` — hữu ích khi khảo sát buổi tối (nghiệm thu đèn sáng).

### 4.4. Responsive breakpoint

| Range        | Layout                        |
| ------------ | ----------------------------- |
| `< 480px`    | Mobile dọc (chế độ chính)     |
| `480–768px`  | Mobile ngang / tablet nhỏ     |
| `> 768px`    | Tablet / desktop (panel trái) |

---

## 5. ID / Class JavaScript đang phụ thuộc — KHÔNG ĐƯỢC ĐỔI

Các selector này được JS gọi trực tiếp, đổi sẽ làm vỡ logic:

```
#map                    #controlsBtn            #controlsModal
#markerPopupForm        #searchInput            #error
#pages                  #githubPathInput        #addMarkerBtn
#markerNameInput        #markerCabinetInput     #markerPoleTypeInput
#markerLampTypeInput    #markerNoteInput        #markerCoordsText
#markerImagePreview     #markerImageInput       #markerBaseSelect
#exportCadModal         #exportCadType          #exportCadBasename
#exportCadCountInfo     #githubContents         #searchResults
#backToTopBtn           #markerPopupSave        #markerPopupCancel
#markerImageButton      #markerCurrentLocationBtn
#markerPickOnMapBtn     #exportCadBtn           #showMoreBtn
```

Có thể đổi class trang trí (`.btn-primary`, `.form-control`...) nhưng phải giữ class Bootstrap cốt lõi **hoặc** thay bằng class mới + cập nhật toàn bộ style tương ứng.

---

## 6. Output mong muốn

- Trả về **1 file HTML duy nhất**, giữ nguyên toàn bộ `<script>` hiện có.
- CSS gộp trong `<style>`, đặt CSS variables ở `:root` để dễ chỉnh màu sau.
- Comment các section CSS bằng tiếng Việt để team dễ maintain.
- Không phá vỡ các ID / class JavaScript đang gọi (xem mục 5).

---

## 7. Prompt rút gọn (cho redesign nhanh)

```text
Refresh UI cho web app bản đồ khảo sát chiếu sáng (file đính kèm).
Giữ nguyên toàn bộ JavaScript và các ID/class JS đang gọi.

Backend mới:
- Dữ liệu marker: Google Sheets (qua Apps Script Web App)
- Ảnh: GitHub (qua Apps Script proxy, URL ghi lại vào Sheet cột "anh_url")
- Excel chỉ còn dùng để export local + xuất CAD

UI:
- Mobile-first, ưu tiên dùng 1 tay ngoài trời (nắng gắt)
- Thay modal slide-left bằng bottom sheet kéo lên
- Form thêm marker thành bottom sheet 2 mức (nửa / gần full) để vẫn nhìn pin
- Top bar mỏng với search luôn visible (giống Google Maps)
- Bottom action bar: Thêm marker / Vị trí / Tracking / Lưu
- Toast notification thay div #error
- Badge "⏳ N marker chưa đồng bộ" khi mất mạng
- Palette: primary #2563eb, amber #f59e0b, nền trắng đặc (chống chói)
- System font, FontAwesome 4 (đã có), không load font ngoài
- Hỗ trợ dark mode auto + safe-area-inset cho iPhone notch
- Tham chiếu: Google Maps mobile + Mapbox Studio

Output: 1 file HTML hoàn chỉnh, CSS dùng :root variables.
```

---

## 8. Quy trình rollout đề xuất

Không redesign toàn bộ trong 1 lần. Làm theo từng bước để dễ rollback:

### Phase 1 — Migration backend (làm trước khi redesign UI)
1. **Bước 1:** Tạo Google Sheet mới với schema ở mục 3.2, import dữ liệu cũ từ `khaosat.xlsx`.
2. **Bước 2:** Viết Apps Script Web App với các endpoint ở mục 3.3, deploy lấy URL.
3. **Bước 3:** Thay các hàm `processFile`, `saveMarkerData`, `updateGitHubExcel` trong JS bằng các hàm gọi Apps Script.
4. **Bước 4:** Chuyển logic upload ảnh từ frontend → Apps Script (đỡ phải nhập GitHub token).
5. **Bước 5:** Thêm cơ chế offline + queue đồng bộ (`localStorage` + IndexedDB).

### Phase 2 — Redesign UI (sau khi backend ổn định)
6. **Bước 6:** Redesign marker form thành bottom sheet → test ngoài hiện trường 1–2 ngày.
7. **Bước 7:** Thay `#error` div bằng toast notification + badge offline.
8. **Bước 8:** Thay top bar + bottom action bar.
9. **Bước 9:** Thay panel điều khiển bên trái.
10. **Bước 10:** Floating controls + dark mode + polish.
