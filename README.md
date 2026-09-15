# Add-in Word: Chuẩn thể thức theo Nghị định 30/2020/NĐ-CP

Add-in cho Microsoft Word (desktop & Word Online) giúp:

- **Tạo văn bản mới**: chèn khung thể thức chuẩn (Quốc hiệu, Tiêu ngữ, tên cơ
  quan, số ký hiệu, địa danh/ngày tháng, tên loại + trích yếu, nơi nhận, khối
  ký tên) đã canh đúng font Times New Roman, cỡ chữ, canh lề theo NĐ 30.
- **Rà soát & sửa** văn bản đang mở: kiểm tra khổ giấy A4, lề trang, font/cỡ
  chữ toàn văn bản, sự hiện diện và định dạng của Quốc hiệu/Tiêu ngữ, mục
  Nơi nhận — có nút "Áp dụng" để tự sửa từng mục.

> Lưu ý: đây là công cụ hỗ trợ ở mức tự động hoá phần lớn quy tắc trình bày.
> Một chi tiết rất đặc thù của NĐ 30 — đường kẻ ngang dưới dòng
> "Độc lập - Tự do - Hạnh phúc" có độ dài **đúng bằng** độ dài dòng chữ đó —
> được xấp xỉ bằng cách gạch chân trực tiếp dòng chữ. Nếu đơn vị bạn yêu cầu
> đúng tuyệt đối theo hình vẽ trong phụ lục, hãy chỉnh tay thêm sau khi chèn.

## Cấu trúc thư mục

```
nd30-word-addin/
├── manifest.xml       # Khai báo add-in cho Word
├── taskpane.html       # Giao diện task pane
├── taskpane.css
├── taskpane.js         # Toàn bộ logic (Office.js)
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    ├── icon-80.png
    └── icon-128.png
```

Add-in này là **add-in tĩnh** (không cần Node/backend khi chạy thật), nhưng
Word bắt buộc các add-in phải được tải qua **HTTPS**, nên bạn cần host 5 file
trên (`taskpane.html`, `.css`, `.js`, và thư mục `assets`) ở một địa chỉ
HTTPS rồi trỏ `manifest.xml` vào đó. Có 2 cách bên dưới — chọn 1 trong 2.

---

## Cách 1 — Nhanh nhất: dùng GitHub Pages (khuyên dùng)

1. Tạo một repository GitHub mới (public), ví dụ `nd30-word-addin`.
2. Copy toàn bộ nội dung thư mục này vào repo, commit & push.
3. Vào **Settings → Pages** của repo, chọn nguồn là branch `main`, thư mục
   `/ (root)`, bấm Save. Sau ít phút bạn sẽ có địa chỉ dạng:
   `https://<tên-user>.github.io/nd30-word-addin/`
4. Mở file `manifest.xml`, thay **tất cả** chỗ `https://YOUR_DOMAIN` thành
   địa chỉ ở bước 3 (bỏ dấu `/` cuối), ví dụ:
   `https://<tên-user>.github.io/nd30-word-addin`
5. Lưu lại `manifest.xml`.

## Cách 2 — Chạy thử nhanh trên máy (không cần deploy)

Yêu cầu máy có Node.js đã cài sẵn.

```bash
# Cài chứng chỉ HTTPS phát triển dùng riêng cho Office Add-in
npx office-addin-dev-certs install

# Chạy 1 web server tĩnh HTTPS ngay tại thư mục add-in này (cổng 3000)
npx http-server . -p 3000 --ssl \
  --cert "$(npx office-addin-dev-certs verify-store | grep -oE '/.*\.crt')" \
  --key  "$(npx office-addin-dev-certs verify-store | grep -oE '/.*\.key')"
```

Nếu lệnh lấy đường dẫn cert phức tạp, cách đơn giản hơn: dùng gói
`office-addin-debugging`/`generator-office` thay vì tự chạy — nhưng cách 1
(GitHub Pages) vẫn là cách ít trục trặc nhất cho người không rành dev.

Sau khi có server chạy ở `https://localhost:3000`, sửa `manifest.xml` để
`https://YOUR_DOMAIN` → `https://localhost:3000`.

---

## Sideload add-in vào Word

### Word trên Windows/Mac (desktop)
1. Mở Word → tab **Chèn (Insert)** → **Add-ins** → **My Add-ins** →
   **Upload My Add-in** (biểu tượng mũi tên lên, góc trên).
2. Chọn file `manifest.xml` đã sửa domain ở bước trên.
3. Add-in xuất hiện trong task pane, hoặc nút "Chuẩn thể thức" trên tab
   Trang chủ (Home) tuỳ phiên bản Word.

### Word trên web (Word Online / Microsoft 365)
1. Mở tài liệu trên `office.com` hoặc SharePoint.
2. **Insert → Add-ins → Upload My Add-in** → chọn `manifest.xml`.

### Dùng chung cho cả tổ chức (khuyên dùng khi triển khai chính thức)
Nếu muốn mọi người trong cơ quan dùng mà không cần sideload thủ công từng
máy, admin Microsoft 365 có thể triển khai qua **Microsoft 365 Admin Center
→ Settings → Integrated apps → Upload custom apps**, dùng chính file
`manifest.xml` này (đã trỏ domain HTTPS chính thức, nên host cố định thay vì
GitHub Pages cá nhân trong trường hợp này).

---

## Cách dùng

1. Mở tài liệu Word (mới hoặc có sẵn) → mở task pane **Chuẩn thể thức NĐ 30**.
2. Tab **"Tạo văn bản mới"**: chọn loại văn bản, điền tên cơ quan/số hiệu/địa
   danh/trích yếu/nơi nhận/chức vụ ký → bấm **Chèn văn bản mẫu vào tài liệu**.
   Khung thể thức được chèn ở đầu tài liệu, phần nội dung để trống dạng
   placeholder để bạn viết tiếp.
3. Tab **"Rà soát & sửa"**: dùng cho văn bản đã soạn sẵn (kể cả văn bản dán
   từ nguồn khác) — bấm **Rà soát toàn bộ & hiện báo cáo** để xem các lỗi
   thể thức, sau đó bấm **Áp dụng** ở từng mục để tự động sửa (lề trang,
   font chữ, Quốc hiệu/Tiêu ngữ).

## Giới hạn hiện tại / có thể mở rộng thêm

- Chưa tự động sinh **chữ ký số / hộp thoại nhập ngày ký** hay số hiệu văn
  bản tự tăng dần theo sổ công văn (cần kết nối hệ thống quản lý văn bản nội
  bộ mới làm được phần này).
- Rà soát hiện dựa trên tìm kiếm văn bản (text match) cho Quốc hiệu/Tiêu ngữ
  và mục Nơi nhận; nếu người dùng gõ sai chính tả nhiều, add-in có thể không
  nhận diện được.
- Có thể mở rộng thêm: kiểm tra khoảng cách dòng/đoạn (line spacing, space
  before/after), kiểm tra đánh số trang, chèn header/footer, hoặc thêm các
  loại văn bản quy phạm pháp luật (có bố cục phức tạp hơn: Điều/Khoản/Điểm).

## Sửa lỗi thường gặp

- **"Không thể tải add-in" / lỗi HTTPS**: chứng chỉ tự ký (Cách 2) chưa được
  máy tin cậy — chạy lại `npx office-addin-dev-certs install` rồi khởi động
  lại Word. Với Cách 1 (GitHub Pages) không gặp lỗi này vì đã là HTTPS thật.
- **Bấm nút không có phản hồi**: mở DevTools của task pane (chuột phải trong
  task pane → *Inspect*, hoặc trên Word Online dùng F12 của trình duyệt) để
  xem lỗi console — thường do domain trong `manifest.xml` chưa khớp domain
  thật đang host file.
- **Bảng ẩn viền vẫn hiện viền mờ**: đó là đường kẻ lưới hỗ trợ hiển thị của
  Word (gridlines), chỉ hiện trên màn hình, không in ra. Có thể tắt ở
  **Layout → View Gridlines**.
