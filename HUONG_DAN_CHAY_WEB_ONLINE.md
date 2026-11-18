# 🌐 HƯỚNG DẪN SỬ DỤNG VÀ ĐƯA TOOL MRIR LÊN WEB CHẠY ONLINE

Hệ thống **Tự động Tổng hợp & Tô màu Hồ sơ MRIR** hiện đã được nâng cấp toàn diện với **Giao diện Web hiện đại (Streamlit Web App)**, cho phép người dùng thao tác trực tiếp trên trình duyệt Web hoặc chia sẻ ra Internet để chạy online từ bất kỳ thiết bị nào (máy tính khác, điện thoại, máy tính bảng).

---

## 🚀 PHẦN 1: KHỞI ĐỘNG VÀ SỬ DỤNG GIAO DIỆN WEB TRÊN MÁY TÍNH

1. Trong thư mục dự án, **nhấp đúp chuột vào file:**
   👉 **`Chay_Web_MRIR.bat`**
2. Hệ thống sẽ tự động khởi động máy chủ Web và mở trình duyệt tại địa chỉ:
   `http://localhost:8501`
3. Giao diện Web cung cấp đầy đủ các tính năng:
   * **Tab 1: 🌐 Tải File Trực Tuyến (Web Upload):**
     * Tải lên file Excel (`.xlsm` hoặc `.xlsx`) của lô hàng.
     * Tải lên các file chứng chỉ PDF (chọn nhiều file cùng lúc) hoặc tải 1 file `.zip` chứa toàn bộ chứng chỉ.
     * Nhấn nút lớn màu xanh: **🚀 BẮT ĐẦU XỬ LÝ & TẠO FILE PDF HOÀN CHỈNH**.
     * Xem nhật ký xử lý trực tiếp (Live Stream Logs).
     * Xem trước file PDF kết quả ngay trên trang Web và nhấn nút **📥 TẢI VỀ FILE PDF KẾT QUẢ**.
   * **Tab 2: 📁 Xử lý Theo Thư Mục Máy Chủ (Server Folder):**
     * Nhập đường dẫn thư mục `Input` và `Output` trên máy để xử lý tự động toàn bộ các lô hàng con.

---

## 🌍 PHẦN 2: 3 GIẢI PHÁP ĐƯA TOOL LÊN WEB ĐỂ CHẠY ONLINE

### ⭐ GIẢI PHÁP 1: Chia Sẻ Online Tức Thì Qua Cloudflare Tunnel (Khuyên dùng nhất - 0đ)

> **Tại sao nên dùng giải pháp này?**
> * **Hoàn toàn miễn phí 100%**, không tốn tiền thuê máy chủ VPS hay tên miền.
> * **Không cần mở cổng mạng (Port forwarding)**, an toàn tuyệt đối, vượt qua mọi tường lửa công ty/nhà mạng.
> * **Tận dụng tối đa phần cứng máy tính và Microsoft Excel COM** của Windows để xuất trang chuẩn 100% và chạy OCR cực nhanh.

**Cách thực hiện chỉ với 2 bước:**
1. Khởi động Web: Nhấp đúp vào **`Chay_Web_MRIR.bat`**.
2. Chia sẻ ra Internet: Nhấp đúp vào **`Chia_Se_Online_Cloudflare.bat`**.
   * Công cụ sẽ tự động tạo một đường link bảo mật HTTPS công khai, ví dụ:
     `https://mrir-pipe-portal-example.trycloudflare.com`
   * Bạn chỉ cần gửi đường link này cho đồng nghiệp hoặc mở trên điện thoại, máy tính khác là có thể tải file lên và xử lý từ bất kỳ đâu qua Internet!

---

### 🚀 GIẢI PHÁP 2: Triển Khai Lên Streamlit Community Cloud (Miễn Phí 24/7)

Nếu bạn muốn tool chạy 24/7 trên Cloud mà không cần bật máy tính cá nhân:

1. **Đưa mã nguồn lên GitHub:**
   * Tạo một tài khoản tại [github.com](https://github.com) (nếu chưa có).
   * Tạo một Repository mới (ví dụ: `mrir-web-tool`) và đẩy mã nguồn trong thư mục này lên GitHub.
2. **Kết nối và Triển khai trên Streamlit Cloud:**
   * Truy cập [share.streamlit.io](https://share.streamlit.io) và đăng nhập bằng tài khoản GitHub.
   * Nhấn nút **New app**.
   * Chọn Repository `mrir-web-tool`.
   * Tại mục **Main file path**, nhập: `app_web.py`.
   * Nhấn nút **Deploy**!
3. Sau khoảng 2-3 phút, trang web sẽ có một tên miền riêng (ví dụ: `https://mrir-tool.streamlit.app`).

> **Lưu ý quan trọng cho Cloud:** Vì máy chủ Streamlit Cloud chạy Linux không có sẵn bản quyền Microsoft Office, người dùng khi tải file lên có thể đính kèm file PDF trang bìa/Excel của lô hàng (hoặc sử dụng LibreOffice).

---

### 🐳 GIẢI PHÁP 3: Triển Khai Bằng Docker Trên Máy Chủ VPS / Mạng Nội Bộ (Intranet)

Dành cho phòng CNTT triển khai tập trung cho toàn công ty trên máy chủ Linux nội bộ:

1. Dự án đã tạo sẵn file `Dockerfile` tích hợp đầy đủ Python, OpenCV, RapidOCR và **LibreOffice headless** (để tự động chuyển đổi Excel sang PDF trên Linux).
2. Chạy 2 lệnh Docker:
   ```bash
   # Build image
   docker build -t mrir-web-app .

   # Chạy container trên cổng 8501
   docker run -d -p 8501:8501 --name mrir-service --restart always mrir-web-app
   ```
3. Truy cập ứng dụng trong mạng nội bộ hoặc qua IP Public:
   `http://<IP-SERVER>:8501`
