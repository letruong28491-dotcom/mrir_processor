import os
import sys
import shutil
import tempfile
import zipfile
import base64
import time
import streamlit as st

# Import core processor
from mrir_processor import MRIRBatchProcessor

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Hệ Thống Xử Lý & Tô Màu Hồ Sơ MRIR",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện hiện đại, chuyên nghiệp
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0f2b48 0%, #1a5276 100%);
        padding: 24px 30px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.12);
    }
    .main-header h1 {
        color: #ffffff;
        font-size: 26px;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .main-header p {
        color: #d4e6f1;
        font-size: 14px;
        margin: 6px 0 0 0;
    }
    .badge-online {
        display: inline-block;
        background-color: #27ae60;
        color: white;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        margin-left: 10px;
        vertical-align: middle;
    }
    .stat-card {
        background: #f8f9fa;
        border-left: 4px solid #1a5276;
        padding: 15px;
        border-radius: 6px;
        margin-bottom: 15px;
    }
    .log-box {
        font-family: 'Consolas', 'Courier New', monospace;
        background-color: #1e1e1e;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 8px;
        max-height: 400px;
        overflow-y: auto;
        font-size: 13px;
        line-height: 1.5;
        white-space: pre-wrap;
    }
</style>
""", unsafe_allow_html=True)

# Thư mục gốc
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT = os.path.join(BASE_DIR, "Input")
DEFAULT_OUTPUT = os.path.join(BASE_DIR, "Output")
os.makedirs(DEFAULT_INPUT, exist_ok=True)
os.makedirs(DEFAULT_OUTPUT, exist_ok=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>📑 HỆ THỐNG TỰ ĐỘNG TỔNG HỢP & TÔ MÀU HỒ SƠ MRIR <span class="badge-online">WEB ONLINE v2.0</span></h1>
    <p>Tự động hóa 100% quy trình đọc Excel RFI/MRIR, lọc trang bọc ống CWC và tô màu toàn dòng MTC/IRN theo chuẩn nghiệm thu dầu khí & công nghiệp.</p>
</div>
""", unsafe_allow_html=True)

# Sidebar thông tin hệ thống
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/document.png", width=64)
    st.title("Thông Tin Hệ Thống")
    st.info("🔹 **Chế độ:** Web & Cloud Ready\n🔹 **Hệ điều hành:** " + sys.platform.upper())
    
    # Kiểm tra các module
    has_com = False
    if sys.platform == "win32":
        try:
            import win32com.client
            has_com = True
        except Exception:
            pass
            
    st.markdown("### Trạng thái Môi trường")
    st.write(f"• **Microsoft Excel COM:** {'✅ Sẵn sàng' if has_com else '⚠️ Không có (Dùng PDF fallback)'}")
    st.write("• **Computer Vision:** ✅ OpenCV")
    st.write("• **OCR Engine:** ✅ RapidOCR ONNX")
    st.write("• **PDF Processor:** ✅ PyMuPDF (Fitz)")
    
    st.markdown("---")
    st.markdown("### Thư mục Mặc định")
    st.code(f"Input: {DEFAULT_INPUT}\nOutput: {DEFAULT_OUTPUT}")
    
    st.markdown("---")
    st.caption("© 2026 MRIR Automation System - Sẵn sàng chạy Online")

# Tạo các Tabs chức năng
tab_upload, tab_folder, tab_guide = st.tabs([
    "🌐 Tải File Trực Tuyến (Web Upload)",
    "📁 Xử Lý Theo Thư Mục Máy Chủ (Server Folder)",
    "☁️ Hướng Dẫn Đưa Lên Online (Deploy Guide)"
])

# ==========================================================
# TAB 1: WEB UPLOAD
# ==========================================================
with tab_upload:
    st.markdown("### 📤 Tải Lên Hồ Sơ Lô Hàng Để Xử Lý")
    st.write("Dành cho người dùng truy cập từ trình duyệt Web (máy tính hoặc điện thoại): Tải lên file Excel và các file chứng chỉ PDF.")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("1. File Excel Lô hàng")
        uploaded_excel = st.file_uploader(
            "Chọn file Excel của lô (.xlsm hoặc .xlsx):",
            type=["xlsm", "xlsx"],
            key="excel_uploader"
        )
        if uploaded_excel:
            st.success(f"Đã chọn: **{uploaded_excel.name}** ({uploaded_excel.size / 1024:.1f} KB)")
            
        uploaded_excel_pdf = st.file_uploader(
            "(Tùy chọn) File PDF xuất từ Excel nếu chạy trên máy chủ không có Microsoft Excel:",
            type=["pdf"],
            key="excel_pdf_uploader",
            help="Nếu server chạy Linux/Cloud không có Microsoft Excel, bạn có thể tải kèm file PDF đã xuất từ Excel tại đây."
        )

    with col2:
        st.subheader("2. File Chứng chỉ PDF (Cert)")
        upload_mode = st.radio(
            "Cách tải chứng chỉ:",
            ["Tải nhiều file PDF cùng lúc", "Tải 1 file nén .ZIP (chứa toàn bộ chứng chỉ)"],
            horizontal=True
        )
        
        uploaded_certs = []
        uploaded_zip = None
        
        if upload_mode == "Tải nhiều file PDF cùng lúc":
            uploaded_certs = st.file_uploader(
                "Chọn các file chứng chỉ PDF (CWC, IRN CWC, MTC PVP):",
                type=["pdf"],
                accept_multiple_files=True,
                key="certs_uploader"
            )
            if uploaded_certs:
                st.info(f"Đã chọn **{len(uploaded_certs)}** file PDF chứng chỉ.")
        else:
            uploaded_zip = st.file_uploader(
                "Chọn file ZIP chứa thư mục chứng chỉ:",
                type=["zip"],
                key="zip_uploader"
            )
            if uploaded_zip:
                st.info(f"Đã chọn: **{uploaded_zip.name}** ({uploaded_zip.size / (1024*1024):.2f} MB)")

    st.markdown("---")
    
    # Nút bắt đầu xử lý
    btn_start_upload = st.button("🚀 BẮT ĐẦU XỬ LÝ & TẠO FILE PDF HOÀN CHỈNH", type="primary", use_container_width=True)
    
    if btn_start_upload:
        if not uploaded_excel:
            st.error("❌ Vui lòng chọn file Excel của lô hàng trước khi xử lý!")
        elif not uploaded_certs and not uploaded_zip:
            st.error("❌ Vui lòng chọn các file chứng chỉ PDF hoặc tải lên file ZIP!")
        else:
            # Tạo thư mục làm việc tạm thời
            temp_work_dir = tempfile.mkdtemp(prefix="mrir_web_")
            cert_temp_dir = os.path.join(temp_work_dir, "Cert_Web")
            os.makedirs(cert_temp_dir, exist_ok=True)
            
            # Lưu file Excel
            saved_excel_path = os.path.join(temp_work_dir, uploaded_excel.name)
            with open(saved_excel_path, "wb") as f:
                f.write(uploaded_excel.getbuffer())
                
            # Lưu file PDF Excel (nếu có)
            if uploaded_excel_pdf:
                saved_excel_pdf = os.path.splitext(saved_excel_path)[0] + ".pdf"
                with open(saved_excel_pdf, "wb") as f:
                    f.write(uploaded_excel_pdf.getbuffer())
                    
            # Lưu các chứng chỉ
            if uploaded_zip:
                zip_path = os.path.join(temp_work_dir, "upload.zip")
                with open(zip_path, "wb") as f:
                    f.write(uploaded_zip.getbuffer())
                with zipfile.ZipFile(zip_path, 'r') as zf:
                    zf.extractall(cert_temp_dir)
            else:
                for c_file in uploaded_certs:
                    target_p = os.path.join(cert_temp_dir, c_file.name)
                    with open(target_p, "wb") as f:
                        f.write(c_file.getbuffer())

            # Giao diện Live Logs
            st.markdown("#### 📝 Nhật ký Xử lý Trực tiếp:")
            log_container = st.empty()
            progress_bar = st.progress(0.0)
            
            logs = []
            def web_logger(msg):
                logs.append(msg)
                # Hiển thị log
                log_container.markdown(f'<div class="log-box">{"<br>".join(logs[-15:])}</div>', unsafe_allow_html=True)
                
            processor = MRIRBatchProcessor(logger=web_logger)
            
            try:
                web_logger("🚀 Khởi động luồng xử lý Web...")
                progress_bar.progress(0.15)
                
                # Chạy xử lý
                final_pdf = processor.process_batch(temp_work_dir, DEFAULT_OUTPUT)
                progress_bar.progress(1.0)
                
                st.success("🎉 XỬ LÝ THÀNH CÔNG HỒ SƠ MRIR!")
                
                # Hiển thị kết quả & tải file
                if os.path.exists(final_pdf):
                    file_size_mb = os.path.getsize(final_pdf) / (1024 * 1024)
                    file_name = os.path.basename(final_pdf)
                    
                    st.markdown(f"""
                    <div class="stat-card">
                        <h4>📄 Thông tin File Kết quả</h4>
                        <b>Tên file:</b> {file_name}<br>
                        <b>Dung lượng:</b> {file_size_mb:.2f} MB<br>
                        <b>Đường dẫn lưu trên máy:</b> <code>{final_pdf}</code>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    with open(final_pdf, "rb") as f:
                        pdf_bytes = f.read()
                        
                    # Nút tải file
                    st.download_button(
                        label=f"📥 TẢI VỀ FILE PDF KẾT QUẢ ({file_name})",
                        data=pdf_bytes,
                        file_name=file_name,
                        mime="application/pdf",
                        type="primary"
                    )
                    
                    # Xem trước file PDF trực tiếp trên web
                    st.markdown("#### 👁️ Xem Trước PDF Trực Tiếp:")
                    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="750" type="application/pdf" style="border-radius: 8px; border: 1px solid #ddd;"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                    
            except Exception as e:
                progress_bar.progress(1.0)
                st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")
                web_logger(f"LỖI: {str(e)}")

# ==========================================================
# TAB 2: FOLDER BATCH (LOCAL / SERVER FOLDER)
# ==========================================================
with tab_folder:
    st.markdown("### 📁 Xử Lý Hồ Sơ Theo Thư Mục Máy Chủ")
    st.write("Sử dụng khi bạn đã sao chép các lô hàng vào thư mục `Input` trên máy hoặc ổ đĩa mạng nội bộ.")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        input_dir_val = st.text_input("Đường dẫn Thư mục Input:", value=DEFAULT_INPUT)
    with col_f2:
        output_dir_val = st.text_input("Đường dẫn Thư mục Output:", value=DEFAULT_OUTPUT)
        
    scan_all_opt = st.checkbox("Tự động quét và xử lý TẤT CẢ các thư mục lô hàng con tìm thấy trong Input", value=True)
    
    btn_start_folder = st.button("🚀 XỬ LÝ THEO THƯ MỤC TRÊN MÁY", key="btn_folder", type="primary")
    
    if btn_start_folder:
        if not os.path.exists(input_dir_val):
            st.error(f"❌ Thư mục Input không tồn tại: {input_dir_val}")
        else:
            st.markdown("#### 📝 Nhật ký Xử lý:")
            f_log_container = st.empty()
            f_progress = st.progress(0.1)
            f_logs = []
            
            def folder_logger(msg):
                f_logs.append(msg)
                f_log_container.markdown(f'<div class="log-box">{"<br>".join(f_logs[-15:])}</div>', unsafe_allow_html=True)
                
            f_processor = MRIRBatchProcessor(logger=folder_logger)
            try:
                if scan_all_opt:
                    res_files = f_processor.process_all_batches(input_dir_val, output_dir_val)
                else:
                    res = f_processor.process_batch(input_dir_val, output_dir_val)
                    res_files = [res]
                    
                f_progress.progress(1.0)
                st.success(f"🎉 Đã hoàn tất xử lý! Tạo thành công {len(res_files)} file kết quả.")
                
                for r_path in res_files:
                    if os.path.exists(r_path):
                        r_name = os.path.basename(r_path)
                        with open(r_path, "rb") as rf:
                            st.download_button(
                                label=f"📥 Tải về: {r_name} ({os.path.getsize(r_path)/(1024*1024):.2f} MB)",
                                data=rf.read(),
                                file_name=r_name,
                                mime="application/pdf",
                                key=f"dl_{r_name}"
                            )
            except Exception as e:
                f_progress.progress(1.0)
                st.error(f"❌ Lỗi: {str(e)}")
                folder_logger(f"LỖI: {str(e)}")

# ==========================================================
# TAB 3: DEPLOY GUIDE (ĐƯA LÊN ONLINE)
# ==========================================================
with tab_guide:
    st.markdown("### ☁️ HƯỚNG DẪN ĐƯA TOOL LÊN WEB ONLINE")
    st.write("Dưới đây là 3 giải pháp tốt nhất để đưa Tool MRIR này lên môi trường Web để bất kỳ ai cũng có thể truy cập qua internet:")
    
    with st.expander("⭐ CÁCH 1: Chia Sẻ Online Tức Thì Bằng Cloudflare Tunnel (Khuyên dùng - 0đ, Nhanh nhất)", expanded=True):
        st.markdown("""
        **Ưu điểm vượt trội:**
        - **Miễn phí 100%**, không cần mua tên miền, không cần cấu hình router / mở cổng mạng (Port Forwarding).
        - **Bảo toàn 100% tính năng Microsoft Excel COM** vì ứng dụng chạy trực tiếp trên máy tính Windows của bạn.
        - Tốc độ xử lý cực nhanh nhờ phần cứng của máy tính.
        - Nhận ngay một đường link bảo mật HTTPS: `https://xxx.trycloudflare.com` gửi cho đồng nghiệp dùng ngay.
        
        **Các bước thực hiện:**
        1. Chạy ứng dụng Web trên máy: Nhấp đúp vào file `Chay_Web_MRIR.bat`.
        2. Tải công cụ Cloudflare Tunnel miễn phí (chỉ cần file `cloudflared.exe` duy nhất).
        3. Nhấp đúp vào file `Chia_Se_Online_Cloudflare.bat` (đã được tạo sẵn trong thư mục).
        4. Cửa sổ dòng lệnh sẽ hiển thị đường link public dạng: `https://random-name.trycloudflare.com`.
        5. Gửi đường link này cho bất kỳ ai qua Zalo, Teams, Email... họ có thể mở trên trình duyệt điện thoại hoặc máy tính để tải file lên và xử lý online!
        """)
        
    with st.expander("🚀 CÁCH 2: Triển Khai Lên Streamlit Community Cloud (Miễn phí 24/7 trên Cloud)"):
        st.markdown("""
        **Ưu điểm:** Chạy trực tiếp 24/7 trên đám mây của Streamlit không cần bật máy tính cá nhân.
        
        **Lưu ý:** Vì môi trường Cloud chạy hệ điều hành Linux (không có Microsoft Excel bản quyền), bạn có thể:
        - Sử dụng tính năng tải kèm file PDF xuất từ Excel đã tích hợp sẵn trong giao diện Tab 1.
        
        **Các bước triển khai:**
        1. Đưa toàn bộ thư mục dự án này lên một kho lưu trữ **GitHub** (chế độ Public hoặc Private).
        2. Truy cập [share.streamlit.io](https://share.streamlit.io) và đăng nhập bằng tài khoản GitHub.
        3. Nhấn **New app** -> Chọn Repository GitHub của bạn.
        4. Ở mục **Main file path**, điền: `app_web.py`.
        5. Nhấn **Deploy**! Sau 2 phút, ứng dụng sẽ có một tên miền riêng (ví dụ: `https://mrir-tool.streamlit.app`).
        """)

    with st.expander("🐳 CÁCH 3: Triển Khai Qua Docker Trên VPS / Server Linux Nội Bộ"):
        st.markdown("""
        **Dành cho Doanh nghiệp / Quản trị mạng:**
        Dự án đã chuẩn bị sẵn file `Dockerfile` tích hợp Python 3.11, PyMuPDF, OpenCV, RapidOCR và **LibreOffice headless** (để tự động xuất Excel sang PDF trên Linux).
        
        **Lệnh build và chạy Docker:**
        ```bash
        # 1. Build image Docker
        docker build -t mrir-web-app .

        # 2. Khởi chạy container trên cổng 8501
        docker run -d -p 8501:8501 --name mrir-service mrir-web-app
        ```
        Truy cập hệ thống tại: `http://<IP-CỦA-SERVER>:8501`.
        """)
