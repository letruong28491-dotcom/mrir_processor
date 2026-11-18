import os
import sys
import shutil
import tempfile
import zipfile
import base64
import time
from datetime import datetime
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

# Custom CSS cho giao diện hiện đại, trực quan
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0f2b48 0%, #1a5276 100%);
        padding: 22px 28px;
        border-radius: 12px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.12);
    }
    .main-header h1 {
        color: #ffffff;
        font-size: 25px;
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
    .summary-card {
        background: #ffffff;
        border-left: 5px solid #27ae60;
        padding: 15px 20px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 20px;
    }
    .dest-card {
        background: #eef7fc;
        border: 1px solid #b8daff;
        border-left: 5px solid #007bff;
        padding: 15px 20px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    .file-item {
        background: #fdfefe;
        border: 1px solid #e1e8ed;
        border-radius: 8px;
        padding: 12px 18px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .log-box {
        font-family: 'Consolas', 'Courier New', monospace;
        background-color: #1e1e1e;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 8px;
        max-height: 420px;
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

# Khởi tạo session state lưu kết quả xử lý
if "last_result_pdf" not in st.session_state:
    st.session_state["last_result_pdf"] = None
if "last_logs" not in st.session_state:
    st.session_state["last_logs"] = []

# Header
st.markdown("""
<div class="main-header">
    <h1>📑 HỆ THỐNG TỰ ĐỘNG TỔNG HỢP & TÔ MÀU HỒ SƠ MRIR <span class="badge-online">ONLINE 24/7</span></h1>
    <p>Tải lên trọn gói thư mục lô hàng để tự động xuất Excel, lọc trang bọc ống CWC và tô màu toàn dòng MTC/IRN theo chuẩn nghiệm thu.</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/folder-invoices.png", width=64)
    st.title("Thông Tin Hệ Thống")
    st.info("🔹 **Chế độ:** Web Cloud Online\n🔹 **Hệ điều hành:** " + sys.platform.upper())
    
    # Kiểm tra môi trường
    has_com = False
    if sys.platform == "win32":
        try:
            import win32com.client
            has_com = True
        except Exception:
            pass
            
    st.markdown("### ⚙️ Trạng thái Hệ thống")
    st.write(f"• **Chuyển đổi Excel:** {'✅ MS Excel COM' if has_com else '✅ LibreOffice Headless'}")
    st.write("• **Computer Vision:** ✅ OpenCV")
    st.write("• **OCR Engine:** ✅ RapidOCR ONNX")
    st.write("• **PDF Engine:** ✅ PyMuPDF (Fitz)")
    
    st.markdown("---")
    st.markdown("### 📂 Vị Trí Lưu File Máy Chủ:")
    st.code(f"{DEFAULT_OUTPUT}", language="text")
    
    st.markdown("---")
    st.caption("© 2026 MRIR Automation - Ready for Cloud & Production")

# Tạo 4 Tab chức năng
tab_folder_upload, tab_output_manager, tab_manual, tab_server = st.tabs([
    "📦 1. Tải Lên Thư Mục Lô Hàng (Xử Lý)",
    "📁 2. Kho File Kết Quả Đã Xuất (Output)",
    "📑 3. Tải Lẻ Từng File",
    "🖥️ 4. Xử Lý Thư Mục Máy Chủ"
])

# ==========================================================
# TAB 1: TẢI TRỌN GÓI THƯ MỤC LÔ HÀNG
# ==========================================================
with tab_folder_upload:
    st.markdown("### 📥 Đưa Trọn Gói Thư Mục Lô Hàng Vào Hệ Thống")
    st.write("Bạn chỉ cần đưa **1 thư mục lô hàng duy nhất** (y hệt như cách dùng tool trên máy tính). Hệ thống tự động quét nhận diện file Excel và các chứng chỉ PDF.")
    
    # Hộp thông tin vị trí lưu file giải đáp thắc mắc người dùng
    st.markdown("""
    <div class="dest-card">
        <b>💡 Vị trí lưu file sau khi xuất xong:</b><br>
        1. <b>Trên máy chủ Web:</b> File PDF hoàn chỉnh được lưu tự động tại thư mục <code>Output/</code>.<br>
        2. <b>Về máy tính của bạn:</b> Sau khi xử lý xong, màn hình sẽ hiện nút <b>📥 TẢI VỀ FILE PDF KẾT QUẢ</b>. Khi bấm nút này, file sẽ tự động tải về thư mục <b>Downloads (Tải về)</b> hoặc vị trí lưu tùy chọn trên máy tính của bạn.
    </div>
    """, unsafe_allow_html=True)
    
    upload_type = st.radio(
        "Hình thức tải thư mục lô hàng:",
        [
            "📁 Cách 1: Tải file nén .ZIP của Thư mục Lô Hàng (Khuyên dùng - nhanh & gọn nhất)",
            "📂 Cách 2: Kéo thả tất cả các file trong thư mục lô hàng vào đây (Ctrl+A chọn hết)"
        ],
        horizontal=False
    )
    
    work_dir = None
    batch_name = "Lô hàng"
    found_excel = None
    found_pdfs = []
    
    if "Cách 1" in upload_type:
        st.info("💡 **Cách làm:** Nhấp chuột phải vào Thư mục Lô hàng trên máy -> chọn **Compress to ZIP file** (hoặc Send to -> Compressed folder), sau đó kéo thả file `.zip` đó vào ô bên dưới.")
        uploaded_zip = st.file_uploader(
            "Kéo thả hoặc Chọn file .ZIP của Thư mục Lô hàng:",
            type=["zip"],
            key="zip_batch_uploader"
        )
        if uploaded_zip:
            temp_dir = tempfile.mkdtemp(prefix="mrir_batch_zip_")
            zip_dest = os.path.join(temp_dir, uploaded_zip.name)
            with open(zip_dest, "wb") as f:
                f.write(uploaded_zip.getbuffer())
                
            with zipfile.ZipFile(zip_dest, 'r') as z:
                z.extractall(temp_dir)
                
            work_dir = temp_dir
            batch_name = os.path.splitext(uploaded_zip.name)[0]
    else:
        st.info("💡 **Cách làm:** Mở thư mục Lô hàng trên máy, bấm `Ctrl + A` chọn toàn bộ các file (Excel + PDF), rồi kéo thả chung vào ô bên dưới.")
        uploaded_files = st.file_uploader(
            "Kéo thả toàn bộ các file trong thư mục lô hàng vào đây:",
            type=["xlsm", "xlsx", "pdf"],
            accept_multiple_files=True,
            key="multi_files_uploader"
        )
        if uploaded_files:
            temp_dir = tempfile.mkdtemp(prefix="mrir_batch_multi_")
            cert_sub_dir = os.path.join(temp_dir, "Cert_Files")
            os.makedirs(cert_sub_dir, exist_ok=True)
            
            for uf in uploaded_files:
                if uf.name.lower().endswith(('.xlsm', '.xlsx')):
                    dest = os.path.join(temp_dir, uf.name)
                else:
                    dest = os.path.join(cert_sub_dir, uf.name)
                with open(dest, "wb") as f:
                    f.write(uf.getbuffer())
                    
            work_dir = temp_dir
            batch_name = "Lô Hàng Đã Tải Lên"

    # Kiểm tra dữ liệu trong thư mục đã nhận
    if work_dir:
        checker = MRIRBatchProcessor()
        found_excel = checker.find_excel_file(work_dir)
        found_cert_dir = checker.find_cert_dir(work_dir, found_excel)
        
        if found_cert_dir and os.path.exists(found_cert_dir):
            for root, dirs, files in os.walk(found_cert_dir):
                for f in files:
                    if f.lower().endswith(".pdf"):
                        found_pdfs.append(f)
                        
        st.markdown(f"""
        <div class="summary-card">
            <h4>📋 KẾT QUẢ QUÉT THƯ MỤC ĐẦU VÀO:</h4>
            • <b>Tên lô hàng:</b> <code>{batch_name}</code><br>
            • <b>File Excel nhận diện:</b> {f'<span style="color:green; font-weight:bold;">✅ {os.path.basename(found_excel)}</span>' if found_excel else '<span style="color:red; font-weight:bold;">❌ Không tìm thấy file Excel (.xlsm/.xlsx)</span>'}<br>
            • <b>Thư mục chứng chỉ:</b> {f'<span style="color:green; font-weight:bold;">✅ {os.path.basename(found_cert_dir)}</span>' if found_cert_dir else '<span style="color:red; font-weight:bold;">❌ Không tìm thấy</span>'}<br>
            • <b>Số lượng file PDF chứng chỉ tìm thấy:</b> <b>{len(found_pdfs)} files</b> (CWC, IRN CWC, MTC PVP)
        </div>
        """, unsafe_allow_html=True)
        
        btn_process = st.button("🚀 BẮT ĐẦU XỬ LÝ LÔ HÀNG & TẠO FILE PDF HOÀN CHỈNH", type="primary", use_container_width=True)
        
        if btn_process:
            if not found_excel:
                st.error("❌ Không thể xử lý vì thư mục tải lên chưa có file Excel (.xlsm hoặc .xlsx)!")
            elif len(found_pdfs) == 0:
                st.error("❌ Không thể xử lý vì không tìm thấy file chứng chỉ PDF nào trong thư mục!")
            else:
                st.markdown("#### 📝 Nhật ký Xử lý Trực tiếp:")
                log_box = st.empty()
                p_bar = st.progress(0.05)
                runtime_logs = []
                
                def live_logger(msg):
                    runtime_logs.append(msg)
                    log_box.markdown(f'<div class="log-box">{"<br>".join(runtime_logs[-16:])}</div>', unsafe_allow_html=True)
                    
                processor = MRIRBatchProcessor(logger=live_logger)
                
                try:
                    live_logger(f"🚀 Bắt đầu xử lý trọn gói lô hàng: {batch_name}...")
                    p_bar.progress(0.15)
                    
                    # Tiến hành xử lý và lưu vào Output
                    res_pdf = processor.process_batch(work_dir, DEFAULT_OUTPUT)
                    p_bar.progress(1.0)
                    
                    # Lưu vào session state
                    st.session_state["last_result_pdf"] = res_pdf
                    st.session_state["last_logs"] = runtime_logs
                    
                    st.success("🎉 XỬ LÝ THÀNH CÔNG! ĐÃ TẠO FILE PDF TỔNG HỢP HOÀN TẤT.")
                    
                except Exception as ex:
                    p_bar.progress(1.0)
                    st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {str(ex)}")
                    live_logger(f"LỖI: {str(ex)}")

    # Hiển thị kết quả xử lý nếu đã có
    if st.session_state["last_result_pdf"] and os.path.exists(st.session_state["last_result_pdf"]):
        res_pdf = st.session_state["last_result_pdf"]
        pdf_size_mb = os.path.getsize(res_pdf) / (1024 * 1024)
        pdf_basename = os.path.basename(res_pdf)
        
        st.markdown(f"""
        <div class="summary-card" style="border-left-color: #2980b9;">
            <h4>📄 FILE PDF HOÀN CHỈNH ĐÃ SẴN SÀNG</h4>
            • <b>Tên file kết quả:</b> <code>{pdf_basename}</code><br>
            • <b>Dung lượng:</b> <b>{pdf_size_mb:.2f} MB</b><br>
            • <b>Vị trí lưu trên máy chủ:</b> <code>{res_pdf}</code><br>
            • <b>Trạng thái:</b> Đã ghép Excel, lọc trang CWC, đính kèm phiếu xuất ống IRN CWC và tô màu toàn dòng MTC/IRN PVP.
        </div>
        """, unsafe_allow_html=True)
        
        with open(res_pdf, "rb") as rf:
            pdf_bytes = rf.read()
            
        # Nút tải file về máy tính cá nhân
        st.download_button(
            label=f"📥 TẢI VỀ FILE PDF KẾT QUẢ VỀ MÁY TÍNH CỦA BẠN ({pdf_basename})",
            data=pdf_bytes,
            file_name=pdf_basename,
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )
        st.caption("👉 Sau khi bấm nút tải về, file sẽ nằm trong thư mục **Downloads (Tải về)** của máy tính bạn.")
        
        # Xem trước PDF
        st.markdown("#### 👁️ Xem Trước Bản In PDF Trực Tiếp:")
        b64_str = base64.b64encode(pdf_bytes).decode('utf-8')
        st.markdown(f'<iframe src="data:application/pdf;base64,{b64_str}" width="100%" height="800" type="application/pdf" style="border-radius: 8px; border: 1px solid #ccc;"></iframe>', unsafe_allow_html=True)

# ==========================================================
# TAB 2: KHO FILE KẾT QUẢ ĐÃ XUẤT (OUTPUT MANAGER)
# ==========================================================
with tab_output_manager:
    st.markdown("### 📁 Kho Lưu Trữ & Quản Lý File Kết Quả Đã Xuất (Output)")
    st.write(f"Tất cả các file PDF sau khi xuất hoàn thành đều được lưu trữ tập trung tại thư mục **`Output`** trên hệ thống.")
    
    # Lấy danh sách file trong thư mục Output
    out_files = []
    if os.path.exists(DEFAULT_OUTPUT):
        for f in os.listdir(DEFAULT_OUTPUT):
            if f.lower().endswith(".pdf"):
                full_p = os.path.join(DEFAULT_OUTPUT, f)
                f_stat = os.stat(full_p)
                out_files.append({
                    "name": f,
                    "path": full_p,
                    "size_mb": f_stat.st_size / (1024 * 1024),
                    "mtime": datetime.fromtimestamp(f_stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S"),
                    "raw_time": f_stat.st_mtime
                })
        out_files.sort(key=lambda x: x["raw_time"], reverse=True)
        
    if not out_files:
        st.info("ℹ️ Hiện chưa có file PDF nào trong thư mục Output. Hãy tải lên và xử lý lô hàng ở Tab 1 để tạo file kết quả.")
    else:
        st.success(f"🔍 Tìm thấy **{len(out_files)} file PDF** kết quả trong kho lưu trữ Output:")
        
        # Nút tải toàn bộ dưới dạng file ZIP
        col_dl_all, col_refresh = st.columns([2, 1])
        with col_dl_all:
            # Tạo file zip tạm chứa toàn bộ các file output
            zip_all_path = os.path.join(tempfile.gettempdir(), "Tat_Ca_File_MRIR_Highlighted.zip")
            with zipfile.ZipFile(zip_all_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for of in out_files:
                    zip_out.write(of["path"], of["name"])
            with open(zip_all_path, "rb") as zf:
                st.download_button(
                    label="📦 TẢI VỀ TOÀN BỘ CÁC FILE TRONG KHO (.ZIP)",
                    data=zf.read(),
                    file_name="Tat_Ca_File_MRIR_Highlighted.zip",
                    mime="application/zip"
                )
        with col_refresh:
            if st.button("🔄 Làm mới danh sách"):
                st.rerun()
                
        st.markdown("---")
        
        # Danh sách từng file kết quả
        for idx, of in enumerate(out_files, 1):
            col_info, col_btn = st.columns([3, 1])
            with col_info:
                st.markdown(f"**{idx}. 📄 `{of['name']}`**")
                st.caption(f"Dung lượng: **{of['size_mb']:.2f} MB** | Thời gian xuất: **{of['mtime']}**")
            with col_btn:
                with open(of["path"], "rb") as cur_f:
                    st.download_button(
                        label="📥 Tải về máy",
                        data=cur_f.read(),
                        file_name=of["name"],
                        mime="application/pdf",
                        key=f"dl_tab2_{idx}"
                    )
            st.divider()

# ==========================================================
# TAB 3: TẢI LẺ TỪNG FILE
# ==========================================================
with tab_manual:
    st.markdown("### 📑 Tải Lên Riêng Lẻ Từng Phần")
    st.write("Sử dụng khi bạn muốn tải riêng rẽ file Excel và chọn các file PDF thành phần.")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        manual_excel = st.file_uploader("1. File Excel (.xlsm/.xlsx):", type=["xlsm", "xlsx"], key="m_excel")
        manual_excel_pdf = st.file_uploader("(Tùy chọn) File PDF xuất từ Excel:", type=["pdf"], key="m_excel_pdf")
    with col_m2:
        manual_certs = st.file_uploader("2. Các file PDF chứng chỉ:", type=["pdf"], accept_multiple_files=True, key="m_certs")
        
    btn_manual_run = st.button("🚀 Xử Lý Dữ Liệu Tải Lẻ", type="primary", key="btn_manual")
    if btn_manual_run:
        if not manual_excel or not manual_certs:
            st.error("❌ Vui lòng chọn đầy đủ cả file Excel và ít nhất một file PDF chứng chỉ!")
        else:
            m_temp = tempfile.mkdtemp(prefix="mrir_manual_")
            m_cert_dir = os.path.join(m_temp, "Cert_Manual")
            os.makedirs(m_cert_dir, exist_ok=True)
            
            saved_xl = os.path.join(m_temp, manual_excel.name)
            with open(saved_xl, "wb") as f:
                f.write(manual_excel.getbuffer())
                
            if manual_excel_pdf:
                saved_xl_pdf = os.path.splitext(saved_xl)[0] + ".pdf"
                with open(saved_xl_pdf, "wb") as f:
                    f.write(manual_excel_pdf.getbuffer())
                    
            for cf in manual_certs:
                with open(os.path.join(m_cert_dir, cf.name), "wb") as f:
                    f.write(cf.getbuffer())
                    
            st.markdown("#### 📝 Tiến trình Xử lý:")
            m_log_box = st.empty()
            m_logs = []
            def m_logger(msg):
                m_logs.append(msg)
                m_log_box.markdown(f'<div class="log-box">{"<br>".join(m_logs[-15:])}</div>', unsafe_allow_html=True)
                
            m_processor = MRIRBatchProcessor(logger=m_logger)
            try:
                m_res = m_processor.process_batch(m_temp, DEFAULT_OUTPUT)
                st.success("🎉 Xử lý thành công!")
                with open(m_res, "rb") as mf:
                    st.download_button(
                        label=f"📥 Tải về: {os.path.basename(m_res)}",
                        data=mf.read(),
                        file_name=os.path.basename(m_res),
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Lỗi: {e}")

# ==========================================================
# TAB 4: XỬ LÝ THEO THƯ MỤC TRÊN MÁY CHỦ (SERVER FOLDER)
# ==========================================================
with tab_server:
    st.markdown("### 🖥️ Xử Lý Theo Thư Mục Cục Bộ Trên Máy / Server")
    st.write("Dành cho trường hợp chạy trên máy tính nội bộ hoặc máy chủ có sẵn thư mục `Input`.")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        s_input = st.text_input("Đường dẫn Thư mục Input:", value=DEFAULT_INPUT)
    with col_s2:
        s_output = st.text_input("Đường dẫn Thư mục Output:", value=DEFAULT_OUTPUT)
        
    s_scan_all = st.checkbox("Tự động quét và chạy TẤT CẢ các thư mục lô hàng con", value=True)
    btn_server_run = st.button("🚀 Chạy Xử Lý Thư Mục Máy Chủ", key="btn_server", type="primary")
    
    if btn_server_run:
        if not os.path.exists(s_input):
            st.error(f"Thư mục không tồn tại: {s_input}")
        else:
            s_log_box = st.empty()
            s_logs = []
            def s_logger(msg):
                s_logs.append(msg)
                s_log_box.markdown(f'<div class="log-box">{"<br>".join(s_logs[-15:])}</div>', unsafe_allow_html=True)
            s_proc = MRIRBatchProcessor(logger=s_logger)
            try:
                if s_scan_all:
                    s_results = s_proc.process_all_batches(s_input, s_output)
                else:
                    s_results = [s_proc.process_batch(s_input, s_output)]
                st.success(f"🎉 Hoàn thành xử lý {len(s_results)} lô hàng!")
            except Exception as e:
                st.error(f"Lỗi: {e}")
