import os
import sys
import re
import cv2
import numpy as np
import openpyxl
import pymupdf as fitz

# Ensure console supports utf-8
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

def default_logger(msg):
    try:
        print(msg, flush=True)
    except Exception:
        pass

class MRIRBatchProcessor:
    def __init__(self, logger=None):
        self.log = logger or default_logger
        self.yellow_color = (1.0, 1.0, 0.0)

    def find_excel_file(self, batch_dir):
        """Tìm file Excel (.xlsm hoặc .xlsx) trong thư mục lô hàng."""
        if os.path.isfile(batch_dir) and batch_dir.lower().endswith(('.xlsm', '.xlsx')):
            return batch_dir
            
        candidates = []
        for root, dirs, files in os.walk(batch_dir):
            for f in files:
                if f.startswith("~$"):
                    continue
                if f.lower().endswith(('.xlsm', '.xlsx')):
                    full_p = os.path.join(root, f)
                    score = 0
                    if "MRIR" in f.upper():
                        score += 10
                    if "RFI" in f.upper():
                        score += 5
                    candidates.append((score, full_p))
        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1]
        return None

    def find_cert_dir(self, batch_dir, excel_path=None):
        """Tìm thư mục chứa chứng chỉ PDF (Cert_...)."""
        base = os.path.dirname(excel_path) if excel_path and os.path.isfile(excel_path) else batch_dir
        
        # 1. Tìm thư mục con có chữ 'Cert'
        for entry in os.listdir(base):
            p = os.path.join(base, entry)
            if os.path.isdir(p) and "CERT" in entry.upper():
                return p
                
        # 2. Kiểm tra nếu ngay trong base có các file .pdf
        pdf_count = len([f for f in os.listdir(base) if f.lower().endswith(".pdf")])
        if pdf_count > 3:
            return base
            
        # 3. Quét đệ quy tìm thư mục chứa nhiều file pdf nhất
        best_dir = None
        max_pdfs = 0
        for root, dirs, files in os.walk(base):
            pdfs = [f for f in files if f.lower().endswith('.pdf')]
            if len(pdfs) > max_pdfs:
                max_pdfs = len(pdfs)
                best_dir = root
        return best_dir

    def export_excel_sheets(self, excel_path, output_pdf_path):
        """Xuất các sheet RFI, MRIR, Sheet1 từ Excel sang PDF chuẩn (Hỗ trợ đa nền tảng Windows/Linux/Cloud)."""
        # 1. Kiểm tra nếu file đích đã tồn tại hợp lệ
        if os.path.exists(output_pdf_path) and os.path.getsize(output_pdf_path) > 0:
            self.log(f" -> Dùng file PDF Excel đã có sẵn: {os.path.basename(output_pdf_path)}")
            return output_pdf_path

        # Kiểm tra nếu trong cùng thư mục Excel có sẵn file .pdf cùng tên
        same_name_pdf = os.path.splitext(excel_path)[0] + ".pdf"
        if os.path.exists(same_name_pdf) and os.path.getsize(same_name_pdf) > 0:
            import shutil
            shutil.copy2(same_name_pdf, output_pdf_path)
            self.log(f" -> Đã lấy file PDF Excel có sẵn từ thư mục lô: {os.path.basename(same_name_pdf)}")
            return output_pdf_path

        # 2. Thử gọi Microsoft Excel COM (nếu đang chạy trên Windows)
        if sys.platform == "win32":
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                self.log(f" -> Đang gọi Microsoft Excel COM để xuất các sheet RFI, MRIR, Sheet1...")
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    wb_com = excel.Workbooks.Open(os.path.abspath(excel_path))
                    sheet_names = [s.Name for s in wb_com.Sheets]
                    target_sheets = []
                    
                    rfi_sheet = next((s for s in sheet_names if "RFI" in s.upper()), None)
                    if rfi_sheet: target_sheets.append(rfi_sheet)
                    
                    mrir_sheet = next((s for s in sheet_names if "MRIR" in s.upper()), None)
                    if mrir_sheet: target_sheets.append(mrir_sheet)
                    
                    ref_sheet = next((s for s in sheet_names if "SHEET1" in s.upper() or "TRA" in s.upper()), None)
                    if ref_sheet: target_sheets.append(ref_sheet)
                    
                    if target_sheets:
                        self.log(f"    Chọn các sheet: {target_sheets}")
                        wb_com.Sheets(target_sheets).Select()
                        excel.ActiveSheet.ExportAsFixedFormat(0, os.path.abspath(output_pdf_path))
                        wb_com.Close(False)
                        self.log(f" -> Xuất Excel sang PDF thành công qua Excel COM: {output_pdf_path}")
                        return output_pdf_path
                finally:
                    try: excel.Quit()
                    except Exception: pass
                    try: pythoncom.CoUninitialize()
                    except Exception: pass
            except Exception as e_com:
                self.log(f" ⚠️ Không thể xuất qua Excel COM ({e_com}). Đang chuyển sang phương thức dự phòng...")

        # 3. Thử gọi LibreOffice (Headless soffice - hỗ trợ Linux, Docker, Windows)
        import subprocess
        import shutil
        soffice_cmd = None
        for candidate in ["soffice", "libreoffice", r"C:\Program Files\LibreOffice\program\soffice.exe"]:
            if shutil.which(candidate) or os.path.exists(candidate):
                soffice_cmd = candidate
                break
        if soffice_cmd:
            try:
                self.log(f" -> Đang chuyển đổi Excel sang PDF bằng LibreOffice headless...")
                out_dir = os.path.dirname(output_pdf_path)
                subprocess.run([
                    soffice_cmd, "--headless", "--convert-to", "pdf",
                    excel_path, "--outdir", out_dir
                ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                gen_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(excel_path))[0] + ".pdf")
                if os.path.exists(gen_pdf) and gen_pdf != output_pdf_path:
                    os.replace(gen_pdf, output_pdf_path)
                if os.path.exists(output_pdf_path):
                    self.log(f" -> Xuất Excel sang PDF thành công qua LibreOffice: {output_pdf_path}")
                    return output_pdf_path
            except Exception as e_soffice:
                self.log(f" ⚠️ LibreOffice convert thất bại: {e_soffice}")

        # 4. Nếu không có cả COM lẫn LibreOffice
        raise RuntimeError(
            "Không thể xuất file Excel sang PDF tự động vì môi trường hiện tại không có Microsoft Excel COM hoặc LibreOffice!\n"
            "Giải pháp:\n"
            "1. Nếu chạy trên Windows: Vui lòng cài Microsoft Excel.\n"
            "2. Nếu chạy trên Linux/Docker: Cài đặt LibreOffice (`apt-get install -y libreoffice`).\n"
            "3. Hoặc tải lên kèm sẵn file PDF trang bìa/Excel của lô hàng."
        )

    def parse_pipe_data(self, excel_path):
        """Đọc danh sách ống, IRN, STT CWC từ Sheet1 và Serial No từ sheet MRIR."""
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        
        # Tìm sheet Bảng tra (Sheet1)
        ref_sheet_name = next((s for s in wb.sheetnames if "SHEET1" in s.upper() or "TRA" in s.upper()), None)
        if not ref_sheet_name:
            ref_sheet_name = wb.sheetnames[0]
        ws_ref = wb[ref_sheet_name]
        
        pipes = []
        for r in range(3, ws_ref.max_row + 1):
            tag_val = ws_ref.cell(r, 2).value
            if not tag_val or str(tag_val).strip() == "" or str(tag_val).startswith("#"):
                continue
            tag = str(tag_val).strip()
            irn = str(ws_ref.cell(r, 5).value or "").strip()
            pvc_cert = str(ws_ref.cell(r, 6).value or "").strip()
            cwc_irn = str(ws_ref.cell(r, 7).value or "").strip()
            cwc_stt_val = ws_ref.cell(r, 8).value
            try:
                cwc_stt = int(cwc_stt_val)
            except (ValueError, TypeError):
                cwc_stt = 0
                
            pipes.append({
                "tag": tag,
                "irn": irn,
                "pvc_cert": pvc_cert,
                "cwc_irn": cwc_irn,
                "cwc_stt": cwc_stt
            })
            
        # Tìm sheet MRIR để lấy số Seri
        mrir_sheet_name = next((s for s in wb.sheetnames if "MRIR" in s.upper()), None)
        mrir_code = ""
        if mrir_sheet_name:
            ws_mrir = wb[mrir_sheet_name]
            # Quét tìm mã MRIR
            for r in range(1, 10):
                for c in range(1, 25):
                    v = str(ws_mrir.cell(r, c).value or "")
                    if "MRIR-" in v:
                        mrir_code = v.strip()
                        break
                if mrir_code:
                    break
                    
            # Đọc mapping Serial No và Tag Name
            for r in range(8, ws_mrir.max_row + 1):
                serial_val = ws_mrir.cell(r, 2).value
                tag_val = ws_mrir.cell(r, 18).value
                if not serial_val and not tag_val:
                    continue
                serial = str(serial_val or "").strip()
                tag = str(tag_val or "").strip()
                for p in pipes:
                    if p["tag"] == tag or (tag and tag in p["tag"]):
                        p["serial"] = serial
                        
        if not mrir_code:
            # Fallback lấy từ tên file Excel
            base_name = os.path.splitext(os.path.basename(excel_path))[0]
            m_match = re.search(r"MRIR-[\w-]+", base_name)
            mrir_code = m_match.group(0) if m_match else base_name
            
        return pipes, mrir_code

    def extract_table_lines(self, page):
        """Trích xuất tọa độ các đường kẻ ngang của bảng bằng OpenCV."""
        pix = page.get_pixmap(dpi=150)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
        else:
            img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        bw = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, -2)
        h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (80, 1))
        h_lines = cv2.morphologyEx(bw, cv2.MORPH_OPEN, h_kernel)
        row_sums = np.sum(h_lines[:, 200:1000], axis=1)
        line_ys = np.where(row_sums > 300 * 255)[0]
        
        grouped_ys = []
        cur = []
        for y in line_ys:
            if not cur or y == cur[-1] + 1:
                cur.append(y)
            else:
                grouped_ys.append(int(np.mean(cur)))
                cur = [y]
        if cur:
            grouped_ys.append(int(np.mean(cur)))
            
        raw_lines = [y for y in grouped_ys if 180 <= y <= 1670]
        
        # Đảm bảo danh sách luôn đủ 41 đường kẻ cho 40 dòng
        expected_rows = 40
        if len(raw_lines) == expected_rows + 1:
            return raw_lines
            
        deltas = [raw_lines[i+1] - raw_lines[i] for i in range(len(raw_lines)-1) if 30 <= (raw_lines[i+1] - raw_lines[i]) <= 45]
        row_h = (sum(deltas) / len(deltas)) if deltas else 36.25
        first_y = raw_lines[0] if raw_lines else 205
        return [int(first_y + i * row_h) for i in range(expected_rows + 1)]

    def get_cwc_row_quad(self, page, y0_px, y1_px):
        """Tính toán Quad highlight cho dòng bảng bọc ống CWC theo góc xoay trang.
        Bổ sung khoảng đệm (inset) 1.5 pt để đảm bảo highlight không bao giờ lấn qua đường kẻ viền
        và không bao giờ chạm/dính sang dòng ống liền kề.
        """
        scale = 72 / 150
        inset_pt = 1.5
        y0_pt = y0_px * scale + inset_pt
        y1_pt = y1_px * scale - inset_pt
        if y1_pt <= y0_pt:
            y0_pt = y0_px * scale
            y1_pt = y1_px * scale
            
        y_start = 18.0
        y_end = 577.0
        
        if page.rotation == 270:
            x_start = 841.92 - y1_pt
            x_end = 841.92 - y0_pt
        elif page.rotation == 90:
            x_start = y0_pt
            x_end = y1_pt
        else:
            return fitz.Quad(
                fitz.Point(y_start, y0_pt),
                fitz.Point(y_end, y0_pt),
                fitz.Point(y_start, y1_pt),
                fitz.Point(y_end, y1_pt)
            )
            
        return fitz.Quad(
            fitz.Point(x_end, y_start),
            fitz.Point(x_end, y_end),
            fitz.Point(x_start, y_start),
            fitz.Point(x_start, y_end)
        )

    def find_cwc_cover_rect(self, page):
        """Định vị ô Certificate No./ Số chứng chỉ trên trang bìa CWC (Trang 1 và Trang 2)."""
        clip_rect = fitz.Rect(430, 20, 565, 55)
        found_rect = None
        # 1. Text blocks
        for b in page.get_text('blocks'):
            if 20 <= b[1] <= 55 and 430 <= b[0] <= 565:
                clean = b[4].replace(' ', '').upper()
                if ('277' in clean and 'CWC' in clean) or ('CERT3.2' in clean) or ('3LPE+CWC' in clean):
                    found_rect = fitz.Rect(b[:4])
                    break
        # 2. Text words
        if not found_rect:
            for w in page.get_text('words', clip=clip_rect):
                clean = w[4].replace(' ', '').upper()
                if ('277' in clean and 'CWC' in clean) or ('CERT3.2' in clean) or ('3LPE+CWC' in clean):
                    found_rect = fitz.Rect(w[:4])
                    break
        # 3. Cropped OCR
        if not found_rect:
            try:
                from rapidocr_onnxruntime import RapidOCR
                if not hasattr(self, '_ocr_engine'):
                    self._ocr_engine = RapidOCR()
                scale = 72.0 / 150.0
                pix = page.get_pixmap(dpi=150, clip=clip_rect)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))
                res, _ = self._ocr_engine(img)
                if res:
                    for box, txt, sc in res:
                        clean = txt.replace(' ', '').upper()
                        if ('277' in clean and 'CWC' in clean) or ('CERT3.2' in clean) or ('3LPE+CWC' in clean):
                            box = np.array(box)
                            x0 = clip_rect.x0 + float(np.min(box[:, 0])) * scale
                            x1 = clip_rect.x0 + float(np.max(box[:, 0])) * scale
                            y0 = clip_rect.y0 + float(np.min(box[:, 1])) * scale
                            y1 = clip_rect.y0 + float(np.max(box[:, 1])) * scale
                            found_rect = fitz.Rect(x0, y0, x1, y1)
                            break
            except Exception:
                pass
        # 4. Fallback standard rect
        if not found_rect:
            found_rect = fitz.Rect(440.0, 32.0, 554.0, 43.5)
        return found_rect

    def get_cwc_cover_quad(self, page, rect):
        """Chuyển đổi hình chữ nhật tọa độ hiển thị sang Quad theo góc xoay của trang."""
        if page.rotation == 270:
            ul = fitz.Point(841.92 - rect.y0, rect.x0)
            ur = fitz.Point(841.92 - rect.y0, rect.x1)
            ll = fitz.Point(841.92 - rect.y1, rect.x0)
            lr = fitz.Point(841.92 - rect.y1, rect.x1)
            return fitz.Quad(ul, ur, ll, lr)
        elif page.rotation == 90:
            ul = fitz.Point(rect.y0, 595.32 - rect.x0)
            ur = fitz.Point(rect.y0, 595.32 - rect.x1)
            ll = fitz.Point(rect.y1, 595.32 - rect.x0)
            lr = fitz.Point(rect.y1, 595.32 - rect.x1)
            return fitz.Quad(ul, ur, ll, lr)
        else:
            return fitz.Quad(
                fitz.Point(rect.x0, rect.y0),
                fitz.Point(rect.x1, rect.y0),
                fitz.Point(rect.x0, rect.y1),
                fitz.Point(rect.x1, rect.y1)
            )

    def find_irn_header_rect(self, page):
        """Định vị ô Released Note No./ Số BB trên trang IRN đứng."""
        clip_rect = fitz.Rect(480, 35, 575, 75)
        found_rect = None
        # 1. Check blocks
        for b in page.get_text('blocks'):
            if 35 <= b[1] <= 70 and 480 <= b[0] <= 575:
                if 'IRN' in b[4].upper():
                    y_mid = (b[1] + b[3]) / 2
                    found_rect = fitz.Rect(b[0] - 1.0, y_mid - 4.5, b[2] + 1.0, y_mid + 4.5)
                    break
        # 2. Check words
        if not found_rect:
            for w in page.get_text('words', clip=clip_rect):
                if 'IRN' in w[4].upper():
                    y_mid = (w[1] + w[3]) / 2
                    found_rect = fitz.Rect(w[0] - 1.0, y_mid - 4.5, w[2] + 1.0, y_mid + 4.5)
                    break
        # 3. Cropped OCR
        if not found_rect:
            try:
                from rapidocr_onnxruntime import RapidOCR
                if not hasattr(self, '_ocr_engine'):
                    self._ocr_engine = RapidOCR()
                scale = 72.0 / 150.0
                pix = page.get_pixmap(dpi=150, clip=clip_rect)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))
                res, _ = self._ocr_engine(img)
                if res:
                    for box, txt, sc in res:
                        if 'IRN' in txt.replace(' ', '').upper():
                            box = np.array(box)
                            x0 = clip_rect.x0 + float(np.min(box[:, 0])) * scale
                            x1 = clip_rect.x0 + float(np.max(box[:, 0])) * scale
                            y0 = clip_rect.y0 + float(np.min(box[:, 1])) * scale
                            y1 = clip_rect.y0 + float(np.max(box[:, 1])) * scale
                            y_mid = (y0 + y1) / 2
                            found_rect = fitz.Rect(x0 - 1.0, y_mid - 4.5, x1 + 1.0, y_mid + 4.5)
                            break
            except Exception:
                pass
        # 4. Fallback standard rect
        if not found_rect:
            found_rect = fitz.Rect(503.0, 49.0, 558.0, 59.0)
        return found_rect

    def find_mtc_header_rect(self, page):
        """Định vị ô CERTIFICATE NUMBER trên trang MTC ngang."""
        clip_rect = fitz.Rect(620, 80, 750, 115)
        found_rect = None
        # 1. Check words
        for w in page.get_text('words', clip=clip_rect):
            if 'MTC' in w[4].upper():
                m = re.search(r'[\w]+-MTC-[\w]+', w[4])
                if m:
                    code = m.group(0).split('_')[0]
                    char_w = (w[2] - w[0]) / max(len(w[4]), 1)
                    matched_w = len(code) * char_w
                    found_rect = fitz.Rect(w[0] - 1.0, w[1] - 1.0, min(w[2], w[0] + matched_w + 2.0), w[3] + 3.0)
                    break
        # 2. Check blocks
        if not found_rect:
            for b in page.get_text('blocks'):
                if 80 <= b[1] <= 125 and 620 <= b[0] <= 750:
                    first_line = b[4].split('\n')[0].strip()
                    if 'MTC' in first_line.upper():
                        m = re.search(r'[\w]+-MTC-[\w]+', first_line)
                        code = m.group(0).split('_')[0] if m else 'MTC'
                        char_w = (b[2] - b[0]) / max(len(first_line), 1)
                        matched_w = len(code) * char_w
                        found_rect = fitz.Rect(b[0] - 1.0, b[1] - 1.0, min(b[2], b[0] + matched_w + 2.0), b[1] + 11.0)
                        break
        # 3. Cropped OCR
        if not found_rect:
            try:
                from rapidocr_onnxruntime import RapidOCR
                if not hasattr(self, '_ocr_engine'):
                    self._ocr_engine = RapidOCR()
                scale = 72.0 / 150.0
                pix = page.get_pixmap(dpi=150, clip=clip_rect)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))
                res, _ = self._ocr_engine(img)
                if res:
                    for box, txt, sc in res:
                        if 'MTC' in txt.replace(' ', '').upper():
                            box = np.array(box)
                            x0 = clip_rect.x0 + float(np.min(box[:, 0])) * scale
                            x1 = clip_rect.x0 + float(np.max(box[:, 0])) * scale
                            y0 = clip_rect.y0 + float(np.min(box[:, 1])) * scale
                            y1 = clip_rect.y0 + float(np.max(box[:, 1])) * scale
                            found_rect = fitz.Rect(x0, y0, x1, y1)
                            break
            except Exception:
                pass
        # 4. Fallback standard rect
        if not found_rect:
            found_rect = fitz.Rect(639.0, 91.0, 695.0, 101.0)
        return found_rect

    def process_batch(self, batch_dir, output_dir, temp_dir=None):
        """Thực hiện xử lý toàn bộ một lô hồ sơ và tạo file PDF gộp hoàn chỉnh."""
        self.log("==================================================")
        self.log(f"BẮT ĐẦU XỬ LÝ LÔ HỒ SƠ: {batch_dir}")
        self.log("==================================================")
        
        # 1. Tìm file Excel và thư mục Chứng chỉ
        excel_path = self.find_excel_file(batch_dir)
        if not excel_path:
            raise FileNotFoundError(f"Không tìm thấy file Excel nào trong {batch_dir}")
        self.log(f" -> Đã tìm thấy file Excel: {os.path.basename(excel_path)}")
        
        cert_dir = self.find_cert_dir(batch_dir, excel_path)
        if not cert_dir or not os.path.exists(cert_dir):
            raise FileNotFoundError(f"Không tìm thấy thư mục chứng chỉ PDF trong {batch_dir}")
        self.log(f" -> Đã tìm thấy thư mục chứng chỉ: {os.path.basename(cert_dir)}")
        
        # Chuẩn bị thư mục tạm và output
        if not temp_dir:
            temp_dir = os.path.join(batch_dir, "temp_proc")
        os.makedirs(temp_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        
        # 2. Đọc thông tin các ống từ Excel
        pipes, mrir_code = self.parse_pipe_data(excel_path)
        self.log(f" -> Đã đọc thông tin {len(pipes)} ống. Mã hồ sơ: {mrir_code}")
        
        # 3. Xuất Excel sang PDF (hoặc dùng lại nếu đã có)
        excel_pdf_path = os.path.join(temp_dir, f"{os.path.splitext(os.path.basename(excel_path))[0]}.pdf")
        if not os.path.exists(excel_pdf_path):
            self.export_excel_sheets(excel_path, excel_pdf_path)
        else:
            self.log(f" -> Dùng file PDF Excel có sẵn tại: {excel_pdf_path}")
            
        # 4. Tìm các file PDF chứng chỉ thành phần
        # 4.1 Danh sách File Phiếu kiểm tra xuất ống (IRN CWC)
        cwc_irn_set = set(p["cwc_irn"] for p in pipes if p.get("cwc_irn"))
        cwc_irn_files = []
        for f in sorted(os.listdir(cert_dir)):
            if f.lower().endswith(".pdf"):
                matched = False
                for irn in cwc_irn_set:
                    if irn and irn.lower() in f.lower():
                        cwc_irn_files.append(os.path.join(cert_dir, f))
                        matched = True
                        break
                if not matched and "IRN-CWC" in f.upper() and not cwc_irn_set:
                    cwc_irn_files.append(os.path.join(cert_dir, f))
        # Loại trùng lặp và giữ thứ tự sắp xếp
        seen_irn_cwc = set()
        unique_cwc_irn_files = []
        for p_f in cwc_irn_files:
            if p_f not in seen_irn_cwc:
                seen_irn_cwc.add(p_f)
                unique_cwc_irn_files.append(p_f)
                
        if unique_cwc_irn_files:
            self.log(f" -> Tìm thấy {len(unique_cwc_irn_files)} Phiếu kiểm tra xuất ống CWC:")
            for p_f in unique_cwc_irn_files:
                self.log(f"    + {os.path.basename(p_f)}")
        else:
            self.log(" [!] Cảnh báo: Không tìm thấy Phiếu kiểm tra xuất ống CWC.")
            
        # 4.2 File Chứng chỉ bọc ống CWC
        cwc_cert_name = pipes[0].get("pvc_cert", "") if pipes else ""
        cwc_pdf_file = None
        for f in os.listdir(cert_dir):
            if f.lower().endswith(".pdf"):
                if cwc_cert_name and cwc_cert_name.lower() in f.lower():
                    cwc_pdf_file = os.path.join(cert_dir, f)
                    break
                elif "3LPE+CWC" in f.upper() or "CWC" in f.upper():
                    if "IRN" not in f.upper():
                        cwc_pdf_file = os.path.join(cert_dir, f)
                        break
        if cwc_pdf_file:
            self.log(f" -> File Chứng chỉ bọc ống CWC: {os.path.basename(cwc_pdf_file)}")
        else:
            self.log(" [!] Cảnh báo: Không tìm thấy file chứng chỉ bọc ống CWC.")
            
        # 4.3 Files MTC PVP
        pvp_irn_set = set(p["irn"] for p in pipes if p.get("irn"))
        pvp_files = []
        for f in sorted(os.listdir(cert_dir)):
            if f.lower().endswith(".pdf"):
                # Kiểm tra xem tên file có chứa IRN nào không
                for irn in pvp_irn_set:
                    if irn.lower() in f.lower():
                        pvp_files.append((irn, os.path.join(cert_dir, f)))
                        break
        # Loại bỏ trùng lặp nếu có
        seen_paths = set()
        unique_pvp_files = []
        for irn, p_path in pvp_files:
            if p_path not in seen_paths:
                seen_paths.add(p_path)
                unique_pvp_files.append((irn, p_path))
        self.log(f" -> Tìm thấy {len(unique_pvp_files)} file MTC PVP tương ứng với các IRN của lô.")
        
        # 5. Xây dựng tài liệu PDF gộp
        merged_doc = fitz.open()
        
        # A. Ghép Excel (4 trang) - Không tô màu
        excel_doc = fitz.open(excel_pdf_path)
        merged_doc.insert_pdf(excel_doc)
        self.log(f" -> [Mục 1] Đã ghép {len(excel_doc)} trang Excel (RFI, MRIR, Sheet1 - không tô màu).")
        
        # B. Xử lý & Ghép Chứng chỉ bọc ống CWC (Chỉ giữ 2 trang bìa + các trang bảng có ống + trang ký cuối)
        cwc_pages_map = {} # map table_pno -> page in merged_doc
        if cwc_pdf_file:
            cwc_doc = fitz.open(cwc_pdf_file)
            total_cwc_pages = len(cwc_doc)
            
            # Tính toán các trang bảng cần giữ lại
            # Bảng bắt đầu từ trang PDF thứ 3 (index 2). Mỗi trang có 40 số thứ tự (STT).
            table_pages_needed = set()
            for p in pipes:
                stt = p.get("cwc_stt", 0)
                if stt > 0:
                    tbl_p = (stt - 1) // 40 + 1
                    pdf_p = 2 + tbl_p
                    if pdf_p <= total_cwc_pages:
                        table_pages_needed.add((tbl_p, pdf_p - 1)) # (table_pno, 0-based pdf_index)
            
            sorted_needed = sorted(list(table_pages_needed), key=lambda x: x[0])
            
            # Danh sách trang cần ghép: Trang 1, 2 (Bìa 3.2), các trang bảng có ống, và Trang cuối cùng (Trang ký)
            pages_to_insert = [0, 1] if total_cwc_pages >= 2 else [0]
            for tbl_p, pdf_idx in sorted_needed:
                if pdf_idx not in pages_to_insert and pdf_idx < (total_cwc_pages - 1):
                    pages_to_insert.append(pdf_idx)
            # Thêm trang ký cuối cùng
            last_idx = total_cwc_pages - 1
            if last_idx not in pages_to_insert:
                pages_to_insert.append(last_idx)
                
            self.log(f" -> [Mục 2] Lọc file CWC {os.path.basename(cwc_pdf_file)}: Giữ lại {len(pages_to_insert)} trang (Bìa 1-2, các trang bảng có ống, và trang ký cuối).")
            
            # Ghép từng trang vào merged_doc và ghi nhận vị trí
            cwc_cover_pages = []
            for pdf_idx in pages_to_insert:
                new_idx = len(merged_doc)
                merged_doc.insert_pdf(cwc_doc, from_page=pdf_idx, to_page=pdf_idx)
                if pdf_idx in (0, 1):
                    cwc_cover_pages.append(new_idx)
                # Kiểm tra nếu là trang bảng có ống
                for tbl_p, p_idx in sorted_needed:
                    if p_idx == pdf_idx:
                        cwc_pages_map[tbl_p] = new_idx
                        
            # Tô màu số chứng chỉ ở Trang 1 và Trang 2 CWC
            for c_idx in cwc_cover_pages:
                cover_p = merged_doc[c_idx]
                r_cert = self.find_cwc_cover_rect(cover_p)
                if r_cert:
                    quad = self.get_cwc_cover_quad(cover_p, r_cert)
                    annot = cover_p.add_highlight_annot(quad)
                    annot.set_colors(stroke=self.yellow_color)
                    annot.update()
            if cwc_cover_pages:
                self.log(f" -> Đã tô màu số chứng chỉ ở {len(cwc_cover_pages)} trang bìa CWC (Trang 1 và Trang 2).")
                        
            # Tô màu chính xác các dòng ống trong CWC
            cwc_highlight_count = 0
            for tbl_p, page_idx in cwc_pages_map.items():
                p_page = merged_doc[page_idx]
                lines = self.extract_table_lines(p_page)
                base_stt = (tbl_p - 1) * 40 + 1
                
                # Tìm các ống trên trang này
                page_pipes = [p for p in pipes if base_stt <= p.get("cwc_stt", 0) <= (base_stt + 39)]
                for pipe in page_pipes:
                    stt = pipe["cwc_stt"]
                    row_idx = stt - base_stt
                    y0_px = lines[row_idx]
                    y1_px = lines[row_idx + 1]
                    quad = self.get_cwc_row_quad(p_page, y0_px, y1_px)
                    annot = p_page.add_highlight_annot(quad)
                    annot.set_colors(stroke=self.yellow_color)
                    annot.update()
                    cwc_highlight_count += 1
            self.log(f" -> Đã tô màu chính xác {cwc_highlight_count} dòng ống trong bảng bọc ống CWC.")
            
        # C. Ghép Phiếu kiểm tra xuất ống CWC (IRN CWC) - VỊ TRÍ MỚI: NGAY SAU FILE KÝ CWC
        if unique_cwc_irn_files:
            self.log(f" -> [Mục 3] Ghép {len(unique_cwc_irn_files)} Phiếu kiểm tra xuất ống CWC (ngay sau file ký CWC, không tô màu):")
            for irn_f in unique_cwc_irn_files:
                doc_irn = fitz.open(irn_f)
                merged_doc.insert_pdf(doc_irn)
                self.log(f"    + Đã đính kèm: {os.path.basename(irn_f)} ({len(doc_irn)} trang)")

        # D. Ghép và tô màu cả dòng cho các file MTC PVP
        start_mtc_page = len(merged_doc)
        irn_count = 0
        mtc_count = 0
        
        def get_irn_row_rect(r):
            # Trong bảng IRN (chiều dọc), mỗi dòng cao ~15-16 pt.
            # Highlight cao 11 pt căn giữa chữ sẽ chừa khoảng đệm an toàn 2.5 pt trên/dưới,
            # đảm bảo không bao giờ chạm hay lấn qua đường viền sang số ống khác.
            y_mid = (r.y0 + r.y1) / 2
            return fitz.Rect(33.0, y_mid - 5.5, 564.0, y_mid + 5.5)

        def get_mtc_row_rect(r):
            # Trong bảng MTC (chiều ngang), khoảng cách giữa 2 số ống liền kề chỉ ~7-8 pt.
            # Chiều cao highlight được giới hạn chính xác theo chiều cao chữ (r.y0 - 0.5 đến r.y1 + 0.5),
            # và chiều rộng chỉ bao phủ bảng thông tin ống bên trái (x = 38.0 đến 215.0),
            # tuyệt đối không lấn sang dòng ống bên trên/dưới hay cột thành phần hóa học bên phải.
            y0 = r.y0 - 0.5
            y1 = r.y1 + 0.5
            return fitz.Rect(38.0, y0, 215.0, y1)
            
        for pvp_idx, (irn_name, f_path) in enumerate(unique_pvp_files, 1):
            self.log(f" -> Đang xử lý file MTC/IRN [{pvp_idx}/{len(unique_pvp_files)}]: {os.path.basename(f_path)}")
            f_doc = fitz.open(f_path)
            n_pages = len(f_doc)
            cur_p_offset = len(merged_doc)
            merged_doc.insert_pdf(f_doc)
            
            irn_pipes = [p for p in pipes if p["irn"] == irn_name]
            
            for p_offset in range(n_pages):
                p = merged_doc[cur_p_offset + p_offset]
                txt = p.get_text("text")
                
                # Phân loại tuyệt đối dựa trên hướng trang vật lý:
                # Trang nằm ngang (Landscape: width > height) LUÔN LUÔN là chứng chỉ MTC.
                # Trang đứng (Portrait: height > width) LUÔN LUÔN là biên bản IRN.
                is_landscape = (p.rect.width > p.rect.height)
                is_irn = not is_landscape
                is_mtc = is_landscape
                
                # Tô màu tiêu đề biểu mẫu (Header) 1 lần duy nhất cho mỗi trang
                if is_irn:
                    irn_hdr_rect = self.find_irn_header_rect(p)
                    if irn_hdr_rect:
                        annot = p.add_highlight_annot(irn_hdr_rect)
                        annot.set_colors(stroke=self.yellow_color)
                        annot.update()
                elif is_mtc:
                    mtc_hdr_rect = self.find_mtc_header_rect(p)
                    if mtc_hdr_rect:
                        annot = p.add_highlight_annot(mtc_hdr_rect)
                        annot.set_colors(stroke=self.yellow_color)
                        annot.update()
                
                ocr_results = None
                    
                for pipe in irn_pipes:
                    r_list = p.search_for(pipe["tag"])
                    if not r_list and is_irn and "serial" in pipe:
                        r_list = p.search_for(pipe["serial"])
                    if not r_list:
                        for alt in [
                            pipe["tag"].replace("181-A", "I81-A"),
                            pipe["tag"].replace("181-A", "18I-A"),
                            pipe["tag"].replace("181-A", "I8I-A"),
                            pipe["tag"].replace("181-A0", "181-A0I"),
                            pipe["tag"].replace("181-A0", "I81-A0I"),
                            pipe["tag"].replace("1", "I"),
                            pipe.get("serial", "").replace("OM0", "OMQ"),
                            pipe.get("serial", "").replace("OM1", "OMI"),
                        ]:
                            if alt:
                                r_list = p.search_for(alt)
                                if r_list:
                                    break
                                    
                    # Nếu tìm text không thấy và trang scan thiếu text layer (len text < 50), chạy OCR
                    if not r_list and len(txt.strip()) < 50:
                        if ocr_results is None:
                            try:
                                from rapidocr_onnxruntime import RapidOCR
                                if not hasattr(self, '_ocr_engine'):
                                    self._ocr_engine = RapidOCR()
                                pix = p.get_pixmap(dpi=150)
                                ocr_results, _ = self._ocr_engine(pix.tobytes("png"))
                            except Exception:
                                ocr_results = []
                        if ocr_results:
                            scale = 72 / 150
                            for box, ocr_word, score in ocr_results:
                                box_x0 = box[0][0] * scale
                                # Lọc đúng cột Pipe No trên từng loại trang
                                if is_mtc and not (45 <= box_x0 <= 115):
                                    continue
                                if is_irn and not (50 <= box_x0 <= 180):
                                    continue
                                if pipe["tag"] in ocr_word or (pipe.get("serial") and pipe["serial"] in ocr_word):
                                    box_rect = fitz.Rect(box[0][0]*scale, box[0][1]*scale, box[2][0]*scale, box[2][1]*scale)
                                    r_list.append(box_rect)
                                    break
                        
                    highlighted_pipe = False
                    for r in r_list:
                        if highlighted_pipe:
                            break
                        if is_irn and (50 <= r.x0 <= 180):
                            row_r = get_irn_row_rect(r)
                            annot = p.add_highlight_annot(row_r)
                            annot.set_colors(stroke=self.yellow_color)
                            annot.update()
                            irn_count += 1
                            highlighted_pipe = True
                        elif is_mtc and (45 <= r.x0 <= 115):
                            row_r = get_mtc_row_rect(r)
                            annot = p.add_highlight_annot(row_r)
                            annot.set_colors(stroke=self.yellow_color)
                            annot.update()
                            mtc_count += 1
                            highlighted_pipe = True
                            
        self.log(f" -> Đã ghép {len(merged_doc) - start_mtc_page} trang MTC PVP.")
        self.log(f" -> Đã tô màu CẢ DÒNG cho {irn_count} dòng IRN và {mtc_count} dòng MTC PVP.")
        
        # 6. Lưu file PDF hoàn tất
        # Đặt tên file xuất
        clean_code = re.sub(r'[^\w-]', '_', mrir_code)
        if not clean_code.endswith("_Highlighted"):
            output_filename = f"{clean_code}_Highlighted.pdf"
        else:
            output_filename = f"{clean_code}.pdf"
            
        final_pdf_path = os.path.join(output_dir, output_filename)
        try:
            merged_doc.save(final_pdf_path)
        except Exception as e:
            if "Permission denied" in str(e) or "cannot remove" in str(e):
                alt_path = final_pdf_path.replace(".pdf", "_v2.pdf")
                self.log(f"⚠️ Cảnh báo: '{os.path.basename(final_pdf_path)}' đang được mở trong phần mềm đọc PDF (Foxit/Acrobat).")
                self.log(f" -> Tự động lưu sang file mới: {os.path.basename(alt_path)}")
                merged_doc.save(alt_path)
                final_pdf_path = alt_path
            else:
                raise e
        
        self.log("==================================================")
        self.log(f"✅ HOÀN TẤT XỬ LÝ LÔ HỒ SƠ!")
        self.log(f" -> File kết quả: {final_pdf_path}")
        self.log(f" -> Tổng số trang: {len(merged_doc)} trang.")
        self.log(f" -> Dung lượng: {os.path.getsize(final_pdf_path) / (1024*1024):.2f} MB")
        self.log("==================================================")
        return final_pdf_path

    def process_all_batches(self, root_input_dir, output_dir):
        """Tự động tìm kiếm và xử lý tất cả các thư mục lô hàng bên trong thư mục gốc."""
        subdirs = [os.path.join(root_input_dir, d) for d in os.listdir(root_input_dir) if os.path.isdir(os.path.join(root_input_dir, d))]
        batch_dirs = []
        for d in subdirs:
            if self.find_excel_file(d):
                batch_dirs.append(d)
                
        if not batch_dirs:
            # Kiểm tra nếu chính thư mục gốc là 1 lô
            if self.find_excel_file(root_input_dir):
                batch_dirs.append(root_input_dir)
                
        if not batch_dirs:
            raise FileNotFoundError(f"Không tìm thấy thư mục lô hàng hợp lệ nào trong: {root_input_dir}")
            
        self.log(f"🔍 Tìm thấy {len(batch_dirs)} lô hàng cần xử lý:")
        for b in batch_dirs:
            self.log(f"   + {os.path.basename(b)}")
            
        results = []
        for idx, b in enumerate(batch_dirs, 1):
            self.log(f"\n>>> [Lô {idx}/{len(batch_dirs)}]: {os.path.basename(b)}")
            res = self.process_batch(b, output_dir)
            results.append(res)
            
        self.log("\n🎉 ĐÃ HOÀN TẤT XỬ LÝ TOÀN BỘ CÁC LÔ HÀNG!")
        return results

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Tool tự động tổng hợp và tô màu hồ sơ MRIR")
    parser.add_argument("--input", "-i", type=str, default="Input", help="Đường dẫn thư mục lô hàng hoặc thư mục Input tổng")
    parser.add_argument("--output", "-o", type=str, default="Output", help="Đường dẫn thư mục lưu kết quả PDF")
    parser.add_argument("--all", "-a", action="store_true", help="Quét và chạy tất cả các thư mục lô hàng con trong thư mục input")
    args = parser.parse_args()
    
    processor = MRIRBatchProcessor()
    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output)
    
    if args.all or (os.path.isdir(input_path) and not processor.find_excel_file(input_path)):
        processor.process_all_batches(input_path, output_path)
    else:
        processor.process_batch(input_path, output_path)
