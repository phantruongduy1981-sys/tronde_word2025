import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT

# ==================== 1. CẤU HÌNH GIAO DIỆN & CSS ====================
st.set_page_config(
    page_title="Trộn Đề THPT Minh Đức",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

HEADER_COLOR = "#00695c"
BUTTON_COLOR = "#d32f2f"
BG_INFO = "#e8f5e9"
BG_WARNING = "#fffde7"
CARD_BG = "#f1f8e9"

st.markdown(f"""
<style>
    .main-header {{ background-color: {HEADER_COLOR}; color: white; padding: 15px; text-align: center; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
    .main-header h1 {{ font-size: 22px; font-weight: 800; margin: 0; text-transform: uppercase; letter-spacing: 1px; }}
    .main-header p {{ margin-top: 5px; font-size: 13px; opacity: 0.9; font-weight: 500; }}
    .config-card {{ background-color: {CARD_BG}; border: 1px solid #c5e1a5; border-radius: 8px; padding: 15px; text-align: center; margin-bottom: 10px; }}
    .config-label {{ color: {HEADER_COLOR}; font-weight: bold; font-size: 14px; margin-bottom: 5px; display: block; }}
    .stButton>button {{ background-color: {BUTTON_COLOR}; color: white; border: none; border-radius: 8px; font-weight: bold; width: 100%; height: 50px; font-size: 16px; transition: 0.3s; }}
    .stButton>button:hover {{ background-color: #b71c1c; box-shadow: 0 4px 8px rgba(0,0,0,0.2); }}
    .info-box {{ background-color: {BG_INFO}; border-left: 5px solid #66bb6a; padding: 15px; border-radius: 5px; margin-bottom: 15px; }}
    .warning-box {{ background-color: {BG_WARNING}; border-left: 5px solid #fbc02d; padding: 15px; border-radius: 5px; margin-top: 15px; }}
    .footer {{ text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #888; font-size: 13px; }}
</style>
""", unsafe_allow_html=True)

# ==================== 2. CÁC HÀM XỬ LÝ CỐT LÕI (CORE UTILS) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

def get_text_from_node(node):
    texts = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def has_complex_content(node):
    """Kiểm tra xem dòng có chứa công thức toán/ảnh không để tránh tách dòng"""
    if node.getElementsByTagNameNS(M_NS, "oMath") or node.getElementsByTagNameNS(M_NS, "oMathPara"): return True
    if node.getElementsByTagNameNS(W_NS, "object") or node.getElementsByTagNameNS(W_NS, "drawing"): return True
    return False

def check_structure_errors(blocks):
    full_text = "\n".join([get_text_from_node(b) for b in blocks])
    errors = []
    if not re.search(r'Câu\s*1[\.:]', full_text, re.IGNORECASE):
        errors.append("❌ Lỗi: Không tìm thấy 'Câu 1'. File phải bắt đầu bằng Câu 1.")
    return errors

def is_marked_correct(paragraph):
    """Kiểm tra đáp án đúng (Màu đỏ/Gạch chân)"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for run in runs:
        rPr = run.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr: continue
        rPr = rPr[0]
        colors = rPr.getElementsByTagNameNS(W_NS, "color")
        for c in colors:
            val = c.getAttributeNS(W_NS, "val")
            if val and (val.upper() in ['FF0000', 'RED']): return True
        u_tags = rPr.getElementsByTagNameNS(W_NS, "u")
        for u in u_tags:
            val = u.getAttributeNS(W_NS, "val")
            if val and val != 'none': return True
    return False

def clean_formatting(paragraph):
    """Xóa định dạng đỏ/gạch chân"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for run in runs:
        rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr_list: continue
        rPr = rPr_list[0]
        for c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(c)
        for u in rPr.getElementsByTagNameNS(W_NS, "u"): rPr.removeChild(u)

def update_label_safely(paragraph, new_label):
    """Cập nhật nhãn mà không làm mất công thức"""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    
    # Chỉ tìm ở các node đầu tiên để tránh sửa nhầm nội dung bên trong
    for t in t_nodes:
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        
        # Regex A. B. C. D.
        if re.match(r'^\s*[A-D][\.:\)]', txt, re.IGNORECASE):
            sub = re.sub(r'^\s*[A-D][\.:\)]', new_label, txt, count=1)
            t.firstChild.nodeValue = sub
            return
            
        # Regex a) b) c) d)
        elif re.match(r'^\s*[a-d][\.:\)]', txt, re.IGNORECASE):
            sub = re.sub(r'^\s*[a-d][\.:\)]', new_label, txt, count=1)
            t.firstChild.nodeValue = sub
            return
            
        # Regex Câu X (SỬA LỖI 2 DẤU CHẤM TẠI ĐÂY)
        # [\.:]* sẽ ăn hết các dấu chấm thừa
        elif re.match(r'^\s*Câu\s*\d+', txt, re.IGNORECASE):
            sub = re.sub(r'^\s*Câu\s*\d+[\.:]*', new_label, txt, count=1, flags=re.IGNORECASE)
            t.firstChild.nodeValue = sub
            return

def split_options_if_safe(question_blocks):
    """Tách dòng nếu an toàn (không có công thức)"""
    normalized_blocks = []
    for block in question_blocks:
        if has_complex_content(block):
            normalized_blocks.append(block)
            continue
            
        txt = get_text_from_node(block)
        matches = list(re.finditer(r'(?:\s|^)([A-D])[\.:]\s', txt))
        
        if len(matches) > 1:
            idxs = [m.start() for m in matches]
            idxs.append(len(txt))
            for i in range(len(matches)):
                start = idxs[i]; end = idxs[i+1]
                sub_text = txt[start:end].strip()
                new_blk = block.cloneNode(True)
                # Set text đơn giản
                runs = new_blk.getElementsByTagNameNS(W_NS, "r")
                if runs:
                    r0 = runs[0]
                    for t in r0.getElementsByTagNameNS(W_NS, "t"): r0.removeChild(t)
                    doc = new_blk.ownerDocument
                    new_t = doc.createElementNS(W_NS, "w:t")
                    new_t.setAttribute("xml:space", "preserve")
                    new_t.appendChild(doc.createTextNode(sub_text))
                    r0.appendChild(new_t)
                    for k in range(1, len(runs)): new_blk.removeChild(runs[k])
                normalized_blocks.append(new_blk)
        else:
            normalized_blocks.append(block)
    return normalized_blocks

# ==================== 3. CÁC HÀM XỬ LÝ TỪNG PHẦN (RIÊNG BIỆT) ====================

def process_part1_mcq(question_blocks, q_idx):
    """Xử lý PHẦN 1: Trắc nghiệm"""
    # 1. Tách dòng nếu cần
    normalized_blocks = split_options_if_safe(question_blocks)
    
    intro = []
    options = []
    
    for b in normalized_blocks:
        txt = get_text_from_node(b)
        if re.match(r'^\s*[A-D][\.:]', txt):
            is_c = is_marked_correct(b)
            options.append({'node': b, 'correct': is_c, 'text': txt})
        else:
            intro.append(b)
    
    labels_mcq = ["A.", "B.", "C.", "D."]
    correct_char = "X"
    
    # 2. Trộn đáp án (chỉ khi đủ >=2 phương án)
    if len(options) >= 2:
        random.shuffle(options)
        for i, opt in enumerate(options):
            lbl = labels_mcq[i] if i < 4 else "*"
            clean_formatting(opt['node'])
            update_label_safely(opt['node'], lbl)
            if opt['correct']: correct_char = lbl[0]
        result_blocks = intro + [o['node'] for o in options]
    else:
        # Giữ nguyên nếu không nhận diện được
        for opt in options:
            if opt['correct']:
                m = re.match(r'^\s*([A-D])', opt['text'], re.IGNORECASE)
                if m: correct_char = m.group(1).upper()
            clean_formatting(opt['node'])
        result_blocks = intro + [o['node'] for o in options]

    # 3. Cập nhật nhãn câu hỏi (Fix 1 dấu chấm)
    if intro: 
        update_label_safely(intro[0], f"Câu {q_idx}.")
        
    return result_blocks, correct_char

def process_part2_tf(question_blocks, q_idx):
    """Xử lý PHẦN 2: Đúng/Sai"""
    intro = []
    options = []
    
    for b in question_blocks:
        txt = get_text_from_node(b)
        if re.match(r'^\s*[a-d][\.:\)]', txt, re.IGNORECASE):
            is_c = is_marked_correct(b)
            clean_formatting(b)
            options.append({'node': b, 'text': txt})
        else:
            intro.append(b)
            
    # Tách d) ra
    d_node = None
    others = []
    for o in options:
        if re.match(r'^\s*d[\.:\)]', o['text'], re.IGNORECASE): d_node = o
        else: others.append(o)
    
    # Trộn a,b,c
    random.shuffle(others)
    
    # Ghép lại
    final_opts = others + ([d_node] if d_node else [])
    
    # ĐÁNH LẠI NHÃN a) b) c) d) (Fix lỗi chưa đánh lại nhãn)
    labels_tf = ["a)", "b)", "c)", "d)"]
    for k, opt in enumerate(final_opts):
        lbl = labels_tf[k] if k < 4 else "*"
        update_label_safely(opt['node'], lbl)
    
    if intro: 
        update_label_safely(intro[0], f"Câu {q_idx}.")
        
    result_blocks = intro + [o['node'] for o in final_opts]
    return result_blocks, "X" # Phần 2 không có đáp án đơn A/B/C/D

def process_part3_fill(question_blocks, q_idx):
    """Xử lý PHẦN 3: Trả lời ngắn"""
    answer_text = "X"
    blocks_to_keep = []
    found_ans = False
    
    for block in reversed(question_blocks):
        txt = get_text_from_node(block)
        match = re.search(r'(?:ĐS|DS|Đáp số)[:\s]+(.*)', txt, re.IGNORECASE)
        if match and not found_ans:
            if is_marked_correct(block):
                answer_text = match.group(1).strip()
                found_ans = True
                continue # Xóa dòng ĐS
        blocks_to_keep.insert(0, block)
        
    if blocks_to_keep:
        update_label_safely(blocks_to_keep[0], f"Câu {q_idx}.")
        
    return blocks_to_keep, answer_text

# ==================== 4. LOGIC TRỘN ĐỀ TỔNG HỢP ====================

def process_docx(file_bytes, num_exams, start_id, shuffle_mode):
    input_io = io.BytesIO(file_bytes)
    resources = {}; xml_content = ""
    try:
        with zipfile.ZipFile(input_io, 'r') as zin:
            for filename in zin.namelist():
                if filename == "word/document.xml": xml_content = zin.read(filename).decode('utf-8')
                else: resources[filename] = zin.read(filename)
    except Exception as e: return None, None, None, [f"Lỗi đọc file: {str(e)}"]

    if not xml_content: return None, None, None, ["Không tìm thấy document.xml"]

    dom = minidom.parseString(xml_content)
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    all_blocks = [child for child in list(body.childNodes) 
                  if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]]
    errors = check_structure_errors(all_blocks)
    
    # Chia Phần
    parts = []; current_part = []
    for block in all_blocks:
        txt = get_text_from_node(block)
        if re.match(r'^\s*PHẦN\s*\d+', txt, re.IGNORECASE):
            if current_part: parts.append(current_part)
            current_part = [block]
        else: current_part.append(block)
    if current_part: parts.append(current_part)
    if not parts: parts = [all_blocks]

    final_zip_io = io.BytesIO()
    excel_data_list = []
    
    with zipfile.ZipFile(final_zip_io, 'w', zipfile.ZIP_DEFLATED) as master_zip:
        for ver in range(num_exams):
            current_code = str(start_id + ver)
            curr_dom = minidom.parseString(xml_content)
            curr_body = curr_dom.getElementsByTagNameNS(W_NS, "body")[0]
            while curr_body.firstChild: curr_body.removeChild(curr_body.firstChild)
            
            exam_blocks = []; exam_key = {"Mã đề": current_code}; global_q_idx = 1
            
            for part_blocks in parts:
                questions = []; intro_part = []; curr_q = []; is_q = False
                cloned_blocks = [b.cloneNode(True) for b in part_blocks]
                
                # Gom nhóm câu hỏi
                for b in cloned_blocks:
                    txt = get_text_from_node(b)
                    if re.match(r'^\s*Câu\s*\d+', txt, re.IGNORECASE):
                        if curr_q: questions.append(curr_q)
                        curr_q = [b]; is_q = True
                    elif re.match(r'^\s*PHẦN', txt, re.IGNORECASE):
                        if curr_q: questions.append(curr_q)
                        curr_q = []; intro_part.append(b); is_q = False
                    else:
                        if is_q: curr_q.append(b)
                        else: intro_part.append(b)
                if curr_q: questions.append(curr_q)
                
                # TRỘN THỨ TỰ CÂU HỎI (FIX LỖI QUÊN TRỘN CÂU HỎI)
                random.shuffle(questions)
                
                # Xác định Mode
                part_txt = get_text_from_node(intro_part[0]) if intro_part else ""
                mode = "MCQ"
                if shuffle_mode == "auto":
                    if "PHẦN 2" in part_txt.upper(): mode = "TF"
                    elif "PHẦN 3" in part_txt.upper(): mode = "FILL"
                elif shuffle_mode == "tf": mode = "TF"
                
                # Xử lý từng câu hỏi
                part_nodes = []
                for q in questions:
                    if mode == "MCQ":
                        nodes, ans = process_part1_mcq(q, global_q_idx)
                    elif mode == "TF":
                        nodes, ans = process_part2_tf(q, global_q_idx)
                    elif mode == "FILL":
                        nodes, ans = process_part3_fill(q, global_q_idx)
                    else:
                        nodes, ans = process_part1_mcq(q, global_q_idx)
                    
                    part_nodes.extend(nodes)
                    exam_key[str(global_q_idx)] = ans
                    global_q_idx += 1
                
                exam_blocks.extend(intro_part + part_nodes)

            for b in exam_blocks: curr_body.appendChild(b)
            
            new_xml = curr_dom.toxml()
            sub_io = io.BytesIO()
            with zipfile.ZipFile(sub_io, 'w', zipfile.ZIP_DEFLATED) as sub_z:
                sub_z.writestr("word/document.xml", new_xml.encode('utf-8'))
                for name, content in resources.items(): sub_z.writestr(name, content)
            
            master_zip.writestr(f"De_Thi/De_{current_code}.docx", sub_io.getvalue())
            excel_data_list.append(exam_key)
            
        # Export Excel & Word
        if excel_data_list:
            df = pd.DataFrame(excel_data_list)
            cols = ["Mã đề"] + sorted([c for c in df.columns if c != "Mã đề"], key=lambda x: int(x) if x.isdigit() else 999)
            df = df[cols]
            excel_io = io.BytesIO()
            with pd.ExcelWriter(excel_io, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='DapAn')
            master_zip.writestr("Dap_An_Excel.xlsx", excel_io.getvalue())
            
            # Word Answer Key
            doc = Document()
            doc.add_heading('BẢNG ĐÁP ÁN', 0)
            for data in excel_data_list:
                doc.add_heading(f'Mã đề: {data.get("Mã đề", "")}', level=1)
                qs = {int(k): v for k, v in data.items() if k.isdigit()}
                sorted_q = sorted(qs.items())
                if not sorted_q: continue
                table = doc.add_table(rows=1, cols=10)
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                for i in range(10): table.rows[0].cells[i].text = "Câu-ĐA"
                for idx, (q_num, ans) in enumerate(sorted_q):
                    col = idx % 10
                    if col == 0: table.add_row()
                    table.rows[-1].cells[col].text = f"{q_num}-{ans}"
                doc.add_paragraph("\n")
            w_io = io.BytesIO()
            doc.save(w_io)
            master_zip.writestr("Dap_An_Word.docx", w_io.getvalue())
        
    return final_zip_io.getvalue(), None, None, errors

# ==================== 5. GIAO DIỆN MAIN ====================
def main():
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG TRUNG HỌC PHỔ THÔNG MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025 - PRO VERSION 5.0 (Final Fix)</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        with st.expander("📄 Hướng dẫn Quy cách Ghi Đáp Án", expanded=True):
            st.markdown("""
            <div class="info-box">
                <div class="info-header">📌 QUY ĐỊNH BẮT BUỘC:</div>
                <ul>
                    <li><b>Câu hỏi:</b> Bắt đầu bằng <code>Câu 1.</code> (1 dấu chấm).</li>
                    <li><b>Phần 1 (A.B.C.D) & Phần 2 (a.b.c.d):</b><br>
                        👉 Đáp án đúng phải <span style="color:red; font-weight:bold">TÔ ĐỎ</span> hoặc <u>GẠCH CHÂN</u>.<br>
                        👉 Nếu không tô đỏ/gạch chân, App sẽ hiểu là không có đáp án.
                    </li>
                    <li><b>Phần 3 (Trả lời ngắn):</b><br>
                        👉 Ghi <span style="color:red; font-weight:bold">ĐS: Kết quả</span> và TÔ ĐỎ ở cuối câu hỏi.<br>
                        <i>Ví dụ: <span style="color:red">ĐS: 2025</span></i>
                    </li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
            sample_url = "https://docs.google.com/document/d/1i1b-By6EA_HO8fWgMYG9iXZPGannmWdg/export?format=docx"
            st.link_button("📥 Tải File Mẫu Mới Nhất", sample_url, use_container_width=True)

        st.markdown('<div class="section-title"><span class="step-circle">1</span> Chọn file đề Word (*.docx)</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("", type=["docx"], label_visibility="collapsed")
        if uploaded_file: st.success(f"✅ Đã tải: {uploaded_file.name}")
        
    with col2:
        st.markdown('<div class="section-title"><span class="step-circle">2</span> Cấu hình</div>', unsafe_allow_html=True)
        shuffle_opt = st.radio("", ["🔄 Tự động", "📝 Trắc nghiệm", "✅ Đúng/Sai"], index=0, label_visibility="collapsed")
        mode_map = {"🔄 Tự động": "auto", "📝 Trắc nghiệm": "mcq", "✅ Đúng/Sai": "tf"}
        selected_mode = mode_map[shuffle_opt]

        st.write("")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="config-card"><span class="config-label">SỐ LƯỢNG ĐỀ</span></div>', unsafe_allow_html=True)
            num_exams = st.number_input("Số lượng", 1, 50, 4, label_visibility="collapsed")
        with c2:
            st.markdown('<div class="config-card"><span class="config-label">MÃ BẮT ĐẦU</span></div>', unsafe_allow_html=True)
            start_id = st.number_input("Mã bắt đầu", 1, 9999, 1001, label_visibility="collapsed")
            
        st.caption(f"📍 Sẽ tạo các mã đề từ: {start_id} ➝ {start_id + num_exams - 1}")
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🚀 Trộn đề & Tải Trọn Bộ (ZIP)"):
            if not uploaded_file: st.error("⚠️ Vui lòng chọn file đề trước!")
            else:
                with st.spinner("⏳ Đang xử lý..."):
                    try:
                        file_bytes = uploaded_file.read()
                        uploaded_file.seek(0)
                        zip_data, _, _, errors = process_docx(file_bytes, num_exams, start_id, selected_mode)
                        
                        if errors:
                            for e in errors: st.error(e)
                        
                        if zip_data:
                            st.download_button("📦 Tải Trọn Bộ (ZIP)", zip_data, f"Tron_Bo_De_{uploaded_file.name}.zip", "application/zip", type="primary")
                            st.balloons()
                            
                    except Exception as e: st.error(f"Lỗi: {str(e)}")

    st.markdown("""<div class="footer">© 2025 Phan Trường Duy - THPT Minh Đức<br>PRO VERSION 5.0</div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
