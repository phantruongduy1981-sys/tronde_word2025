import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# ==================== CẤU HÌNH TRANG & CSS ====================
st.set_page_config(
    page_title="Trộn Đề THPT Minh Đức",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# MÀU SẮC GIAO DIỆN
HEADER_COLOR = "#00695c"
BUTTON_COLOR = "#d32f2f"
BG_INFO = "#e8f5e9"
BG_WARNING = "#fffde7"
CARD_BG = "#f1f8e9" 

st.markdown(f"""
<style>
    /* 1. Header thu gọn (60% chiều cao cũ) */
    .main-header {{
        background-color: {HEADER_COLOR};
        color: white;
        padding: 15px; 
        text-align: center;
        border-radius: 8px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }}
    .main-header h1 {{
        font-size: 22px; 
        font-weight: 800;
        margin: 0;
        text-transform: uppercase;
        color: white;
        letter-spacing: 1px;
    }}
    .main-header p {{ 
        margin-top: 5px; 
        font-size: 13px; 
        opacity: 0.9; 
        font-weight: 500; 
    }}
    
    /* 2. Style cho thẻ Cấu hình (Input card) */
    .config-card {{
        background-color: {CARD_BG};
        border: 1px solid #c5e1a5;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        margin-bottom: 10px;
    }}
    .config-label {{
        color: {HEADER_COLOR};
        font-weight: bold;
        font-size: 14px;
        margin-bottom: 5px;
        display: block;
    }}

    .info-box {{ background-color: {BG_INFO}; border-left: 5px solid #66bb6a; padding: 15px; border-radius: 5px; margin-bottom: 15px; }}
    .info-header {{ font-weight: bold; color: #2e7d32; font-size: 16px; margin-bottom: 10px; }}
    .warning-box {{ background-color: {BG_WARNING}; border-left: 5px solid #fbc02d; padding: 15px; border-radius: 5px; margin-top: 15px; }}
    .warning-header {{ font-weight: bold; color: #f57f17; margin-bottom: 8px; }}

    .stButton>button {{
        background-color: {BUTTON_COLOR};
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: bold;
        width: 100%;
        height: 50px;
        font-size: 16px;
        transition: all 0.3s;
    }}
    .stButton>button:hover {{ background-color: #b71c1c; box-shadow: 0 4px 8px rgba(0,0,0,0.2); }}

    .step-circle {{
        display: inline-block; width: 28px; height: 28px; background-color: {HEADER_COLOR};
        color: white; border-radius: 50%; text-align: center; line-height: 28px; font-weight: bold; margin-right: 10px;
    }}
    .section-title {{ font-size: 18px; font-weight: bold; color: {HEADER_COLOR}; margin-bottom: 15px; display: flex; align-items: center; }}
    
    .footer {{ text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #888; font-size: 13px; }}
</style>
""", unsafe_allow_html=True)

# ==================== CORE LOGIC ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text_from_node(node):
    texts = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def check_structure_errors(blocks):
    full_text = "\n".join([get_text_from_node(b) for b in blocks])
    errors = []
    if not re.search(r'Câu\s*1[\.:]', full_text, re.IGNORECASE):
        errors.append("❌ Lỗi: Không tìm thấy 'Câu 1'. File phải bắt đầu bằng Câu 1.")
    return errors

def is_answer_marked(paragraph):
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
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for run in runs:
        rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr_list: continue
        rPr = rPr_list[0]
        for c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(c)
        for u in rPr.getElementsByTagNameNS(W_NS, "u"): rPr.removeChild(u)

def update_label_in_node(paragraph, new_label):
    """Cập nhật nhãn (Câu 1., A., a))."""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    first_node = None
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue.strip():
            first_node = t; break
    if not first_node: return
    
    txt = first_node.firstChild.nodeValue
    
    # 1. Xử lý A. B. C. D.
    if re.match(r'^\s*[A-D][\.:\)]', txt, re.IGNORECASE):
        sub = re.sub(r'^\s*[A-D][\.:\)]', new_label, txt, count=1)
        first_node.firstChild.nodeValue = sub
        
    # 2. Xử lý a) b) c) d)
    elif re.match(r'^\s*[a-d][\.:\)]', txt, re.IGNORECASE):
        sub = re.sub(r'^\s*[a-d][\.:\)]', new_label, txt, count=1)
        first_node.firstChild.nodeValue = sub
        
    # 3. Xử lý Câu X (Thay thành Câu X. theo yêu cầu)
    elif re.match(r'^\s*Câu\s*\d+', txt, re.IGNORECASE):
        sub = re.sub(r'^\s*Câu\s*\d+[\.:]?', new_label, txt, count=1, flags=re.IGNORECASE)
        first_node.firstChild.nodeValue = sub

def extract_part3_answer(question_blocks):
    answer_text = "X"
    blocks_to_keep = []
    found_ans = False
    for block in reversed(question_blocks):
        txt = get_text_from_node(block)
        match = re.search(r'ĐS[:\s]+(.*)', txt, re.IGNORECASE) or re.search(r'DS[:\s]+(.*)', txt, re.IGNORECASE)
        if match and not found_ans:
            if is_answer_marked(block):
                answer_text = match.group(1).strip()
                found_ans = True
                continue 
        blocks_to_keep.insert(0, block)
    return blocks_to_keep, answer_text

def shuffle_questions(questions, mode="MCQ"):
    indices = list(range(len(questions)))
    random.shuffle(indices)
    shuffled_output = []
    key_map = {} 
    labels_mcq = ["A.", "B.", "C.", "D."]
    labels_tf = ["a)", "b)", "c)", "d)"]
    
    for new_idx, old_idx in enumerate(indices):
        q_blocks = questions[old_idx] 
        
        # --- PHẦN 3 ---
        if mode == "FILL":
            cleaned_blocks, ans_text = extract_part3_answer(q_blocks)
            if cleaned_blocks: 
                # Yêu cầu: Câu 1.
                update_label_in_node(cleaned_blocks[0], f"Câu {new_idx + 1}. ")
            shuffled_output.extend(cleaned_blocks)
            key_map[new_idx + 1] = ans_text
            continue

        # --- PHẦN 1 & 2 ---
        intro = []; options = []
        for b in q_blocks:
            txt = get_text_from_node(b)
            is_opt = False
            if mode == "MCQ" and re.match(r'^\s*[A-D][\.:]', txt): is_opt = True
            elif mode == "TF" and re.match(r'^\s*[a-d][\)]', txt): is_opt = True
            
            if is_opt:
                is_correct = is_answer_marked(b)
                clean_formatting(b)
                options.append({'node': b, 'correct': is_correct, 'text': txt})
            else: intro.append(b)
        
        correct_char = ""
        
        if options:
            if mode == "MCQ":
                random.shuffle(options)
                for i, opt in enumerate(options):
                    lbl = labels_mcq[i] if i < 4 else "*"
                    update_label_in_node(opt['node'], lbl)
                    if opt['correct']: correct_char = lbl[0]
            
            elif mode == "TF":
                # Giữ d) ở cuối
                d_node = None
                others = []
                for opt in options:
                    if re.match(r'^\s*d[\)]', opt['text'], re.IGNORECASE): d_node = opt
                    else: others.append(opt)
                
                random.shuffle(others)
                final_opts = others
                if d_node: final_opts.append(d_node)
                
                for i, opt in enumerate(final_opts):
                    lbl = labels_tf[i] if i < 4 else "*"
                    update_label_in_node(opt['node'], lbl)
                options = final_opts

        # Cập nhật nhãn Câu (1 dấu chấm)
        if intro: 
            update_label_in_node(intro[0], f"Câu {new_idx + 1}. ")
            
        if mode == "MCQ": key_map[new_idx + 1] = correct_char if correct_char else "X"
            
        shuffled_output.extend(intro)
        for o in options: shuffled_output.extend([o['node']])
        
    return shuffled_output, key_map

# --- TẠO FILE ĐÁP ÁN WORD ---
def create_word_answer_key(excel_data_list):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    doc.add_heading('BẢNG ĐÁP ÁN CHI TIẾT', 0)
    
    for data in excel_data_list:
        exam_code = data.get("Mã đề", "Unknown")
        p = doc.add_paragraph()
        p.add_run(f'Mã đề: {exam_code}').bold = True
        
        questions = {int(k): v for k, v in data.items() if k.isdigit()}
        sorted_q = sorted(questions.items())
        
        if not sorted_q: continue
        
        table = doc.add_table(rows=1, cols=10)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_cells = table.rows[0].cells
        for i in range(10):
            hdr_cells[i].text = "Câu-ĐA"
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            
        row_cells = None
        for idx, (q_num, ans) in enumerate(sorted_q):
            col_idx = idx % 10
            if col_idx == 0: row_cells = table.add_row().cells
            row_cells[col_idx].text = f"{q_num}-{ans}"
            
        doc.add_paragraph("\n")

    f = io.BytesIO()
    doc.save(f)
    return f.getvalue()

def process_docx(file_bytes, num_exams, start_id, shuffle_mode):
    input_io = io.BytesIO(file_bytes)
    resources = {}; xml_content = ""
    try:
        with zipfile.ZipFile(input_io, 'r') as zin:
            for filename in zin.namelist():
                if filename == "word/document.xml": xml_content = zin.read(filename).decode('utf-8')
                else: resources[filename] = zin.read(filename)
    except Exception as e: return None, None, [f"Lỗi đọc file: {str(e)}"]

    if not xml_content: return None, None, ["Không tìm thấy document.xml"]

    dom = minidom.parseString(xml_content)
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    all_blocks = [child for child in list(body.childNodes) 
                  if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]]
    for b in all_blocks: body.removeChild(b)
            
    errors = check_structure_errors(all_blocks)
    
    parts = []; current_part = []
    for block in all_blocks:
        txt = get_text_from_node(block)
        if re.match(r'^\s*PHẦN\s*\d+', txt, re.IGNORECASE):
            if current_part: parts.append(current_part)
            current_part = [block]
        else: current_part.append(block)
    if current_part: parts.append(current_part)
    if not parts: parts = [all_blocks]

    # --- TẠO 1 FILE ZIP CHỨA TẤT CẢ ---
    final_zip_io = io.BytesIO()
    excel_data_list = []
    
    with zipfile.ZipFile(final_zip_io, 'w', zipfile.ZIP_DEFLATED) as master_zip:
        
        # 1. Tạo các đề thi
        for ver in range(num_exams):
            current_code = str(start_id + ver)
            curr_dom = minidom.parseString(xml_content)
            curr_body = curr_dom.getElementsByTagNameNS(W_NS, "body")[0]
            while curr_body.firstChild: curr_body.removeChild(curr_body.firstChild)
            
            exam_blocks = []; exam_key = {"Mã đề": current_code}; global_q_idx = 1
            
            for part_blocks in parts:
                questions = []; intro_part = []; curr_q = []; is_q = False
                cloned_blocks = [b.cloneNode(True) for b in part_blocks]
                
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
                
                part_txt = get_text_from_node(intro_part[0]) if intro_part else ""
                current_mode = "MCQ"
                if shuffle_mode == "auto":
                    if "PHẦN 2" in part_txt.upper(): current_mode = "TF"
                    elif "PHẦN 3" in part_txt.upper(): current_mode = "FILL"
                elif shuffle_mode == "tf": current_mode = "TF"
                
                shuffled_nodes, key_map = shuffle_questions(questions, mode=current_mode)
                
                final_nodes = []; local_count = 0
                for k, v in key_map.items(): exam_key[str(global_q_idx + k - 1)] = v
                for node in shuffled_nodes:
                    txt = get_text_from_node(node)
                    if re.match(r'^\s*Câu\s*\d+', txt):
                        # Cập nhật nhãn Câu toàn cục (1 dấu chấm)
                        update_label_in_node(node, f"Câu {global_q_idx + local_count}. ")
                        local_count += 1
                    final_nodes.append(node)
                
                global_q_idx += len(key_map)
                exam_blocks.extend(intro_part)
                exam_blocks.extend(final_nodes)

            for b in exam_blocks: curr_body.appendChild(b)
            
            new_xml = curr_dom.toxml()
            sub_io = io.BytesIO()
            with zipfile.ZipFile(sub_io, 'w', zipfile.ZIP_DEFLATED) as sub_z:
                sub_z.writestr("word/document.xml", new_xml.encode('utf-8'))
                for name, content in resources.items(): sub_z.writestr(name, content)
            
            # Ghi file đề vào thư mục De_Thi trong Zip
            master_zip.writestr(f"De_Thi/De_{current_code}.docx", sub_io.getvalue())
            excel_data_list.append(exam_key)
            
        # 2. Tạo Excel đáp án
        df = pd.DataFrame(excel_data_list)
        cols = ["Mã đề"] + sorted([c for c in df.columns if c != "Mã đề"], key=lambda x: int(x) if x.isdigit() else 999)
        df = df[cols] if not df.empty else df
        excel_io = io.BytesIO()
        try:
            with pd.ExcelWriter(excel_io, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='DapAn')
            master_zip.writestr("Dap_An_Excel.xlsx", excel_io.getvalue())
        except: pass

        # 3. Tạo Word Đáp án
        word_bytes = create_word_answer_key(excel_data_list)
        master_zip.writestr("Dap_An_Word.docx", word_bytes)
        
    return final_zip_io.getvalue(), None, None, errors

# ==================== MAIN ====================
def main():
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG TRUNG HỌC PHỔ THÔNG MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025 - PRO VERSION</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        with st.expander("📄 Hướng dẫn & File Mẫu", expanded=True):
            # LINK FILE MỚI CỦA BẠN
            sample_url = "https://docs.google.com/document/d/1i1b-By6EA_HO8fWgMYG9iXZPGannmWdg/export?format=docx"
            st.link_button("📥 Tải File Mẫu", sample_url, use_container_width=True)
            st.markdown("""
            <div class="info-box" style="margin-top:15px;">
                <div class="info-header">📌 Cấu trúc:</div>
                <ul><li><b>PHẦN 1:</b> Trắc nghiệm (A.B.C.D)</li><li><b>PHẦN 2:</b> Đúng/Sai (a)b)c)d))</li><li><b>PHẦN 3:</b> Trả lời ngắn</li></ul>
            </div>
            <div class="warning-box">
                <div class="warning-header">⚠️ Lưu ý:</div>
                <ul><li>Câu hỏi: <code>Câu 1.</code> (1 dấu chấm)</li><li>Đáp án d) (Phần 2) luôn cố định ở cuối.</li></ul>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<div class="section-title"><span class="step-circle">1</span> Chọn file đề Word (*.docx)</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("", type=["docx"], label_visibility="collapsed")
        if uploaded_file: st.success(f"✅ Đã tải: {uploaded_file.name}")
        
    with col2:
        st.markdown('<div class="section-title"><span class="step-circle">2</span> Cấu hình</div>', unsafe_allow_html=True)
        shuffle_opt = st.radio("", ["🔄 Tự động", "📝 Trắc nghiệm", "✅ Đúng/Sai"], index=0, label_visibility="collapsed")
        mode_map = {"🔄 Tự động": "auto", "📝 Trắc nghiệm": "mcq", "✅ Đúng/Sai": "tf"}
        selected_mode = mode_map[shuffle_opt]

        st.write("")
        # --- INPUT ĐƯỢC TÔ MÀU ĐẸP ---
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="config-card"><span class="config-label">SỐ LƯỢNG ĐỀ</span></div>', unsafe_allow_html=True)
            num_exams = st.number_input("Số lượng", 1, 50, 4, label_visibility="collapsed")
        with c2:
            st.markdown('<div class="config-card"><span class="config-label">MÃ BẮT ĐẦU</span></div>', unsafe_allow_html=True)
            start_id = st.number_input("Mã bắt đầu", 1, 9999, 1001, label_visibility="collapsed")
            
        st.caption(f"📍 Sẽ tạo các mã đề từ: {start_id} ➝ {start_id + num_exams - 1}")
        st.markdown("<br>", unsafe_allow_html=True)

        # NÚT TẢI TRỌN BỘ
        if st.button("🚀 Trộn đề & Tải Trọn Bộ (ZIP)"):
            if not uploaded_file: st.error("⚠️ Vui lòng chọn file đề trước!")
            else:
                with st.spinner("⏳ Đang xử lý..."):
                    try:
                        file_bytes = uploaded_file.read()
                        uploaded_file.seek(0)
                        
                        # Chỉ trả về 1 biến zip_data
                        zip_data, _, _, errors = process_docx(file_bytes, num_exams, start_id, selected_mode)
                        
                        if errors:
                            for e in errors: st.error(e)
                        
                        if zip_data:
                            st.download_button(
                                "📦 Tải Trọn Bộ (ZIP)", 
                                zip_data, 
                                f"Tron_Bo_De_{uploaded_file.name}.zip", 
                                "application/zip",
                                type="primary"
                            )
                            st.balloons()
                            
                    except Exception as e: st.error(f"Lỗi: {str(e)}")

    st.markdown("""<div class="footer">© 2025 Phan Trường Duy - THPT Minh Đức<br>PRO VERSION 2.2</div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
