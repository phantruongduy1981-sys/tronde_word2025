import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG & CSS ====================
st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Màu sắc chủ đạo
PRIMARY_COLOR = "#00796b" 
BG_YELLOW = "#fff9c4"
TEXT_RED = "#d32f2f"

st.markdown(f"""
<style>
    /* Header chính */
    .main-header {{
        background-color: {PRIMARY_COLOR};
        color: white;
        padding: 20px;
        text-align: center;
        border-radius: 8px;
        margin-bottom: 20px;
    }}
    .main-header h1 {{
        font-size: 28px;
        font-weight: bold;
        margin: 0;
        text-transform: uppercase;
        color: white;
    }}
    .main-header p {{
        margin: 5px 0 0 0;
        font-size: 16px;
        opacity: 0.9;
    }}

    /* Khung Hướng dẫn */
    .info-box {{
        background-color: #f1f8e9;
        border: 1px solid #c5e1a5;
        border-radius: 5px;
        padding: 15px;
        margin-bottom: 15px;
        font-size: 14px;
        line-height: 1.6;
    }}
    .info-title {{
        color: {PRIMARY_COLOR};
        font-weight: bold;
        font-size: 16px;
        margin-bottom: 10px;
        text-transform: uppercase;
        border-bottom: 2px solid {PRIMARY_COLOR};
        display: inline-block;
    }}
    .part-label {{
        font-weight: bold;
        color: #2e7d32;
    }}
    
    /* Khung Cảnh báo/Quy ước */
    .warning-box {{
        background-color: {BG_YELLOW};
        border: 1px solid #fff59d;
        border-radius: 5px;
        padding: 15px;
        margin-top: 15px;
        font-size: 14px;
        color: #f57f17;
    }}
    .warning-title {{
        font-weight: bold;
        color: #ef6c00;
        margin-bottom: 5px;
        text-transform: uppercase;
    }}
    
    /* Upload Box */
    .upload-container {{
        border: 2px dashed #b2dfdb;
        border-radius: 5px;
        padding: 20px;
        text-align: center;
        background-color: #fafafa;
    }}

    /* Step circles */
    .step-circle {{
        display: inline-block;
        width: 25px;
        height: 25px;
        background-color: {PRIMARY_COLOR};
        color: white;
        border-radius: 50%;
        text-align: center;
        line-height: 25px;
        font-weight: bold;
        margin-right: 8px;
    }}
    .step-header {{
        font-size: 16px;
        font-weight: bold;
        color: {PRIMARY_COLOR};
        margin-bottom: 10px;
        display: flex;
        align-items: center;
    }}

    /* Nút bấm */
    .stButton>button {{
        background-color: {PRIMARY_COLOR};
        color: white;
        border: none;
        border-radius: 5px;
        font-weight: bold;
        width: 100%;
        height: 45px;
    }}
    .stButton>button:hover {{
        background-color: #004d40;
        color: white;
    }}
    
    a[target="_blank"] {{ text-decoration: none; }}
    
    .error-msg {{
        background-color: #ffebee;
        color: {TEXT_RED};
        padding: 10px;
        border-radius: 5px;
        border: 1px solid #ffcdd2;
        margin-top: 10px;
        font-size: 14px;
    }}
    
    .footer {{
        text-align: center;
        font-size: 12px;
        color: #757575;
        margin-top: 30px;
        padding-top: 10px;
        border-top: 1px solid #eeeeee;
    }}
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC XỬ LÝ WORD (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text_from_node(node):
    """Lấy text thuần từ XML node"""
    texts = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def check_structure_errors(blocks):
    full_text = "\n".join([get_text_from_node(b) for b in blocks])
    errors = []
    if not re.search(r'Câu\s*1[\.:]', full_text, re.IGNORECASE):
        errors.append("Không tìm thấy 'Câu 1'. File cần bắt đầu bằng Câu 1.")
    return errors

def is_answer_marked(paragraph):
    """Kiểm tra đáp án có được tô đỏ/gạch chân không"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for run in runs:
        rPr = run.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr: continue
        rPr = rPr[0]
        # Check màu đỏ
        colors = rPr.getElementsByTagNameNS(W_NS, "color")
        for c in colors:
            val = c.getAttributeNS(W_NS, "val")
            if val and (val.upper() in ['FF0000', 'RED']): return True
        # Check gạch chân
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

def update_label_in_node(paragraph, new_label):
    """Cập nhật nhãn Câu X. hoặc A. B. C. D."""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    
    first_node = None
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue.strip():
            first_node = t
            break
            
    if not first_node: return

    txt = first_node.firstChild.nodeValue
    # Regex thay thế
    if re.match(r'^\s*[A-D][\.:\)]', txt, re.IGNORECASE):
        sub = re.sub(r'^\s*[A-D][\.:\)]', new_label, txt, count=1)
        first_node.firstChild.nodeValue = sub
    elif re.match(r'^\s*[a-d][\.:\)]', txt, re.IGNORECASE):
        sub = re.sub(r'^\s*[a-d][\.:\)]', new_label, txt, count=1)
        first_node.firstChild.nodeValue = sub
    elif re.match(r'^\s*Câu\s*\d+', txt, re.IGNORECASE):
        sub = re.sub(r'^\s*Câu\s*\d+[\.:]?', new_label, txt, count=1, flags=re.IGNORECASE)
        first_node.firstChild.nodeValue = sub

def extract_part3_answer(question_blocks):
    """
    Tìm và trích xuất đáp án phần 3 (ĐS: ...)
    Đồng thời xóa dòng đáp án đó khỏi block để in đề
    Trả về: (Cleaned Blocks, Answer String)
    """
    answer_text = "X"
    blocks_to_keep = []
    
    found_ans = False
    
    # Duyệt ngược để tìm đáp án (thường ở cuối câu)
    for block in reversed(question_blocks):
        txt = get_text_from_node(block)
        
        # Regex tìm "ĐS: giá_trị" hoặc "DS: giá_trị"
        match = re.search(r'ĐS[:\s]+(.*)', txt, re.IGNORECASE)
        if not match:
            match = re.search(r'DS[:\s]+(.*)', txt, re.IGNORECASE) # Dự phòng trường hợp gõ ko dấu
            
        if match and not found_ans:
            # Kiểm tra xem có tô đỏ không
            if is_answer_marked(block):
                answer_text = match.group(1).strip()
                found_ans = True
                # Không thêm block này vào blocks_to_keep => Tức là XÓA nó khỏi đề thi
                continue 
        
        blocks_to_keep.insert(0, block) # Thêm vào đầu danh sách để giữ đúng thứ tự
        
    return blocks_to_keep, answer_text

def shuffle_questions(questions, mode="MCQ"):
    """Hàm trộn câu hỏi và đáp án"""
    indices = list(range(len(questions)))
    random.shuffle(indices)
    
    shuffled_output = []
    key_map = {} # {local_idx: answer_value}
    
    labels_mcq = ["A.", "B.", "C.", "D."]
    labels_tf = ["a)", "b)", "c)", "d)"]
    
    for new_idx, old_idx in enumerate(indices):
        q_blocks = questions[old_idx] 
        
        # XỬ LÝ RIÊNG CHO PHẦN 3 (Trả lời ngắn)
        if mode == "FILL":
            # Trích xuất đáp án và xóa dòng ĐS khỏi đề
            cleaned_blocks, ans_text = extract_part3_answer(q_blocks)
            
            # Cập nhật nhãn Câu X
            if cleaned_blocks:
                update_label_in_node(cleaned_blocks[0], f"Câu {new_idx + 1}.")
            
            shuffled_output.extend(cleaned_blocks)
            key_map[new_idx + 1] = ans_text
            continue

        # XỬ LÝ CHO MCQ và TF (Phần 1, 2)
        intro = []
        options = []
        
        for b in q_blocks:
            txt = get_text_from_node(b)
            is_opt = False
            if mode == "MCQ" and re.match(r'^\s*[A-D][\.:]', txt): is_opt = True
            elif mode == "TF" and re.match(r'^\s*[a-d][\)]', txt): is_opt = True
            
            if is_opt:
                is_correct = is_answer_marked(b)
                clean_formatting(b)
                options.append({'node': b, 'correct': is_correct})
            else:
                intro.append(b)
                
        correct_char = ""
        if options:
            if mode == "MCQ":
                random.shuffle(options)
                for i, opt in enumerate(options):
                    lbl = labels_mcq[i] if i < 4 else "*"
                    update_label_in_node(opt['node'], lbl)
                    if opt['correct']: correct_char = lbl[0]
            elif mode == "TF":
                random.shuffle(options)
                for i, opt in enumerate(options):
                    lbl = labels_tf[i] if i < 4 else "*"
                    update_label_in_node(opt['node'], lbl)
        
        if intro:
            update_label_in_node(intro[0], f"Câu {new_idx + 1}.")
            
        if mode == "MCQ":
            key_map[new_idx + 1] = correct_char if correct_char else "X"
            
        shuffled_output.extend(intro)
        for o in options: shuffled_output.extend([o['node']])
        
    return shuffled_output, key_map

def process_docx(file_bytes, num_exams, shuffle_mode):
    input_io = io.BytesIO(file_bytes)
    resources = {}
    xml_content = ""
    
    try:
        with zipfile.ZipFile(input_io, 'r') as zin:
            for filename in zin.namelist():
                if filename == "word/document.xml":
                    xml_content = zin.read(filename).decode('utf-8')
                else:
                    resources[filename] = zin.read(filename)
    except Exception as e:
        return None, None, [f"Lỗi đọc file: {str(e)}"]

    if not xml_content:
        return None, None, ["Không tìm thấy nội dung document.xml"]

    # Parse XML gốc
    dom = minidom.parseString(xml_content)
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    
    all_blocks = []
    for child in list(body.childNodes):
        if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]:
            all_blocks.append(child)
            body.removeChild(child) 
            
    errors = check_structure_errors(all_blocks)
    
    # Phân chia PHẦN
    parts = [] 
    current_part = []
    for block in all_blocks:
        txt = get_text_from_node(block)
        if re.match(r'^\s*PHẦN\s*\d+', txt, re.IGNORECASE):
            if current_part: parts.append(current_part)
            current_part = [block]
        else:
            current_part.append(block)
    if current_part: parts.append(current_part)
    
    if not parts: parts = [all_blocks]

    output_zip_io = io.BytesIO()
    excel_data = []
    
    with zipfile.ZipFile(output_zip_io, 'w', zipfile.ZIP_DEFLATED) as zout:
        
        for ver in range(num_exams):
            exam_code = f"10{ver+1}"
            
            curr_dom = minidom.parseString(xml_content)
            curr_body = curr_dom.getElementsByTagNameNS(W_NS, "body")[0]
            while curr_body.firstChild: curr_body.removeChild(curr_body.firstChild)
            
            exam_blocks = []
            exam_key = {"Mã đề": exam_code}
            global_q_idx = 1
            
            for part_blocks in parts:
                questions = []
                intro_part = []
                curr_q = []
                is_q = False
                
                cloned_part_blocks = [b.cloneNode(True) for b in part_blocks]
                
                for b in cloned_part_blocks:
                    txt = get_text_from_node(b)
                    if re.match(r'^\s*Câu\s*\d+', txt, re.IGNORECASE):
                        if curr_q: questions.append(curr_q)
                        curr_q = [b]
                        is_q = True
                    elif re.match(r'^\s*PHẦN', txt, re.IGNORECASE):
                        if curr_q: questions.append(curr_q)
                        curr_q = []
                        intro_part.append(b)
                        is_q = False
                    else:
                        if is_q: curr_q.append(b)
                        else: intro_part.append(b)
                if curr_q: questions.append(curr_q)
                
                # Xác định Mode
                part_txt = get_text_from_node(intro_part[0]) if intro_part else ""
                current_mode = "MCQ"
                if shuffle_mode == "auto":
                    if "PHẦN 2" in part_txt.upper(): current_mode = "TF"
                    elif "PHẦN 3" in part_txt.upper(): current_mode = "FILL" # Mode mới cho Phần 3
                elif shuffle_mode == "tf": current_mode = "TF"
                
                # Thực hiện trộn
                shuffled_q_nodes, key_map = shuffle_questions(questions, mode=current_mode)
                
                final_q_nodes = []
                local_count = 0
                
                # Map đáp án
                for k, v in key_map.items():
                    exam_key[str(global_q_idx + k - 1)] = v
                    
                # Fix lại label Câu X lần cuối cho khớp global index
                for node in shuffled_q_nodes:
                    txt = get_text_from_node(node)
                    if re.match(r'^\s*Câu\s*\d+', txt):
                        update_label_in_node(node, f"Câu {global_q_idx + local_count}.")
                        local_count += 1
                    final_q_nodes.append(node)
                
                global_q_idx += len(key_map)
                exam_blocks.extend(intro_part)
                exam_blocks.extend(final_q_nodes)

            for b in exam_blocks:
                curr_body.appendChild(b)
                
            new_xml = curr_dom.toxml()
            
            sub_io = io.BytesIO()
            with zipfile.ZipFile(sub_io, 'w', zipfile.ZIP_DEFLATED) as sub_z:
                sub_z.writestr("word/document.xml", new_xml.encode('utf-8'))
                for name, content in resources.items():
                    sub_z.writestr(name, content)
            
            zout.writestr(f"De_{exam_code}.docx", sub_io.getvalue())
            excel_data.append(exam_key)
            
    df = pd.DataFrame(excel_data)
    cols = ["Mã đề"] + sorted([c for c in df.columns if c != "Mã đề"], key=lambda x: int(x) if x.isdigit() else 999)
    df = df[cols] if not df.empty else df
    
    excel_io = io.BytesIO()
    with pd.ExcelWriter(excel_io, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='DapAn')
        
    return output_zip_io.getvalue(), excel_io.getvalue(), errors

# ==================== GIAO DIỆN CHÍNH ====================
def main():
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG THPT MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025</p>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        with st.expander("📄 Hướng dẫn & File mẫu (Bấm để xem)", expanded=True):
            st.markdown(f"""
            <div class="info-box">
                <div class="info-title">📌 CẤU TRÚC FILE CHUẨN:</div>
                <p>
                    <span class="part-label">PHẦN 1:</span> Trắc nghiệm nhiều lựa chọn (A. B. C. D.)<br>
                    <i>(Trộn cả câu hỏi và phương án)</i>
                </p>
                <p>
                    <span class="part-label">PHẦN 2:</span> Trắc nghiệm Đúng/Sai (a) b) c) d))<br>
                    <i>(Trộn câu hỏi, trộn ý a,b,c - giữ d cố định)</i>
                </p>
                <p>
                    <span class="part-label">PHẦN 3:</span> Trắc nghiệm trả lời ngắn<br>
                    <i>(Chỉ trộn thứ tự câu hỏi)</i>
                </p>
                
                <div class="warning-box">
                    <div class="warning-title">⚠️ QUY ƯỚC ĐÁP ÁN (BẮT BUỘC):</div>
                    <ul style="margin-bottom: 0; padding-left: 20px;">
                        <li><b>Quy tắc chung:</b> Bắt đầu câu hỏi bằng <code>Câu 1.</code>, <code>Câu 2.</code>...</li>
                        <li>
                            <b>Phần 1 & 2:</b> Đáp án đúng phải <span style="color:red; font-weight:bold">TÔ ĐỎ</span> hoặc <u>GẠCH CHÂN</u>.
                        </li>
                        <li>
                            <b>Phần 3:</b> Ghi <span style="color:red; font-weight:bold">ĐS: Giá trị</span> ở cuối câu và <span style="color:red; font-weight:bold">TÔ ĐỎ</span>.<br>
                            <i>(Ví dụ: <span style="color:red">ĐS: -5</span> hoặc <span style="color:red">ĐS: 10.5</span>)</i>
                        </li>
                    </ul>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            sample_url = "https://docs.google.com/document/d/1mnQqyUqQMRSbhxLDP_E_CswHvHlXziFU/export?format=docx"
            st.link_button("📥 Tải File Mẫu (Google Docs)", sample_url, use_container_width=True)

        st.markdown('<div class="step-header"><span class="step-circle">1</span> Chọn file đề Word (*.docx)</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("", type=["docx"], label_visibility="collapsed")
        
        if not uploaded_file:
            st.markdown("""
            <div class="upload-container">
                <div style="font-size: 20px;">☁️</div>
                <div style="font-weight: bold;">Drag and drop file here</div>
                <div style="color: #999; font-size: 12px;">Limit 200MB per file • DOCX</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.success(f"✅ Đã tải lên: {uploaded_file.name}")

        st.button("● Kiểm tra cấu trúc & lỗi", type="primary")

    with col_right:
        st.markdown('<div class="step-header"><span class="step-circle">2</span> Chọn kiểu trộn</div>', unsafe_allow_html=True)
        shuffle_opt = st.radio(
            "",
            ["Tự động (Phần 1, 2, 3)", "Trắc nghiệm", "Đúng/Sai"],
            index=0,
            label_visibility="collapsed"
        )
        mode_map = {"Tự động (Phần 1, 2, 3)": "auto", "Trắc nghiệm": "mcq", "Đúng/Sai": "tf"}
        selected_mode = mode_map[shuffle_opt]

        st.markdown("---")

        st.markdown('<div class="step-header"><span class="step-circle">3</span> Số mã đề cần tạo</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([2, 3])
        with c1:
            num_exams = st.number_input("", min_value=1, max_value=50, value=4)
        with c2:
            st.markdown("<div style='padding-top: 10px; color:#666; font-size:13px'>● 1 mã → File Word<br>● Nhiều mã → File ZIP</div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("🚀 Trộn đề & Tải xuống", type="primary"):
            if not uploaded_file:
                st.markdown(f'<div class="error-msg">⚠️ Vui lòng chọn file đề Word trước!</div>', unsafe_allow_html=True)
            else:
                with st.spinner("⏳ Đang xử lý..."):
                    try:
                        file_bytes = uploaded_file.read()
                        uploaded_file.seek(0)
                        
                        zip_data, excel_data, errors = process_docx(file_bytes, num_exams, selected_mode)
                        
                        if errors:
                            for e in errors:
                                st.markdown(f'<div class="error-msg">{e}</div>', unsafe_allow_html=True)
                        
                        if zip_data:
                            col_dl1, col_dl2 = st.columns(2)
                            with col_dl1:
                                st.download_button(
                                    label="📦 Tải Đề (ZIP/Docx)",
                                    data=zip_data,
                                    file_name=f"Bo_De_Thi.zip",
                                    mime="application/zip",
                                    type="primary"
                                )
                            with col_dl2:
                                st.download_button(
                                    label="📊 Tải Đáp Án (Excel)",
                                    data=excel_data,
                                    file_name="Dap_An.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            st.balloons()
                            
                    except Exception as e:
                         st.markdown(f'<div class="error-msg">Lỗi xử lý: {str(e)}</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="footer">
        © 2025 Phan Trường Duy - THPT Minh Đức<br>
        Hệ thống quản lý trộn đề thi trắc nghiệm
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
