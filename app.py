import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG (ĐÃ CHỈNH SANG WIDE) ====================

st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🏫",
    layout="wide",  # <--- THAY ĐỔI QUAN TRỌNG: Sử dụng toàn bộ chiều ngang
    initial_sidebar_state="collapsed"
)

# Custom CSS - PHIÊN BẢN BIG SIZE CHO MÀN HÌNH RỘNG
st.markdown("""
<style>
    /* 1. Cấu hình chung cho layout Wide */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        max-width: 95% !important; /* Sử dụng 95% chiều ngang màn hình */
    }

    /* 2. Header Card - Phóng to cực đại */
    .header-card {
        background: linear-gradient(180deg, #ffffff 0%, #e0f2fe 100%);
        border: 2px solid #brent;
        border-radius: 20px;
        padding: 30px 20px; /* Tăng padding */
        text-align: center;
        box-shadow: 0 6px 10px rgba(0,0,0,0.1);
        margin-bottom: 25px;
    }
    
    .header-card h1 {
        color: #d93025; 
        font-size: 4rem !important; /* Tăng size chữ tên trường cực lớn */
        font-weight: 900;
        text-transform: uppercase;
        margin: 0 !important;
        line-height: 1.2;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.15);
    }
    
    .header-card h2 {
        color: #0d9488;
        font-size: 2.2rem !important; /* Tăng size tên ứng dụng */
        font-weight: bold;
        margin: 10px 0 0 0 !important;
    }

    /* 3. Vùng Hướng dẫn - Chữ to dễ đọc */
    .instruction-container {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        border-radius: 15px;
        padding: 25px;
        font-size: 1.25rem !important; /* Chữ hướng dẫn to hơn (20px) */
        line-height: 1.6;
    }
    .instruction-container strong {
        color: #0369a1;
        font-weight: 700;
    }
    .instruction-container li {
        margin-bottom: 8px;
    }

    /* 4. Tăng kích thước các thành phần Input (Radio, Number, File Uploader) */
    
    /* Chữ nhãn (Label) của các ô nhập liệu */
    .stMarkdown p, .stRadio label, .stNumberInput label, .stFileUploader label {
        font-size: 1.3rem !important; /* Tăng size nhãn */
        font-weight: 600;
    }
    
    /* Text bên trong Radio button */
    .stRadio div[role="radiogroup"] label p {
        font-size: 1.2rem !important;
    }

    /* Ô nhập số (Number Input) to hơn */
    div[data-baseweb="input"] {
        height: 3rem; /* Cao hơn */
        border-radius: 10px;
    }
    div[data-baseweb="input"] input {
        font-size: 1.2rem !important; /* Số bên trong to hơn */
        text-align: center;
    }

    /* 5. Nút bấm (Button) - To và nổi bật */
    .stButton > button {
        width: 100%;
        background: linear-gradient(90deg, #0d9488, #14b8a6);
        color: white;
        border: none;
        padding: 1rem 2rem; /* Nút dày hơn */
        font-size: 1.5rem !important; /* Chữ trong nút to */
        font-weight: 800;
        border-radius: 12px;
        box-shadow: 0 4px 10px rgba(13, 148, 136, 0.3);
        text-transform: uppercase;
        margin-top: 10px;
    }
    .stButton > button:hover {
        background: linear-gradient(90deg, #0f766e, #0d9488);
        transform: scale(1.01);
    }
    
    /* Thông báo lỗi/thành công to hơn */
    .stAlert {
        font-size: 1.1rem !important;
    }

    /* Footer */
    .footer {
        text-align: center;
        color: #64748b;
        padding: 2rem 0;
        border-top: 2px dashed #cbd5e1;
        margin-top: 3rem;
        font-size: 1.1rem; /* Footer to hơn */
    }
    .footer strong {
        color: #0d9488;
        font-size: 1.2rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC TRỘN ĐỀ (GIỮ NGUYÊN 100%) ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def shuffle_array(arr):
    """Fisher-Yates shuffle"""
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out


def get_text(block):
    """Lấy text từ một block"""
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()


def style_run_blue_bold(run):
    """Tô xanh đậm một run"""
    doc = run.ownerDocument
    
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list:
        rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list:
        color_el = color_list[0]
    else:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
    color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)


def update_mcq_label(paragraph, new_label):
    """Cập nhật nhãn A. B. C. D."""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes:
        return
    
    new_letter = new_label[0].upper()
    new_punct = "."
    
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue:
            continue
        
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([A-D])([\.\)])?', txt, re.IGNORECASE)
        if not m:
            continue
        
        leading_space = m.group(1) or ""
        old_punct = m.group(3) or ""
        after_match = txt[m.end():]
        
        if old_punct:
            t.firstChild.nodeValue = leading_space + new_letter + new_punct + after_match
        else:
            t.firstChild.nodeValue = leading_space + new_letter + after_match
            
            found_punct = False
            for j in range(i + 1, len(t_nodes)):
                t2 = t_nodes[j]
                if not t2.firstChild or not t2.firstChild.nodeValue:
                    continue
                txt2 = t2.firstChild.nodeValue
                if re.match(r'^[\.\)]', txt2):
                    t2.firstChild.nodeValue = new_punct + txt2[1:]
                    found_punct = True
                    break
                elif re.match(r'^\s*$', txt2):
                    continue
                else:
                    break
            
            if not found_punct:
                t.firstChild.nodeValue = leading_space + new_letter + new_punct + after_match
        
        run = t.parentNode
        if run and run.localName == "r":
            style_run_blue_bold(run)
        break


def update_tf_label(paragraph, new_label):
    """Cập nhật nhãn a) b) c) d)"""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes:
        return
    
    new_letter = new_label[0].lower()
    new_punct = ")"
    
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue:
            continue
        
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([a-d])(\))?', txt, re.IGNORECASE)
        if not m:
            continue
        
        leading_space = m.group(1) or ""
        old_punct = m.group(3) or ""
        after_match = txt[m.end():]
        
        if old_punct:
            t.firstChild.nodeValue = leading_space + new_letter + new_punct + after_match
        else:
            t.firstChild.nodeValue = leading_space + new_letter + after_match
            
            found_punct = False
            for j in range(i + 1, len(t_nodes)):
                t2 = t_nodes[j]
                if not t2.firstChild or not t2.firstChild.nodeValue:
                    continue
                txt2 = t2.firstChild.nodeValue
                if re.match(r'^\)', txt2):
                    found_punct = True
                    break
                elif re.match(r'^\s*$', txt2):
                    continue
                else:
                    break
            
            if not found_punct:
                t.firstChild.nodeValue = leading_space + new_letter + new_punct + after_match
        
        run = t.parentNode
        if run and run.localName == "r":
            style_run_blue_bold(run)
        break


def update_question_label(paragraph, new_label):
    """Cập nhật nhãn Câu X."""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes:
        return
    
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue:
            continue
        
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)(\.)?', txt, re.IGNORECASE)
        if not m:
            continue
        
        leading_space = m.group(1) or ""
        after_match = txt[m.end():]
        
        t.firstChild.nodeValue = leading_space + new_label + after_match
        
        run = t.parentNode
        if run and run.localName == "r":
            style_run_blue_bold(run)
        
        for j in range(i + 1, len(t_nodes)):
            t2 = t_nodes[j]
            if not t2.firstChild or not t2.firstChild.nodeValue:
                continue
            txt2 = t2.firstChild.nodeValue
            if re.match(r'^[\s0-9\.]*$', txt2) and txt2.strip():
                t2.firstChild.nodeValue = ""
            elif re.match(r'^\s*$', txt2):
                continue
            else:
                break
        break


def find_part_index(blocks, part_number):
    """Tìm dòng PHẦN n"""
    pattern = re.compile(rf'PHẦN\s*{part_number}\b', re.IGNORECASE)
    for i, block in enumerate(blocks):
        text = get_text(block)
        if pattern.search(text):
            return i
    return -1


def parse_questions_in_range(blocks, start, end):
    """Tách câu hỏi trong phạm vi"""
    part_blocks = blocks[start:end]
    intro = []
    questions = []
    
    i = 0
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text):
            break
        intro.append(part_blocks[i])
        i += 1
    
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text):
            group = [part_blocks[i]]
            i += 1
            while i < len(part_blocks):
                t2 = get_text(part_blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2):
                    break
                if re.match(r'^PHẦN\s*\d\b', t2, re.IGNORECASE):
                    break
                group.append(part_blocks[i])
                i += 1
            questions.append(group)
        else:
            intro.append(part_blocks[i])
            i += 1
    
    return intro, questions


def shuffle_mcq_options(question_blocks):
    """Trộn phương án A B C D"""
    indices = []
    for i, block in enumerate(question_blocks):
        text = get_text(block)
        if re.match(r'^\s*[A-D][\.\)]', text, re.IGNORECASE):
            indices.append(i)
    
    if len(indices) < 2:
        return question_blocks
    
    options = [question_blocks[idx] for idx in indices]
    shuffled = shuffle_array(options)
    
    min_idx = min(indices)
    max_idx = max(indices)
    before = question_blocks[:min_idx]
    after = question_blocks[max_idx + 1:]
    
    return before + shuffled + after


def relabel_mcq_options(question_blocks):
    """Đánh lại nhãn A B C D"""
    letters = ["A", "B", "C", "D"]
    option_blocks = []
    
    for block in question_blocks:
        text = get_text(block)
        if re.match(r'^\s*[A-D][\.\)]', text, re.IGNORECASE):
            option_blocks.append(block)
    
    for idx, block in enumerate(option_blocks):
        letter = letters[idx] if idx < len(letters) else letters[-1]
        update_mcq_label(block, f"{letter}.")


def shuffle_tf_options(question_blocks):
    """Trộn phương án a b c (giữ d cố định)"""
    option_indices = {}
    
    for i, block in enumerate(question_blocks):
        text = get_text(block)
        m = re.match(r'^\s*([a-d])\)', text, re.IGNORECASE)
        if m:
            option_indices[m.group(1).lower()] = i
    
    abc_idx = [option_indices.get(k) for k in ["a", "b", "c"] if option_indices.get(k) is not None]
    
    if len(abc_idx) < 2:
        return question_blocks
    
    abc_nodes = [question_blocks[idx] for idx in abc_idx]
    shuffled_abc = shuffle_array(abc_nodes)
    
    all_idx = [v for v in option_indices.values() if v is not None]
    min_idx = min(all_idx)
    max_idx = max(all_idx)
    
    before = question_blocks[:min_idx]
    after = question_blocks[max_idx + 1:]
    
    d_node = question_blocks[option_indices["d"]] if "d" in option_indices else None
    
    middle = shuffled_abc.copy()
    if d_node:
        middle.append(d_node)
    
    return before + middle + after


def relabel_tf_options(question_blocks):
    """Đánh lại nhãn a b c d"""
    letters = ["a", "b", "c", "d"]
    option_blocks = []
    
    for block in question_blocks:
        text = get_text(block)
        if re.match(r'^\s*[a-d]\)', text, re.IGNORECASE):
            option_blocks.append(block)
    
    for idx, block in enumerate(option_blocks):
        letter = letters[idx] if idx < len(letters) else letters[-1]
        update_tf_label(block, f"{letter})")


def relabel_questions(questions):
    """Đánh lại số câu 1, 2, 3..."""
    for i, q_blocks in enumerate(questions):
        if not q_blocks:
            continue
        first_block = q_blocks[0]
        update_question_label(first_block, f"Câu {i + 1}.")


def process_part(blocks, start, end, part_type):
    """Xử lý một PHẦN"""
    intro, questions = parse_questions_in_range(blocks, start, end)
    
    if part_type == "PHAN1":
        processed_questions = [shuffle_mcq_options(q) for q in questions]
    elif part_type == "PHAN2":
        processed_questions = [shuffle_tf_options(q) for q in questions]
    else:
        processed_questions = [q.copy() for q in questions]
    
    shuffled_questions = shuffle_array(processed_questions)
    relabel_questions(shuffled_questions)
    
    if part_type == "PHAN1":
        for q in shuffled_questions:
            relabel_mcq_options(q)
    elif part_type == "PHAN2":
        for q in shuffled_questions:
            relabel_tf_options(q)
    
    result = intro.copy()
    for q in shuffled_questions:
        result.extend(q)
    
    return result


def process_all_as_mcq(blocks):
    """Xử lý toàn bộ như trắc nghiệm MCQ"""
    intro, questions = parse_questions_in_range(blocks, 0, len(blocks))
    
    processed_questions = [shuffle_mcq_options(q) for q in questions]
    shuffled_questions = shuffle_array(processed_questions)
    relabel_questions(shuffled_questions)
    
    for q in shuffled_questions:
        relabel_mcq_options(q)
    
    result = intro.copy()
    for q in shuffled_questions:
        result.extend(q)
    
    return result


def process_all_as_tf(blocks):
    """Xử lý toàn bộ như đúng/sai"""
    intro, questions = parse_questions_in_range(blocks, 0, len(blocks))
    
    processed_questions = [shuffle_tf_options(q) for q in questions]
    shuffled_questions = shuffle_array(processed_questions)
    relabel_questions(shuffled_questions)
    
    for q in shuffled_questions:
        relabel_tf_options(q)
    
    result = intro.copy()
    for q in shuffled_questions:
        result.extend(q)
    
    return result


def shuffle_docx(file_bytes, shuffle_mode="auto"):
    """Trộn file DOCX, trả về bytes"""
    input_buffer = io.BytesIO(file_bytes)
    
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        doc_xml = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        
        body_list = dom.getElementsByTagNameNS(W_NS, "body")
        if not body_list:
            raise Exception("Không tìm thấy w:body trong document.xml")
        body = body_list[0]
        
        blocks = []
        for child in body.childNodes:
            if child.nodeType == child.ELEMENT_NODE:
                if child.localName in ["p", "tbl"]:
                    blocks.append(child)
        
        if shuffle_mode == "mcq":
            new_blocks = process_all_as_mcq(blocks)
        elif shuffle_mode == "tf":
            new_blocks = process_all_as_tf(blocks)
        else:
            part1_idx = find_part_index(blocks, 1)
            part2_idx = find_part_index(blocks, 2)
            part3_idx = find_part_index(blocks, 3)
            
            new_blocks = []
            cursor = 0
            
            if part1_idx >= 0:
                new_blocks.extend(blocks[cursor:part1_idx + 1])
                cursor = part1_idx + 1
                
                end1 = part2_idx if part2_idx >= 0 else len(blocks)
                part1_processed = process_part(blocks, cursor, end1, "PHAN1")
                new_blocks.extend(part1_processed)
                cursor = end1
            
            if part2_idx >= 0:
                new_blocks.append(blocks[part2_idx])
                start2 = part2_idx + 1
                end2 = part3_idx if part3_idx >= 0 else len(blocks)
                part2_processed = process_part(blocks, start2, end2, "PHAN2")
                new_blocks.extend(part2_processed)
                cursor = end2
            
            if part3_idx >= 0:
                new_blocks.append(blocks[part3_idx])
                start3 = part3_idx + 1
                end3 = len(blocks)
                part3_processed = process_part(blocks, start3, end3, "PHAN3")
                new_blocks.extend(part3_processed)
                cursor = end3
            
            if part1_idx == -1 and part2_idx == -1 and part3_idx == -1:
                new_blocks = process_all_as_mcq(blocks)
        
        other_nodes = []
        for child in list(body.childNodes):
            if child.nodeType == child.ELEMENT_NODE:
                if child.localName not in ["p", "tbl"]:
                    other_nodes.append(child)
            body.removeChild(child)
        
        for block in new_blocks:
            body.appendChild(block)
        
        for node in other_nodes:
            body.appendChild(node)
        
        new_xml = dom.toxml()
        
        output_buffer = io.BytesIO()
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_xml.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
        
        return output_buffer.getvalue()


def create_zip_multiple(file_bytes, base_name, num_versions, shuffle_mode, start_code):
    """Tạo ZIP chứa nhiều mã đề với mã bắt đầu tùy chỉnh"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
        for i in range(num_versions):
            shuffled = shuffle_docx(file_bytes, shuffle_mode)
            # Tên file theo mã đề bắt đầu
            current_code = start_code + i
            filename = f"{base_name}_{current_code}.docx"
            zout.writestr(filename, shuffled)
    
    return zip_buffer.getvalue()


# ==================== GIAO DIỆN STREAMLIT (FULL WIDTH) ====================

def main():
    # Header MỚI - FULL WIDTH & BIG SIZE
    st.markdown("""
    <div class="header-card">
        <h1>TRƯỜNG TRUNG HỌC PHỔ THÔNG MINH ĐỨC</h1>
        <h2>ỨNG DỤNG TRỘN ĐỀ WORD 2025</h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Hướng dẫn - BIG TEXT
    with st.expander("📋 Hướng dẫn & Cấu trúc file (Nhấn để xem)", expanded=False):
        st.markdown("""
        <div class="instruction-container">
            <strong>Cấu trúc file Word chuẩn:</strong>
            <ul>
                <li><strong>PHẦN 1:</strong> Trắc nghiệm (A. B. C. D.) – Trộn câu hỏi + phương án</li>
                <li><strong>PHẦN 2:</strong> Đúng/Sai (a) b) c) d)) – Trộn câu hỏi + trộn a,b,c (giữ d cố định)</li>
                <li><strong>PHẦN 3:</strong> Trả lời ngắn – Chỉ trộn thứ tự câu hỏi</li>
            </ul>
            <strong>Quy tắc & Đáp án:</strong>
            <ul>
                <li>Mỗi câu bắt đầu bằng <code>Câu 1.</code>, <code>Câu 2.</code>...</li>
                <li>Phương án MCQ: <code>A.</code> <code>B.</code> <code>C.</code> <code>D.</code> (viết hoa + dấu chấm)</li>
                <li>Phương án Đúng/Sai: <code>a)</code> <code>b)</code> <code>c)</code> <code>d)</code> (viết thường + dấu ngoặc)</li>
                <li><strong>Đáp án Phần 1 & 2:</strong> Gạch chân hoặc tô màu đáp án đúng.</li>
                <li><strong>Đáp án Phần 3:</strong> Ghi đáp án/lời giải ngay sau câu hỏi. Có thể tô màu/gạch chân để làm nổi bật.</li>
            </ul>
            <p style="margin-top: 15px;">📥 <a href="https://docs.google.com/document/d/1i1b-By6EA_HO8fWgMYG9iXZPGannmWdg/edit?usp=drive_link&ouid=112824050529887271694&rtpof=true&sd=true" target="_blank">Tải file mẫu tại đây</a></p>
        </div>
        """, unsafe_allow_html=True)
    
    st.write("") # Spacer

    # 1. Upload file (Dùng cột để căn chỉnh nếu cần, nhưng ở wide mode để mặc định cũng tốt)
    st.subheader("1️⃣ Chọn file đề Word (.docx)")
    uploaded_file = st.file_uploader("", type=["docx"])
    
    if uploaded_file:
        st.success(f"✅ Đã chọn: **{uploaded_file.name}**")
    
    st.divider()
    
    # 2. Cấu hình (Chia cột cho thoáng)
    col_left, col_right = st.columns([1, 1], gap="large")
    
    with col_left:
        st.subheader("2️⃣ Kiểu trộn")
        shuffle_mode = st.radio(
            "Chọn chế độ:",
            options=["auto", "mcq", "tf"],
            format_func=lambda x: {
                "auto": "🔄 Tự động (Theo từng Phần)",
                "mcq": "📝 Trắc nghiệm (Toàn bộ)",
                "tf": "✅ Đúng/Sai (Toàn bộ)"
            }[x],
            index=0
        )

    with col_right:
        st.subheader("3️⃣ Cấu hình mã đề")
        c1, c2 = st.columns(2)
        with c1:
            num_versions = st.number_input("Số lượng đề:", min_value=1, max_value=50, value=4)
        with c2:
            start_code = st.number_input("Mã bắt đầu:", min_value=0, value=101)
        
        if num_versions > 1:
            st.info(f"📦 Tạo {num_versions} đề: Từ mã {start_code} ➝ {start_code + num_versions - 1}")
        else:
            st.info(f"📄 Tạo 1 đề: Mã {start_code}")

    st.write("") # Spacer

    # 4. Nút trộn đề - TO VÀ RÕ
    if st.button("🎲 BẮT ĐẦU TRỘN ĐỀ", type="primary", use_container_width=True):
        if not uploaded_file:
            st.warning("⚠️ Vui lòng chọn file Word trước khi trộn!")
        else:
            try:
                with st.spinner("🚀 Đang xử lý dữ liệu..."):
                    file_bytes = uploaded_file.read()
                    base_name = re.sub(r'[^\w\s-]', '', uploaded_file.name.replace(".docx", "")).strip() or "De"
                    
                    if num_versions == 1:
                        result = shuffle_docx(file_bytes, shuffle_mode)
                        filename = f"{base_name}_{start_code}.docx"
                        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    else:
                        result = create_zip_multiple(file_bytes, base_name, num_versions, shuffle_mode, start_code)
                        filename = f"{base_name}_From_{start_code}.zip"
                        mime = "application/zip"
                
                st.balloons() # Hiệu ứng chúc mừng
                st.success("✅ TRỘN ĐỀ THÀNH CÔNG! BẤM NÚT DƯỚI ĐỂ TẢI.")
                st.download_button(label=f"📥 TẢI XUỐNG {filename}", data=result, file_name=filename, mime=mime, use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")
    
    # Footer
    st.markdown("""
    <div class="footer">
        <p><strong>TRƯỜNG TRUNG HỌC PHỔ THÔNG MINH ĐỨC</strong></p>
        <p>Zalo hỗ trợ kỹ thuật: <strong>038994070</strong></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
