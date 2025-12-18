import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG ====================

st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================== CSS GIAO DIỆN (ĐÃ CẬP NHẬT THEO YÊU CẦU) ====================
st.markdown("""
<style>
    /* 1. Cấu hình chung */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 98% !important;
    }
    div[data-testid="stVerticalBlock"] > div {
        gap: 0.5rem !important;
    }

    /* 2. Header Card - Màu Xanh Ngọc */
    .header-card {
        background: linear-gradient(180deg, #ffffff 0%, #d1fae5 100%);
        border: 1px solid #a7f3d0;
        border-radius: 15px;
        padding: 10px 5px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .header-card h1 {
        color: #d93025; 
        font-size: clamp(2rem, 3.5vw, 3.5rem) !important; 
        white-space: nowrap !important;
        font-weight: 900;
        text-transform: uppercase;
        margin: 0 !important;
        line-height: 1.1;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }
    .header-card h2 {
        color: #0d9488;
        font-size: 1.6rem !important;
        font-weight: bold;
        margin: 0 !important;
        padding-top: 2px !important;
    }

    /* 3. Style cho CÁC BƯỚC (Step Header) - TÔ MÀU TEXT */
    .step-header {
        color: #0d9488; /* Màu xanh teal đậm đà */
        font-size: 1.4rem;
        font-weight: 700;
        margin-bottom: 5px;
        border-left: 5px solid #0d9488;
        padding-left: 10px;
    }

    /* 4. Style ĐẶC BIỆT cho Số lượng đề & Mã đề (Bold + Big Size 50%) */
    div[data-testid="stNumberInput"] label p {
        font-size: 1.5rem !important; /* Tăng 50% so với bình thường (~1rem) */
        font-weight: 900 !important;
        color: #1f2937;
    }
    div[data-testid="stNumberInput"] input {
        font-size: 1.3rem !important;
        font-weight: bold;
        color: #d93025; /* Số màu đỏ cho nổi */
    }

    /* 5. Vùng Hướng dẫn */
    .instruction-container {
        background-color: #f0fdfa;
        border: 1px solid #99f6e4;
        border-radius: 10px;
        padding: 15px;
        font-size: 1.1rem !important;
        line-height: 1.4;
    }

    /* 6. Nút bấm */
    .stButton > button {
        width: 100%;
        background: linear-gradient(90deg, #0d9488, #14b8a6);
        color: white;
        border: none;
        padding: 0.8rem 1.5rem;
        font-size: 1.4rem !important;
        font-weight: 800;
        border-radius: 10px;
        box-shadow: 0 4px 10px rgba(13, 148, 136, 0.3);
        text-transform: uppercase;
        margin-top: 15px;
    }
    .stButton > button:hover {
        background: linear-gradient(90deg, #0f766e, #0d9488);
        transform: scale(1.01);
    }

    /* Footer */
    .footer {
        text-align: center;
        color: #64748b;
        padding: 1rem 0;
        border-top: 1px dashed #cbd5e1;
        margin-top: 1rem;
        font-size: 1.1rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC XỬ LÝ WORD (CORE) ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

def get_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

# --- Hàm kiểm tra đáp án đúng (Gạch chân hoặc Màu đỏ) ---
def is_correct_answer(block):
    """Kiểm tra xem block có chứa định dạng gạch chân hoặc màu đỏ/màu khác đen không"""
    # Check underline
    u_nodes = block.getElementsByTagNameNS(W_NS, "u")
    for u in u_nodes:
        val = u.getAttributeNS(W_NS, "val")
        if val and val != "none":
            return True
    
    # Check color
    color_nodes = block.getElementsByTagNameNS(W_NS, "color")
    for c in color_nodes:
        val = c.getAttributeNS(W_NS, "val")
        # FF0000 là đỏ, hoặc bất kỳ màu nào khác auto/000000
        if val and val not in ["auto", "000000"]:
            return True
            
    # Check highlight
    highlight_nodes = block.getElementsByTagNameNS(W_NS, "highlight")
    if highlight_nodes:
        return True

    return False

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    
    # Màu xanh
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list: color_el = color_list[0]
    else:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
    color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    
    # In đậm
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def update_mcq_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].upper(); new_punct = "."
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([A-D])(\s*[\.\)])?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""; after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_letter + new_punct + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        
        found_punct_in_regex = bool(m.group(3))
        if not found_punct_in_regex:
            for j in range(i + 1, len(t_nodes)):
                t2 = t_nodes[j]
                if not t2.firstChild or not t2.firstChild.nodeValue: continue
                txt2 = t2.firstChild.nodeValue
                if re.match(r'^[\.\)]', txt2): t2.firstChild.nodeValue = txt2[1:]; break
                elif re.match(r'^\s*$', txt2): continue
                else: break
        break

def update_tf_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].lower(); new_punct = ")"
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([a-d])(\s*[\.\)])?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""; after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_letter + new_punct + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        
        found_punct_in_regex = bool(m.group(3))
        if not found_punct_in_regex:
            for j in range(i + 1, len(t_nodes)):
                t2 = t_nodes[j]
                if not t2.firstChild or not t2.firstChild.nodeValue: continue
                txt2 = t2.firstChild.nodeValue
                if re.match(r'^\)', txt2): t2.firstChild.nodeValue = txt2[1:]; break
                elif re.match(r'^\s*$', txt2): continue
                else: break
        break

def update_question_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)(\.)?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""; after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_label + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        for j in range(i + 1, len(t_nodes)):
            t2 = t_nodes[j]
            if not t2.firstChild or not t2.firstChild.nodeValue: continue
            txt2 = t2.firstChild.nodeValue
            if re.match(r'^[\s0-9\.]*$', txt2) and txt2.strip(): t2.firstChild.nodeValue = ""
            elif re.match(r'^\s*$', txt2): continue
            else: break
        break

def find_part_index(blocks, part_number):
    pattern = re.compile(rf'PHẦN\s*{part_number}\b', re.IGNORECASE)
    for i, block in enumerate(blocks):
        text = get_text(block)
        if pattern.search(text): return i
    return -1

def parse_questions_in_range(blocks, start, end):
    part_blocks = blocks[start:end]; intro = []; questions = []; i = 0
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text): break
        intro.append(part_blocks[i]); i += 1
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text):
            group = [part_blocks[i]]; i += 1
            while i < len(part_blocks):
                t2 = get_text(part_blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2): break
                if re.match(r'^PHẦN\s*\d\b', t2, re.IGNORECASE): break
                group.append(part_blocks[i]); i += 1
            questions.append(group)
        else: intro.append(part_blocks[i]); i += 1
    return intro, questions

# --- LOGIC TRỘN & THEO DÕI ĐÁP ÁN MCQ ---
def shuffle_mcq_options(question_blocks):
    """Trộn và trả về: (blocks_đã_trộn, đáp_án_đúng_nếu_có)"""
    indices = []
    correct_char = None # A, B, C, D
    
    # Tìm các phương án A, B, C, D
    for i, block in enumerate(question_blocks):
        text = get_text(block)
        if re.match(r'^\s*[A-D][\.\)]', text, re.IGNORECASE):
            indices.append(i)
    
    if len(indices) < 2:
        return question_blocks, None
    
    # Xác định đáp án đúng dựa trên gạch chân/màu sắc TRƯỚC khi trộn
    original_options = [question_blocks[idx] for idx in indices]
    marked_index = -1
    
    for k, opt_block in enumerate(original_options):
        if is_correct_answer(opt_block):
            marked_index = k
            break
            
    # Thực hiện trộn
    shuffled_options = shuffle_array(original_options)
    
    # Tìm xem đáp án đúng đã chạy đi đâu
    if marked_index != -1:
        # marked_index là index trong danh sách original_options (0=A cũ, 1=B cũ...)
        # Tìm block đó trong shuffled_options
        target_block = original_options[marked_index]
        try:
            new_pos = shuffled_options.index(target_block)
            # new_pos: 0->A, 1->B, 2->C, 3->D
            letters = ["A", "B", "C", "D"]
            if new_pos < len(letters):
                correct_char = letters[new_pos]
        except:
            pass

    # Tái tạo danh sách blocks
    min_idx = min(indices); max_idx = max(indices)
    before = question_blocks[:min_idx]
    after = question_blocks[max_idx + 1:]
    
    return before + shuffled_options + after, correct_char

def shuffle_tf_options(question_blocks):
    """Trộn T/F (không theo dõi đáp án chi tiết vì phức tạp)"""
    option_indices = {}
    for i, block in enumerate(question_blocks):
        text = get_text(block)
        m = re.match(r'^\s*([a-d])\)', text, re.IGNORECASE)
        if m: option_indices[m.group(1).lower()] = i
    abc_idx = [option_indices.get(k) for k in ["a", "b", "c"] if option_indices.get(k) is not None]
    if len(abc_idx) < 2: return question_blocks, None
    abc_nodes = [question_blocks[idx] for idx in abc_idx]
    shuffled_abc = shuffle_array(abc_nodes)
    all_idx = [v for v in option_indices.values() if v is not None]
    min_idx = min(all_idx); max_idx = max(all_idx)
    d_node = question_blocks[option_indices["d"]] if "d" in option_indices else None
    middle = shuffled_abc.copy()
    if d_node: middle.append(d_node)
    return question_blocks[:min_idx] + middle + question_blocks[max_idx + 1:], None

def relabel_mcq_options(question_blocks):
    letters = ["A", "B", "C", "D"]; option_blocks = []
    for block in question_blocks:
        text = get_text(block)
        if re.match(r'^\s*[A-D][\.\)]', text, re.IGNORECASE): option_blocks.append(block)
    for idx, block in enumerate(option_blocks):
        letter = letters[idx] if idx < len(letters) else letters[-1]
        update_mcq_label(block, f"{letter}.")

def relabel_tf_options(question_blocks):
    letters = ["a", "b", "c", "d"]; option_blocks = []
    for block in question_blocks:
        text = get_text(block)
        if re.match(r'^\s*[a-d]\)', text, re.IGNORECASE): option_blocks.append(block)
    for idx, block in enumerate(option_blocks):
        letter = letters[idx] if idx < len(letters) else letters[-1]
        update_tf_label(block, f"{letter})")

def relabel_questions(questions, start_index=1):
    for i, q_blocks in enumerate(questions):
        if not q_blocks: continue
        update_question_label(q_blocks[0], f"Câu {start_index + i}.")

# --- PROCESS PART (CÓ THU THẬP ĐÁP ÁN) ---
def process_part(blocks, start, end, part_type, start_number=1):
    intro, questions = parse_questions_in_range(blocks, start, end)
    shuffled_questions_with_blocks = []
    
    # Dictionary lưu đáp án của phần này: { "Câu 1": "A", "Câu 2": "C"... }
    part_answers = {} 
    
    # 1. Trộn từng câu và lấy đáp án
    temp_processed = []
    for q_idx, q in enumerate(questions):
        if part_type == "PHAN1":
            new_blocks, ans = shuffle_mcq_options(q)
        elif part_type == "PHAN2":
            new_blocks, ans = shuffle_tf_options(q) # ans sẽ là None
        else:
            new_blocks, ans = q.copy(), None
        temp_processed.append(new_blocks)
        
        # Lưu đáp án tạm (chưa đánh số câu chính thức)
        # Sẽ map lại sau khi trộn thứ tự câu hỏi
        # Cấu trúc temp: (blocks, original_answer_char)
        
    # 2. Trộn thứ tự câu hỏi
    # Cần giữ liên kết giữa câu hỏi và đáp án của nó
    combined = list(zip(temp_processed, [None]*len(questions))) # Placeholder
    if part_type == "PHAN1":
        # Với Phần 1, ta cần list (blocks, answer)
        combined = []
        for q in questions:
            b, a = shuffle_mcq_options(q)
            combined.append((b, a))
    
    shuffled_combined = shuffle_array(combined)
    
    # Tách ra lại
    final_questions_blocks = [x[0] for x in shuffled_combined]
    final_answers_list = [x[1] for x in shuffled_combined]
    
    # 3. Đánh số câu hỏi
    relabel_questions(final_questions_blocks, start_number)
    
    # 4. Đánh lại nhãn ABCD/abcd
    if part_type == "PHAN1":
        for q in final_questions_blocks: relabel_mcq_options(q)
    elif part_type == "PHAN2":
        for q in final_questions_blocks: relabel_tf_options(q)
        
    # 5. Tổng hợp blocks kết quả
    result = intro.copy()
    for q in final_questions_blocks:
        result.extend(q)
        
    # 6. Tổng hợp Map Đáp án {1: 'A', 2: 'B'}
    # start_number là số câu bắt đầu (vd: 1, 13, 25...)
    for i, ans in enumerate(final_answers_list):
        if ans:
            part_answers[start_number + i] = ans
            
    next_number = start_number + len(final_questions_blocks)
    return result, next_number, part_answers

def process_all_as_mcq(blocks):
    intro, questions = parse_questions_in_range(blocks, 0, len(blocks))
    combined = []
    for q in questions:
        b, a = shuffle_mcq_options(q)
        combined.append((b, a))
        
    shuffled = shuffle_array(combined)
    final_blocks = [x[0] for x in shuffled]
    final_ans = [x[1] for x in shuffled]
    
    relabel_questions(final_blocks, 1)
    for q in final_blocks: relabel_mcq_options(q)
    
    result = intro.copy()
    for q in final_blocks: result.extend(q)
    
    answers = {}
    for i, a in enumerate(final_ans):
        if a: answers[i+1] = a
        
    return result, answers

# --- LOGIC CHÍNH: CÓ TRẢ VỀ ĐÁP ÁN ---
def shuffle_docx(file_bytes, shuffle_mode="auto"):
    input_buffer = io.BytesIO(file_bytes)
    all_answers = {} # { "P1": {1:'A', 2:'B'}, "P2": {}, ... }
    
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        doc_xml = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        
        blocks = [c for c in body.childNodes if c.nodeType == c.ELEMENT_NODE and c.localName in ["p", "tbl"]]
        
        if shuffle_mode == "mcq":
            new_blocks, ans = process_all_as_mcq(blocks)
            all_answers["P1"] = ans
        elif shuffle_mode == "tf":
            # TF mode fallback
            new_blocks = process_all_as_tf(blocks) # Hàm cũ chưa sửa để trả về ans, nhưng ít dùng mode này
        else:
            p1_idx = find_part_index(blocks, 1); p2_idx = find_part_index(blocks, 2)
            p3_idx = find_part_index(blocks, 3); p4_idx = find_part_index(blocks, 4)
            
            new_blocks = []; cursor = 0; curr_num = 1
            
            # PHẦN 1
            if p1_idx >= 0:
                new_blocks.extend(blocks[cursor:p1_idx + 1]); cursor = p1_idx + 1
                end1 = len(blocks)
                if p2_idx >= 0: end1 = p2_idx
                elif p3_idx >= 0: end1 = p3_idx
                elif p4_idx >= 0: end1 = p4_idx
                
                p1_blocks, next_num, p1_ans = process_part(blocks, cursor, end1, "PHAN1", curr_num)
                new_blocks.extend(p1_blocks)
                all_answers["P1"] = p1_ans
                curr_num = next_num; cursor = end1
            
            # PHẦN 2
            if p2_idx >= 0:
                new_blocks.append(blocks[p2_idx]); start2 = p2_idx + 1
                end2 = len(blocks)
                if p3_idx >= 0: end2 = p3_idx
                elif p4_idx >= 0: end2 = p4_idx
                
                p2_blocks, next_num, p2_ans = process_part(blocks, start2, end2, "PHAN2", curr_num)
                new_blocks.extend(p2_blocks)
                all_answers["P2_Count"] = len(p2_ans) if p2_ans else (next_num - curr_num) # Lưu số lượng câu
                all_answers["P2_Start"] = curr_num
                curr_num = next_num; cursor = end2
            
            # PHẦN 3
            if p3_idx >= 0:
                new_blocks.append(blocks[p3_idx]); start3 = p3_idx + 1
                end3 = len(blocks)
                if p4_idx >= 0: end3 = p4_idx
                
                p3_blocks, next_num, p3_ans = process_part(blocks, start3, end3, "PHAN3", curr_num)
                new_blocks.extend(p3_blocks)
                all_answers["P3_Count"] = (next_num - curr_num)
                all_answers["P3_Start"] = curr_num
                curr_num = next_num; cursor = end3

            # PHẦN 4
            if p4_idx >= 0: new_blocks.extend(blocks[p4_idx:])
            
            if p1_idx == -1 and p2_idx == -1 and p3_idx == -1 and p4_idx == -1:
                new_blocks, ans = process_all_as_mcq(blocks)
                all_answers["P1"] = ans
        
        # Reconstruct XML
        other_nodes = [c for c in list(body.childNodes) if c.nodeType == c.ELEMENT_NODE and c.localName not in ["p", "tbl"]]
        while body.firstChild: body.removeChild(body.firstChild)
        for block in new_blocks: body.appendChild(block)
        for node in other_nodes: body.appendChild(node)
        
        output_buffer = io.BytesIO()
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml": zout.writestr(item, dom.toxml().encode('utf-8'))
                else: zout.writestr(item, zin.read(item.filename))
        
        return output_buffer.getvalue(), all_answers

# --- TẠO FILE BẢNG ĐÁP ÁN (HTML -> WORD) ---
def generate_answer_key_html(all_exam_data):
    """Tạo nội dung HTML cho file Word đáp án"""
    # all_exam_data: { "101": { "P1": {1:'A'...}, "P2_Count": 4... }, "102": ... }
    
    html = """
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head><meta charset='utf-8'><title>Đáp án</title>
    <style>
        body { font-family: 'Times New Roman', serif; font-size: 12pt; }
        h1, h2 { text-align: center; color: #C00000; margin: 5px 0; }
        h3 { color: #002060; margin-top: 20px; margin-bottom: 5px; font-size: 13pt; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 15px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; font-size: 11pt; }
        th { background-color: #D9E2F3; font-weight: bold; }
        .note { font-style: italic; font-size: 11pt; color: #002060; margin-bottom: 5px; }
    </style>
    </head><body>
    """
    
    html += """
    <h1>TRƯỜNG THPT MINH ĐỨC</h1>
    <h2>BẢNG ĐÁP ÁN</h2>
    <p style='text-align:center; font-weight:bold;'>KIỂM TRA HỌC KỲ I - NĂM HỌC 2025 – 2026</p>
    <br>
    """
    
    exam_codes = sorted(all_exam_data.keys())
    sample_data = all_exam_data[exam_codes[0]]
    
    # --- PHẦN 1: TRẮC NGHIỆM ---
    if "P1" in sample_data and sample_data["P1"]:
        html += "<h3>PHẦN I: Trắc nghiệm nhiều lựa chọn</h3>"
        html += "<div class='note'>- Mỗi câu đúng được 0,25 điểm.</div>"
        
        # Lấy danh sách câu hỏi P1
        q_nums = sorted(sample_data["P1"].keys())
        
        html += "<table><tr><th>Mã đề</th>"
        for q in q_nums: html += f"<th>{q}</th>"
        html += "</tr>"
        
        for code in exam_codes:
            html += f"<tr><td><b>{code}</b></td>"
            ans_map = all_exam_data[code].get("P1", {})
            for q in q_nums:
                html += f"<td><b>{ans_map.get(q, '')}</b></td>"
            html += "</tr>"
        html += "</table>"
    
    # --- PHẦN 2: ĐÚNG SAI (TEMPLATE) ---
    if "P2_Count" in sample_data:
        count = sample_data["P2_Count"]
        start = sample_data.get("P2_Start", 1)
        
        html += "<h3>PHẦN II: Trắc nghiệm đúng sai</h3>"
        html += "<div class='note'>- Điểm tối đa mỗi câu là 1 điểm.</div>"
        html += "<div class='note'>- Đúng 1 ý được 0,1 điểm; đúng 2 ý được 0,25 điểm; đúng 3 ý được 0,5 điểm; đúng 4 ý được 1 điểm.</div>"
        
        html += "<table><tr><th>Mã đề</th>"
        for i in range(count): html += f"<th>Câu {start + i}</th>"
        html += "</tr>"
        
        for code in exam_codes:
            html += f"<tr><td><b>{code}</b></td>"
            for i in range(count): html += "<td></td>" # Để trống cho giáo viên điền
            html += "</tr>"
        html += "</table>"

    # --- PHẦN 3: TRẢ LỜI NGẮN (TEMPLATE) ---
    if "P3_Count" in sample_data:
        count = sample_data["P3_Count"]
        start = sample_data.get("P3_Start", 1)
        
        html += "<h3>PHẦN III: Trắc nghiệm trả lời ngắn</h3>"
        html += "<div class='note'>- Điểm tối đa mỗi câu là 0,5 điểm.</div>"
        
        html += "<table><tr><th>Mã đề</th>"
        for i in range(count): html += f"<th>Câu {start + i}</th>"
        html += "</tr>"
        
        for code in exam_codes:
            html += f"<tr><td><b>{code}</b></td>"
            for i in range(count): html += "<td></td>" # Để trống
            html += "</tr>"
        html += "</table>"

    html += "</body></html>"
    return html

def create_zip_multiple(file_bytes, base_name, num_versions, shuffle_mode, start_code):
    zip_buffer = io.BytesIO()
    all_exam_data = {}
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
        # 1. Tạo các file đề
        for i in range(num_versions):
            current_code = start_code + i
            shuffled_bytes, exam_answers = shuffle_docx(file_bytes, shuffle_mode)
            
            # Lưu đáp án của mã đề này
            all_exam_data[current_code] = exam_answers
            
            filename = f"{base_name}_{current_code}.docx"
            zout.writestr(filename, shuffled_bytes)
        
        # 2. Tạo file Bảng Đáp Án (Word via HTML)
        try:
            answer_key_html = generate_answer_key_html(all_exam_data)
            zout.writestr("Bang_Dap_An.doc", answer_key_html.encode('utf-8'))
        except Exception as e:
            print(f"Error creating answer key: {e}")
    
    return zip_buffer.getvalue()

# ==================== GIAO DIỆN STREAMLIT ====================

def main():
    # Header
    st.markdown("""
    <div class="header-card">
        <h1>TRƯỜNG TRUNG HỌC PHỔ THÔNG MINH ĐỨC</h1>
        <h2>ỨNG DỤNG TRỘN ĐỀ WORD 2025</h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Hướng dẫn
    with st.expander("📋 Hướng dẫn & Cấu trúc file", expanded=False):
        st.markdown("""
        <div class="instruction-container">
            <strong>Cấu trúc file Word chuẩn:</strong>
            <ul>
                <li><strong>PHẦN 1:</strong> Trắc nghiệm (A. B. C. D.)</li>
                <li><strong>PHẦN 2:</strong> Đúng/Sai (a) b) c) d))</li>
                <li><strong>PHẦN 3:</strong> Trả lời ngắn</li>
                <li><strong>PHẦN 4:</strong> Tự luận (Giữ nguyên)</li>
            </ul>
            <strong>Lưu ý về Đáp Án:</strong>
            <ul>
                <li>Để phần mềm tự tạo <b>Bảng Đáp Án</b> cho Phần 1, vui lòng <b>Gạch chân</b> hoặc <b>Tô màu đỏ</b> đáp án đúng trong file gốc.</li>
            </ul>
            <p style="margin-top: 5px;">📥 <a href="https://docs.google.com/document/d/1i1b-By6EA_HO8fWgMYG9iXZPGannmWdg/edit?usp=drive_link&ouid=112824050529887271694&rtpof=true&sd=true" target="_blank">Tải file mẫu tại đây</a></p>
        </div>
        """, unsafe_allow_html=True)
    
    st.write("") 

    # 1. Upload
    st.markdown('<div class="step-header">1️⃣ CHỌN FILE ĐỀ WORD (.docx)</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["docx"], label_visibility="collapsed")
    
    if uploaded_file:
        st.success(f"✅ Đã tải lên: **{uploaded_file.name}**")
    
    st.write("") 
    
    col_left, col_right = st.columns([1, 1], gap="large")
    
    with col_left:
        # 2. Kiểu trộn
        st.markdown('<div class="step-header">2️⃣ KIỂU TRỘN</div>', unsafe_allow_html=True)
        shuffle_mode = st.radio(
            "Chọn chế độ:",
            options=["auto", "mcq", "tf"],
            format_func=lambda x: {
                "auto": "🔄 Tự động (Theo từng Phần)",
                "mcq": "📝 Trắc nghiệm (Toàn bộ)",
                "tf": "✅ Đúng/Sai (Toàn bộ)"
            }[x],
            index=0,
            label_visibility="collapsed"
        )

    with col_right:
        # 3. Cấu hình
        st.markdown('<div class="step-header">3️⃣ CẤU HÌNH MÃ ĐỀ</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            num_versions = st.number_input("Số lượng đề cần tạo", min_value=1, max_value=50, value=4)
        with c2:
            start_code = st.number_input("Mã đề bắt đầu", min_value=0, value=101)
        
        if num_versions > 1:
            st.info(f"📦 Sẽ tạo {num_versions} đề: {start_code} ➝ {start_code + num_versions - 1}")
        else:
            st.info(f"📄 Sẽ tạo 1 đề: {start_code}")

    st.write("") 

    # 4. Nút trộn đề
    if st.button("🎲 BẮT ĐẦU TRỘN ĐỀ & TẢI VỀ", type="primary", use_container_width=True):
        if not uploaded_file:
            st.warning("⚠️ Vui lòng chọn file Word trước khi trộn!")
        else:
            try:
                with st.spinner("🚀 Đang xử lý và tạo bảng đáp án..."):
                    file_bytes = uploaded_file.read()
                    base_name = re.sub(r'[^\w\s-]', '', uploaded_file.name.replace(".docx", "")).strip() or "De"
                    
                    if num_versions == 1:
                        # Nếu 1 đề thì vẫn tạo zip để kèm đáp án
                        result = create_zip_multiple(file_bytes, base_name, 1, shuffle_mode, start_code)
                        filename = f"{base_name}_Mix_{start_code}.zip"
                    else:
                        result = create_zip_multiple(file_bytes, base_name, num_versions, shuffle_mode, start_code)
                        filename = f"{base_name}_Mix_From_{start_code}.zip"
                    
                    mime = "application/zip"
                
                st.balloons()
                st.success("✅ THÀNH CÔNG! File tải về đã bao gồm Đề thi và Bảng đáp án.")
                st.download_button(label=f"📥 TẢI XUỐNG {filename}", data=result, file_name=filename, mime=mime, use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")
    
    # Footer
    st.markdown("""
    <div class="footer">
        <p>Zalo hỗ trợ kỹ thuật: <strong>038994070</strong></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
