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

# ==================== CSS GIAO DIỆN (GIỮ NGUYÊN) ====================
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
    .step-header {
        color: #0d9488;
        font-size: 1.4rem;
        font-weight: 700;
        margin-bottom: 5px;
        border-left: 5px solid #0d9488;
        padding-left: 10px;
        background-color: #f0fdfa;
        border-radius: 0 5px 5px 0;
    }
    div[data-testid="stNumberInput"] label p {
        font-size: 1.5rem !important;
        font-weight: 900 !important;
        color: #0d9488;
    }
    div[data-testid="stNumberInput"] input {
        font-size: 1.5rem !important;
        font-weight: bold;
        color: #d93025; 
        height: 3rem;
    }
    .instruction-container {
        background-color: #f0fdfa;
        border: 1px solid #99f6e4;
        border-radius: 10px;
        padding: 15px;
        font-size: 1.1rem !important;
        line-height: 1.4;
    }
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

def is_marked_correct(node):
    """Kiểm tra gạch chân, màu đỏ, highlight"""
    u_nodes = node.getElementsByTagNameNS(W_NS, "u")
    for u in u_nodes:
        val = u.getAttributeNS(W_NS, "val")
        if val and val != "none": return True
    
    color_nodes = node.getElementsByTagNameNS(W_NS, "color")
    for c in color_nodes:
        val = c.getAttributeNS(W_NS, "val")
        if val and val not in ["auto", "000000"]: return True
            
    highlight_nodes = node.getElementsByTagNameNS(W_NS, "highlight")
    for h in highlight_nodes:
        val = h.getAttributeNS(W_NS, "val")
        if val and val != "none": return True
    
    shd_nodes = node.getElementsByTagNameNS(W_NS, "shd")
    for s in shd_nodes:
        val = s.getAttributeNS(W_NS, "fill")
        if val and val not in ["auto", "FFFFFF", "000000"]: return True

    return False

# --- HÀM LỌC ĐÁP ÁN P3: CẮT BỎ CÁC TỪ THỪA ---
def extract_highlighted_text(blocks):
    """Lấy text đáp án và làm sạch"""
    extracted_text = []
    for block in blocks:
        runs = block.getElementsByTagNameNS(W_NS, "r")
        for run in runs:
            if is_marked_correct(run): 
                t_nodes = run.getElementsByTagNameNS(W_NS, "t")
                for t in t_nodes:
                    if t.firstChild and t.firstChild.nodeValue:
                        extracted_text.append(t.firstChild.nodeValue)
    
    full_text = "".join(extracted_text).strip()
    
    # Lọc bỏ "Câu 1.", "Câu 1:", "ĐS:", "Đáp số:", "KQ:"...
    full_text = re.sub(r'^(Câu\s*\d+[\.\:]\s*)?', '', full_text, flags=re.IGNORECASE)
    full_text = re.sub(r'^(ĐS|Đáp số|Đáp án|KQ|Kết quả)[\.\:]?\s*', '', full_text, flags=re.IGNORECASE)
    
    return full_text.strip()

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list: color_el = color_list[0]
    else:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
    color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    
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

def shuffle_mcq_options(question_blocks):
    indices = []; correct_char = None
    for i, block in enumerate(question_blocks):
        text = get_text(block)
        if re.match(r'^\s*[A-D][\.\)]', text, re.IGNORECASE): indices.append(i)
    if len(indices) < 2: return question_blocks, None
    original_options = [question_blocks[idx] for idx in indices]
    marked_index = -1
    for k, opt_block in enumerate(original_options):
        if is_marked_correct(opt_block):
            marked_index = k; break
    shuffled_options = shuffle_array(original_options)
    if marked_index != -1:
        target_block = original_options[marked_index]
        try:
            new_pos = shuffled_options.index(target_block)
            letters = ["A", "B", "C", "D"]
            if new_pos < len(letters): correct_char = letters[new_pos]
        except: pass
    min_idx = min(indices); max_idx = max(indices)
    return question_blocks[:min_idx] + shuffled_options + question_blocks[max_idx + 1:], correct_char

def shuffle_tf_options(question_blocks):
    option_indices = {}
    for i, block in enumerate(question_blocks):
        text = get_text(block)
        m = re.match(r'^\s*([a-d])\)', text, re.IGNORECASE)
        if m: option_indices[m.group(1).lower()] = i
    abc_keys = [k for k in ['a', 'b', 'c'] if k in option_indices]
    if len(abc_keys) < 2: return question_blocks, ""
    items_to_shuffle = []
    for k in abc_keys:
        blk = question_blocks[option_indices[k]]
        stat = 'Đ' if is_marked_correct(blk) else 'S'
        items_to_shuffle.append({'block': blk, 'status': stat})
    shuffled_items = shuffle_array(items_to_shuffle)
    d_item = None
    if 'd' in option_indices:
        blk = question_blocks[option_indices['d']]
        stat = 'Đ' if is_marked_correct(blk) else 'S'
        d_item = {'block': blk, 'status': stat}
    final_ordered_items = shuffled_items.copy()
    if d_item: final_ordered_items.append(d_item)
    ans_parts = [item['status'] for item in final_ordered_items]
    ans_str = "-".join(ans_parts)
    target_indices = sorted([v for k,v in option_indices.items() if k in ['a','b','c','d']])
    new_blocks = question_blocks.copy()
    for i, target_idx in enumerate(target_indices):
        if i < len(final_ordered_items):
            new_blocks[target_idx] = final_ordered_items[i]['block']
    return new_blocks, ans_str

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

def process_part(blocks, start, end, part_type, start_number=1):
    intro, questions = parse_questions_in_range(blocks, start, end)
    questions_data = [] 
    for q in questions:
        if part_type == "PHAN1":
            new_q_blocks, ans = shuffle_mcq_options(q)
            questions_data.append((new_q_blocks, ans))
        elif part_type == "PHAN2":
            new_q_blocks, ans = shuffle_tf_options(q)
            questions_data.append((new_q_blocks, ans))
        elif part_type == "PHAN3":
            ans = extract_highlighted_text(q) 
            questions_data.append((q, ans))
        else:
            questions_data.append((q, None))
    shuffled_data = shuffle_array(questions_data)
    final_questions_blocks = [x[0] for x in shuffled_data]
    final_answers_list = [x[1] for x in shuffled_data]
    relabel_questions(final_questions_blocks, start_number)
    if part_type == "PHAN1":
        for q in final_questions_blocks: relabel_mcq_options(q)
    elif part_type == "PHAN2":
        for q in final_questions_blocks: relabel_tf_options(q)
    result = intro.copy()
    for q in final_questions_blocks: result.extend(q)
    part_answers = {}
    for i, ans in enumerate(final_answers_list):
        if ans: part_answers[start_number + i] = ans
    next_number = start_number + len(final_questions_blocks)
    return result, next_number, part_answers

def process_all_as_mcq(blocks):
    intro, questions = parse_questions_in_range(blocks, 0, len(blocks))
    questions_data = []
    for q in questions:
        b, a = shuffle_mcq_options(q)
        questions_data.append((b, a))
    shuffled = shuffle_array(questions_data)
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

def shuffle_docx(file_bytes, shuffle_mode="auto"):
    input_buffer = io.BytesIO(file_bytes)
    all_answers = {} 
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        doc_xml = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        blocks = [c for c in body.childNodes if c.nodeType == c.ELEMENT_NODE and c.localName in ["p", "tbl"]]
        
        if shuffle_mode == "mcq":
            new_blocks, ans = process_all_as_mcq(blocks)
            all_answers["P1"] = ans
        elif shuffle_mode == "tf":
            new_blocks, _ = process_all_as_mcq(blocks)
        else:
            p1_idx = find_part_index(blocks, 1); p2_idx = find_part_index(blocks, 2)
            p3_idx = find_part_index(blocks, 3); p4_idx = find_part_index(blocks, 4)
            new_blocks = []; cursor = 0; curr_num = 1
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
            if p2_idx >= 0:
                new_blocks.append(blocks[p2_idx]); start2 = p2_idx + 1
                end2 = len(blocks)
                if p3_idx >= 0: end2 = p3_idx
                elif p4_idx >= 0: end2 = p4_idx
                p2_blocks, next_num, p2_ans = process_part(blocks, start2, end2, "PHAN2", curr_num)
                new_blocks.extend(p2_blocks)
                all_answers["P2"] = p2_ans 
                all_answers["P2_Start"] = (curr_num - len(p2_ans)) if p2_ans else curr_num
                all_answers["P2_Count"] = len(p2_ans) if p2_ans else 0
                curr_num = next_num; cursor = end2
            if p3_idx >= 0:
                new_blocks.append(blocks[p3_idx]); start3 = p3_idx + 1
                end3 = len(blocks)
                if p4_idx >= 0: end3 = p4_idx
                p3_blocks, next_num, p3_ans = process_part(blocks, start3, end3, "PHAN3", curr_num)
                new_blocks.extend(p3_blocks)
                all_answers["P3"] = p3_ans
                all_answers["P3_Start"] = (curr_num - len(p3_ans)) if p3_ans else curr_num
                all_answers["P3_Count"] = len(p3_ans) if p3_ans else 0
                curr_num = next_num; cursor = end3
            if p4_idx >= 0: new_blocks.extend(blocks[p4_idx:])
            if p1_idx == -1 and p2_idx == -1 and p3_idx == -1 and p4_idx == -1:
                new_blocks, ans = process_all_as_mcq(blocks)
                all_answers["P1"] = ans
        
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

def generate_answer_key_html(all_exam_data):
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
    if not all_exam_data: return html + "</body></html>"
    exam_codes = sorted(all_exam_data.keys())
    sample_data = all_exam_data[exam_codes[0]]
    if "P1" in sample_data and sample_data["P1"]:
        html += "<h3>PHẦN I: Trắc nghiệm nhiều lựa chọn</h3>"
        html += "<div class='note'>- Mỗi câu đúng được 0,25 điểm.</div>"
        q_nums = sorted(sample_data["P1"].keys())
        html += "<table><tr><th>Mã đề</th>"
        for q in q_nums: html += f"<th>{q}</th>"
        html += "</tr>"
        for code in exam_codes:
            html += f"<tr><td><b>{code}</b></td>"
            ans_map = all_exam_data[code].get("P1", {})
            for q in q_nums: html += f"<td><b>{ans_map.get(q, '')}</b></td>"
            html += "</tr>"
        html += "</table>"
    if "P2" in sample_data and sample_data["P2"]:
        count = sample_data.get("P2_Count", 0)
        q_nums_p2 = sorted(sample_data["P2"].keys())
        html += "<h3>PHẦN II: Trắc nghiệm đúng sai</h3>"
        html += "<div class='note'>- Điểm tối đa mỗi câu là 1 điểm.</div>"
        html += "<div class='note'>- Đúng 1 ý được 0,1 điểm; đúng 2 ý được 0,25 điểm; đúng 3 ý được 0,5 điểm; đúng 4 ý được 1 điểm.</div>"
        html += "<table><tr><th>Mã đề</th>"
        for q in q_nums_p2: html += f"<th>Câu {q}</th>"
        html += "</tr>"
        for code in exam_codes:
            html += f"<tr><td><b>{code}</b></td>"
            ans_map = all_exam_data[code].get("P2", {})
            for q in q_nums_p2: html += f"<td><b>{ans_map.get(q, '')}</b></td>"
            html += "</tr>"
        html += "</table>"
    if "P3" in sample_data and sample_data["P3"]:
        q_nums_p3 = sorted(sample_data["P3"].keys())
        html += "<h3>PHẦN III: Trắc nghiệm trả lời ngắn</h3>"
        html += "<div class='note'>- Điểm tối đa mỗi câu là 0,5 điểm.</div>"
        html += "<table><tr><th>Mã đề</th>"
        for q in q_nums_p3: html += f"<th>Câu {q}</th>"
        html += "</tr>"
        for code in exam_codes:
            html += f"<tr><td><b>{code}</b></td>"
            ans_map = all_exam_data[code].get("P3", {})
            for q in q_nums_p3: html += f"<td><b>{ans_map.get(q, '')}</b></td>"
            html += "</tr>"
        html += "</table>"
    html += "</body></html>"
    return html

def create_zip_multiple(file_bytes, base_name, num_versions, shuffle_mode, start_code):
    zip_buffer = io.BytesIO()
    all_exam_data = {}
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
        for i in range(num_versions):
            current_code = start_code + i
            shuffled_bytes, exam_answers = shuffle_docx(file_bytes, shuffle_mode)
            all_exam_data[current_code] = exam_answers
            filename = f"{base_name}_{current_code}.docx"
            zout.writestr(filename, shuffled_bytes)
        try:
            answer_key_html = generate_answer_key_html(all_exam_data)
            zout.writestr("Bang_Dap_An.doc", answer_key_html.encode('utf-8'))
        except Exception as e: print(f"Error creating answer key: {e}")
    return zip_buffer.getvalue()

# ==================== GIAO DIỆN STREAMLIT ====================
def main():
    st.markdown("""
    <div class="header-card">
        <h1>TRƯỜNG TRUNG HỌC PHỔ THÔNG MINH ĐỨC</h1>
        <h2>ỨNG DỤNG TRỘN ĐỀ WORD 2025</h2>
    </div>
    """, unsafe_allow_html=True)
    
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
            <strong>HƯỚNG DẪN TẠO ĐÁP ÁN TỰ ĐỘNG:</strong>
            <ul>
                <li><strong>Phần 1 & 2:</strong> Vui lòng <b>Gạch chân</b> hoặc <b>Tô màu đỏ</b> ý đúng.</li>
                <li><strong>Phần 3:</strong> Vui lòng <b>Gạch chân</b> hoặc <b>Tô màu đỏ</b> nội dung đáp án.</li>
            </ul>
            <p style="margin-top: 5px;">📥 <a href="https://docs.google.com/document/d/1A3bm_KNbl0vmnuYDfWdkqifS30RD-mLh/edit?usp=sharing&ouid=112824050529887271694&rtpof=true&sd=true" target="_blank">Tải file mẫu tại đây</a></p>
        </div>
        """, unsafe_allow_html=True)
    
    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.markdown('<div class="step-header">1️⃣ CHỌN FILE ĐỀ & KIỂU TRỘN</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("", type=["docx"], label_visibility="collapsed")
        if uploaded_file:
            st.success(f"✅ Đã tải lên: **{uploaded_file.name}**")
        
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
        st.markdown('<div class="step-header">2️⃣ CẤU HÌNH MÃ ĐỀ</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            num_versions = st.number_input("Số lượng đề", min_value=1, max_value=50, value=4)
        with c2:
            start_code = st.number_input("Mã đề bắt đầu", min_value=0, value=101)
        
        if num_versions > 1:
            st.info(f"📦 Tạo {num_versions} đề: {start_code} ➝ {start_code + num_versions - 1}")
        else:
            st.info(f"📄 Tạo 1 đề: {start_code}")

    st.markdown('<div class="step-header">3️⃣ THỰC HIỆN</div>', unsafe_allow_html=True)
    if st.button("🎲 BẮT ĐẦU TRỘN ĐỀ & TẢI VỀ", type="primary", use_container_width=True):
        if not uploaded_file:
            st.warning("⚠️ Vui lòng chọn file Word trước khi trộn!")
        else:
            try:
                with st.spinner("🚀 Đang xử lý và tạo bảng đáp án..."):
                    file_bytes = uploaded_file.read()
                    base_name = re.sub(r'[^\w\s-]', '', uploaded_file.name.replace(".docx", "")).strip() or "De"
                    
                    if num_versions == 1:
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
    
    st.markdown("""
    <div class="footer">
        <p>Zalo hỗ trợ kỹ thuật: <strong>038994070</strong></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
