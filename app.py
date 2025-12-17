import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG ====================
st.set_page_config(
    page_title="Trộn Đề Word Pro - THPT Minh Đức",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    .header-banner {
        background-color: #00796b;
        color: white;
        padding: 1rem 1rem;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header-banner h1 { 
        font-size: 2rem; 
        font-weight: 800; 
        margin: 0; 
        color: white; 
        text-transform: uppercase; 
    }
    .header-banner p { 
        font-size: 0.9rem; 
        opacity: 0.9; 
        margin-top: 5px; 
        margin-bottom: 0;
    }
    .step-label { 
        font-weight: bold; 
        font-size: 1.1rem; 
        color: #004d40; 
        margin-bottom: 0.5rem; 
        display: flex; 
        align-items: center; 
    }
    .step-circle { 
        background-color: #009688; 
        color: white; 
        width: 30px; 
        height: 30px; 
        border-radius: 50%; 
        display: inline-flex; 
        align-items: center; 
        justify-content: center; 
        margin-right: 10px; 
    }
    div[data-testid="stFileUploader"] { 
        border: 1px dashed #009688; 
        border-radius: 10px; 
        padding: 10px; 
    }
    .footer { 
        position: fixed; 
        bottom: 0; 
        left: 0; 
        width: 100%; 
        background-color: #f1f5f9; 
        text-align: center; 
        padding: 8px; 
        font-size: 0.8rem; 
        color: #64748b; 
        border-top: 1px solid #e2e8f0; 
        z-index: 999; 
    }
    div[role="radiogroup"] { gap: 10px; }
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC XỬ LÝ XML & ĐÁP ÁN ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def check_is_marked_answer(block):
    """
    Kiểm tra xem block có được đánh dấu là đáp án không.
    Cải tiến: Bắt được nhiều trường hợp màu sắc và gạch chân hơn.
    """
    runs = block.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr_list = r.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr_list: continue
        rPr = rPr_list[0]
        
        # 1. Kiểm tra Gạch chân (u)
        u_tags = rPr.getElementsByTagNameNS(W_NS, "u")
        if u_tags:
            val = u_tags[0].getAttributeNS(W_NS, "val")
            # Chấp nhận mọi kiểu gạch chân trừ 'none'
            if val and val.lower() != 'none':
                return True

        # 2. Kiểm tra Màu sắc (color)
        colors = rPr.getElementsByTagNameNS(W_NS, "color")
        if colors:
            val = colors[0].getAttributeNS(W_NS, "val")
            # Chấp nhận mọi màu trừ 'auto' và đen tuyền
            if val and val.lower() not in ['auto', '000000', 'black']:
                return True
                
        # 3. Kiểm tra Highlight / Shading (Nền)
        if rPr.getElementsByTagNameNS(W_NS, "highlight") or rPr.getElementsByTagNameNS(W_NS, "shd"):
             return True
             
        # 4. Kiểm tra Bold (Phòng hờ nếu người dùng dùng Bold)
        # Bỏ comment dòng dưới nếu muốn chấp nhận cả in đậm là đáp án
        # if rPr.getElementsByTagNameNS(W_NS, "b"): return True
        
    return False

def remove_answer_formatting(block):
    runs = block.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr_list = r.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr_list: continue
        rPr = rPr_list[0]
        for tag in ["color", "u", "highlight", "shd", "b"]: # Xóa cả Bold nếu có
            nodes = rPr.getElementsByTagNameNS(W_NS, tag)
            for node in nodes:
                rPr.removeChild(node)

def clean_short_answer_content(block):
    runs = block.getElementsByTagNameNS(W_NS, "r")
    full_text = get_text(block)
    match = re.search(r'(ĐS|Đáp số|DS)\s*[:\.]\s*(.*)', full_text, re.IGNORECASE)
    found_answer = None
    if match:
        found_answer = match.group(2).strip()
        for r in runs:
            t_nodes = r.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild and t.firstChild.nodeValue:
                    val = t.firstChild.nodeValue
                    if re.search(r'(ĐS|Đáp số|DS)', val, re.IGNORECASE):
                        m_run = re.search(r'(.*?)(ĐS|Đáp số|DS)', val, re.IGNORECASE)
                        if m_run:
                            t.firstChild.nodeValue = m_run.group(1)
                        else:
                            t.firstChild.nodeValue = ""
    return found_answer

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if not color_list:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
    else: color_el = color_list[0]
    color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    if not rPr.getElementsByTagNameNS(W_NS, "b"):
        rPr.appendChild(doc.createElementNS(W_NS, "w:b"))

def update_label(paragraph, regex_pattern, new_label_func):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    found = False
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        m = re.match(regex_pattern, txt, re.IGNORECASE)
        if m:
            new_text = new_label_func(m, txt)
            t.firstChild.nodeValue = new_text
            run = t.parentNode
            if run and run.localName == "r": style_run_blue_bold(run)
            for j in range(i + 1, len(t_nodes)):
                t2 = t_nodes[j]
                if t2.firstChild and t2.firstChild.nodeValue:
                    if re.match(r'^[\.\)\s]*$', t2.firstChild.nodeValue):
                        t2.firstChild.nodeValue = ""
                    else: break
            found = True
            break
    return found

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

def process_mcq_question(question_blocks):
    options = []
    others = []
    option_indices = []
    for i, block in enumerate(question_blocks):
        txt = get_text(block)
        if re.match(r'^\s*[A-D][\.\)]', txt, re.IGNORECASE):
            options.append(block)
            option_indices.append(i)
        else:
            others.append(block)
            
    if len(options) < 2: return question_blocks, ""
    
    # Logic tìm đáp án đúng được cải thiện
    correct_original_idx = -1
    for idx, opt in enumerate(options):
        if check_is_marked_answer(opt):
            correct_original_idx = idx
            break
            
    # Clean formatting
    for opt in options:
        remove_answer_formatting(opt)
        
    tagged_options = []
    for idx, opt in enumerate(options):
        is_correct = (idx == correct_original_idx)
        tagged_options.append((opt, is_correct))
        
    shuffled_tagged = shuffle_array(tagged_options)
    shuffled_blocks = [x[0] for x in shuffled_tagged]
    
    new_correct_char = ""
    chars = ["A", "B", "C", "D"]
    for idx, (opt, is_correct) in enumerate(shuffled_tagged):
        if is_correct:
            new_correct_char = chars[idx] if idx < 4 else "?"
            
    for idx, block in enumerate(shuffled_blocks):
        lbl = chars[idx] if idx < 4 else chars[-1]
        update_label(block, r'^(\s*)([A-D])([\.\)])?', lambda m, txt: m.group(1) + lbl + ".")
        
    min_idx = min(option_indices)
    max_idx = max(option_indices)
    final_blocks = question_blocks[:min_idx] + shuffled_blocks + question_blocks[max_idx+1:]
    return final_blocks, new_correct_char

def process_tf_question(question_blocks):
    opt_map = {}
    for i, block in enumerate(question_blocks):
        txt = get_text(block)
        m = re.match(r'^\s*([a-d])\)', txt, re.IGNORECASE)
        if m:
            key = m.group(1).lower()
            opt_map[key] = (block, i)
    if len(opt_map) < 2: return question_blocks, ""
    items = []
    for k in ['a', 'b', 'c', 'd']:
        if k in opt_map:
            blk = opt_map[k][0]
            is_true = check_is_marked_answer(blk)
            remove_answer_formatting(blk)
            items.append({'key': k, 'block': blk, 'val': is_true})
    to_shuffle = [x for x in items if x['key'] in ['a','b','c']]
    fixed = [x for x in items if x['key'] == 'd']
    shuffled_abc = shuffle_array(to_shuffle)
    final_order = shuffled_abc + fixed
    ans_parts = []
    labels = ['a', 'b', 'c', 'd']
    for idx, item in enumerate(final_order):
        ans_parts.append("D" if item['val'] else "S")
        lbl = labels[idx]
        update_label(item['block'], r'^(\s*)([a-d])(\))?', lambda m, txt: m.group(1) + lbl + ")")
    final_ans_str = "-".join(ans_parts)
    indices = [opt_map[k][1] for k in opt_map]
    min_idx, max_idx = min(indices), max(indices)
    new_opt_blocks = [x['block'] for x in final_order]
    final_blocks = question_blocks[:min_idx] + new_opt_blocks + question_blocks[max_idx+1:]
    return final_blocks, final_ans_str

def process_short_ans_question(question_blocks):
    answer_val = ""
    for block in question_blocks:
        extracted = clean_short_answer_content(block)
        if extracted:
            answer_val = extracted
    return question_blocks, answer_val

def parse_questions_in_range(blocks, start, end):
    part_blocks = blocks[start:end]
    intro = []
    questions = []
    i = 0
    while i < len(part_blocks):
        if re.match(r'^Câu\s*\d+\b', get_text(part_blocks[i])): break
        intro.append(part_blocks[i])
        i += 1
    while i < len(part_blocks):
        txt = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', txt):
            grp = [part_blocks[i]]
            i += 1
            while i < len(part_blocks):
                t2 = get_text(part_blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2) or re.match(r'^PHẦN\s*\d', t2, re.IGNORECASE): break
                grp.append(part_blocks[i])
                i += 1
            questions.append(grp)
        else:
            i += 1
    return intro, questions

def process_full_docx(file_bytes, code_name):
    input_buffer = io.BytesIO(file_bytes)
    zin = zipfile.ZipFile(input_buffer, 'r')
    doc_xml = zin.read("word/document.xml").decode('utf-8')
    dom = minidom.parseString(doc_xml)
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    blocks = []
    for child in body.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]:
            blocks.append(child)
    p1_idx = -1
    p2_idx = -1
    p3_idx = -1
    for i, b in enumerate(blocks):
        t = get_text(b)
        if re.search(r'PHẦN\s*1', t, re.IGNORECASE): p1_idx = i
        elif re.search(r'PHẦN\s*2', t, re.IGNORECASE): p2_idx = i
        elif re.search(r'PHẦN\s*3', t, re.IGNORECASE): p3_idx = i
    exam_answers = {"P1": {}, "P2": {}, "P3": {}}
    new_blocks = []
    cursor = 0
    def handle_part(start_idx, end_idx, part_type):
        intro, questions = parse_questions_in_range(blocks, start_idx, end_idx)
        processed_qs = []
        for q_blocks in questions:
            if part_type == 1: b_new, ans = process_mcq_question(q_blocks)
            elif part_type == 2: b_new, ans = process_tf_question(q_blocks)
            else: b_new, ans = process_short_ans_question(q_blocks)
            processed_qs.append((b_new, ans))
        shuffled_qs_with_ans = shuffle_array(processed_qs)
        final_part_blocks = intro.copy()
        for idx, (q_blks, ans) in enumerate(shuffled_qs_with_ans):
            if q_blks:
                update_label(q_blks[0], r'^(\s*)(Câu\s*)(\d+)(\.)?', lambda m, t: f"{m.group(1)}Câu {idx+1}.")
            final_part_blocks.extend(q_blks)
            exam_answers[f"P{part_type}"][idx+1] = ans
        return final_part_blocks
    if p1_idx != -1:
        new_blocks.extend(blocks[cursor:p1_idx+1])
        cursor = p1_idx + 1
        end = p2_idx if p2_idx != -1 else (p3_idx if p3_idx != -1 else len(blocks))
        new_blocks.extend(handle_part(cursor, end, 1))
        cursor = end
    if p2_idx != -1:
        new_blocks.append(blocks[p2_idx])
        cursor = p2_idx + 1
        end = p3_idx if p3_idx != -1 else len(blocks)
        new_blocks.extend(handle_part(cursor, end, 2))
        cursor = end
    if p3_idx != -1:
        new_blocks.append(blocks[p3_idx])
        cursor = p3_idx + 1
        end = len(blocks)
        new_blocks.extend(handle_part(cursor, end, 3))
        cursor = end
    if p1_idx == -1 and p2_idx == -1 and p3_idx == -1:
        new_blocks.extend(handle_part(0, len(blocks), 1))
    for child in list(body.childNodes):
        if child.nodeType == child.ELEMENT_NODE and child.localName not in ["sectPr"]:
            body.removeChild(child)
    for b in new_blocks:
        body.appendChild(b)
    new_xml = dom.toxml()
    out_io = io.BytesIO()
    with zipfile.ZipFile(out_io, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_xml.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))
    return out_io.getvalue(), exam_answers

def generate_answer_key_doc(all_exam_data):
    html = """<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Đáp Án</title><style>body { font-family: 'Times New Roman', serif; font-size: 12pt; } table { border-collapse: collapse; width: 100%; margin-bottom: 20px; } th, td { border: 1px solid black; padding: 5px; text-align: center; } th { background-color: #f2f2f2; font-weight: bold; } h2, h3 { text-align: center; margin: 10px 0; color: #003366; } .part-title { font-weight: bold; margin-top: 20px; color: #003366; }</style></head><body>"""
    html += """<div style='text-align:center; font-weight:bold;'><p style='font-size:14pt; margin:0'>TRƯỜNG THPT MINH ĐỨC</p><p style='margin:0'>BẢNG ĐÁP ÁN TRỘN ĐỀ</p></div><hr>"""
    
    sorted_codes = sorted(all_exam_data.keys())
    if not sorted_codes: return b""
    
    # Tính số lượng câu hỏi tối đa cho mỗi phần (để tránh lỗi nếu mã đề nào đó parse thiếu)
    p1_max = max((len(all_exam_data[c]["P1"]) for c in sorted_codes), default=0)
    p2_max = max((len(all_exam_data[c]["P2"]) for c in sorted_codes), default=0)
    p3_max = max((len(all_exam_data[c]["P3"]) for c in sorted_codes), default=0)

    if p1_max > 0:
        html += "<div class='part-title'>PHẦN I: Trắc nghiệm nhiều lựa chọn</div>"
        html += "<table><tr><th>Mã đề</th>"
        for i in range(1, p1_max + 1): html += f"<th>{i}</th>"
        html += "</tr>"
        for code in sorted_codes:
            html += f"<tr><td><b>{code}</b></td>"
            ans_dict = all_exam_data[code]["P1"]
            for i in range(1, p1_max + 1): html += f"<td>{ans_dict.get(i, '')}</td>"
            html += "</tr>"
        html += "</table>"

    if p2_max > 0:
        html += "<div class='part-title'>PHẦN II: Trắc nghiệm đúng sai</div>"
        html += "<table><tr><th>Mã đề</th>"
        for i in range(1, p2_max + 1): html += f"<th>Câu {i}</th>"
        html += "</tr>"
        for code in sorted_codes:
            html += f"<tr><td><b>{code}</b></td>"
            ans_dict = all_exam_data[code]["P2"]
            for i in range(1, p2_max + 1): html += f"<td>{ans_dict.get(i, '')}</td>"
            html += "</tr>"
        html += "</table>"

    if p3_max > 0:
        html += "<div class='part-title'>PHẦN III: Trả lời ngắn</div>"
        html += "<table><tr><th>Mã đề</th>"
        for i in range(1, p3_max + 1): html += f"<th>Câu {i}</th>"
        html += "</tr>"
        for code in sorted_codes:
            html += f"<tr><td><b>{code}</b></td>"
            ans_dict = all_exam_data[code]["P3"]
            for i in range(1, p3_max + 1): html += f"<td>{ans_dict.get(i, '')}</td>"
            html += "</tr>"
        html += "</table>"
        
    html += "</body></html>"
    return html.encode('utf-8')

# ==================== MAIN UI ====================

def main():
    st.markdown("""
    <div class="header-banner">
        <h1>TRƯỜNG THPT MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025 - PRO VERSION</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1.5, 1], gap="large")
    
    with col1:
        with st.expander("📄 Hướng dẫn & File Mẫu", expanded=True):
            st.info("""
            **Quy tắc nhận diện đáp án:**
            1. **MCQ & Đúng/Sai:** Đáp án phải được **Gạch chân** hoặc **Tô màu** (hệ thống đã cải tiến để nhận diện màu nhạy hơn).
            2. **Trả lời ngắn (Phần 3):** Ghi `ĐS: kết quả` (tô màu).
            """)
            file_mau_url = "https://docs.google.com/document/d/1i1b-By6EA_HO8fWgMYG9iXZPGannmWdg/export?format=docx"
            st.link_button("📥 Tải File Word Mẫu", file_mau_url, help="Bấm để tải file mẫu chuẩn về máy", type="secondary")
        
        st.markdown('<div class="step-label"><span class="step-circle">1</span>Upload file gốc (có đáp án)</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Chọn file docx", type=["docx"], label_visibility="collapsed")
        if uploaded_file: st.success(f"✅ Đã nhận: {uploaded_file.name}")

    with col2:
        st.markdown('<div class="step-label"><span class="step-circle">2</span>Chọn kiểu trộn</div>', unsafe_allow_html=True)
        shuffle_mode = st.radio("Chọn kiểu trộn", options=["auto", "mcq", "tf"], format_func=lambda x: {"auto": "🔄 Tự động (Phần 1, 2, 3)", "mcq": "📝 Trắc nghiệm", "tf": "✅ Đúng/Sai"}[x], label_visibility="collapsed")
        
        st.markdown('<div style="height: 20px"></div>', unsafe_allow_html=True)
        st.markdown('<div class="step-label"><span class="step-circle">3</span>Số mã đề cần tạo</div>', unsafe_allow_html=True)
        num_versions = st.number_input("Số lượng mã đề", min_value=1, max_value=50, value=4, label_visibility="collapsed")
        st.caption(f"📍 Tạo ra {num_versions} mã đề khác nhau từ file gốc.")
        st.markdown('<div style="height: 20px"></div>', unsafe_allow_html=True)
        
        start_code = 1001
        
        if st.button("🚀 Trộn đề & Tạo đáp án", type="primary", use_container_width=True):
            if not uploaded_file:
                st.error("Chưa chọn file!")
            else:
                try:
                    with st.spinner("Đang xử lý..."):
                        original_bytes = uploaded_file.read()
                        uploaded_file.seek(0)
                        base_name = uploaded_file.name.replace(".docx", "")
                        zip_buffer = io.BytesIO()
                        all_exam_data = {}
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
                            for i in range(num_versions):
                                code = start_code + i
                                new_docx, answers = process_full_docx(original_bytes, str(code))
                                fname = f"{base_name}_MaDe_{code}.docx"
                                zout.writestr(fname, new_docx)
                                all_exam_data[code] = answers
                            key_docx_bytes = generate_answer_key_doc(all_exam_data)
                            zout.writestr(f"DAP_AN_CHITIET_{base_name}.doc", key_docx_bytes)
                        final_zip = zip_buffer.getvalue()
                    st.success("Thành công! Tải về bên dưới:")
                    st.download_button(label="📥 Tải trọn bộ (Đề + Đáp án)", data=final_zip, file_name=f"TronDe_{base_name}_Full.zip", mime="application/zip", type="primary")
                except Exception as e:
                    st.error(f"Lỗi: {str(e)}")
                    st.exception(e)

    st.markdown('<div class="footer">© 2025 Phan Trường Duy - THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
