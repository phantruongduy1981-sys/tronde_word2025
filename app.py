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
        padding: 0.8rem;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header-banner h1 { font-size: 1.8rem; font-weight: 800; margin: 0; color: white; text-transform: uppercase; }
    .step-label { font-weight: bold; font-size: 1.1rem; color: #004d40; margin-bottom: 0.5rem; display: flex; align-items: center; }
    .step-circle { background-color: #009688; color: white; width: 30px; height: 30px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; margin-right: 10px; }
    div[data-testid="stFileUploader"] { border: 1px dashed #009688; border-radius: 10px; padding: 10px; }
    .footer { position: fixed; bottom: 0; left: 0; width: 100%; background-color: #f1f5f9; text-align: center; padding: 5px; font-size: 0.7rem; color: #64748b; border-top: 1px solid #e2e8f0; z-index: 999; }
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC XỬ LÝ XML (CORE - FIX LỖI) ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_full_text_from_paragraph(paragraph):
    """Lấy toàn bộ text của đoạn văn để regex matching"""
    texts = []
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts)

def check_is_marked_answer(block):
    """Kiểm tra dấu hiệu đáp án (Gạch chân, Màu sắc, Highlight)"""
    runs = block.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr_list = r.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr_list: continue
        rPr = rPr_list[0]
        
        # Check Underline
        u_tags = rPr.getElementsByTagNameNS(W_NS, "u")
        if u_tags:
            val = u_tags[0].getAttributeNS(W_NS, "val")
            if val and val.lower() != 'none': return True

        # Check Color
        colors = rPr.getElementsByTagNameNS(W_NS, "color")
        if colors:
            val = colors[0].getAttributeNS(W_NS, "val")
            if val and val.lower() not in ['auto', '000000', 'black']: return True
                
        # Check Highlight/Shading
        if rPr.getElementsByTagNameNS(W_NS, "highlight") or rPr.getElementsByTagNameNS(W_NS, "shd"):
             return True
    return False

def remove_answer_formatting(block):
    """Xóa định dạng đáp án"""
    runs = block.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr_list = r.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr_list: continue
        rPr = rPr_list[0]
        for tag in ["color", "u", "highlight", "shd", "b"]:
            nodes = rPr.getElementsByTagNameNS(W_NS, tag)
            for node in nodes:
                rPr.removeChild(node)

def clean_short_answer_content(block):
    """Xử lý phần 3: Lấy đáp án và Xóa text 'ĐS:...' một cách an toàn"""
    runs = block.getElementsByTagNameNS(W_NS, "r")
    full_text = get_full_text_from_paragraph(block)
    
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
                        m_part = re.search(r'(.*?)(ĐS|Đáp số|DS)', val, re.IGNORECASE)
                        if m_part:
                            t.firstChild.nodeValue = m_part.group(1)
                            t.setAttribute("xml:space", "preserve") 
    return found_answer

def style_run_blue_bold(run):
    """Tô đậm xanh cho run chứa label"""
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

def update_label_smart(paragraph, regex_pattern, new_label_text):
    """
    CẬP NHẬT LABEL THÔNG MINH (FIX LỖI DÍNH CHỮ)
    """
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return False
    
    found = False
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        
        m = re.match(regex_pattern, txt, re.IGNORECASE)
        if m:
            # Giữ lại khoảng trắng đầu dòng (nếu có)
            original_match = m.group(0)
            leading_space = m.group(1) if m.lastindex and m.lastindex >= 1 else ""
            
            # Cắt bỏ phần label cũ, ghép label mới vào
            remaining_text = txt[len(original_match):]
            
            # Label mới = (Khoảng trắng cũ) + (Label mới từ code) + (Text còn lại)
            # new_label_text từ hàm gọi đã bao gồm dấu chấm (VD: "A.")
            t.firstChild.nodeValue = leading_space + new_label_text + remaining_text
            
            # QUAN TRỌNG: Giữ nguyên khoảng cách
            t.setAttribute("xml:space", "preserve")
            
            run = t.parentNode
            if run and run.localName == "r": style_run_blue_bold(run)
            
            # Xóa các node rác (dấu chấm/ngoặc thừa) NHƯNG KHÔNG XÓA KHOẢNG TRẮNG
            for j in range(i + 1, len(t_nodes)):
                t2 = t_nodes[j]
                if t2.firstChild and t2.firstChild.nodeValue:
                    val = t2.firstChild.nodeValue
                    # FIX LỖI DÍNH CHỮ: Chỉ xóa nếu là dấu chấm hoặc ngoặc. 
                    # Nếu chứa khoảng trắng (\s), TUYỆT ĐỐI KHÔNG XÓA.
                    if re.match(r'^[\.\)]+$', val): 
                        t2.firstChild.nodeValue = ""
                    else:
                        break # Gặp chữ hoặc khoảng trắng thì dừng ngay
            found = True
            break
    return found

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

# --- XỬ LÝ TỪNG LOẠI CÂU HỎI ---

def process_mcq_question(question_blocks):
    options = []
    others = []
    option_indices = []
    
    # Regex lỏng hơn để bắt được "A.", "A .", "A)"
    # Regex: (Space)(A-D)(Space)(Dot or Paren or Empty)
    label_regex = r'^\s*[A-D]\s*[\.\)]?'
    
    for i, block in enumerate(question_blocks):
        txt = get_full_text_from_paragraph(block)
        if re.match(label_regex, txt, re.IGNORECASE):
            options.append(block)
            option_indices.append(i)
        else:
            others.append(block)
            
    if len(options) < 2: return question_blocks, ""
    
    correct_original_idx = -1
    for idx, opt in enumerate(options):
        if check_is_marked_answer(opt):
            correct_original_idx = idx
            break
            
    for opt in options: remove_answer_formatting(opt)
        
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
        # Regex update phải khớp với Regex nhận diện
        # Group 1: Leading space. Group 2: Label char. Group 3: Space/Punct
        update_label_smart(block, r'^(\s*)([A-D])(\s*[\.\)])?', f"{lbl}.")
        
    min_idx, max_idx = min(option_indices), max(option_indices)
    final_blocks = question_blocks[:min_idx] + shuffled_blocks + question_blocks[max_idx+1:]
    return final_blocks, new_correct_char

def process_tf_question(question_blocks):
    opt_map = {}
    label_regex = r'^\s*([a-d])\s*[\.\)]'
    
    for i, block in enumerate(question_blocks):
        txt = get_full_text_from_paragraph(block)
        m = re.match(label_regex, txt, re.IGNORECASE)
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
        update_label_smart(item['block'], r'^(\s*)([a-d])(\s*[\.\)])?', f"{lbl})")
        
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
        if extracted: answer_val = extracted
    return question_blocks, answer_val

# --- PARSING ---

def parse_questions_in_range(blocks, start, end):
    part_blocks = blocks[start:end]
    intro = []
    questions = []
    i = 0
    # Regex nhận diện câu hỏi: Chấp nhận "Câu 1.", "Câu 1:", "Câu 1"
    q_regex = r'^Câu\s*\d+'
    
    while i < len(part_blocks):
        if re.match(q_regex, get_full_text_from_paragraph(part_blocks[i]), re.IGNORECASE): break
        intro.append(part_blocks[i])
        i += 1
    
    while i < len(part_blocks):
        txt = get_full_text_from_paragraph(part_blocks[i])
        if re.match(q_regex, txt, re.IGNORECASE):
            grp = [part_blocks[i]]
            i += 1
            while i < len(part_blocks):
                t2 = get_full_text_from_paragraph(part_blocks[i])
                if re.match(q_regex, t2, re.IGNORECASE) or re.search(r'PHẦN\s*\d', t2, re.IGNORECASE): break
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
            
    p1_idx = -1; p2_idx = -1; p3_idx = -1
    for i, b in enumerate(blocks):
        t = get_full_text_from_paragraph(b)
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
                # Update "Câu X" (Group 1: space, Group 2: Câu, Group 3: number)
                update_label_smart(q_blks[0], r'^(\s*)(Câu\s*)(\d+)(\s*[\.:])?', f"Câu {idx+1}.")
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
    for b in new_blocks: body.appendChild(b)
        
    new_xml = dom.toxml()
    out_io = io.BytesIO()
    with zipfile.ZipFile(out_io, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_xml.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))
    return out_io.getvalue(), exam_answers

# --- TẠO FILE ĐÁP ÁN ---
def generate_answer_key_doc(all_exam_data):
    html = """<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Đáp Án</title><style>body { font-family: 'Times New Roman', serif; font-size: 12pt; } table { border-collapse: collapse; width: 100%; margin-bottom: 20px; } th, td { border: 1px solid black; padding: 5px; text-align: center; } th { background-color: #f2f2f2; font-weight: bold; } h2, h3 { text-align: center; margin: 10px 0; color: #003366; } .part-title { font-weight: bold; margin-top: 20px; color: #003366; }</style></head><body>"""
    html += """<div style='text-align:center; font-weight:bold;'><p style='font-size:14pt; margin:0'>TRƯỜNG THPT MINH ĐỨC</p><p style='margin:0'>BẢNG ĐÁP ÁN TRỘN ĐỀ</p></div><hr>"""
    
    sorted_codes = sorted(all_exam_data.keys())
    if not sorted_codes: return b""
    
    p1_max = 0; p2_max = 0; p3_max = 0
    for code in sorted_codes:
        if len(all_exam_data[code]["P1"]) > p1_max: p1_max = len(all_exam_data[code]["P1"])
        if len(all_exam_data[code]["P2"]) > p2_max: p2_max = len(all_exam_data[code]["P2"])
        if len(all_exam_data[code]["P3"]) > p3_max: p3_max = len(all_exam_data[code]["P3"])

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
            col_btn_1, col_btn_2 = st.columns([1, 1])
            with col_btn_2:
                file_mau_url = "https://docs.google.com/document/d/1i1b-By6EA_HO8fWgMYG9iXZPGannmWdg/export?format=docx"
                st.link_button("📥 Tải File Mẫu", file_mau_url, help="Tải file mẫu chuẩn", type="primary", use_container_width=True)
            st.success("""
            **📌 Cấu trúc:**
            * PHẦN 1: Trắc nghiệm (A. B. C. D.)
            * PHẦN 2: Đúng/Sai (a) b) c) d))
            * PHẦN 3: Trả lời ngắn
            """)
            st.warning("""
            **⚠️ Lưu ý:**
            * Câu hỏi: `Câu 1.`, `Câu 2.`...
            * Đáp án: **Gạch chân** hoặc **Tô màu**.
            * Phần 3: Ghi **ĐS: Kết quả** và tô đỏ.
            """)
        
        st.markdown('<div class="step-label"><span class="step-circle">1</span>Upload file gốc</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Chọn file docx", type=["docx"], label_visibility="collapsed")
        if uploaded_file: st.success(f"✅ Đã nhận: {uploaded_file.name}")

    with col2:
        st.markdown('<div class="step-label"><span class="step-circle">2</span>Cấu hình</div>', unsafe_allow_html=True)
        shuffle_mode = st.radio("Kiểu trộn", options=["auto", "mcq", "tf"], format_func=lambda x: {"auto": "🔄 Tự động", "mcq": "📝 Trắc nghiệm", "tf": "✅ Đúng/Sai"}[x], label_visibility="collapsed")
        st.markdown('<div style="height: 10px"></div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1: num_versions = st.number_input("Số lượng đề", min_value=1, max_value=50, value=4)
        with c2: start_code = st.number_input("Mã bắt đầu", value=1001, step=1)
        st.caption(f"📍 Mã đề: {start_code} ⮕ {start_code + num_versions - 1}")
        st.markdown('<div style="height: 20px"></div>', unsafe_allow_html=True)
        
        if st.button("🚀 Trộn đề & Tạo đáp án", type="primary", use_container_width=True):
            if not uploaded_file:
                st.error("Chưa chọn file!")
            else:
                try:
                    with st.spinner("Đang xử lý (Fix lỗi dính chữ)..."):
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
                    st.download_button(label="📥 Tải trọn bộ (ZIP)", data=final_zip, file_name=f"TronDe_{base_name}_Full.zip", mime="application/zip", type="primary")
                except Exception as e:
                    st.error(f"Lỗi: {str(e)}")
                    st.exception(e)

    st.markdown('<div class="footer">© 2025 Phan Trường Duy - THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
