import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT

# ==================== 1. CẤU HÌNH GIAO DIỆN ====================
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

# ==================== 2. HÀM CỐT LÕI (CORE UTILS) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

def get_text_from_node(node):
    """Lấy text thuần từ node"""
    texts = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def set_text_to_node(node, new_text):
    """Gán text mới vào node (An toàn)"""
    runs = node.getElementsByTagNameNS(W_NS, "r")
    if not runs: return
    
    # Dùng run đầu tiên để chứa text mới
    target_run = runs[0]
    
    # Xóa hết t cũ trong run đầu
    for t in target_run.getElementsByTagNameNS(W_NS, "t"):
        target_run.removeChild(t)
        
    # Tạo t mới
    doc = node.ownerDocument
    new_t = doc.createElementNS(W_NS, "w:t")
    new_t.setAttribute("xml:space", "preserve")
    new_t.appendChild(doc.createTextNode(new_text))
    target_run.appendChild(new_t)
    
    # Xóa các run thừa phía sau (để tránh text cũ còn sót)
    # Lưu ý: Chỉ xóa nếu run đó chỉ chứa text (không xóa run chứa ảnh/công thức)
    # Ở đây ta làm đơn giản: Clone node thì an toàn hơn.
    for i in range(1, len(runs)):
        node.removeChild(runs[i])

def has_complex_content(node):
    """Kiểm tra xem dòng có chứa công thức toán/ảnh không"""
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
    """Cập nhật nhãn (A., B., a), Câu 1...)"""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    
    # Gom text để check pattern
    full_text = ""
    for t in t_nodes: 
        if t.firstChild: full_text += t.firstChild.nodeValue
            
    # Pattern 1: A. B. C. D.
    if re.match(r'^\s*[A-D][\.:\)]', full_text, re.IGNORECASE):
        # Thay thế ký tự đầu
        new_text = re.sub(r'^\s*[A-D][\.:\)]', new_label, full_text, count=1)
        set_text_to_node(paragraph, new_text)
        return

    # Pattern 2: a) b) c) d)
    if re.match(r'^\s*[a-d][\.:\)]', full_text, re.IGNORECASE):
        new_text = re.sub(r'^\s*[a-d][\.:\)]', new_label, full_text, count=1)
        set_text_to_node(paragraph, new_text)
        return
        
    # Pattern 3: Câu X
    if re.match(r'^\s*Câu\s*\d+', full_text, re.IGNORECASE):
        new_text = re.sub(r'^\s*Câu\s*\d+[\.:]*', new_label, full_text, count=1, flags=re.IGNORECASE)
        set_text_to_node(paragraph, new_text)
        return

# ==================== 3. HÀM XỬ LÝ NÂNG CAO (FIX LỖI) ====================

def normalize_part1_options(question_blocks):
    """
    Tách dòng thông minh cho Phần 1.
    Nếu 1 dòng chứa "A. ... B. ...", cắt thành 2 dòng riêng biệt.
    """
    normalized = []
    for block in question_blocks:
        # Nếu có công thức -> Giữ nguyên (Safe Mode)
        if has_complex_content(block):
            normalized.append(block)
            continue
            
        txt = get_text_from_node(block)
        
        # Tìm tất cả các nhãn A., B., C., D. trong dòng
        # Regex: Bắt đầu dòng hoặc sau khoảng trắng, là A/B/C/D + dấu chấm/ngoặc
        matches = list(re.finditer(r'(?:^|\s)([A-D])[\.:]\s', txt))
        
        if len(matches) > 1:
            # Có nhiều đáp án trên 1 dòng -> Cắt
            indices = [m.start() for m in matches]
            indices.append(len(txt)) # Sentinel
            
            for i in range(len(matches)):
                start = indices[i]
                end = indices[i+1]
                # Lấy text con (ví dụ: "A. 5")
                sub_text = txt[start:end].strip()
                
                # Clone block gốc để giữ style cơ bản
                new_block = block.cloneNode(True)
                set_text_to_node(new_block, sub_text)
                normalized.append(new_block)
        else:
            normalized.append(block)
            
    return normalized

def process_part1_mcq(question_blocks, q_idx):
    """Xử lý PHẦN 1: Trắc nghiệm (Fix lỗi thiếu/dư)"""
    # 1. Tách dòng
    blocks = normalize_part1_options(question_blocks)
    
    intro = []
    options = []
    
    for b in blocks:
        txt = get_text_from_node(b)
        if re.match(r'^\s*[A-D][\.:]', txt):
            is_c = is_marked_correct(b)
            # Lưu lại text gốc để debug nếu cần
            options.append({'node': b, 'correct': is_c})
        else:
            intro.append(b)
    
    labels_mcq = ["A.", "B.", "C.", "D."]
    correct_char = "X"
    
    # 2. Logic Trộn
    # Chỉ trộn nếu có >= 2 đáp án. 
    # Nếu có > 4 đáp án (do lỗi file gốc), chỉ lấy 4 cái đầu tiên? 
    # -> Tạm thời trộn tất cả, nhưng chỉ gán nhãn cho 4 cái đầu.
    
    if len(options) >= 2:
        random.shuffle(options)
        
        # Gán lại nhãn
        for i, opt in enumerate(options):
            # Nếu vượt quá 4 đáp án, đánh dấu *
            lbl = labels_mcq[i] if i < 4 else "*"
            
            clean_formatting(opt['node'])
            update_label_safely(opt['node'], lbl)
            
            if opt['correct'] and i < 4:
                correct_char = lbl[0]
        
        result_blocks = intro + [o['node'] for o in options]
    else:
        # Fallback: Giữ nguyên
        for opt in options:
            clean_formatting(opt['node'])
        result_blocks = intro + [o['node'] for o in options]

    # Cập nhật câu
    if intro: 
        update_label_safely(intro[0], f"Câu {q_idx}. ")
        
    return result_blocks, correct_char

def process_part2_tf(question_blocks, q_idx):
    """Xử lý PHẦN 2: Đúng/Sai (Fix lỗi thứ tự nhãn)"""
    intro = []
    options = []
    
    for b in question_blocks:
        txt = get_text_from_node(b)
        # Nhận diện a) b) c) d)
        if re.match(r'^\s*[a-d][\.:\)]', txt, re.IGNORECASE):
            is_c = is_marked_correct(b)
            clean_formatting(b)
            options.append({'node': b, 'text': txt})
        else:
            intro.append(b)
            
    # Tách d)
    d_node = None
    others = []
    for o in options:
        # Check xem dòng này có phải là d) gốc không
        if re.match(r'^\s*d[\.:\)]', o['text'], re.IGNORECASE):
            d_node = o
        else:
            others.append(o)
            
    # Trộn a,b,c
    random.shuffle(others)
    
    # Ghép lại
    final_opts = others + ([d_node] if d_node else [])
    
    # CƯỠNG CHẾ GÁN NHÃN LẠI TỪ ĐẦU
    labels_tf = ["a)", "b)", "c)", "d)"]
    for k, opt in enumerate(final_opts):
        lbl = labels_tf[k] if k < 4 else "*"
        # Gọi hàm update label để sửa text đầu dòng (ví dụ "c) ..." thành "a) ...")
        update_label_safely(opt['node'], lbl)
    
    if intro:
        update_label_safely(intro[0], f"Câu {q_idx}. ")
        
    result_blocks = intro + [o['node'] for o in final_opts]
    return result_blocks, "X"

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
                continue 
        blocks_to_keep.insert(0, block)
        
    if blocks_to_keep:
        update_label_safely(blocks_to_keep[0], f"Câu {q_idx}. ")
        
    return blocks_to_keep, answer_text

# ==================== 4. LOGIC TỔNG HỢP ====================

def process_docx(file_bytes, num_exams, start_id, shuffle_mode):
    input_io = io.BytesIO(file_bytes)
    resources = {}; xml_content = ""
    try:
        with zipfile.ZipFile(input_io, 'r') as zin:
            for filename in zin.namelist():
                if filename == "word/document.xml": xml_content = zin.read(filename).decode('utf-8')
                else: resources[filename] = zin.read(filename)
    except Exception as e: return None, None, None, [f"Lỗi đọc file: {str(e)}"]

    if not xml_content: return None, None, None, ["File lỗi."]

    dom = minidom.parseString(xml_content)
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    all_blocks = [child for child in list(body.childNodes) 
                  if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]]
    errors = check_structure_errors(all_blocks)
    
    # Chia phần
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
                
                # TRỘN CÂU HỎI
                random.shuffle(questions)
                
                part_txt = get_text_from_node(intro_part[0]) if intro_part else ""
                current_mode = "MCQ"
                if shuffle_mode == "auto":
                    if "PHẦN 2" in part_txt.upper(): current_mode = "TF"
                    elif "PHẦN 3" in part_txt.upper(): current_mode = "FILL"
                elif shuffle_mode == "tf": current_mode = "TF"
                
                part_nodes = []
                for q in questions:
                    if current_mode == "MCQ":
                        nodes, ans = process_part1_mcq(q, global_q_idx)
                    elif current_mode == "TF":
                        nodes, ans = process_part2_tf(q, global_q_idx)
                    elif current_mode == "FILL":
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
            
        if excel_data_list:
            df = pd.DataFrame(excel_data_list)
            cols = ["Mã đề"] + sorted([c for c in df.columns if c != "Mã đề"], key=lambda x: int(x) if x.isdigit() else 999)
            df = df[cols]
            excel_io = io.BytesIO()
            with pd.ExcelWriter(excel_io, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='DapAn')
            master_zip.writestr("Dap_An_Excel.xlsx", excel_io.getvalue())
            
            word_bytes = create_word_answer_key(excel_data_list)
            master_zip.writestr("Dap_An_Word.docx", word_bytes)
        
    return final_zip_io.getvalue(), None, None, errors

# ==================== MAIN UI ====================
def main():
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG TRUNG HỌC PHỔ THÔNG MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025 - PRO VERSION 6.0</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        with st.expander("📄 Hướng dẫn & Quy cách", expanded=True):
            sample_url = "https://docs.google.com/document/d/1i1b-By6EA_HO8fWgMYG9iXZPGannmWdg/export?format=docx"
            st.link_button("📥 Tải File Mẫu Chuẩn", sample_url, use_container_width=True)
            st.markdown("""
            <div class="info-box">
                <div class="info-header">📌 QUY ĐỊNH BẮT BUỘC:</div>
                <ul>
                    <li><b>Câu hỏi:</b> Bắt đầu bằng <code>Câu 1.</code> (1 dấu chấm).</li>
                    <li><b>Phần 1 & 2:</b> Đáp án đúng <span style="color:red; font-weight:bold">TÔ ĐỎ</span> hoặc <u>GẠCH CHÂN</u>.</li>
                    <li><b>Phần 3:</b> Ghi <span style="color:red; font-weight:bold">ĐS: Kết quả</span> và tô đỏ.</li>
                </ul>
            </div>
            <div class="warning-box">
                <div class="warning-header">⚠️ Khắc phục lỗi thường gặp:</div>
                <ul>
                    <li><b>Lỗi Phần 1 thiếu đáp án:</b> Do viết gộp dòng (A. 5 B. 6). App đã tự động tách, nhưng tốt nhất nên xuống dòng.</li>
                    <li><b>Lỗi Phần 2 sai thứ tự:</b> App sẽ tự động đánh lại a,b,c,d từ trên xuống dưới.</li>
                </ul>
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

    st.markdown("""<div class="footer">© 2025 Phan Trường Duy - THPT Minh Đức<br>PRO VERSION 6.0</div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
