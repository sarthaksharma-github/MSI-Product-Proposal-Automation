# ============================================================
#  MSI SERVICES — SLIDE AUTOMATION TOOL v3.1
#  Streamlit App
# ============================================================

import streamlit as st
import pandas as pd
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Emu
from PIL import Image as PILImage
import io, copy, os, traceback, re, json, base64

# ══════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="MSI Slide Automation Tool",
    page_icon="📊",
    layout="wide"  # Use wide layout to better accommodate mapping UI & Live Preview side-by-side
)

# ══════════════════════════════════════════════════════════════
#  STYLING
# ══════════════════════════════════════════════════════════════

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
  }
  .msi-header {
    background: #8B6F4E;
    padding: 22px 30px;
    border-radius: 10px 10px 0 0;
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 0;
  }
  .msi-logo {
    font-size: 26px;
    font-weight: 600;
    color: #FFFFFF;
    letter-spacing: 2px;
  }
  .msi-header-title {
    margin: 0;
    font-size: 15px;
    font-weight: 600;
    color: #FFFFFF;
  }
  .msi-header-sub {
    margin: 2px 0 0;
    font-size: 11px;
    color: #E4D6CA;
    font-weight: 300;
  }
  .msi-footer {
    background: #F9F6F3;
    border: 1px solid #E4D6CA;
    border-radius: 0 0 10px 10px;
    padding: 10px 30px;
    font-size: 10.5px;
    color: #9A8070;
    text-align: center;
    margin-top: 0;
  }
  .msi-badge {
    display: inline-block;
    background: #FFF3E8;
    border: 1px solid #E4D6CA;
    border-radius: 20px;
    padding: 2px 8px;
    font-size: 10.5px;
    color: #8B6F4E;
    font-weight: 500;
    margin-left: 6px;
  }
  .stButton > button, .stDownloadButton > button {
    background-color: #8B6F4E !important;
    color: white !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 28px !important;
    font-size: 14px !important;
    width: 100%;
  }
  .stButton > button:hover, .stDownloadButton > button:hover {
    background-color: #6E5840 !important;
  }
  div[data-testid="stSelectbox"] label,
  div[data-testid="stFileUploader"] label,
  div[data-testid="stTextInput"] label,
  div[data-testid="stNumberInput"] label {
    font-weight: 600;
    color: #2C1F14;
    font-size: 13px;
  }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════

st.markdown("""
<div class="msi-header">
  <div class="msi-logo">MSI</div>
  <div>
    <p class="msi-header-title">Slide Automation Tool <span style="font-size:11px;opacity:0.7;font-weight:400;">v3.1</span></p>
    <p class="msi-header-sub">Sales Support Operations &nbsp;·&nbsp; Making Dream Surfaces Attainable</p>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  CORE ENGINE
# ══════════════════════════════════════════════════════════════

def clone_slide(prs, source_slide):
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    sp_tree = new_slide.shapes._spTree
    while len(sp_tree) > 2:
        sp_tree.remove(sp_tree[-1])
    source_sp_tree = source_slide.shapes._spTree
    for el in list(source_sp_tree)[2:]:
        sp_tree.append(copy.deepcopy(el))
    return new_slide

def get_excel_images(excel_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    image_map = {}
    for img in ws._images:
        try:
            row = img.anchor._from.row + 1
            image_map[row] = img._data()
        except:
            pass
    return image_map

def process_text_frame(tf, placeholders):
    for para in tf.paragraphs:
        for key, val in placeholders.items():
            if key in para.text:
                new_text = para.text.replace(key, str(val))
                if para.runs:
                    para.runs[0].text = new_text
                    for i in range(1, len(para.runs)):
                        para.runs[i].text = ""

def replace_text_in_shape(shape, placeholders):
    if shape.has_text_frame:
        process_text_frame(shape.text_frame, placeholders)
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                process_text_frame(cell.text_frame, placeholders)

def find_image_placeholder(slide):
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.lower()
            if "insert product" in text or "picture here" in text:
                return shape
    return None

def insert_image_into_placeholder(slide, img_bytes, placeholder):
    left, top, width, height = placeholder.left, placeholder.top, placeholder.width, placeholder.height
    placeholder.fill.background()
    for para in placeholder.text_frame.paragraphs:
        for run in para.runs:
            run.text = ""

    img = PILImage.open(io.BytesIO(img_bytes)).convert('RGB')
    img_w, img_h = img.size
    target_ratio, img_ratio = width / height, img_w / img_h

    if img_ratio > target_ratio:
        new_h = int(img_w / target_ratio)
        padded = PILImage.new('RGB', (img_w, new_h), (255, 255, 255))
        padded.paste(img, (0, (new_h - img_h) // 2))
    else:
        new_w = int(img_h * target_ratio)
        padded = PILImage.new('RGB', (new_w, img_h), (255, 255, 255))
        padded.paste(img, (((new_w - img_w) // 2), 0))

    output = io.BytesIO()
    padded.save(output, format='JPEG', quality=95)
    output.seek(0)
    slide.shapes.add_picture(output, left, top, width, height)

def safe_format(val, is_currency=False, is_percent=False):
    try:
        cleaned = str(val).replace('$', '').replace(',', '').replace('%', '').strip()
        num = 0.0 if (not cleaned or cleaned.lower() == 'nan') else float(cleaned)
        if is_currency:
            return f"${num:,.2f}" if num % 1 != 0 else f"${num:,.0f}"
        if is_percent:
            return f"{num:.0%}"
        return f"{num:,.0f}"
    except:
        return str(val)

def safe_text(val):
    v = str(val).strip()
    return "" if v.lower() == 'nan' else v

def extract_placeholders_from_pptx(pptx_bytes):
    """Scans the PowerPoint template for all bracketed tags matching [TAG]."""
    placeholders = set()
    pattern = re.compile(r'\[([^\]]+)\]')
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for match in pattern.findall(para.text):
                            placeholders.add(f"[{match.strip()}]")
                elif shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            for para in cell.text_frame.paragraphs:
                                for match in pattern.findall(para.text):
                                    placeholders.add(f"[{match.strip()}]")
    except Exception as e:
        pass
    return sorted(list(placeholders))

CONFIG_FILE = "mapping_config.json"

def load_mapping_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}

def save_mapping_config(config):
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4)
    except:
        pass

def render_slide_preview(mapping_dict, excel_row=None, image_bytes=None, is_template_mode=False):
    # 1. Classify tags
    title_tag = None
    sku_tag = None
    feature_tags = []
    benefit_tags = []
    metadata_tags = []
    
    for tag in mapping_dict.keys():
        tag_lower = tag.lower()
        if not title_tag and ("name" in tag_lower or "title" in tag_lower):
            title_tag = tag
        elif not sku_tag and ("num" in tag_lower or "sku" in tag_lower or "id" in tag_lower or "#" in tag_lower):
            sku_tag = tag
        elif "feat" in tag_lower:
            feature_tags.append(tag)
        elif "ben" in tag_lower:
            benefit_tags.append(tag)
        else:
            metadata_tags.append(tag)
            
    feature_tags.sort()
    benefit_tags.sort()
    metadata_tags.sort()
    
    all_tags = list(mapping_dict.keys())
    if not title_tag and all_tags:
        title_tag = all_tags[0]
        
    def get_val(tag):
        if is_template_mode:
            return tag
        col = mapping_dict[tag].get("column", "")
        fmt = mapping_dict[tag].get("format", "Text")
        if not col or excel_row is None:
            return f"({tag} Unmapped)"
        raw_val = excel_row.get(col, "")
        if pd.isna(raw_val):
            return ""
        if fmt == "Currency":
            return safe_format(raw_val, is_currency=True)
        elif fmt == "Percentage":
            return safe_format(raw_val, is_percent=True)
        elif fmt == "Integer":
            return safe_format(raw_val)
        return safe_text(raw_val)
        
    title_val = get_val(title_tag) if title_tag else "Sample Product Title"
    sku_val = get_val(sku_tag) if sku_tag else "SKU-000000"
    
    features_html = ""
    for f_tag in feature_tags:
        val = get_val(f_tag)
        if val:
            features_html += f"<li>{val}</li>"
    if not features_html:
        features_html = "<li>Key Product Feature #1</li><li>Key Product Feature #2</li>"
        
    benefits_html = ""
    for b_tag in benefit_tags:
        val = get_val(b_tag)
        if val:
            benefits_html += f"<li>{val}</li>"
    if not benefits_html:
        benefits_html = "<li>Customer Benefit #1</li><li>Customer Benefit #2</li>"

    details_html = ""
    for m_tag in metadata_tags:
        val = get_val(m_tag)
        label = m_tag.replace("[", "").replace("]", "").replace("_", " ").title()
        details_html += f"""
        <div style="display: flex; justify-content: space-between; padding: 4px 0; border-bottom: 1px solid #F0EAE1;">
            <span style="font-weight: 600; color: #5C483A; font-size: 12px;">{label}</span>
            <span style="color: #2C1F14; font-size: 12px; font-weight: 500;">{val or '—'}</span>
        </div>
        """
        
    if is_template_mode:
        img_html = """
        <div style="width: 100%; height: 260px; background-color: #EFECE8; border: 2px dashed #C8B8AA; border-radius: 6px; display: flex; align-items: center; justify-content: center; color: #8B6F4E; font-weight: 600; font-size: 13px;">
            [PRODUCT IMAGE PLACEHOLDER]
        </div>
        """
    else:
        if image_bytes:
            base64_str = base64.b64encode(image_bytes).decode('utf-8')
            img_html = f"""
            <img src="data:image/jpeg;base64,{base64_str}" style="width: 100%; height: 260px; object-fit: contain; background-color: #FFFFFF; border-radius: 6px; border: 1px solid #E4D6CA; padding: 4px;" />
            """
        else:
            img_html = """
            <div style="width: 100%; height: 260px; background-color: #F9F6F3; border: 1px solid #E4D6CA; border-radius: 6px; display: flex; align-items: center; justify-content: center; color: #C4B0A0; font-style: italic; font-size: 12px;">
                No Image Found for Row
            </div>
            """
            
    html_content = f"""
    <div style="font-family: 'DM Sans', sans-serif; background: #FFFFFF; border: 1px solid #E4D6CA; border-radius: 10px; box-shadow: 0 4px 12px rgba(44, 31, 20, 0.06); padding: 20px; max-width: 100%; margin: 10px auto;">
        <div style="display: flex; justify-content: space-between; align-items: baseline; border-bottom: 2px solid #8B6F4E; padding-bottom: 8px; margin-bottom: 15px;">
            <h3 style="margin: 0; color: #2C1F14; font-size: 18px; font-weight: 600;">{title_val}</h3>
            <span style="font-size: 11px; background: #FFF3E8; border: 1px solid #E4D6CA; color: #8B6F4E; padding: 2px 8px; border-radius: 20px; font-weight: 600;">{sku_val}</span>
        </div>
        
        <div style="display: flex; gap: 20px; flex-wrap: wrap;">
            <div style="flex: 1.2; min-width: 280px; display: flex; flex-direction: column; gap: 15px;">
                <div>
                    <h4 style="margin: 0 0 6px 0; color: #8B6F4E; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid #E4D6CA; padding-bottom: 3px;">Details</h4>
                    {details_html or '<div style="color:#C4B0A0; font-size:12px; font-style:italic;">No metadata mapped</div>'}
                </div>
                
                <div style="display: flex; gap: 15px;">
                    <div style="flex: 1;">
                        <h4 style="margin: 0 0 6px 0; color: #8B6F4E; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Key Features</h4>
                        <ul style="margin: 0; padding-left: 16px; color: #5C483A; font-size: 11.5px; line-height: 1.4;">
                            {features_html}
                        </ul>
                    </div>
                    <div style="flex: 1;">
                        <h4 style="margin: 0 0 6px 0; color: #8B6F4E; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Benefits</h4>
                        <ul style="margin: 0; padding-left: 16px; color: #5C483A; font-size: 11.5px; line-height: 1.4;">
                            {benefits_html}
                        </ul>
                    </div>
                </div>
            </div>
            
            <div style="flex: 0.8; min-width: 220px; display: flex; align-items: center; justify-content: center;">
                {img_html}
            </div>
        </div>
    </div>
    """
    st.markdown(html_content, unsafe_allow_html=True)

def run_automation(excel_bytes, pptx_bytes, from_row, to_row, mapping_dict, image_tag):
    df = pd.read_excel(io.BytesIO(excel_bytes))

    start_idx = max(0, from_row - 2)
    end_idx = min(to_row - 1, len(df))
    df_subset = df.iloc[start_idx:end_idx].copy()

    if df_subset.empty:
        raise ValueError("Selected row range contains no data.")

    image_map = get_excel_images(excel_bytes)
    prs = Presentation(io.BytesIO(pptx_bytes))
    if len(prs.slides) != 1:
        raise ValueError("Template must have exactly 1 slide.")

    source_slide = prs.slides[0]
    for _ in range(len(df_subset) - 1):
        clone_slide(prs, source_slide)

    for i, (original_idx, row) in enumerate(df_subset.iterrows()):
        slide = prs.slides[i]
        excel_row_num = original_idx + 2

        replacements = {}
        for tag, cfg in mapping_dict.items():
            col = cfg.get("column", "")
            fmt = cfg.get("format", "Text")
            if not col:
                continue
            val = row.get(col, "")
            if pd.isna(val):
                val = ""
            
            if fmt == "Currency":
                formatted_val = safe_format(val, is_currency=True)
            elif fmt == "Percentage":
                formatted_val = safe_format(val, is_percent=True)
            elif fmt == "Integer":
                formatted_val = safe_format(val)
            else:
                formatted_val = safe_text(val)
                
            replacements[tag] = formatted_val

        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

        placeholder = None
        if image_tag and image_tag != '(Detect by "insert product" / "picture here" text)':
            for shape in slide.shapes:
                if shape.has_text_frame and image_tag in shape.text_frame.text:
                    placeholder = shape
                    break
        else:
            placeholder = find_image_placeholder(slide)

        if excel_row_num in image_map:
            if placeholder:
                insert_image_into_placeholder(slide, image_map[excel_row_num], placeholder)
        else:
            if placeholder:
                for para in placeholder.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = ""
                placeholder.fill.background()

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue(), len(df_subset)


# ══════════════════════════════════════════════════════════════
#  TEMPLATE LOADER
# ══════════════════════════════════════════════════════════════

TEMPLATES_DIR = "templates"
EXCEL_TEMPLATES_DIR = "excel_templates"

def get_available_templates():
    """Returns list of .pptx filenames found in the templates/ folder."""
    if not os.path.exists(TEMPLATES_DIR):
        return []
    return [f for f in os.listdir(TEMPLATES_DIR) if f.endswith('.pptx')]

def load_template(filename):
    """Reads and returns bytes of a template from the templates/ folder."""
    path = os.path.join(TEMPLATES_DIR, filename)
    with open(path, 'rb') as f:
        return f.read()

def get_available_excel_templates():
    """Returns list of .xlsx filenames found in the excel_templates/ folder."""
    if not os.path.exists(EXCEL_TEMPLATES_DIR):
        return []
    return [f for f in os.listdir(EXCEL_TEMPLATES_DIR) if f.endswith('.xlsx')]

def load_excel_template(filename):
    """Reads and returns bytes of an excel template from the excel_templates/ folder."""
    path = os.path.join(EXCEL_TEMPLATES_DIR, filename)
    with open(path, 'rb') as f:
        return f.read()


# ══════════════════════════════════════════════════════════════
#  UI — STEP 1: TEMPLATE SELECTION
# ══════════════════════════════════════════════════════════════

st.markdown("#### Step 1 — Select Template")
st.caption("Choose the PowerPoint template to use. Templates are pre-loaded — no upload needed.")

available_templates = get_available_templates()

if not available_templates:
    st.warning("No templates found in the `templates/` folder. Please add at least one `.pptx` template to the repo.")
    st.stop()

selected_template = st.selectbox(
    "Template",
    options=available_templates,
    format_func=lambda x: x.replace('.pptx', '').replace('_', ' ')
)

st.divider()

# ══════════════════════════════════════════════════════════════
#  UI — STEP 2: EXCEL UPLOAD
# ══════════════════════════════════════════════════════════════

st.markdown("#### Step 2 — Upload Product Data")
st.caption("Upload your Excel file. Any filename is accepted.")

excel_templates = get_available_excel_templates()
if excel_templates:
    st.markdown("<p style='font-size: 13px; font-weight: 600; color: #2C1F14; margin-bottom: 5px;'>Need an Excel template?</p>", unsafe_allow_html=True)
    t_col1, t_col2 = st.columns([3, 1])
    with t_col1:
        selected_excel = st.selectbox(
            "Select Excel Template",
            options=excel_templates,
            format_func=lambda x: x.replace('.xlsx', '').replace('_', ' '),
            label_visibility="collapsed"
        )
    with t_col2:
        if selected_excel:
            st.download_button(
                label="📥 Download",
                data=load_excel_template(selected_excel),
                file_name=selected_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_excel_template"
            )
    st.markdown("<br>", unsafe_allow_html=True)

excel_file = st.file_uploader(
    "Excel File (.xlsx)",
    type=["xlsx"],
    help="Product data with embedded images. Filename does not matter."
)

st.divider()

# ══════════════════════════════════════════════════════════════
#  UI — STEP 2.5: INTERACTIVE COLUMN MAPPING & LIVE PREVIEW
# ══════════════════════════════════════════════════════════════

mapping_dict = {}
image_tag = '(Detect by "insert product" / "picture here" text)'

if excel_file is not None:
    st.markdown("#### Step 2.5 — Map Columns & Preview")
    st.caption("Customize how placeholders on the slide link to Excel columns. Preview changes in real-time.")
    
    excel_bytes = excel_file.getvalue()
    pptx_bytes = load_template(selected_template)
    
    detected_tags = extract_placeholders_from_pptx(pptx_bytes)
    
    try:
        df = pd.read_excel(io.BytesIO(excel_bytes))
        excel_columns = list(df.columns)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()
        
    if not detected_tags:
        st.info("No placeholders (like `[TAG]`) found in the template. Processing as-is.")
    else:
        saved_config = load_mapping_config()
        saved_mappings = saved_config.get("mappings", {})
        saved_image_tag = saved_config.get("image_tag", '(Detect by "insert product" / "picture here" text)')
        
        map_col, prev_col = st.columns([1, 1.2])
        
        with map_col:
            st.markdown("**Placeholder Mappings**")
            
            for tag in detected_tags:
                tag_clean = tag.replace("[", "").replace("]", "").replace("_", " ").lower()
                
                default_col = None
                default_format = "Text"
                if tag in saved_mappings:
                    default_col = saved_mappings[tag].get("column")
                    default_format = saved_mappings[tag].get("format", "Text")
                    
                if default_col not in excel_columns:
                    default_col = None
                    for col in excel_columns:
                        col_lower = str(col).lower()
                        if tag_clean in col_lower or col_lower in tag_clean:
                            default_col = col
                            break
                            
                    if not default_col:
                        if "num" in tag_clean and any(x in str(c).lower() for c in excel_columns for x in ["item #", "item number", "sku", "id"]):
                            default_col = next(c for c in excel_columns if any(x in str(c).lower() for x in ["item #", "item number", "sku", "id"]))
                        elif "name" in tag_clean and any(x in str(c).lower() for c in excel_columns for x in ["name", "title", "description"]):
                            default_col = next(c for c in excel_columns if any(x in str(c).lower() for x in ["name", "title", "description"]))
                            
                if default_format == "Text" and tag not in saved_mappings:
                    tag_l = tag.lower()
                    if any(x in tag_l for x in ["retail", "cost", "price", "rtl", "amt", "value"]):
                        default_format = "Currency"
                    elif any(x in tag_l for x in ["imu", "percent", "pct", "%"]):
                        default_format = "Percentage"
                    elif any(x in tag_l for x in ["units", "qty", "count", "num"]):
                        default_format = "Integer"
                
                col_idx = excel_columns.index(default_col) if default_col in excel_columns else 0
                
                c1, c2 = st.columns([2, 1])
                with c1:
                    mapped_col_name = st.selectbox(
                        f"Map {tag}",
                        options=excel_columns,
                        index=col_idx,
                        key=f"map_{tag}"
                    )
                with c2:
                    format_options = ["Text", "Currency", "Percentage", "Integer"]
                    fmt_idx = format_options.index(default_format) if default_format in format_options else 0
                    mapped_format = st.selectbox(
                        "Format",
                        options=format_options,
                        index=fmt_idx,
                        key=f"fmt_{tag}"
                    )
                
                mapping_dict[tag] = {"column": mapped_col_name, "format": mapped_format}
                
            st.markdown("---")
            st.markdown("**Image Configuration**")
            
            image_tag_options = ['(Detect by "insert product" / "picture here" text)'] + detected_tags
            img_idx = 0
            if saved_image_tag in image_tag_options:
                img_idx = image_tag_options.index(saved_image_tag)
                
            image_tag = st.selectbox(
                "Image Placeholder Tag",
                options=image_tag_options,
                index=img_idx,
                help="Specify which text tag represents the image container, or let the app scan automatically."
            )
            
            if st.button("💾 Save Mapping Configuration"):
                config_to_save = {
                    "mappings": mapping_dict,
                    "image_tag": image_tag
                }
                save_mapping_config(config_to_save)
                st.success("Configuration saved successfully!")
                
        with prev_col:
            st.markdown("**Live Slide Preview**")
            preview_mode = st.radio(
                "Preview Mode",
                options=["Template View (Placeholders)", "Data View (First Product Row)"],
                horizontal=True
            )
            
            is_template = (preview_mode == "Template View (Placeholders)")
            first_row = df.iloc[0].to_dict() if len(df) > 0 else None
            
            image_map = get_excel_images(excel_bytes)
            first_excel_row = 2
            first_row_image = image_map.get(first_excel_row)
            if not first_row_image and image_map:
                first_row_image = image_map[min(image_map.keys())]
                
            render_slide_preview(
                mapping_dict=mapping_dict,
                excel_row=first_row,
                image_bytes=first_row_image,
                is_template_mode=is_template
            )
            
    st.divider()

# ══════════════════════════════════════════════════════════════
#  UI — STEP 3: OUTPUT FILENAME
# ══════════════════════════════════════════════════════════════

st.markdown("#### Step 3 — Output Filename")
st.caption("Name your output file. `.pptx` is added automatically.")

output_name = st.text_input(
    "Output Filename (Optional)",
    value="",
    placeholder="e.g. May_Fountain_Submissions"
)

st.divider()

# ══════════════════════════════════════════════════════════════
#  UI — STEP 4: ROW RANGE
# ══════════════════════════════════════════════════════════════

st.markdown("#### Step 4 — Row Range *(optional)*")
st.caption("Leave defaults to process all rows. Adjust for a partial batch.")

col1, col2 = st.columns(2)
with col1:
    from_row = st.number_input("From Row", min_value=1, max_value=9999, value=2)
with col2:
    to_row = st.number_input("To Row", min_value=1, max_value=9999, value=9999)

st.divider()

# ══════════════════════════════════════════════════════════════
#  UI — STEP 5: GENERATE
# ══════════════════════════════════════════════════════════════

st.markdown("#### Step 5 — Generate Slides")

if st.button("🚀  Generate Slides"):
    if excel_file is None:
        st.error("⚠️ Please upload your Excel file before generating.")
    else:
        default_fname = f"{selected_template.replace('.pptx', '')}_Output"
        fname = output_name.strip() or default_fname
        if not fname.endswith('.pptx'):
            fname += '.pptx'

        with st.spinner("Building your slides... please wait."):
            try:
                excel_bytes = excel_file.getvalue()
                pptx_bytes = load_template(selected_template)

                result_bytes, count = run_automation(
                    excel_bytes, pptx_bytes,
                    from_row=int(from_row),
                    to_row=int(to_row),
                    mapping_dict=mapping_dict,
                    image_tag=image_tag
                )

                st.success(f"✔️ Done! {count} slide{'s' if count > 1 else ''} generated.")

                st.download_button(
                    label=f"📥  Download {fname}",
                    data=result_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except ValueError as e:
                st.error(f"⚠️ {str(e)}")
            except Exception as e:
                st.error(f"Something went wrong: `{str(e)}` \n\nCheck your file formats and column names, then try again.")

# ══════════════════════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════════════════════

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div class="msi-footer">
  MSI Services &nbsp;·&nbsp; Internal Tool &nbsp;·&nbsp; For issues contact Sales Support
</div>
""", unsafe_allow_html=True)
