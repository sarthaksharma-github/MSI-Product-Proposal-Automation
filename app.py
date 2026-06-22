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
import io, copy, os, traceback

# ══════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="MSI Slide Automation Tool",
    page_icon="📊",
    layout="centered"
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
  .mode-badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    margin-left: 8px;
    vertical-align: middle;
  }
  .mode-product {
    background: #E8F4FD;
    color: #1A6FA8;
    border: 1px solid #B3D7F0;
  }
  .mode-store {
    background: #FFF0E0;
    color: #C05A00;
    border: 1px solid #FFD099;
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
#  SHARED HELPERS
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
    # If the value is a whole-number float (e.g. 7005.0), convert to int first
    # so it renders as "7005" instead of "7005.0".
    try:
        f = float(val)
        if f == int(f):
            val = int(f)
    except (ValueError, TypeError):
        pass
    v = str(val).strip()
    return "" if v.lower() == 'nan' else v


def fit_image_to_box(img_bytes, width, height):
    """Pad image to match the target aspect ratio, return JPEG BytesIO."""
    img = PILImage.open(io.BytesIO(img_bytes)).convert('RGB')
    img_w, img_h = img.size
    target_ratio = width / height
    img_ratio = img_w / img_h

    if img_ratio > target_ratio:
        new_h = int(img_w / target_ratio)
        padded = PILImage.new('RGB', (img_w, new_h), (255, 255, 255))
        padded.paste(img, (0, (new_h - img_h) // 2))
    else:
        new_w = int(img_h * target_ratio)
        padded = PILImage.new('RGB', (new_w, img_h), (255, 255, 255))
        padded.paste(img, ((new_w - img_w) // 2, 0))

    output = io.BytesIO()
    padded.save(output, format='JPEG', quality=95)
    output.seek(0)
    return output


# ══════════════════════════════════════════════════════════════
#  ENGINE A — PRODUCT SUBMISSION (original)
# ══════════════════════════════════════════════════════════════

def get_excel_images_by_row(excel_bytes):
    """Map images keyed by Excel row number (1-based). Used for product template."""
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

    output = fit_image_to_box(img_bytes, width, height)
    slide.shapes.add_picture(output, left, top, width, height)


def run_automation_product(excel_bytes, pptx_bytes, from_row, to_row):
    """Original engine: one image per row, product-submission template."""
    df = pd.read_excel(io.BytesIO(excel_bytes))

    start_idx = max(0, from_row - 2)
    end_idx   = min(to_row - 1, len(df))
    df_subset = df.iloc[start_idx:end_idx].copy()

    if df_subset.empty:
        raise ValueError("Selected row range contains no data.")

    image_map  = get_excel_images_by_row(excel_bytes)
    prs        = Presentation(io.BytesIO(pptx_bytes))
    if len(prs.slides) != 1:
        raise ValueError("Template must have exactly 1 slide.")

    source_slide = prs.slides[0]
    for _ in range(len(df_subset) - 1):
        clone_slide(prs, source_slide)

    for i, (original_idx, row) in enumerate(df_subset.iterrows()):
        slide        = prs.slides[i]
        excel_row_num = original_idx + 2

        replacements = {
            "[ITEM_NUM]":   safe_text(row.get('Item #', '')),
            "[ITEM_NAME]":  safe_text(row.get('Name', '')),
            "[RETAIL]":     safe_format(row.get('Retail', row.get('Unit Retail', 0)), is_currency=True),
            "[COST]":       safe_format(row.get('Unit Cost', 0), is_currency=True),
            "[IMU]":        safe_format(row.get('IMU%', 0), is_percent=True),
            "[PROJ_UNITS]": safe_format(row.get('Projected Sales Units', 0)),
            "[PROJ_RTL]":   safe_format(row.get('Projected Sales Rtl', 0), is_currency=True),
            "[FEAT_1]":     safe_text(row.get('Key Product Feature #1', '')),
            "[FEAT_2]":     safe_text(row.get('Key Product Feature #2', '')),
            "[FEAT_3]":     safe_text(row.get('Key Product Feature #3', '')),
            "[BEN_1]":      safe_text(row.get('Associated Customer Benefit #1', '')),
            "[BEN_2]":      safe_text(row.get('Associated Customer Benefit #2', '')),
            "[BEN_3]":      safe_text(row.get('Associated Customer Benefit #3', '')),
            "[CATEGORY]":   safe_text(row.get('Category', '')),
            "[DES]":        safe_text(row.get('Description', '')),
            "[MAT]":        safe_text(row.get('Material', '')),
            "[DIM]":        safe_text(row.get('Dimensions', '')),
            "[COLOR]":      safe_text(row.get('Color', '')),
            "[COUNTRY]":    safe_text(row.get('Country', '')),
        }

        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

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
#  ENGINE B — STORE EXECUTION (new)
# ══════════════════════════════════════════════════════════════

# Column indices (0-based) for each image placeholder tag in the store template.
# Excel layout:  A=Store Number, B=Image 1, C=Image 2, D=Image 3, E=Sales Units
_STORE_IMAGE_COLS = {
    "[Image 1]": 1,  # column B
    "[Image 2]": 2,  # column C
    "[Image 3]": 3,  # column D
}


def get_excel_images_by_cell(excel_bytes):
    """
    Map images keyed by (row_0based, col_0based) matching openpyxl anchor.
    Returns {(row, col): bytes}
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    image_map = {}
    for img in ws._images:
        try:
            r = img.anchor._from.row  # 0-based
            c = img.anchor._from.col  # 0-based
            image_map[(r, c)] = img._data()
        except:
            pass
    return image_map


def find_named_image_placeholder(slide, tag):
    """
    Find a text-box whose text exactly matches `tag` (case-insensitive),
    e.g. '[Image 1]'.
    """
    tag_lower = tag.strip().lower()
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip().lower()
            if text == tag_lower:
                return shape
    return None


def insert_image_clearing_placeholder(slide, img_bytes, placeholder):
    """Replace a named placeholder text-box with a fitted image."""
    left   = placeholder.left
    top    = placeholder.top
    width  = placeholder.width
    height = placeholder.height

    placeholder.fill.background()
    for para in placeholder.text_frame.paragraphs:
        for run in para.runs:
            run.text = ""

    output = fit_image_to_box(img_bytes, width, height)
    slide.shapes.add_picture(output, left, top, width, height)


def run_automation_store(excel_bytes, pptx_bytes, from_row, to_row):
    """
    Store-execution engine:
      - One slide per data row
      - Replaces [Store Number] and [UNITS] text placeholders
      - Replaces [Image 1], [Image 2], [Image 3] with images from columns B/C/D
    """
    df = pd.read_excel(io.BytesIO(excel_bytes))

    start_idx = max(0, from_row - 2)
    end_idx   = min(to_row - 1, len(df))
    df_subset = df.iloc[start_idx:end_idx].copy()

    if df_subset.empty:
        raise ValueError("Selected row range contains no data.")

    image_map = get_excel_images_by_cell(excel_bytes)

    prs = Presentation(io.BytesIO(pptx_bytes))
    if len(prs.slides) != 1:
        raise ValueError("Template must have exactly 1 slide.")

    source_slide = prs.slides[0]
    for _ in range(len(df_subset) - 1):
        clone_slide(prs, source_slide)

    for i, (original_idx, row) in enumerate(df_subset.iterrows()):
        slide = prs.slides[i]
        # openpyxl row 0 = header row, row 1 = first data row
        openpyxl_row_0 = original_idx + 1

        replacements = {
            "[Store Number]": safe_text(row.get('Store Number', '')),
            "[UNITS]":        safe_format(row.get('Sales Units', 0)),
        }

        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

        # Insert the three images by column
        for tag, col_0 in _STORE_IMAGE_COLS.items():
            placeholder = find_named_image_placeholder(slide, tag)
            img_bytes   = image_map.get((openpyxl_row_0, col_0))

            if img_bytes and placeholder:
                insert_image_clearing_placeholder(slide, img_bytes, placeholder)
            elif placeholder:
                # No image — silently clear the placeholder
                placeholder.fill.background()
                for para in placeholder.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = ""

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue(), len(df_subset)


# ══════════════════════════════════════════════════════════════
#  AUTO-DETECT TEMPLATE MODE
# ══════════════════════════════════════════════════════════════

_STORE_COLUMNS = {"store number", "image 1", "image 2", "image 3", "sales units"}


def detect_template_mode(excel_bytes):
    """
    Returns 'store' if the Excel file looks like a store-execution sheet,
    otherwise returns 'product'.
    """
    try:
        df   = pd.read_excel(io.BytesIO(excel_bytes), nrows=0)
        cols = {c.strip().lower() for c in df.columns}
        if _STORE_COLUMNS.issubset(cols):
            return "store"
    except:
        pass
    return "product"


# ══════════════════════════════════════════════════════════════
#  TEMPLATE LOADER
# ══════════════════════════════════════════════════════════════

TEMPLATES_DIR       = "templates"
EXCEL_TEMPLATES_DIR = "excel_templates"


def get_available_templates():
    if not os.path.exists(TEMPLATES_DIR):
        return []
    return [f for f in os.listdir(TEMPLATES_DIR) if f.endswith('.pptx')]


def load_template(filename):
    path = os.path.join(TEMPLATES_DIR, filename)
    with open(path, 'rb') as f:
        return f.read()


def get_available_excel_templates():
    if not os.path.exists(EXCEL_TEMPLATES_DIR):
        return []
    return [f for f in os.listdir(EXCEL_TEMPLATES_DIR) if f.endswith('.xlsx')]


def load_excel_template(filename):
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
st.caption(
    "Upload your Excel file. The tool **auto-detects** whether to use the "
    "**Product Submission** or **Store Execution** engine based on your column headers."
)

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
                label="⬇️ Download",
                data=load_excel_template(selected_excel),
                file_name=selected_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_excel_template"
            )
    st.markdown("<br>", unsafe_allow_html=True)

excel_file = st.file_uploader(
    "Excel File (.xlsx)",
    type=["xlsx"],
    help="Product data or store execution data with embedded images."
)

# Show detected mode badge once file is uploaded
if excel_file is not None:
    _peek = excel_file.read()
    excel_file.seek(0)
    detected_mode = detect_template_mode(_peek)
    if detected_mode == "store":
        st.markdown(
            "✅ Detected: <span class='mode-badge mode-store'>🏪 Store Execution Mode</span>"
            "&nbsp; Columns expected: <code>Store Number · Image 1 · Image 2 · Image 3 · Sales Units</code>",
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            "✅ Detected: <span class='mode-badge mode-product'>📦 Product Submission Mode</span>"
            "&nbsp; Standard product data sheet",
            unsafe_allow_html=True
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
                excel_bytes = excel_file.read()
                pptx_bytes  = load_template(selected_template)
                mode        = detect_template_mode(excel_bytes)

                if mode == "store":
                    result_bytes, count = run_automation_store(
                        excel_bytes, pptx_bytes,
                        from_row=int(from_row),
                        to_row=int(to_row)
                    )
                    mode_label = "Store Execution"
                else:
                    result_bytes, count = run_automation_product(
                        excel_bytes, pptx_bytes,
                        from_row=int(from_row),
                        to_row=int(to_row)
                    )
                    mode_label = "Product Submission"

                st.success(
                    f"✅ Done! {count} slide{'s' if count > 1 else ''} generated "
                    f"using the **{mode_label}** engine."
                )

                st.download_button(
                    label=f"⬇️  Download {fname}",
                    data=result_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except ValueError as e:
                st.error(f"⚠️ {str(e)}")
            except Exception as e:
                st.error(
                    f"Something went wrong: `{str(e)}`\n\n"
                    "Check your file formats and column names, then try again."
                )
                with st.expander("Show traceback"):
                    st.code(traceback.format_exc())

# ══════════════════════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════════════════════

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div class="msi-footer">
  MSI Services &nbsp;·&nbsp; Internal Tool &nbsp;·&nbsp; For issues contact Sales Support
</div>
""", unsafe_allow_html=True)
