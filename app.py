# ============================================================
#  MSI SERVICES — SLIDE AUTOMATION TOOL v3.1
#  Streamlit App — Dynamic Column Mapping Edition
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
  .stButton > button {
    background-color: #8B6F4E !important;
    color: white !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 28px !important;
    font-size: 14px !important;
    width: 100%;
  }
  .stButton > button:hover {
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
    left, top, width, height = (
        placeholder.left, placeholder.top,
        placeholder.width, placeholder.height
    )
    placeholder.fill.background()
    for para in placeholder.text_frame.paragraphs:
        for run in para.runs:
            run.text = ""

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
    slide.shapes.add_picture(output, left, top, width, height)


# ══════════════════════════════════════════════════════════════
#  SMART FORMAT  ← NEW in v3.1
#
#  Replaces safe_format() + safe_text() with a single unified
#  function. Rules (evaluated in order):
#
#  1. Empty / NaN                 → ""
#  2. Already has a % symbol      → treat as percent string
#  3. Numeric AND 0 < value < 1   → percentage  (e.g. 0.45 → "45%")
#  4. Numeric AND has decimals    → currency    (e.g. 12.99 → "$12.99")
#  5. Numeric AND whole number    → comma int   (e.g. 1500 → "1,500")
#  6. Everything else             → plain text
# ══════════════════════════════════════════════════════════════

def smart_format(val):
    v = str(val).strip()

    # Rule 1 — empty or NaN
    if not v or v.lower() == 'nan':
        return ""

    # Rule 2 — already marked as percent by the source data
    if '%' in v:
        try:
            num = float(v.replace('%', '').replace(',', '').strip())
            # Source stores as "45%" → we want "45%", not "4500%"
            # So divide by 100 only if num > 1 (it's already a whole-percent)
            ratio = num / 100 if num > 1 else num
            return f"{ratio:.0%}"
        except ValueError:
            return v

    # Strip currency symbols to attempt numeric parse
    cleaned = v.replace('$', '').replace(',', '').strip()

    try:
        num = float(cleaned)

        # Rule 3 — fraction between 0 and 1 → percentage
        if 0 < num < 1:
            return f"{num:.0%}"

        # Rule 4 — has a non-zero decimal component → currency
        if num % 1 != 0:
            return f"${num:,.2f}"

        # Rule 5 — whole number → plain formatted integer
        return f"{num:,.0f}"

    except ValueError:
        # Rule 6 — non-numeric text → return as-is
        return v


# ══════════════════════════════════════════════════════════════
#  DYNAMIC REPLACEMENTS BUILDER  ← NEW in v3.1
#
#  Reads every column from the DataFrame row and constructs
#  the replacements dict automatically.
#
#  Convention: Excel column "Unit Retail"  →  PPT tag [Unit Retail]
#
#  The PPT template is now the source of truth for tag names.
#  Whatever placeholder text exists in the template (e.g. [Name],
#  [Unit Cost]) must exactly match the corresponding Excel header,
#  wrapped in square brackets.
# ══════════════════════════════════════════════════════════════

def build_replacements(row):
    """
    Given a pandas Series (one Excel row), returns a dict of
    { "[Column Name]": formatted_value } for every column.
    """
    replacements = {}
    for col_name, val in row.items():
        tag = f"[{col_name}]"
        replacements[tag] = smart_format(val)
    return replacements


# ══════════════════════════════════════════════════════════════
#  MAIN AUTOMATION RUNNER
# ══════════════════════════════════════════════════════════════

def run_automation(excel_bytes, pptx_bytes, from_row, to_row):
    df = pd.read_excel(io.BytesIO(excel_bytes))

    # from_row / to_row are Excel row numbers (header = row 1, data starts row 2)
    start_idx = max(0, from_row - 2)
    end_idx   = min(to_row - 1, len(df))
    df_subset = df.iloc[start_idx:end_idx].copy()

    if df_subset.empty:
        raise ValueError("Selected row range contains no data.")

    image_map  = get_excel_images(excel_bytes)
    prs        = Presentation(io.BytesIO(pptx_bytes))

    if len(prs.slides) != 1:
        raise ValueError("Template must have exactly 1 slide.")

    source_slide = prs.slides[0]
    for _ in range(len(df_subset) - 1):
        clone_slide(prs, source_slide)

    for i, (original_idx, row) in enumerate(df_subset.iterrows()):
        slide         = prs.slides[i]
        excel_row_num = original_idx + 2   # offset: header=1, first data row=2

        # ── Dynamic replacements (no hardcoding) ──────────────
        replacements = build_replacements(row)

        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

        # ── Image insertion ───────────────────────────────────
        placeholder = find_image_placeholder(slide)

        if excel_row_num in image_map:
            if placeholder:
                insert_image_into_placeholder(
                    slide, image_map[excel_row_num], placeholder
                )
        else:
            # No image for this row — silently clear the placeholder
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

def get_available_templates():
    if not os.path.exists(TEMPLATES_DIR):
        return []
    return [f for f in os.listdir(TEMPLATES_DIR) if f.endswith('.pptx')]

def load_template(filename):
    path = os.path.join(TEMPLATES_DIR, filename)
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

excel_file = st.file_uploader(
    "Excel File (.xlsx)",
    type=["xlsx"],
    help="Product data with embedded images. Filename does not matter."
)

# ── Column preview (shown once Excel is uploaded) ────────────
if excel_file is not None:
    try:
        preview_df = pd.read_excel(excel_file, nrows=0)   # headers only
        excel_file.seek(0)                                  # reset for later read
        tags = [f"[{col}]" for col in preview_df.columns]
        st.info(
            f"**{len(tags)} columns detected.** "
            f"Your template placeholders should match these tags exactly:  \n"
            + "  `" + "`   `".join(tags) + "`"
        )
    except Exception:
        pass   # non-critical — don't block the user

st.divider()

# ══════════════════════════════════════════════════════════════
#  UI — STEP 3: OUTPUT FILENAME
# ══════════════════════════════════════════════════════════════

st.markdown("#### Step 3 — Output Filename")
st.caption("Name your output file. `.pptx` is added automatically.")

output_name = st.text_input(
    "Output Filename",
    value="Final_Submissions",
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
        fname = output_name.strip() or "Final_Submissions"
        if not fname.endswith('.pptx'):
            fname += '.pptx'

        with st.spinner("Building your slides... please wait."):
            try:
                excel_bytes = excel_file.read()
                pptx_bytes  = load_template(selected_template)

                result_bytes, count = run_automation(
                    excel_bytes, pptx_bytes,
                    from_row=int(from_row),
                    to_row=int(to_row)
                )

                st.success(f"✅ Done! {count} slide{'s' if count > 1 else ''} generated.")

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
                    f"Check your file formats and that your template placeholders "
                    f"match the column names shown above."
                )

# ══════════════════════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════════════════════

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div class="msi-footer">
  MSI Services &nbsp;·&nbsp; Internal Tool &nbsp;·&nbsp; For issues contact Sales Support
</div>
""", unsafe_allow_html=True)
