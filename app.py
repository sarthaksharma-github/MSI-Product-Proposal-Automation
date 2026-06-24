# ============================================================
#  MSI SERVICES — SLIDE AUTOMATION TOOL v3.4
#  Streamlit App
# ============================================================

import streamlit as st
import pandas as pd
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Emu
from PIL import Image as PILImage
import io, copy, os, traceback, re, json, base64
import streamlit.components.v1 as components

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
    <p class="msi-header-title">Slide Automation Tool <span style="font-size:11px;opacity:0.7;font-weight:400;">v3.4</span></p>
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
    """Returns {col_0based: [img_bytes sorted by anchor row]} for positional assignment."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    raw = []  # list of (row, col, img_bytes)
    for img in ws._images:
        try:
            row = img.anchor._from.row
            col = img.anchor._from.col
            raw.append((row, col, img._data()))
        except:
            pass
    # Sort by row so product 0 = first image, product 1 = second, etc.
    raw.sort(key=lambda x: x[0])
    col_to_images = {}
    for row, col, img_bytes in raw:
        col_to_images.setdefault(col, []).append(img_bytes)
    return col_to_images

def parse_placeholder_tag(full_tag):
    """Parse a PPTX placeholder tag that may carry an inline type annotation.

    New format (recommended):
        '[IMAGE 1 (Image)]'       → base='IMAGE 1',  type='Image',      symbol=None
        '[NAME (Text)]'           → base='NAME',      type='Text',       symbol=None
        '[RETAIL (Currency $)]'   → base='RETAIL',    type='Currency',   symbol='$'
        '[LENGTH (Integer)]'      → base='LENGTH',    type='Integer',    symbol=None
        '[IMU (Percentage)]'      → base='IMU',       type='Percentage', symbol=None

    Legacy format (backward-compatible, no annotation):
        '[NAME]'                  → base='NAME',      type=None,         symbol=None

    Returns (base: str, col_type: str|None, symbol: str|None)
    """
    inner = full_tag.strip().lstrip('[').rstrip(']').strip()  # e.g. "RETAIL (Currency $)"
    m = re.match(r'^(.+?)\s*\(([^)]+)\)\s*$', inner)
    if m:
        base = m.group(1).strip()
        ann  = m.group(2).strip()
        ann_l = ann.lower()
        if any(x in ann_l for x in ['image', 'img', 'photo', 'picture']):
            return base, 'Image', None
        if 'currency' in ann_l or any(c in ann for c in '$€£¥₹'):
            # Extract the currency symbol (any non-alpha, non-space char after stripping 'currency')
            sym_m = re.search(r'[\.\$€£¥₹]', ann)
            symbol = sym_m.group(0) if sym_m else '$'
            return base, 'Currency', symbol
        if any(x in ann_l for x in ['percent', '%']):
            return base, 'Percentage', None
        if any(x in ann_l for x in ['integer', 'int', 'number', 'num']):
            return base, 'Integer', None
        return base, 'Text', None
    # No annotation — legacy tag like [NAME] or [RETAIL]
    return inner, None, None


def build_auto_mapping(all_pptx_tags, excel_columns):
    """Universal auto-mapping: type is read from the PPTX tag annotation.

    PPTX tag format (recommended):
        [IMAGE 1 (Image)]        → image placeholder, matched to 'Image 1' column
        [NAME (Text)]            → text field,        matched to 'Name' column
        [RETAIL (Currency $)]    → currency ($),      matched to 'HD Retail' / 'Retail' column
        [FEATURE 1 (Text)]       → text field,        matched to 'Feature 1' column

    Legacy tags without annotation (backward-compatible):
        [RETAIL]                 → falls back to tag-name heuristics (retail → Currency)

    Excel column names stay PLAIN (no annotations needed).
    Returns (mapping_dict, image_mappings) compatible with run_automation().
    """
    def _norm(s):
        s = re.sub(r'[\[\]_]', ' ', str(s))
        return re.sub(r'\s+', ' ', s).strip().lower()

    def _num(s):
        m = re.search(r'(\d+)\s*$', s.strip())
        return int(m.group(1)) if m else None

    def _best_excel_match(base_norm, cols):
        """Find the best-matching plain Excel column for a normalised base name."""
        tag_num   = _num(base_norm)
        tag_roots = {w for w in base_norm.split() if not w.isdigit()}
        best_col, best_score = '', -999
        for col in cols:
            col_norm  = _norm(col)
            col_num   = _num(col_norm)
            col_roots = {w for w in col_norm.split() if not w.isdigit()}
            if not any(tr in cr or cr in tr for tr in tag_roots for cr in col_roots):
                continue
            score = 1
            if tag_num is not None and col_num is not None:
                score += 5 if tag_num == col_num else -10
            elif tag_num is not None:
                score -= 1
            if base_norm in col_norm or col_norm in base_norm:
                score += 2
            if score > best_score:
                best_score, best_col = score, col
        return best_col if best_score > 0 else ''

    mapping_dict   = {}
    image_mappings = {}

    for full_tag in all_pptx_tags:
        base, col_type, symbol = parse_placeholder_tag(full_tag)
        base_norm   = _norm(base)
        matched_col = _best_excel_match(base_norm, excel_columns)

        if not matched_col:
            continue

        if col_type == 'Image':
            image_mappings[full_tag] = matched_col

        else:
            # Determine format: annotation wins, then tag-name heuristic
            if col_type in ('Currency', 'Percentage', 'Integer', 'Text'):
                fmt = col_type
            else:
                tag_l = full_tag.lower()
                if any(x in tag_l for x in ['retail', 'cost', 'price', 'rtl', 'amt', 'value']):
                    fmt    = 'Currency'
                    symbol = symbol or '$'
                elif any(x in tag_l for x in ['imu', 'percent', 'pct', '%']):
                    fmt = 'Percentage'
                elif any(x in tag_l for x in ['units', 'qty', 'count']):
                    fmt = 'Integer'
                else:
                    fmt = 'Text'

            mapping_dict[full_tag] = {
                'column': matched_col,
                'format': fmt,
                'symbol': symbol or '$',
            }

    return mapping_dict, image_mappings


def process_text_frame(tf, placeholders):
    for para in tf.paragraphs:
        for key, val in placeholders.items():
            if key in para.text:
                new_text = para.text.replace(key, str(val))
                if para.runs:
                    para.runs[0].text = new_text
                    for i in range(1, len(para.runs)):
                        para.runs[i].text = ""

def purge_empty_paragraphs(shape):
    """Remove paragraphs whose text is empty (blank bullet points from empty data values).
    Keeps at least one paragraph so the shape stays valid."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    txBody = tf._txBody
    all_paras = tf.paragraphs
    if len(all_paras) <= 1:
        return  # Never remove the last paragraph
    for para in list(all_paras):
        # Check all run text concatenated
        run_text = ''.join(r.text or '' for r in para.runs).strip()
        if run_text == '' and len(tf.paragraphs) > 1:
            txBody.remove(para._p)

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

def safe_format(val, is_currency=False, currency_symbol='$', is_percent=False):
    try:
        cleaned = re.sub(r'[^\d.\-]', '', str(val).strip())
        num = 0.0 if (not cleaned or cleaned.lower() == 'nan') else float(cleaned)
        if is_currency:
            return f"{currency_symbol}{num:,.2f}" if num % 1 != 0 else f"{currency_symbol}{num:,.0f}"
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

def render_slide_preview(mapping_dict, image_mappings=None, excel_row=None, image_map=None, excel_row_idx=None, is_template_mode=False):
    """Renders a live HTML preview card. image_mappings = {tag: col_name}, image_map = {(row,col): bytes}."""
    if image_mappings is None:
        image_mappings = {}
    if image_map is None:
        image_map = {}
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
        
    # Build image HTML for each mapped image tag
    def _img_box(label, img_bytes_val):
        if is_template_mode:
            return f'<div style="flex:1;min-width:140px;height:200px;background:#EFECE8;border:2px dashed #C8B8AA;border-radius:6px;display:flex;align-items:center;justify-content:center;color:#8B6F4E;font-weight:600;font-size:12px;text-align:center;padding:6px;">{label}</div>'
        if img_bytes_val:
            b64 = base64.b64encode(img_bytes_val).decode('utf-8')
            return f'<img src="data:image/jpeg;base64,{b64}" style="flex:1;min-width:140px;height:200px;object-fit:contain;background:#FFF;border-radius:6px;border:1px solid #E4D6CA;padding:4px;"/>'
        return f'<div style="flex:1;min-width:140px;height:200px;background:#F9F6F3;border:1px solid #E4D6CA;border-radius:6px;display:flex;align-items:center;justify-content:center;color:#C4B0A0;font-style:italic;font-size:12px;">No image for {label}</div>'

    images_html_parts = []
    if image_mappings:
        for img_tag, col_name in image_mappings.items():
            if is_template_mode:
                images_html_parts.append(_img_box(img_tag, None))
            else:
                # Find which pandas column index this col_name corresponds to
                col_idx = None
                if excel_row is not None and col_name:
                    cols = list(excel_row.keys())
                    if col_name in cols:
                        col_idx = cols.index(col_name)
                img_bytes_val = image_map.get((excel_row_idx, col_idx)) if col_idx is not None and excel_row_idx is not None else None
                images_html_parts.append(_img_box(img_tag, img_bytes_val))
    else:
        images_html_parts.append(_img_box('[IMAGE]', None))

    img_html = f'<div style="display:flex;gap:10px;flex-wrap:wrap;">{" ".join(images_html_parts)}</div>'
            
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
    components.html(
        f"""<!DOCTYPE html><html><head><link href='https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&display=swap' rel='stylesheet'></head><body style='margin:0;padding:0;background:transparent;'>{html_content}</body></html>""",
        height=420,
        scrolling=False
    )

def run_automation(excel_bytes, pptx_bytes, from_row, to_row, mapping_dict, image_mappings):
    """image_mappings: dict of {placeholder_tag: excel_col_name} for each image placeholder."""
    if image_mappings is None:
        image_mappings = {}
    df = pd.read_excel(io.BytesIO(excel_bytes))

    start_idx = max(0, from_row - 2)
    end_idx = min(to_row - 1, len(df))
    df_subset = df.iloc[start_idx:end_idx].copy()

    if df_subset.empty:
        raise ValueError("Selected row range contains no data.")

    # col_to_images: {col_0based: [img_bytes in row order]}
    col_to_images = get_excel_images(excel_bytes)
    df_cols = list(df.columns)

    prs = Presentation(io.BytesIO(pptx_bytes))
    if len(prs.slides) != 1:
        raise ValueError("Template must have exactly 1 slide.")

    source_slide = prs.slides[0]
    for _ in range(len(df_subset) - 1):
        clone_slide(prs, source_slide)

    for i, (original_idx, row) in enumerate(df_subset.iterrows()):
        slide = prs.slides[i]

        # ── Text replacements ──
        replacements = {}
        for tag, cfg in mapping_dict.items():
            col    = cfg.get("column", "")
            fmt    = cfg.get("format", "Text")
            symbol = cfg.get("symbol", "$")
            if not col:
                continue
            val = row.get(col, "")
            if pd.isna(val):
                val = ""
            if fmt == "Currency":
                formatted_val = safe_format(val, is_currency=True, currency_symbol=symbol)
            elif fmt == "Percentage":
                formatted_val = safe_format(val, is_percent=True)
            elif fmt == "Integer":
                formatted_val = safe_format(val)
            else:
                formatted_val = safe_text(val)
            replacements[tag] = formatted_val

        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

        # Remove blank bullet paragraphs left by empty-value replacements
        for shape in slide.shapes:
            purge_empty_paragraphs(shape)

        # ── Image placeholders ── positional: i-th product gets i-th image per col
        if image_mappings:
            for img_tag, col_name in image_mappings.items():
                if not col_name or col_name not in df_cols:
                    continue
                col_idx = df_cols.index(col_name)
                imgs = col_to_images.get(col_idx, [])
                img_bytes = imgs[i] if i < len(imgs) else None

                # Find the placeholder shape that still has the tag text
                ph_shape = None
                for shape in slide.shapes:
                    if shape.has_text_frame and img_tag in shape.text_frame.text:
                        ph_shape = shape
                        break
                if ph_shape:
                    if img_bytes:
                        insert_image_into_placeholder(slide, img_bytes, ph_shape)
                    else:
                        for para in ph_shape.text_frame.paragraphs:
                            for run in para.runs:
                                run.text = ""
                        ph_shape.fill.background()
        else:
            # Fallback: auto-detect by shape text
            placeholder = find_image_placeholder(slide)
            if placeholder:
                all_imgs = [b for imgs in col_to_images.values() for b in imgs]
                img_bytes = all_imgs[i] if i < len(all_imgs) else None
                if img_bytes:
                    insert_image_into_placeholder(slide, img_bytes, placeholder)
                else:
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
st.caption("Choose a pre-loaded template or upload your own custom `.pptx` file.")

available_templates = get_available_templates()

_template_source = st.radio(
    "Template Source",
    options=["Pre-loaded Templates", "Upload Custom Template"],
    horizontal=True,
    key="template_source"
)

pptx_bytes_custom = None  # holds bytes of a custom-uploaded template
selected_template = None

if _template_source == "Pre-loaded Templates":
    if not available_templates:
        st.warning("No templates found in the `templates/` folder. Please upload a custom template instead.")
    else:
        selected_template = st.selectbox(
            "Template",
            options=available_templates,
            format_func=lambda x: x.replace('.pptx', '').replace('_', ' ')
        )
else:
    custom_pptx = st.file_uploader(
        "Upload PowerPoint Template (.pptx)",
        type=["pptx"],
        help="Upload a single-slide .pptx file. Placeholders should use the [TAG] format.",
        key="custom_pptx_upload"
    )
    if custom_pptx:
        pptx_bytes_custom = custom_pptx.getvalue()
        selected_template = custom_pptx.name
        st.success(f"✔ Custom template loaded: **{custom_pptx.name}**")
    else:
        st.info("Upload a `.pptx` file to continue.")

def get_pptx_bytes():
    """Returns the bytes of the currently selected template."""
    if pptx_bytes_custom is not None:
        return pptx_bytes_custom
    if selected_template:
        return load_template(selected_template)
    return None

if selected_template is None and pptx_bytes_custom is None:
    st.stop()

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

# Auto-mapping is built silently once the Excel file is uploaded
mapping_dict = {}
image_mappings = {}  # {img_tag: excel_col_name}

if excel_file is not None:
    try:
        _excel_bytes_am = excel_file.getvalue()
        _pptx_bytes_am = get_pptx_bytes()
        _detected_tags = extract_placeholders_from_pptx(_pptx_bytes_am)
        _df_am = pd.read_excel(io.BytesIO(_excel_bytes_am))
        _excel_columns_am = list(_df_am.columns)
        mapping_dict, image_mappings = build_auto_mapping(_detected_tags, _excel_columns_am)
    except Exception as _e:
        st.warning(f"Auto-mapping could not be built: {_e}")

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
        default_fname = f"{(selected_template or 'Output').replace('.pptx', '')}_Output"
        fname = output_name.strip() or default_fname
        if not fname.endswith('.pptx'):
            fname += '.pptx'

        with st.spinner("Building your slides... please wait."):
            try:
                excel_bytes = excel_file.getvalue()
                pptx_bytes = get_pptx_bytes()

                result_bytes, count = run_automation(
                    excel_bytes, pptx_bytes,
                    from_row=int(from_row),
                    to_row=int(to_row),
                    mapping_dict=mapping_dict,
                    image_mappings=image_mappings
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
