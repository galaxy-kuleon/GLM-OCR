# python-docx Assembly Guide

Technical reference for generating DOCX files from glmocr OCR output using `python-docx` and direct oxml manipulation.

---

## 1. Page Setup

### Points → EMU conversion

```python
from docx.shared import Emu, Pt, Inches, Cm

# 1 inch = 72 pts = 914400 EMU
# pts to EMU:
emu = int(pts * 914400 / 72)

# Or use the Pt helper (for font sizes etc.):
size = Pt(12)  # 12 point font
```

### Set page dimensions from pdfinfo

Parse `pdfinfo` output for "Page size:" line. The format is typically:
```
Page size:      595.276 x 841.89 pts (A4)
```

```python
from docx import Document
from docx.shared import Emu

def set_page_size(section, width_pts, height_pts):
    """Set page size from points (as reported by pdfinfo)."""
    section.page_width = Emu(int(width_pts * 914400 / 72))
    section.page_height = Emu(int(height_pts * 914400 / 72))

def set_margins(section, top_cm=1.27, bottom_cm=1.27, left_cm=1.27, right_cm=1.27):
    """Set small margins to maximize content area."""
    section.top_margin = Cm(top_cm)
    section.bottom_margin = Cm(bottom_cm)
    section.left_margin = Cm(left_cm)
    section.right_margin = Cm(right_cm)

doc = Document()
section = doc.sections[0]
set_page_size(section, 595.276, 841.89)  # A4
set_margins(section)
```

---

## 2. Coordinate System Conversion

### OCR bbox_2d → physical dimensions

glmocr `bbox_2d` is normalized to **0–1000** scale (NOT pixels).

```python
def bbox_to_emu(bbox_2d, page_width_emu, page_height_emu):
    """Convert normalized 0-1000 bbox to EMU coordinates.
    Use for image sizing and general dimension calculations."""
    x1, y1, x2, y2 = bbox_2d
    left = int(x1 * page_width_emu / 1000)
    top = int(y1 * page_height_emu / 1000)
    width = int((x2 - x1) * page_width_emu / 1000)
    height = int((y2 - y1) * page_height_emu / 1000)
    return left, top, width, height

def bbox_width_inches(bbox_2d, page_width_pts):
    """Get bbox width in inches for image sizing."""
    x1, _, x2, _ = bbox_2d
    width_pts = (x2 - x1) * page_width_pts / 1000
    return width_pts / 72.0
```

### OCR bbox_2d → TWIPS (for framePr)

**CRITICAL**: `w:framePr` uses **TWIPS** (1 pt = 20 twips), NOT EMU. Using EMU values in framePr will cause text frames to appear at wrong positions (typically top-left corner).

```python
def bbox_to_twips(bbox_2d, page_width_pts, page_height_pts):
    """Convert normalized 0-1000 bbox to TWIPS for use with w:framePr.
    1 pt = 20 twips."""
    x1, y1, x2, y2 = bbox_2d
    left = int(x1 * page_width_pts / 1000 * 20)
    top = int(y1 * page_height_pts / 1000 * 20)
    width = int((x2 - x1) * page_width_pts / 1000 * 20)
    height = int((y2 - y1) * page_height_pts / 1000 * 20)
    return left, top, width, height
```

---

## 3. Text Elements

### Headings

**IMPORTANT**: python-docx heading styles default to **blue** color. You MUST override to black after every `add_heading()` call.

```python
def add_heading_black(doc, text, level=1, font_name="Times New Roman", ea_font="SimSun"):
    """Add heading with color forced to black (python-docx defaults to blue!)."""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)  # Override blue default
        run.font.name = font_name
        if any('\u4e00' <= c <= '\u9fff' for c in text):
            set_east_asian_font(run, ea_font)
    return heading
```

### Paragraph spacing

**IMPORTANT**: Default python-docx paragraph spacing is too large and causes page count inflation. Always reduce spacing to match typical PDF text density.

```python
def set_tight_spacing(paragraph):
    """Reduce paragraph spacing to match PDF text density."""
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(1)
```

### Paragraphs with formatting

```python
import re

def add_formatted_paragraph(doc, text, font_name="Times New Roman", font_size_pt=11, ea_font="SimSun"):
    """Add a paragraph with Markdown bold/italic parsing and tight spacing."""
    para = doc.add_paragraph()

    # Strip heading markdown prefix
    text = re.sub(r'^#{1,6}\s+', '', text)

    # Tight spacing to prevent page count inflation
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after = Pt(1)

    # Parse markdown inline formatting
    # Pattern handles: **bold**, *italic*, ***bold+italic***
    pattern = r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|([^*]+))'

    for match in re.finditer(pattern, text):
        if match.group(2):  # ***bold+italic***
            run = para.add_run(match.group(2))
            run.bold = True
            run.italic = True
        elif match.group(3):  # **bold**
            run = para.add_run(match.group(3))
            run.bold = True
        elif match.group(4):  # *italic*
            run = para.add_run(match.group(4))
            run.italic = True
        elif match.group(5):  # plain text
            run = para.add_run(match.group(5))

        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        if any('\u4e00' <= c <= '\u9fff' for c in text):
            set_east_asian_font(run, ea_font)

    return para
```

### Bullet / numbered lists

```python
# Detect bullet: content starts with "- "
if content.startswith("- "):
    para = doc.add_paragraph(content[2:], style="List Bullet")
# Detect numbered: content starts with "(N) " or "N. "
elif re.match(r'^(\d+[\.\)]\s|\(\d+\)\s)', content):
    para = doc.add_paragraph(content, style="List Number")
else:
    para = doc.add_paragraph(content)
```

---

## 4. Tables

### Parse HTML table from OCR content (with line-break preservation)

```python
from lxml import html as lxml_html

def get_cell_text_with_breaks(cell_element):
    """Extract text from HTML cell, preserving <br> as newlines.
    CRITICAL: OCR tables may contain <br> for in-cell line breaks."""
    parts = []
    if cell_element.text:
        parts.append(cell_element.text)
    for child in cell_element:
        if child.tag == 'br':
            parts.append('\n')
        else:
            parts.append(child.text_content())
        if child.tail:
            parts.append(child.tail)
    return ''.join(parts).strip()

def parse_html_table(html_content):
    """Parse HTML table string into rows of cells.

    Returns: list of rows, each row is list of dicts:
        {"text": str, "rowspan": int, "colspan": int, "is_header": bool}
    """
    tree = lxml_html.fromstring(html_content)
    rows = []
    for tr in tree.iter("tr"):
        row = []
        for cell in tr:
            if cell.tag in ("td", "th"):
                text = get_cell_text_with_breaks(cell)  # Preserves <br>!
                rowspan = int(cell.get("rowspan", 1))
                colspan = int(cell.get("colspan", 1))
                row.append({
                    "text": text,
                    "rowspan": rowspan,
                    "colspan": colspan,
                    "is_header": cell.tag == "th"
                })
        if row:
            rows.append(row)
    return rows
```

### Build DOCX table (enhanced with styling, line breaks, column widths)

Use an **occupancy grid** to track merged cells. This is more reliable than sentinel values.

**IMPORTANT**: Many table cells show multi-line text in the original PDF due to column-width-forced wrapping, NOT explicit `<br>` tags. The strategy is:
1. Parse `<br>` tags with `get_cell_text_with_breaks()` (handles explicit breaks)
2. Set explicit column widths + `autofit = False` (forces natural word-wrapping)
3. Split cell text on `\n` and use `add_break()` for explicit newlines

```python
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

def set_cell_shading(cell, color_hex):
    """Set cell background color."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shading = parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{color_hex}"/>'
    )
    tcPr.append(shading)

def add_cell_text_with_breaks(cell, text, font_name="Times New Roman", font_size_pt=9,
                               ea_font="SimSun", bold=False, italic=False, color_rgb=None):
    """Add text to a cell, splitting on \\n into separate lines via add_break()."""
    cell.paragraphs[0].clear()
    para = cell.paragraphs[0]
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)

    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i > 0:
            br_run = para.add_run()
            br_run.add_break()
        run = para.add_run(line.strip())
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        if any('\u4e00' <= c <= '\u9fff' for c in line):
            set_east_asian_font(run, ea_font)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if color_rgb:
            run.font.color.rgb = RGBColor(*color_rgb)

def build_table(doc, html_content, font_name="Times New Roman", font_size_pt=9,
                ea_font="SimSun", style_data=None, bbox_2d=None, page_width_pts=595):
    """Build a DOCX table from HTML with styling, line breaks, and column widths.

    style_data: optional dict from VLM style extraction with keys:
        header_row_style: {bold, bg_color_rgb, color_rgb, font_size_pt}
        cell_styles: [{row, col, color_rgb, bold, italic}]
        border_style: "full" | "partial" | "none"
    """
    rows_data = parse_html_table(html_content)
    if not rows_data:
        return None

    max_cols = max(sum(c["colspan"] for c in r) for r in rows_data)
    num_rows = len(rows_data)
    table = doc.add_table(rows=num_rows, cols=max_cols)
    table.style = "Table Grid"

    # Set explicit column widths from bbox proportions to force word-wrapping
    table.autofit = False
    if bbox_2d:
        table_width_pts = (bbox_2d[2] - bbox_2d[0]) * page_width_pts / 1000
        col_width = Pt(table_width_pts / max_cols)
        for col in table.columns:
            col.width = col_width

    # Parse VLM style data
    header_style = (style_data or {}).get("header_row_style", {})
    cell_styles = {(cs["row"], cs["col"]): cs for cs in (style_data or {}).get("cell_styles", [])}

    # Occupancy grid to track merged cells
    occupied = [[False] * max_cols for _ in range(num_rows)]

    for r_idx, row in enumerate(rows_data):
        c_idx = 0
        for cell_data in row:
            while c_idx < max_cols and occupied[r_idx][c_idx]:
                c_idx += 1
            if c_idx >= max_cols:
                break

            cell = table.cell(r_idx, c_idx)
            text = cell_data["text"]

            # Determine cell-level style overrides
            cs = cell_styles.get((r_idx, c_idx), {})
            is_header = cell_data["is_header"] or r_idx == 0
            cell_bold = cs.get("bold", is_header or (is_header and header_style.get("bold", False)))
            cell_italic = cs.get("italic", False)
            cell_color = cs.get("color_rgb")
            cell_font_size = header_style.get("font_size_pt", font_size_pt) if is_header else font_size_pt

            # Apply header background color
            if is_header and header_style.get("bg_color_rgb"):
                bg = header_style["bg_color_rgb"]
                set_cell_shading(cell, f"{bg[0]:02X}{bg[1]:02X}{bg[2]:02X}")

            # Add text with line break preservation
            add_cell_text_with_breaks(
                cell, text, font_name=font_name, font_size_pt=cell_font_size,
                ea_font=ea_font, bold=cell_bold, italic=cell_italic, color_rgb=cell_color
            )

            # Mark occupancy
            for mr in range(r_idx, min(r_idx + cell_data["rowspan"], num_rows)):
                for mc in range(c_idx, min(c_idx + cell_data["colspan"], max_cols)):
                    occupied[mr][mc] = True

            # Merge cells
            end_r = min(r_idx + cell_data["rowspan"] - 1, num_rows - 1)
            end_c = min(c_idx + cell_data["colspan"] - 1, max_cols - 1)
            if end_r > r_idx or end_c > c_idx:
                cell.merge(table.cell(end_r, end_c))

            c_idx += cell_data["colspan"]

    return table
```

### Set table borders

```python
def set_table_borders(table, color="000000", size=4):
    """Set borders on all cells."""
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            borders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'  <w:top w:val="single" w:sz="{size}" w:space="0" w:color="{color}"/>'
                f'  <w:left w:val="single" w:sz="{size}" w:space="0" w:color="{color}"/>'
                f'  <w:bottom w:val="single" w:sz="{size}" w:space="0" w:color="{color}"/>'
                f'  <w:right w:val="single" w:sz="{size}" w:space="0" w:color="{color}"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(borders)
```

### Dense data tables

For tables with many rows/columns (e.g., payroll tables, spreadsheets), use smaller fonts and compact row heights:

```python
def set_compact_table_rows(table, row_height_twips=240):
    """Set minimum row heights for dense data tables."""
    for row in table.rows:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trH = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{row_height_twips}" w:hRule="atLeast"/>')
        trPr.append(trH)

# Usage: build_table(doc, html, font_size_pt=7) for dense tables
# Then: set_compact_table_rows(table)
```

---

## 5. Images

### Cropped image naming — CRITICAL

Image files from OCR are named `cropped_page{N}_idx{M}.jpg` where:
- `N` = page index (0-based)
- `M` = **sequential image counter** from `enumerate()` — **NOT the region's `index` field in the JSON**

You MUST track a separate counter per page that increments only for image-type regions:

```python
image_counter = 0  # Reset to 0 for each page
for region in page_regions:
    if region["label"] == "image":
        img_path = imgs_dir / f"cropped_page{page_idx}_idx{image_counter}.jpg"
        image_counter += 1
        # Use img_path for add_picture()
```

### Insert image with bbox-proportional sizing

```python
from docx.shared import Inches
import pathlib

def add_image(doc, image_path, bbox_2d, page_width_pts):
    """Insert an image, sized proportionally to its bbox width."""
    if not pathlib.Path(image_path).exists():
        doc.add_paragraph(f"[Image missing: {image_path}]")
        return

    # Calculate width from bbox
    x1, _, x2, _ = bbox_2d
    width_ratio = (x2 - x1) / 1000.0
    # Subtract margins (assuming ~1 inch total horizontal margins)
    usable_width_inches = (page_width_pts / 72.0) - 1.0
    img_width = width_ratio * usable_width_inches

    # Cap at usable width
    img_width = min(img_width, usable_width_inches)

    doc.add_picture(str(image_path), width=Inches(img_width))
```

---

## 6. Text Boxes / Floating Frames

For content that needs absolute positioning (floating text, bordered text boxes, captions overlaid on images):

### Preferred: `w:framePr` text frames (Word-compatible, no post-processing)

**Use `w:framePr` for ALL floating/positioned text**, including text boxes with borders. This approach:
- Works directly via python-docx oxml API (no ZIP post-processing)
- Is fully compatible with Microsoft Word
- Supports borders via `w:pBdr` (paragraph borders)
- All paragraphs sharing the same `framePr` position are grouped into one visual frame by Word

```python
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def emu_to_twips(emu):
    """Convert EMU to twips. 1 twip = 914400/72/20 = 635 EMU."""
    return round(emu / 635)

def add_frame_pr(paragraph, x_twips, y_twips, w_twips, h_twips):
    """Add w:framePr to a paragraph to make it a positioned text frame.
    Uses TWIPS (1 pt = 20 twips). NOT EMU."""
    pPr = paragraph._p.get_or_add_pPr()
    frame_pr = OxmlElement('w:framePr')
    frame_pr.set(qn('w:w'), str(w_twips))
    frame_pr.set(qn('w:h'), str(h_twips))
    frame_pr.set(qn('w:hRule'), 'exact')
    frame_pr.set(qn('w:hAnchor'), 'page')
    frame_pr.set(qn('w:vAnchor'), 'page')
    frame_pr.set(qn('w:x'), str(x_twips))
    frame_pr.set(qn('w:y'), str(y_twips))
    frame_pr.set(qn('w:wrap'), 'notBeside')
    pPr.insert(0, frame_pr)

def add_paragraph_borders(paragraph, color_hex="000000", size=4):
    """Add visible borders to a framePr paragraph (simulates text box border).

    Args:
        color_hex: 6-char hex color string (e.g. "FF0000" for red)
        size: border width in 1/8 pt (4 = 0.5pt, 8 = 1pt)
    """
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{size}" w:space="1" w:color="{color_hex}"/>'
        f'<w:left w:val="single" w:sz="{size}" w:space="1" w:color="{color_hex}"/>'
        f'<w:bottom w:val="single" w:sz="{size}" w:space="1" w:color="{color_hex}"/>'
        f'<w:right w:val="single" w:sz="{size}" w:space="1" w:color="{color_hex}"/>'
        f'</w:pBdr>')
    pPr.append(pBdr)
```

Usage example with borders and style JSON:
```python
# Check style JSON for border info
has_border = region_style.get("has_border", False)
border_color = region_style.get("border_color_rgb")  # [R, G, B] or None

p = doc.add_paragraph()
run = p.add_run(content)
# ... set font, size, color, etc. ...

add_frame_pr(p, x_twips, y_twips, w_twips, h_twips)

if has_border:
    bc_hex = f"{border_color[0]:02X}{border_color[1]:02X}{border_color[2]:02X}" if border_color else "000000"
    add_paragraph_borders(p, color_hex=bc_hex)
```

### DO NOT USE: OOXML text box via DOCX post-processing

**CRITICAL WARNING**: This approach produces files that **Microsoft Word CANNOT open**. It is documented here only as a reference for understanding the failure mode. **NEVER use this approach.**

Reasons it fails:
1. `mc:AlternateContent` requires BOTH `mc:Choice` (DrawingML/wps) AND `mc:Fallback` (VML/v:shape). Creating proper VML fallback XML is extremely complex.
2. `lxml etree.tostring()` re-serialization corrupts namespace declarations in document.xml, causing Word to reject the file.
3. Adding `xmlns:a` to the document root element breaks Word (Word rejects non-standard root namespace declarations from python-docx template).

The code below is kept for reference only — **DO NOT USE IT**:

**Text box XML structure** (uses EMU for positioning, `a:ln` for border):

```python
def make_textbox_xml(x_emu, y_emu, w_emu, h_emu, paragraphs_data, doc_pr_id=1,
                     border_color="000000", border_width_emu=12700):
    """Create w:p XML string for a proper OOXML text box.

    paragraphs_data: list of dicts with keys:
        text, font_name, font_size_half_pt, ea_font, color_hex, italic, bold
    border_width_emu: 12700 = 1pt
    """
    txbx_parts = []
    for pd in paragraphs_data:
        fn = pd.get("font_name", "Times New Roman")
        ea = pd.get("ea_font", "")
        rfonts = f'<w:rFonts w:ascii="{fn}" w:hAnsi="{fn}"'
        if ea:
            rfonts += f' w:eastAsia="{ea}"'
        rfonts += '/>'

        rpr_items = [rfonts]
        rpr_items.append(f'<w:sz w:val="{pd.get("font_size_half_pt", 22)}"/>')
        rpr_items.append(f'<w:szCs w:val="{pd.get("font_size_half_pt", 22)}"/>')
        if pd.get("color_hex"):
            rpr_items.append(f'<w:color w:val="{pd["color_hex"]}"/>')
        if pd.get("italic"):
            rpr_items.append('<w:i/>')
        if pd.get("bold"):
            rpr_items.append('<w:b/>')

        text = pd["text"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        rpr_xml = ''.join(rpr_items)
        txbx_parts.append(
            f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
            f'<w:r><w:rPr>{rpr_xml}</w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        )
    txbx_content = ''.join(txbx_parts)

    # border_width_emu: 12700 EMU = 1pt
    return f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="0" distR="0"
                     simplePos="0" relativeHeight="{251659264 + doc_pr_id}"
                     behindDoc="0" locked="0" layoutInCells="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page">
              <wp:posOffset>{x_emu}</wp:posOffset>
            </wp:positionH>
            <wp:positionV relativeFrom="page">
              <wp:posOffset>{y_emu}</wp:posOffset>
            </wp:positionV>
            <wp:extent cx="{w_emu}" cy="{h_emu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="{doc_pr_id}" name="Text Box {doc_pr_id}"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="{w_emu}" cy="{h_emu}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:noFill/>
                    <a:ln w="{border_width_emu}">
                      <a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill>
                    </a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>{txbx_content}</w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square"
                              lIns="91440" tIns="45720" rIns="91440" bIns="45720"
                              anchor="t" anchorCtr="0" upright="1">
                    <a:noAutofit/>
                  </wps:bodyPr>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'''
```

**Post-processing function** to inject text boxes into the saved DOCX:

```python
import zipfile
from lxml import etree

def postprocess_add_textboxes(input_docx, output_docx, textbox_xmls):
    """Add text boxes by modifying document.xml inside the DOCX ZIP."""
    with zipfile.ZipFile(input_docx, 'r') as zin:
        with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == 'word/document.xml':
                    root = etree.fromstring(data)
                    body = root.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')

                    # Ensure namespaces are declared
                    for prefix, uri in {
                        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
                        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
                    }.items():
                        if root.nsmap.get(prefix) is None:
                            root.attrib[f'{{http://www.w3.org/2000/xmlns/}}{prefix}'] = uri

                    ignorable = root.get(
                        '{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable', '')
                    for ns in ['wps', 'wp14']:
                        if ns not in ignorable:
                            ignorable = (ignorable + f' {ns}').strip()
                    root.set(
                        '{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable',
                        ignorable)

                    sect_pr = body.find(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
                    for tb_xml in textbox_xmls:
                        elem = etree.fromstring(tb_xml.encode('utf-8'))
                        if sect_pr is not None:
                            sect_pr.addprevious(elem)
                        else:
                            body.append(elem)

                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
```

### Position and size conversion for text boxes

Text box positioning uses **EMU** (unlike framePr which uses TWIPS):

```python
def bbox_to_textbox_emu(bbox_2d, page_width_emu, page_height_emu, padding=10):
    """Convert normalized 0-1000 bbox to EMU for text box positioning.
    padding: extra space around content in normalized units."""
    x1, y1, x2, y2 = bbox_2d
    x1, y1 = max(0, x1 - padding), max(0, y1 - padding)
    x2, y2 = min(1000, x2 + padding), min(1000, y2 + padding)
    return (int(x1 * page_width_emu / 1000),
            int(y1 * page_height_emu / 1000),
            int((x2 - x1) * page_width_emu / 1000),
            int((y2 - y1) * page_height_emu / 1000))
```

### Preferred for side-by-side WITHOUT borders: Invisible Table

The most reliable approach for side-by-side content (signature blocks, two-column layouts) is an invisible table. This works consistently across Word and LibreOffice.

```python
def add_side_by_side_block(doc, left_lines, right_lines, font_name="Times New Roman",
                            font_size_pt=11, ea_font="SimSun"):
    """Add side-by-side content using an invisible table (most reliable method).

    left_lines: list of strings for left column
    right_lines: list of strings for right column
    """
    num_rows = max(len(left_lines), len(right_lines))
    table = doc.add_table(rows=num_rows, cols=2)
    table.autofit = False

    # Set invisible borders
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcPr.append(parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="none" w:sz="0" w:space="0"/>'
                f'<w:left w:val="none" w:sz="0" w:space="0"/>'
                f'<w:bottom w:val="none" w:sz="0" w:space="0"/>'
                f'<w:right w:val="none" w:sz="0" w:space="0"/>'
                f'</w:tcBorders>'))

    for i in range(num_rows):
        for col_idx, lines in enumerate([left_lines, right_lines]):
            if i < len(lines):
                cell = table.cell(i, col_idx)
                cell.paragraphs[0].clear()
                run = cell.paragraphs[0].add_run(lines[i])
                run.font.name = font_name
                run.font.size = Pt(font_size_pt)
                if any('\u4e00' <= c <= '\u9fff' for c in lines[i]):
                    set_east_asian_font(run, ea_font)

    return table
```

### Note on framePr positioning reliability

`w:framePr` positions may differ slightly between Word and LibreOffice. For **side-by-side layouts without positioning needs** (e.g., signature blocks, two-column text), prefer invisible tables (see below). But for **floating text boxes that need absolute positioning** (especially with borders), `w:framePr` + `w:pBdr` is the only Word-compatible approach.

### Floating text detection

```python
def is_floating_text(bbox_2d, page_regions, current_index):
    """Detect floating text: side-by-side or vertically stacked in text box."""
    x1, y1, x2, y2 = bbox_2d
    for other in page_regions:
        if other["index"] == current_index or not other.get("bbox_2d"):
            continue
        ox1, oy1, ox2, oy2 = other["bbox_2d"]
        # Side-by-side: similar Y, non-overlapping X
        if abs(y1 - oy1) < 50 and (x1 > ox2 or x2 < ox1):
            return True
        # Vertically stacked with overlapping X (text box with multiple lines)
        x_overlap = min(x2, ox2) - max(x1, ox1)
        x_min_width = min(x2 - x1, ox2 - ox1)
        if x_min_width > 0 and x_overlap / x_min_width > 0.5:
            y_gap = abs(y1 - oy2) if y1 > oy2 else abs(oy1 - y2)
            if y_gap < 50 and min(x1, ox1) > 300:  # Right side of page
                return True
    return False
```

---

## 7. Fonts and Colors

### Font mapping

```python
def get_font_name(content, pdf_fonts_info):
    """Determine appropriate font based on content and PDF font info."""
    has_cjk = any('\u4e00' <= c <= '\u9fff' for c in content)

    if has_cjk:
        for font in pdf_fonts_info:
            if any(kw in font.lower() for kw in ['song', 'hei', 'kai', 'ming', 'gothic', 'cjk']):
                return font
        return "SimSun"  # Fallback CJK font
    else:
        for font in pdf_fonts_info:
            if any(kw in font.lower() for kw in ['arial', 'helvetica', 'times', 'calibri']):
                return font
        return "Arial"  # Fallback Latin font
```

### Set East Asian font (requires oxml)

```python
def set_east_asian_font(run, font_name):
    """Set East Asian font using oxml."""
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
```

### Colors

```python
from docx.shared import RGBColor

# Set text color
run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # Red

# Set paragraph shading (background)
def set_paragraph_shading(paragraph, color_hex="FFFF00"):
    """Set paragraph background color."""
    pPr = paragraph._p.get_or_add_pPr()
    shading = parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{color_hex}"/>'
    )
    pPr.append(shading)
```

---

## 8. Font Size Estimation

### WARNING: bbox-based estimation fails for wrapped text

**DO NOT blindly estimate font size from bbox height**. This fails catastrophically for long paragraphs where content has no `\n` but wraps across many visual lines. The bbox covers the entire wrapped block, so dividing by line count (1) gives absurdly large sizes (e.g., 87pt instead of 12pt), causing massive page count inflation.

### Smart font size strategy

Use bbox-based estimation **ONLY** for short text (< 60 chars) with small bbox height (< 30 pts). For everything else, use defaults by `native_label`:

```python
FONT_SIZES = {
    "doc_title": 14,
    "paragraph_title": 13,
    "text": 11,
    "figure_title": 10,
    "vision_footnote": 9,
    "table": 7,  # for dense data tables
}

def smart_font_size(bbox_2d, content, native_label, page_height_pts):
    """Estimate font size safely, falling back to defaults for long text."""
    default = FONT_SIZES.get(native_label, 11)
    if bbox_2d and len(content) < 60:
        height_pts = (bbox_2d[3] - bbox_2d[1]) * page_height_pts / 1000.0
        if height_pts < 30:
            return max(8, min(36, height_pts / 1.2))
    return default
```

---

## 9. Page Breaks

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_page_break(doc):
    """Add a page break."""
    para = doc.add_paragraph()
    run = para.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)
    return para
```

---

## 10. Common Pitfalls

1. **Heading color defaults to blue**: python-docx heading styles (`Heading 1`, `Heading 2`, etc.) render in blue by default. Always set `run.font.color.rgb = RGBColor(0, 0, 0)` after `add_heading()`.

2. **framePr uses TWIPS, not EMU**: The `w:framePr` element expects values in TWIPS (1 pt = 20 twips). Using EMU values causes frames to appear at wrong positions or disappear entirely.

3. **Paragraph spacing causes page inflation**: Default Word spacing is too large for PDF-faithful reproduction. Set `space_before = Pt(1)` and `space_after = Pt(1)` on all paragraphs to match PDF density.

4. **Image idx is sequential counter, NOT JSON region index**: OCR output images are `cropped_page{N}_idx{M}.jpg` where M is a sequential counter from `enumerate()`, not the region's `index` field. Track a separate counter per page.

5. **Font size estimation fails for wrapped text**: Long paragraphs have large bbox heights but the content doesn't contain `\n` for wrapped lines. Dividing bbox height by 1 line gives absurd font sizes. Use the `smart_font_size` pattern instead.

6. **OCR doesn't provide text styling**: Color, bold, italic, underline, font names, and font sizes are NOT available from OCR output. Use reasonable defaults and rely on VLM evaluation feedback for corrections.

7. **Merge sentinel values**: When merging cells, pre-existing text in target cells persists. Use an occupancy grid to track merged cells rather than sentinel values.

8. **oxml namespace**: Always use `nsdecls("w")` when building XML strings. Missing namespaces cause silent failures.

9. **Image paths**: Use `pathlib.Path` for cross-platform compatibility. Always check file existence before `add_picture()`.

10. **EMU integer overflow**: EMU values can be large. Always use `int()` to avoid float values in XML.

11. **Table column width**: By default, python-docx auto-sizes columns. To set explicit widths:
    ```python
    from docx.shared import Cm
    table.columns[0].width = Cm(3)
    ```

12. **Dense data tables**: For tables with many rows/columns, use 7-8pt font, compact cell margins (`space_before=Pt(0)`, `space_after=Pt(0)`), and small row heights (240 twips) to prevent excessive page count.

13. **Ordinal superscripts (15th, 1st, 2nd, 3rd)**: OCR may output ordinals as plain text or LaTeX. Use this pattern:
    ```python
    import re
    ORDINAL_RE = re.compile(r"\b(\d+)(th|st|nd|rd)\b")
    def add_text_with_superscript(para, text, font_size_pt=11):
        """Split text on ordinal suffixes and apply superscript."""
        last_end = 0
        for m in ORDINAL_RE.finditer(text):
            if m.start() > last_end:
                run = para.add_run(text[last_end:m.start()])
                run.font.size = Pt(font_size_pt)
            # Number part
            run = para.add_run(m.group(1))
            run.font.size = Pt(font_size_pt)
            # Suffix part (superscript)
            run = para.add_run(m.group(2))
            run.font.size = Pt(int(font_size_pt * 0.65))
            run.font.superscript = True
            last_end = m.end()
        if last_end < len(text):
            run = para.add_run(text[last_end:])
            run.font.size = Pt(font_size_pt)
    ```

14. **Per-run VLM style application**: When VLM style JSON has a `runs` array for a region, split the content and apply per-run styles:
    ```python
    def apply_runs_style(para, content, runs_data, default_font_size=11):
        """Apply mixed styles from VLM runs array."""
        if not runs_data:
            run = para.add_run(content)
            run.font.size = Pt(default_font_size)
            return
        pos = 0
        for rd in runs_data:
            snippet = rd.get("text_snippet", "")
            idx = content.find(snippet, pos)
            if idx > pos:
                # Add text before this run with default style
                run = para.add_run(content[pos:idx])
                run.font.size = Pt(default_font_size)
            if idx >= 0:
                run = para.add_run(snippet)
                run.font.size = Pt(default_font_size)
                if rd.get("bold"): run.bold = True
                if rd.get("italic"): run.italic = True
                if rd.get("color_rgb"):
                    run.font.color.rgb = RGBColor(*rd["color_rgb"])
                pos = idx + len(snippet)
        if pos < len(content):
            run = para.add_run(content[pos:])
            run.font.size = Pt(default_font_size)
    ```

15. **Keyword-based cell coloring**: For table cells with specific keywords that should be colored (e.g., "全型" in red):
    ```python
    def add_cell_text_with_keyword_color(cell, text, keyword, color_rgb,
                                          font_size_pt=9, bold_keyword=True):
        """Split text on keyword and apply color to matching fragments."""
        cell.paragraphs[0].clear()
        para = cell.paragraphs[0]
        parts = re.split(f"({re.escape(keyword)})", text)
        for part in parts:
            if not part:
                continue
            run = para.add_run(part)
            run.font.size = Pt(font_size_pt)
            if part == keyword:
                run.font.color.rgb = RGBColor(*color_rgb)
                if bold_keyword:
                    run.bold = True
    ```

16. **Side-by-side layouts**: Use invisible tables (Section 6) for side-by-side content without positioning needs. For floating/positioned text (including bordered text boxes), use `w:framePr` + `w:pBdr`. **NEVER use OOXML `wp:anchor` text boxes** — they produce files Word cannot open (see Section 6 critical warning).

17. **OOXML text box post-processing is BROKEN**: ZIP post-processing with `lxml etree.tostring()` corrupts namespace declarations. `mc:AlternateContent` requires VML `mc:Fallback` which is extremely complex to generate correctly. Always use `w:framePr` instead.

18. **Text box borders with framePr**: Use `w:pBdr` (paragraph borders) on framePr paragraphs to simulate text box borders. Check style JSON `has_border` and `border_color_rgb` fields. All paragraphs in a frame should have matching border settings for a clean visual result.
