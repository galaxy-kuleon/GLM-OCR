# XML DSL Reference

## Contents

- [Page element](#page-element)
- [Element types](#element-types) (heading, paragraph, run, table, cell, image, text-frame, side-by-side)
- [Font mapping](#font-mapping)

## Page element

```xml
<page number="1" width-pts="595.276" height-pts="841.89"
      margin-top-cm="1.27" margin-bottom-cm="1.27"
      margin-left-cm="1.27" margin-right-cm="1.27"
      font-latin="Arial" font-cjk="SimSun">
```

## Element types

### heading

```xml
<heading level="1|2|3" alignment="left|center|right"
         font-family="sans|serif|mono"
         space-before-pt="12" space-after-pt="6" bg-color="HEXCOLOR">
  <run ...>text</run>
</heading>
```

### paragraph

```xml
<paragraph alignment="left|center|right|justify"
           space-before-pt="0" space-after-pt="0" line-spacing="1.0"
           indent-left-cm="0" indent-right-cm="0" indent-first-line-cm="0"
           list-level="1" list-type="bullet|number"
           font-family="sans|serif|mono" bg-color="HEXCOLOR"
           style="formula">
  <run ...>text</run>
</paragraph>
```

- `list-level`: nesting depth (1=top). Auto-sets left indent if `indent-left-cm` not specified (0.63cm per level — matches Word's default list indent of ~1/4 inch).
- `list-type`: `bullet` or `number`. Marker text is preserved in run content.
- `style="formula"`: enables LaTeX math processing in runs with `is-math="true"`.

### run (inline text)

```xml
<run font-size-pt="12"
     bold="true|false" italic="true|false" underline="true|false"
     superscript="true|false" subscript="true|false" strikethrough="true|false"
     color-rgb="R,G,B" highlight-color="HEXCOLOR"
     font-name="FontName"
     is-math="true" latex="\frac{a}{b}">
  visible text
</run>
```

| Attribute | Default | Description |
|---|---|---|
| `font-size-pt` | `11` | Font size in points |
| `bold` | `false` | Bold text |
| `italic` | `false` | Italic text |
| `underline` | `false` | Underlined text |
| `superscript` | `false` | Superscript (e.g., E=mc^2) |
| `subscript` | `false` | Subscript (e.g., H₂O) |
| `strikethrough` | `false` | Strikethrough text |
| `color-rgb` | `0,0,0` | Text color as R,G,B (0-255) |
| `highlight-color` | none | Background highlight as hex (e.g., `FFFF00`) |
| `font-name` | inherits page | Override font name |
| `is-math` | `false` | LaTeX math content (do NOT translate) |

### table

```xml
<table rows="3" cols="4" border-style="single|double|none"
       bbox="x1,y1,x2,y2" page-width-pts="595">
  <col-widths>0.25,0.25,0.25,0.25</col-widths>
  <row index="0" is-header="true">
    <cell row="0" col="0" colspan="1" rowspan="1"
          font-size-pt="10" bold="true" italic="false"
          alignment="center" color-rgb="R,G,B"
          bg-color="HEXCOLOR" text-bg-color="HEXCOLOR">
      Cell text
    </cell>
    <!-- or cell with per-run styling: -->
    <cell row="0" col="1" font-size-pt="10">
      <run font-size-pt="10" color-rgb="255,0,0" bold="true">keyword</run>
      <run font-size-pt="10"> normal text</run>
    </cell>
  </row>
</table>
```

- `bbox`: normalized 0-1000 coordinates
- `col-widths`: comma-separated ratios summing to ~1.0
- Cell with `<run>` children: per-run formatting. Cell without: uniform text.

### image

```xml
<image src="ocr-output/input/imgs/cropped_page0_idx0.jpg"
       bbox="x1,y1,x2,y2" page-width-pts="595"
       alignment="left|center|right" />
```

- `src`: relative to workspace
- `bbox`: normalized 0-1000

### text-frame

**Preferred (VLM path)** — uses bbox like images/tables:
```xml
<text-frame bbox="x1,y1,x2,y2" page-width-pts="W" page-height-pts="H"
            has-border="true|false" border-color="HEXCOLOR">
  <paragraph alignment="center">
    <run font-size-pt="11">Floating text</run>
  </paragraph>
</text-frame>
```

**Legacy (OCR path)** — direct twips positioning:
```xml
<text-frame x-twips="N" y-twips="N" width-twips="N" height-twips="N"
            has-border="true|false" border-color="HEXCOLOR">
  ...
</text-frame>
```

- `bbox`: normalized 0-1000 coordinates (same as images/tables). `dsl_to_docx.py` auto-converts to twips.
- 1 pt = 20 twips. Internally uses `w:framePr` + `w:pBdr` in OOXML.

### side-by-side

```xml
<side-by-side cols="2">
  <column>
    <paragraph><run>Left text</run></paragraph>
  </column>
  <column>
    <paragraph><run>Right text</run></paragraph>
  </column>
</side-by-side>
```

Rendered as invisible-border table in DOCX.

### horizontal-rule

```xml
<horizontal-rule />
<horizontal-rule color="FF0000" size="12" />
```

- `color`: hex or `auto` (default)
- `size`: border size in half-points (default `6`)

### page-header / page-footer

```xml
<page-header>
  <paragraph alignment="right">
    <run font-size-pt="9" color-rgb="128,128,128">Header text</run>
  </paragraph>
</page-header>

<page-footer>
  <paragraph alignment="center">
    <run font-size-pt="8">Page 1</run>
  </paragraph>
</page-footer>
```

Applied to the current document section. Contains `<paragraph>` + `<run>` children.

## Font mapping

Default fonts — chosen for maximum compatibility across Windows/macOS/Linux:

| `font-family` | Latin | CJK | Rationale |
|---|---|---|---|
| `serif` | Times New Roman | SimSun (宋體) | Standard serif pair, bundled with all OS |
| `sans` | Arial | SimHei (黑體) | Standard sans pair, highest availability |
| `mono` | Courier New | SimSun (宋體) | Mono latin + readable CJK fallback |

Override per-element with `font-name` attribute on `<run>` for non-standard fonts.
