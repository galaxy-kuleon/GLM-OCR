check: https://code.claude.com/docs/en/skills
and 
https://github.com/anthropics/skills/tree/main/skills/docx

make/develop/design/refine a skill called (pdf-to-docx) which can use something like:

`uv run glmocr parse examples-pdf/ElectronicPayrollFormat.pdf --output ./output/`

and use all the materials in output dir , to assemble a docx file, roughly the usable tools are:

1. pdftocairo (200 dpi)
  - use pdftocairo to convert input pdf to pngs (these pngs are called `input-pdf-rendered-pngs`)
2. other poppler tools for extrating pdf infos
3. `uv run glmocr parse ....` (pdf as input)
4. generate .py code where it has python-docx (and lxml, or the openxml lib inside python-docx) for assembling docx files
5. the way you evaluate the quality of the output docx file:
  - docx -> pdf with `soffice`
  - pdf -> pngs with `pdftocairo` , the pngs here is called `docx-rendered-pngs`
  - combine `output/*/layout_vis/*` and `docx-rendered-pngs` and `input-pdf-rendered-pngs` with proper prompts and instructions, and send to "ollama local + model `qwen3-vl:235b-cloud` (256k context)"
  - as new chat session (but need to append previous loops' checklists and fixed issues something, give qwen3-vl some memories), ask the model to give a check list of issues found, including but not limited to:
    - missing texts
    - missing textboxes
    - layout issues (texts/images/tables mis-aligned, wrong font size/bold/italic/underline, wrong color, wrong table border etc)
    - missing images
    - wrong table row/col merge/split
    - wrong table border
    - > must ask the model NOT TO answer scores and NOT TO give ambiguous review or suggestions.
    - > must ask the model to include precise detailed instructions on how to fix each issue found, the exact content/pixel position/size/font/color/border/align etc.
6. and review the result from 5. , and refine the .py code (which assembles the docx file) accordingly.
7. loop 4->6 until `docx-rendered-pngs` have visually extreme high fidelity with `input-pdf-rendered-pngs`. 

the new skill must be capable to do 1 ~ 7 above automatically, with zero human intervention.

and the input pdf can be in any language, 有各式各樣的圖片、表格、文字方塊，各式各樣的顏色、背景，奇奇怪怪的東西可能都有，複雜的內容.

the output docx must be editable in MS Word, LibreOffice Writer, WPS Writer etc.
