# Optional Steps Reference

Detailed procedures for optional pipeline steps. Read only the section you need.

## Contents

- [Step 5.5: VLM Review](#step-55-vlm-review)
- [Step 6.5: Visual Verification](#step-65-visual-verification)
- [Step 7b: Content Completeness Audit](#step-7b-content-completeness-audit)

---

## Step 5.5: VLM Review

Run the review script:

```bash
uv run --with requests,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/review_dsl.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT"
```

**What this does**: Compares XML DSL against page images using LMStudio.

**Multi-page mode** (<=5 pages): All page PNGs + all XML DSLs sent in one API call for cross-page consistency. Timeout 300s.

**Per-page fallback** (>5 pages or multi-page failure): Each page reviewed independently. Timeout 300s per page.

Output: `$WORKSPACE/dsl/review-page-{N}.json`

### Handle review results

For each page:

1. Read `review-page-{N}.json`
2. If array is empty `[]`, do nothing
3. If array has more than 5 items → STOP, report the page number
4. Only apply fixes for these known issue types:
   - `missing_text` → add the missing content
   - `wrong_style` → update the relevant attributes
   - `wrong_order` → reorder the affected elements
   - `missing_image` → verify image path and add `<image>`
   - `extra_content` → remove the extra element
   - `missing_textframe` → wrap in `<text-frame has-border="true">`
   - `wrong_textframe` → correct text-frame attributes
   - `cross_page_inconsistency` → align style to consistent value
5. If any issue type is unknown or ambiguous → STOP, report the page number
6. Edit `page-{N}.xml` directly
7. Re-run Step 6 once after all fixes
8. Re-run Step 5.5 at most once only if user explicitly asks

---

## Step 6.5: Visual Verification

**Prerequisites**: `soffice` must be available.

```bash
command -v soffice >/dev/null && echo "OK: soffice available" || echo "SKIP: soffice not found"
```

If soffice is not found, return to SKILL.md and continue.

**Run verification:**

```bash
uv run --with requests,Pillow \
  .claude/skills/pdf-to-docx/scripts/verify_docx_visual.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT" \
  --docx "$WORKSPACE/output.docx"
```

**What this does**: Converts output.docx to PDF via soffice, renders to PNGs, VLM compares against original.

Output: `$WORKSPACE/dsl/visual-review-page-{N}.json`

**Page count mismatch handling**: Input PDF may have 2 pages but DOCX may render as 5 due to font/spacing differences. The script handles this internally.

### Handle visual review results

For each page:

1. Read `visual-review-page-{N}.json`
2. If array is empty `[]` → do nothing
3. If array has more than 5 items → STOP, report page number
4. If `page_count_mismatch` is present → STOP, report it
5. Only apply concrete fixes with clear `field` and `suggested_value`
6. If any issue is ambiguous → STOP, report page number
7. After fixing XML, re-run Step 6 once
8. Do NOT re-run Step 6.5 unless user explicitly asks

---

## Step 7b: Content Completeness Audit

Only run if the user explicitly asked for a completeness audit.

Use only as a hinting step. Do not make speculative edits from fuzzy matches. Only edit XML when you can point to a specific missing fragment and the target page.
