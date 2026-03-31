# Route B Reference

Detailed procedures, workspace layout, environment variables, and error handling for Route B (PDF / Image → VLM+OCR Pipeline). Read only the section you need.

## Contents

- [B6 Visual Verification](#b6-visual-verification)
- [Workspace Structure](#workspace-structure)
- [Environment Variables](#environment-variables)
- [Error Handling](#error-handling)

---

## B6 Visual Verification

Only run B6 if BOTH conditions are true:
- The user explicitly asked for visual QA / extra layout polishing
- `soffice` is available

**Step 1: Check soffice:**
```bash
command -v soffice >/dev/null && echo "B6 OK: soffice available" || echo "B6 SKIP: soffice not found"
```

If `B6 SKIP`, return to main SKILL.md and continue to B7.

**Step 2: Run visual verification:**
```bash
uv run --with requests,Pillow \
  .claude/skills/pdf-to-docx/scripts/verify_docx_visual.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT" \
  --docx "$WORKSPACE/output.docx"
```

**Step 3: Read review results and fix:**

Read each `$WORKSPACE/dsl/visual-review-page-{N}.json`. For each non-empty review:

| Issue `type` | Fix action |
|---|---|
| `font_difference` | Edit `<run>` attribute in `page-{N}.xml` |
| `missing_text` | Add content to `page-{N}.xml` |
| `layout_difference` | Edit spacing/margin attributes |
| `page_count_mismatch` | Reduce font sizes or margins |
| Any other | Apply the `suggested_value` from the JSON |

**Step 4:** After fixing XML, re-run B5 (dsl_to_docx.py) **once**. Do NOT re-run B6.

**Stop conditions:**
- If a review JSON is ambiguous or suggests more than 10 fixes on one page → STOP and report the page number
- If no issues found → return to main SKILL.md, continue to B7

---

## Workspace Structure

```
$WORKSPACE/
├── input-images/page-{N}.png       (reference PNGs, ≤1820px wide)
├── image-info.json                 (page count + dimensions)
├── glossary.json                   (merged glossary, if translation)
├── pdf-info.txt                    (PDF route only)
├── pdf-fulltext.txt                (PDF route only)
├── ocr-output/input/               (glm-ocr output)
│   ├── input.json, input.md, imgs/
├── dsl-vlm/page-{N}.xml           (VLM first-draft XML)
├── dsl/page-{N}.xml               (final merged XML)
├── dsl-translated/page-{N}.xml    (translated XML, if --lang)
├── texts.json                      (extracted segments)
├── translations.json               (translated segments)
├── output.docx                     (untranslated DOCX)
├── translated-output.docx          (translated DOCX, if --lang)
└── final-output.docx               (final deliverable)
```

---

## Environment Variables

VLM scripts read these env vars (all optional, have sensible defaults):

| Variable | Default | Description |
|---|---|---|
| `VLM_MODEL_PROFILE` | `strong` | `strong` or `weak` — controls prompts, batch sizes, merge strategy |
| `VLM_ENDPOINT` | `http://localhost:1234/v1/chat/completions` | OpenAI-compatible VLM API endpoint |
| `VLM_MODEL` | `qwen3.5-122b-a10b` | Model name sent in API request |
| `VLM_API_KEY` | `lm-studio` | API key (LMStudio ignores this) |
| `VLM_TIMEOUT` | `600` | Request timeout in seconds |
| `VLM_MAX_TOKENS` | profile-dependent | Max output tokens (strong: 131072, weak: 32768) |
| `VLM_TEMPERATURE` | 0.6 | Sampling temperature (both profiles use 0.6) |
| `VLM_RETRY_DELAY` | `120` | Seconds between retries |
| `VLM_MERGE_IMAGE_WIDTH` | `1200` | Max image width for merge step (px) |

### Weak Model Quick Setup (e.g., Qwen 3.5 35B-A3B)

```bash
export VLM_MODEL_PROFILE=weak
export VLM_MODEL="qwen3.5-35b-a3b"       # or your model name
export VLM_ENDPOINT="http://localhost:1234/v1/chat/completions"
```

This automatically uses:
- Simplified prompt with few-shot example (1856 chars vs 2780)
- Batch size 2 (vs 8)
- Temperature 0.3 (vs 0.6)
- Deterministic merge (no VLM merge call — saves 50% VLM usage)
- Enhanced XML repair for malformed output

---

## Error Handling

| Error | Action |
|---|---|
| Input file not found | Abort with message |
| soffice not found (DOC input) | Abort — required for conversion |
| pandoc not found (MD input) | Abort — required for conversion |
| glmocr fails | Re-run with `--log-level DEBUG`. Check input image quality. |
| VLM API unreachable | Check LMStudio is running. Scripts print FIX suggestions. |
| dsl_to_docx.py fails | Read traceback. Fix XML in `dsl/page-{N}.xml`. Re-run B5. |
| Translation mismatch | Auto-retry missing segments (max 3 retries in translation-prompt.md) |
| Output DOCX is 0 bytes | Check dsl/ XML validity. Re-run B5. |

### Structured Error Recovery

**If any step's verify command prints FAIL:**

1. Read the FAIL message and the FIX suggestion
2. If the fix is a re-run → run the same step command again (retry once)
3. If the fix requires checking a file → read that file, apply the fix
4. If the second attempt fails → STOP and report the error to the user
5. Do NOT skip a failed step and continue to the next step
