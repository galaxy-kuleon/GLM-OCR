# Translation Procedure

This file is the COMPLETE translation procedure. Follow steps T1-T5 exactly. Do not invent alternate steps.

**CRITICAL**

- You perform the translation yourself.
- Do NOT write Python scripts that perform the translation.
- Do NOT call LMStudio or any external API for translation.
- When a step asks you to write JSON, write JSON only. No extra prose inside the JSON file.

## Non-negotiable translation rules

1. Keep the same segment count and the same `id` values.
2. Translate only text content.
3. Never insert new whitespace between a CJK character and an adjacent Latin letter, digit, or ASCII punctuation mark unless the source text already contains that exact whitespace.
4. Preserve normal spaces inside English phrases and between separate English words.
5. Preserve numbers, dates, proper nouns, brand names, and technical codes unless the target language has a standard localization.
6. Never translate mathematical formulas, LaTeX, code snippets, file paths, URLs, or product codes.
7. Preserve bullet or numbered markers at the start of a segment (`1.`, `2.`, `-`, `*`).
8. Apply glossary terms exactly as specified. Glossary overrides default translation.
9. If the source text is empty or whitespace-only, keep it empty.

Examples:

- Good: `使用OpenAI API`
- Bad: `使用 OpenAI API`
- Good: `支援Windows 11與macOS`
- Bad: `支援 Windows 11 與 macOS`
- Good: `版本v2.1已發布`
- Bad: `版本 v2.1 已發布`
- Good: `Call the OpenAI API` (keep normal English internal spaces)

**Prerequisite**: `$WORKSPACE/texts.json` must exist before starting.

## T1: Read inputs

```bash
TOTAL=$(python3 -c "import json; print(json.load(open('$WORKSPACE/texts.json', encoding='utf-8'))['total_segments'])")
echo "Total segments to translate: $TOTAL"

GLOSSARY=""
if [ -f "$WORKSPACE/glossary.json" ]; then
  GLOSSARY=$(python3 -c "import json; print(json.dumps(json.load(open('$WORKSPACE/glossary.json', encoding='utf-8')), ensure_ascii=False))")
  echo "Glossary loaded"
fi
```

Variable sources:

- `TARGET_LANG`: from `--lang`
- `STYLE_NOTES`: from `--style`
- `GLOSSARY`: from `$WORKSPACE/glossary.json` if present

## T2: Determine batches

Use `BATCH_SIZE=30`.

```bash
BATCH_SIZE=30
TOTAL_BATCHES=$(python3 -c "total=int('$TOTAL'); print((total + 29) // 30)")
echo "Total batches: $TOTAL_BATCHES"
```

You must finish every batch from `1` through `TOTAL_BATCHES`.

## T3: Translate every batch and write one JSON file per batch

For each batch number `M` from `1` to `TOTAL_BATCHES`:

1. Extract the batch input file:

```bash
export WORKSPACE
export BATCH_NO="$M"
python3 << 'PYEOF'
import json, os

workspace = os.environ["WORKSPACE"]
batch_no = int(os.environ["BATCH_NO"])
batch_size = 30

with open(os.path.join(workspace, "texts.json"), encoding="utf-8") as f:
    data = json.load(f)

segments = data["segments"]
start = (batch_no - 1) * batch_size
end = min(start + batch_size, len(segments))
batch = segments[start:end]

out_path = os.path.join(workspace, f"batch-input-{batch_no}.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(batch, f, ensure_ascii=False, indent=2)

print(out_path)
print(f"Batch {batch_no}: {len(batch)} segments")
PYEOF
```

2. Read `$WORKSPACE/batch-input-M.json`.
3. Translate that batch yourself.
4. Write `$WORKSPACE/batch-M.json` in this exact format:

```json
{
  "translations": [
    {"id": "p1:heading[0]/run[0]", "translated_text": "Translated text"},
    {"id": "p1:paragraph[0]/run[0]", "translated_text": "More translated text"}
  ]
}
```

Batch requirements:

- Same number of items as the batch input
- Same `id` values, unchanged
- No extra fields
- No missing items
- No extra prose
- Follow the CJK/Latin whitespace rule above

5. Validate the batch file immediately:

```bash
python3 -c "import json; src=json.load(open('$WORKSPACE/batch-input-$BATCH_NO.json', encoding='utf-8')); out=json.load(open('$WORKSPACE/batch-$BATCH_NO.json', encoding='utf-8')); actual=len(out['translations']); expected=len(src); print(f'Batch $BATCH_NO validation: {actual}/{expected}'); assert actual == expected"
```

6. Print progress: `Batch M/TOTAL_BATCHES complete`.

Do not stop early.

## T4: Merge and normalize all batch files into translations.json

After all normal batches are complete, run normalize_translations.py which discovers all
batch-N.json files, handles any variant format (canonical, segments, bare array, etc.),
normalizes every entry, merges them, validates against texts.json, and writes translations.json.

```bash
uv run .claude/skills/anything-to-docx/scripts/normalize_translations.py \
  --workspace "$WORKSPACE"
```

This writes `$WORKSPACE/translations.json` in canonical format. The script:
- Finds and sorts all `batch-N.json` files (excludes `batch-input-*` and `batch-recovery-*`)
- Normalizes variant formats (segments key, bare arrays, text vs translated_text)
- Detects and warns about duplicate IDs, non-dict entries, missing IDs
- Validates total count against `texts.json` if present
- Exits non-zero on validation mismatch or malformed batch files

**If exit code is non-zero**: inspect the warnings on stderr and fix the batch files before continuing.

## T5: Validate completeness

Run this exact check:

```bash
export WORKSPACE
python3 << 'PYEOF'
import json, os, sys

workspace = os.environ["WORKSPACE"]

with open(os.path.join(workspace, "texts.json"), encoding="utf-8") as f:
    texts = json.load(f)
with open(os.path.join(workspace, "translations.json"), encoding="utf-8") as f:
    trans = json.load(f)

expected = texts["total_segments"]
actual = len(trans["translations"])

if actual == expected:
    print(f"T5 OK: {actual}/{expected} segments translated")
    sys.exit(0)

source_ids = {s["id"] for s in texts["segments"]}
trans_ids = {t["id"] for t in trans["translations"]}
missing = [s for s in texts["segments"] if s["id"] not in trans_ids]
missing_ids = [s["id"] for s in missing]

print(f"T5 MISMATCH: {actual}/{expected}. Missing IDs:")
for item in missing_ids:
    print(f"  {item}")

recovery_path = os.path.join(workspace, "batch-input-recovery.json")
with open(recovery_path, "w", encoding="utf-8") as f:
    json.dump(missing, f, ensure_ascii=False, indent=2)

print(f"Recovery batch written to: {recovery_path}")
sys.exit(1)
PYEOF
```

If T5 fails:

1. Read `$WORKSPACE/batch-input-recovery.json`.
2. Translate those missing segments yourself.
3. Write `$WORKSPACE/batch-recovery-1.json` with the same `{"translations": [...]}` format.
4. Merge the recovery batch into `translations.json` with this exact command:

```bash
export WORKSPACE
python3 << 'PYEOF'
import json, os

workspace = os.environ["WORKSPACE"]

with open(os.path.join(workspace, "translations.json"), encoding="utf-8") as f:
    current = json.load(f)
with open(os.path.join(workspace, "batch-recovery-1.json"), encoding="utf-8") as f:
    recovery = json.load(f)

current["translations"].extend(recovery["translations"])

with open(os.path.join(workspace, "translations.json"), "w", encoding="utf-8") as f:
    json.dump(current, f, ensure_ascii=False, indent=2)

print(f"Merged {len(recovery['translations'])} recovery translations")
PYEOF
```

5. Re-run T5.
6. Maximum 3 recovery retries.

**When T5 passes**: return to the calling step (A2c or B7c).
