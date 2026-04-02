"""Shared constants for the anything-to-docx pipeline.

Every tunable parameter lives here as a named constant.
Source scripts import only what they need — no behavioral changes.
"""

# ===========================================================================
# MERGE: Text similarity weights and thresholds
# ===========================================================================

# Word count boundary: texts shorter than this use char-overlap-heavy weighting
MERGE_SHORT_TEXT_WORD_THRESHOLD = 5

# Similarity weighting for short text (< MERGE_SHORT_TEXT_WORD_THRESHOLD words)
MERGE_SHORT_TEXT_WORD_WEIGHT = 0.4    # word-Jaccard weight for short text
MERGE_SHORT_TEXT_CHAR_WEIGHT = 0.6    # char-overlap weight for short text

# Similarity weighting for long text (>= MERGE_SHORT_TEXT_WORD_THRESHOLD words)
MERGE_LONG_TEXT_WORD_WEIGHT = 0.7     # word-Jaccard weight for long text
MERGE_LONG_TEXT_CHAR_WEIGHT = 0.3     # char-overlap weight for long text

# Minimum base similarity before any bonuses are applied
MERGE_BASE_SIM_THRESHOLD = 0.2

# Maximum length ratio (longer/shorter) before a match is rejected
MERGE_LENGTH_RATIO_MAX = 5

# Minimum similarity to override length-ratio rejection
MERGE_LENGTH_RATIO_SIM_OVERRIDE = 0.5

# Bonus added when OCR native_label matches VLM element type
MERGE_LABEL_BONUS = 0.1

# Maximum cursor distance that earns a proximity bonus
MERGE_PROXIMITY_DISTANCE = 3

# Proximity bonus per unit of closeness to cursor
MERGE_PROXIMITY_BONUS_PER_UNIT = 0.05

# Minimum score to accept a text match
MERGE_ACCEPT_THRESHOLD = 0.25

# Search window radius around OCR cursor for greedy matching
MERGE_LOOK_AROUND = 6

# ===========================================================================
# MERGE: Superset / dedup detection for image regions
# ===========================================================================

# Containment ratio above which a smaller region is "inside" a larger one
MERGE_SUPERSET_CONTAINMENT_THRESHOLD = 0.7

# Minimum number of contained regions to declare a superset
MERGE_SUPERSET_MIN_CONTAINED = 2

# ===========================================================================
# MERGE: Duplicate element detection
# ===========================================================================

# Similarity threshold for consecutive duplicate heading/paragraph removal
MERGE_DUPLICATE_SIMILARITY = 0.8

# ===========================================================================
# MERGE: Quality gate (switch to OCR-only when VLM structure is too broken)
# ===========================================================================

# Minimum match rate (replaced/total) before switching to OCR-only fallback
MERGE_QUALITY_GATE_MATCH_RATE = 0.3

# Minimum OCR text regions required to trigger OCR-only fallback
MERGE_QUALITY_GATE_MIN_OCR_REGIONS = 2

# ===========================================================================
# MERGE: Gap filling
# ===========================================================================

# Minimum character count for an unmatched OCR region to be appended
MERGE_GAP_MIN_CHARS = 5

# ===========================================================================
# MERGE: Image IoU thresholds
# ===========================================================================

# Minimum IoU to accept an image bbox match
MERGE_IMAGE_IOU_THRESHOLD = 0.05

# Minimum IoU for table-cell OCR region overlap
MERGE_TABLE_CELL_IOU_THRESHOLD = 0.01

# ===========================================================================
# MERGE: Table cell text replacement
# ===========================================================================

# Minimum similarity for replacing a table cell's text with OCR text
MERGE_TABLE_CELL_SIMILARITY = 0.5

# ===========================================================================
# MERGE: Heading font sizes (level -> pt string)
# ===========================================================================

MERGE_HEADING_FONT_SIZES = {
    1: "18",
    2: "14",
    3: "12",
    4: "11",
    5: "11",
    6: "10",
}

# Default font-size-pt for body text and fallback heading sizes
MERGE_DEFAULT_FONT_SIZE_PT = "11"

# ===========================================================================
# MERGE: Poppler integration (digital PDF enhancement)
# ===========================================================================

# Minimum text similarity for accepting a poppler paragraph match
MERGE_POPPLER_MATCH_THRESHOLD = 0.3

# Poppler heading level determination by font size (pts)
# >18pt → H1, >=14pt → H2, >=12pt → H3, >=11pt+bold → H4
MERGE_POPPLER_H1_MIN_PTS = 18.0
MERGE_POPPLER_H2_MIN_PTS = 14.0
MERGE_POPPLER_H3_MIN_PTS = 12.0
MERGE_POPPLER_H4_MIN_PTS = 11.0  # only if bold

# ===========================================================================
# PAGE: Default page dimensions and margins
# ===========================================================================

# US Letter page dimensions in points (72 pts/inch)
PAGE_DEFAULT_WIDTH_PTS = 612

# US Letter page height in points
PAGE_DEFAULT_HEIGHT_PTS = 792

# Default margin in centimeters (all four sides)
PAGE_DEFAULT_MARGIN_CM = "1.27"

# Default Latin font family
PAGE_DEFAULT_FONT_LATIN = "Arial"

# Default CJK font family
PAGE_DEFAULT_FONT_CJK = "SimSun"

# ===========================================================================
# IMAGE: Processing dimensions and quality
# ===========================================================================

# Max width (px) for page images sent to weak VLM (reduces inference time)
IMAGE_WEAK_PAGE_MAX_WIDTH = 1024

# Max width (px) for layout visualization images sent to weak VLM
IMAGE_LAYOUT_VIS_MAX_WIDTH = 768

# JPEG quality for resized images sent to VLM
IMAGE_JPEG_QUALITY = 75

# ===========================================================================
# VLM: Default model configuration (overridable via env vars)
# ===========================================================================

# Default VLM model name
VLM_DEFAULT_MODEL = "qwen3.5-122b-a10b"

# Default LMStudio-compatible endpoint
VLM_DEFAULT_ENDPOINT = "http://localhost:1234/v1/chat/completions"

# Default API key (LMStudio ignores this but OpenAI-compat requires the header)
VLM_DEFAULT_API_KEY = "lm-studio"

# Request timeout in seconds (20 min — 122B model needs ~2-4 min/page on consumer GPU)
VLM_DEFAULT_TIMEOUT = 1200

# Maximum output tokens (Qwen3.5 supports up to 128K)
VLM_DEFAULT_MAX_TOKENS = 131072

# Delay in seconds between retries (GPU recovery from OOM/thermal throttle)
VLM_DEFAULT_RETRY_DELAY = 120

# Temperature for structured XML output (low = deterministic)
VLM_DEFAULT_TEMPERATURE = 0.6

# ===========================================================================
# VLM: Model profile configurations
# ===========================================================================

# Strong profile: larger batches, standard settings
VLM_PROFILE_STRONG_BATCH_SIZE = 8

# Weak profile: one page at a time for complete coverage
VLM_PROFILE_WEAK_BATCH_SIZE = 1
