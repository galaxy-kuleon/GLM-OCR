#!/usr/bin/env python3
"""
Font Metrics Utility - Extract font metrics and normalize font names.

Uses fontTools to:
1. Map PDF internal font names to proper family names
2. Extract ascender/descender/line gap for accurate line spacing
3. Handle CJK fonts specially

This improves upon simple font_size * 1.2 estimates.
"""

import os
import re
from pathlib import Path
from typing import Optional, Dict, Tuple
from dataclasses import dataclass

try:
    from fontTools.ttLib import TTFont
    HAS_FONTTOOLS = True
except ImportError:
    HAS_FONTTOOLS = False


@dataclass
class FontMetrics:
    """Font metrics extracted from font file."""
    family_name: str
    full_name: str
    ascender: float  # In font units
    descender: float  # In font units (negative)
    line_gap: float  # In font units
    units_per_em: int
    
    @property
    def ascender_pt(self) -> float:
        """Ascender as fraction of em size."""
        return self.ascender / self.units_per_em if self.units_per_em > 0 else 0.8
    
    @property
    def descender_pt(self) -> float:
        """Descender as fraction of em size (negative)."""
        return self.descender / self.units_per_em if self.units_per_em > 0 else -0.2
    
    @property
    def line_height_ratio(self) -> float:
        """Natural line height as multiple of font size."""
        return (self.ascender - self.descender + self.line_gap) / self.units_per_em if self.units_per_em > 0 else 1.2


# Common PDF font name mappings
# PDF internal names -> (family_name, is_bold, is_italic)
FONT_NAME_MAP = {
    # Arial variants
    'ArialMT': ('Arial', False, False),
    'Arial-BoldMT': ('Arial', True, False),
    'Arial-ItalicMT': ('Arial', False, True),
    'Arial-BoldItalicMT': ('Arial', True, True),
    'Arial-Bold': ('Arial', True, False),
    'Arial-Italic': ('Arial', False, True),
    
    # Times variants
    'TimesNewRomanPSMT': ('Times New Roman', False, False),
    'TimesNewRomanPS-BoldMT': ('Times New Roman', True, False),
    'TimesNewRomanPS-ItalicMT': ('Times New Roman', False, True),
    'TimesNewRomanPS-BoldItalicMT': ('Times New Roman', True, True),
    'TimesNewRoman': ('Times New Roman', False, False),
    'TimesNewRoman-Bold': ('Times New Roman', True, False),
    'TimesNewRoman-Italic': ('Times New Roman', False, True),
    
    # Courier variants
    'CourierNewPSMT': ('Courier New', False, False),
    'CourierNewPS-BoldMT': ('Courier New', True, False),
    'CourierNewPS-ItalicMT': ('Courier New', False, True),
    'CourierNewPS-BoldItalicMT': ('Courier New', True, True),
    'Courier': ('Courier', False, False),
    'Courier-Bold': ('Courier', True, False),
    
    # Helvetica variants
    'Helvetica': ('Helvetica', False, False),
    'Helvetica-Bold': ('Helvetica', True, False),
    'Helvetica-Oblique': ('Helvetica', False, True),
    'Helvetica-BoldOblique': ('Helvetica', True, True),
    
    # Liberation fonts (open source equivalents)
    'LiberationSans': ('Liberation Sans', False, False),
    'LiberationSans-Bold': ('Liberation Sans', True, False),
    'LiberationSans-Italic': ('Liberation Sans', False, True),
    'LiberationSerif': ('Liberation Serif', False, False),
    'LiberationSerif-Bold': ('Liberation Serif', True, False),
    'LiberationMono': ('Liberation Mono', False, False),
    
    # Calibri
    'Calibri': ('Calibri', False, False),
    'Calibri-Bold': ('Calibri', True, False),
    'Calibri-Italic': ('Calibri', False, True),
    
    # CJK fonts
    'MSung-Light': ('Ming', False, False),
    'MSungStd-Light': ('Ming', False, False),
    'HeiseiMin-W3': ('Hiragino Mincho', False, False),
    'HeiseiKakuGo-W5': ('Hiragino Sans', False, False),
    'STSong-Light': ('Song', False, False),
    'STSongStd-Light': ('Song', False, False),
    'AdobeSongStd-Light': ('Adobe Song Std', False, False),
    'AdobeMingStd-Light': ('Adobe Ming Std', False, False),
}

# Subset font prefix pattern (e.g., "ABCDEF+ArialMT")
SUBSET_PATTERN = re.compile(r'^[A-Z]{6}\+(.+)$')

# Font search directories
FONT_DIRS = [
    '/System/Library/Fonts',
    '/System/Library/Fonts/Supplemental',
    '/Library/Fonts',
    os.path.expanduser('~/Library/Fonts'),
    '/usr/share/fonts',
    '/usr/local/share/fonts',
]


def normalize_font_name(pdf_font_name: str) -> Tuple[str, bool, bool]:
    """
    Normalize PDF internal font name to proper family name.
    
    Args:
        pdf_font_name: Font name from PDF (e.g., "ABCDEF+ArialMT", "TimesNewRomanPS-BoldMT")
    
    Returns:
        (family_name, is_bold, is_italic)
    """
    # Remove subset prefix (e.g., "ABCDEF+ArialMT" -> "ArialMT")
    match = SUBSET_PATTERN.match(pdf_font_name)
    if match:
        base_name = match.group(1)
    else:
        base_name = pdf_font_name
    
    # Remove any trailing comma and style suffix (e.g., "Calibri,Bold" -> "Calibri-Bold")
    base_name = base_name.replace(',', '-')
    
    # Check direct mapping
    if base_name in FONT_NAME_MAP:
        return FONT_NAME_MAP[base_name]
    
    # Try to infer from name patterns
    is_bold = any(kw in base_name.lower() for kw in ['bold', 'black', 'heavy'])
    is_italic = any(kw in base_name.lower() for kw in ['italic', 'oblique'])
    
    # Extract family name by removing style keywords
    family = base_name
    for kw in ['Bold', 'Italic', 'Oblique', 'Regular', 'Light', 'Medium', 
               'Black', 'Heavy', 'Thin', 'Condensed', 'MT', 'PS', 'PSMT']:
        family = family.replace(kw, '')
    family = family.strip('-').strip()
    
    if not family:
        family = base_name
    
    return (family, is_bold, is_italic)


def is_cjk_font(font_name: str) -> bool:
    """Check if font name indicates a CJK font."""
    cjk_keywords = ['sung', 'ming', 'song', 'hei', 'gothic', 'kai', 'fang',
                    'cjk', 'noto', 'han', 'jp', 'kr', 'sc', 'tc', 'chinese',
                    'japanese', 'korean', 'pingfang', 'hiragino', 'yu']
    name_lower = font_name.lower()
    return any(kw in name_lower for kw in cjk_keywords)


def find_font_file(font_name: str) -> Optional[Path]:
    """
    Find font file on system by name.
    
    Args:
        font_name: Font family name (e.g., "Arial", "Times New Roman")
    
    Returns:
        Path to font file, or None if not found
    """
    if not HAS_FONTTOOLS:
        return None
    
    # Normalize name for matching
    name_lower = font_name.lower().replace(' ', '')
    
    for font_dir in FONT_DIRS:
        if not os.path.exists(font_dir):
            continue
        
        for filename in os.listdir(font_dir):
            if not filename.lower().endswith(('.ttf', '.otf', '.ttc')):
                continue
            
            # Check if filename matches
            file_lower = filename.lower().replace(' ', '').replace('.ttf', '').replace('.otf', '').replace('.ttc', '')
            if name_lower in file_lower or file_lower in name_lower:
                return Path(font_dir) / filename
    
    return None


def get_font_metrics(font_name: str) -> Optional[FontMetrics]:
    """
    Get font metrics from system font file.
    
    Args:
        font_name: Font family name (e.g., "Arial", "Times New Roman")
    
    Returns:
        FontMetrics object, or None if font not found or fontTools unavailable
    """
    if not HAS_FONTTOOLS:
        return None
    
    font_file = find_font_file(font_name)
    if font_file is None:
        return None
    
    try:
        # Handle TTC (TrueType Collection) files
        font_path = str(font_file)
        if font_path.lower().endswith('.ttc'):
            # Try each font in the collection
            from fontTools.ttLib import TTCollection
            ttc = TTCollection(font_path)
            font = None
            for f in ttc.fonts:
                name_table = f.get('name')
                if name_table:
                    for record in name_table.names:
                        if record.nameID == 1 and record.platformID == 3:
                            if font_name.lower() in record.toUnicode().lower():
                                font = f
                                break
                if font:
                    break
            if font is None:
                font = ttc.fonts[0]  # Use first font as fallback
        else:
            font = TTFont(font_path)
        
        # Get name table
        name_table = font.get('name')
        family_name = font_name
        full_name = font_name
        
        if name_table:
            # Name ID 1 = Family name, 4 = Full name
            for record in name_table.names:
                if record.nameID == 1 and record.platformID == 3:  # Windows
                    family_name = record.toUnicode()
                elif record.nameID == 4 and record.platformID == 3:
                    full_name = record.toUnicode()
        
        # Get OS/2 table for metrics
        os2 = font.get('OS/2')
        if os2:
            ascender = os2.sTypoAscender
            descender = os2.sTypoDescender
            line_gap = os2.sTypoLineGap
        else:
            # Fallback to hhea table
            hhea = font.get('hhea')
            if hhea:
                ascender = hhea.ascent
                descender = hhea.descent
                line_gap = hhea.lineGap
            else:
                # Default values
                ascender = 800
                descender = -200
                line_gap = 0
        
        units_per_em = font['head'].unitsPerEm
        
        font.close()
        
        return FontMetrics(
            family_name=family_name,
            full_name=full_name,
            ascender=ascender,
            descender=descender,
            line_gap=line_gap,
            units_per_em=units_per_em
        )
    except Exception as e:
        print(f"Warning: Could not read font metrics for {font_name}: {e}")
        return None


def get_line_height_ratio(font_name: str, font_size_pt: float) -> float:
    """
    Get recommended line height ratio for a font.
    
    Uses font metrics if available, otherwise falls back to defaults.
    
    Args:
        font_name: Font family name
        font_size_pt: Font size in points
    
    Returns:
        Line height as multiple of font size (e.g., 1.2 means 120%)
    """
    metrics = get_font_metrics(font_name)
    
    if metrics:
        # Use font's natural line height
        ratio = metrics.line_height_ratio
        # Clamp to reasonable range
        return max(1.0, min(2.0, ratio))
    
    # Fallback defaults
    if is_cjk_font(font_name):
        # CJK fonts typically need more line spacing
        return 1.4
    else:
        # Standard Latin fonts
        return 1.2


# Cache for font metrics
_metrics_cache: Dict[str, Optional[FontMetrics]] = {}


def get_font_metrics_cached(font_name: str) -> Optional[FontMetrics]:
    """Get font metrics with caching."""
    if font_name not in _metrics_cache:
        _metrics_cache[font_name] = get_font_metrics(font_name)
    return _metrics_cache[font_name]


if __name__ == '__main__':
    # Test font name normalization
    test_names = [
        'ABCDEF+ArialMT',
        'TimesNewRomanPS-BoldMT',
        'Calibri,Bold',
        'LiberationSans-Italic',
        'STSong-Light',
        'HeiseiMin-W3',
    ]
    
    print("Font name normalization:")
    for name in test_names:
        family, bold, italic = normalize_font_name(name)
        styles = []
        if bold:
            styles.append('bold')
        if italic:
            styles.append('italic')
        style_str = ', '.join(styles) if styles else 'regular'
        print(f"  {name} -> {family} ({style_str})")
    
    print("\nFont metrics:")
    test_fonts = ['Arial', 'Times New Roman', 'Courier New']
    for font in test_fonts:
        metrics = get_font_metrics(font)
        if metrics:
            print(f"  {metrics.family_name}:")
            print(f"    Line height ratio: {metrics.line_height_ratio:.2f}")
            print(f"    Ascender: {metrics.ascender_pt:.2f} em")
            print(f"    Descender: {metrics.descender_pt:.2f} em")
        else:
            print(f"  {font}: not found")
