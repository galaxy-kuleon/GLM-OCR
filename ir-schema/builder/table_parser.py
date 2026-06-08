#!/usr/bin/env python3
"""
Table Parser for DocIR

Parses table content from GLM-OCR output into structured DocIR table elements.
Supports multiple formats:
- Markdown tables (| col1 | col2 |)
- Tab-separated values
- Empty content (fallback to minimal structure)
"""

import re
from typing import List, Tuple, Optional
from dataclasses import dataclass


@dataclass
class TableCell:
    """A single table cell."""
    content: str
    row_span: int = 1
    col_span: int = 1
    is_header: bool = False


@dataclass
class TableData:
    """Parsed table structure."""
    rows: List[List[TableCell]]
    num_rows: int
    num_cols: int
    has_header: bool = False
    
    def to_markdown(self) -> str:
        """Convert back to markdown for debugging."""
        lines = []
        for i, row in enumerate(self.rows):
            cells = [cell.content for cell in row]
            lines.append("| " + " | ".join(cells) + " |")
            if i == 0 and self.has_header:
                lines.append("| " + " | ".join(["---"] * len(cells)) + " |")
        return "\n".join(lines)


def parse_markdown_table(content: str) -> Optional[TableData]:
    """
    Parse a markdown table.
    
    Format:
    | Header 1 | Header 2 |
    |----------|----------|
    | Cell 1   | Cell 2   |
    | Cell 3   | Cell 4   |
    """
    lines = content.strip().split('\n')
    if not lines:
        return None
    
    # Filter out empty lines and code fences
    lines = [l.strip() for l in lines if l.strip() and not l.strip().startswith('```')]
    
    # Check if it looks like a markdown table
    table_lines = [l for l in lines if l.startswith('|') and l.endswith('|')]
    if len(table_lines) < 2:
        return None
    
    # Parse rows
    rows = []
    has_header = False
    
    for i, line in enumerate(table_lines):
        # Skip separator lines (|---|---|)
        if re.match(r'^\|[\s\-:|]+\|$', line):
            if i == 1:  # Separator after first row means first row is header
                has_header = True
            continue
        
        # Parse cells
        cells_str = line.strip('|').split('|')
        cells = [TableCell(content=c.strip(), is_header=(i == 0 and has_header)) for c in cells_str]
        rows.append(cells)
    
    if not rows:
        return None
    
    # Determine number of columns (max across all rows)
    num_cols = max(len(row) for row in rows)
    
    # Pad rows to have equal number of columns
    for row in rows:
        while len(row) < num_cols:
            row.append(TableCell(content=""))
    
    return TableData(
        rows=rows,
        num_rows=len(rows),
        num_cols=num_cols,
        has_header=has_header
    )


def parse_tab_separated(content: str) -> Optional[TableData]:
    """Parse tab-separated values."""
    lines = content.strip().split('\n')
    if not lines:
        return None
    
    # Filter out empty lines and code fences
    lines = [l for l in lines if l.strip() and not l.strip().startswith('```')]
    
    # Check if lines have tabs
    tab_lines = [l for l in lines if '\t' in l]
    if len(tab_lines) < 2:
        return None
    
    # Parse rows
    rows = []
    for line in lines:
        if '\t' not in line:
            continue
        cells_str = line.split('\t')
        cells = [TableCell(content=c.strip()) for c in cells_str]
        rows.append(cells)
    
    if not rows:
        return None
    
    num_cols = max(len(row) for row in rows)
    for row in rows:
        while len(row) < num_cols:
            row.append(TableCell(content=""))
    
    return TableData(
        rows=rows,
        num_rows=len(rows),
        num_cols=num_cols,
        has_header=False
    )


def parse_table_content(content: Optional[str]) -> TableData:
    """
    Parse table content from GLM-OCR output.
    
    Tries multiple formats in order:
    1. Markdown table
    2. Tab-separated values
    3. Fallback to minimal structure
    
    Args:
        content: Raw table content from GLM-OCR
    
    Returns:
        TableData with parsed structure
    """
    if not content:
        # Empty content - return minimal structure
        return TableData(
            rows=[[TableCell(content="[Empty table]")]],
            num_rows=1,
            num_cols=1,
            has_header=False
        )
    
    # Try markdown table
    result = parse_markdown_table(content)
    if result:
        return result
    
    # Try tab-separated
    result = parse_tab_separated(content)
    if result:
        return result
    
    # Fallback: treat as single cell
    return TableData(
        rows=[[TableCell(content=content.strip()[:100])]],
        num_rows=1,
        num_cols=1,
        has_header=False
    )


def table_data_to_docir_xml(table_data: TableData) -> str:
    """
    Convert TableData to DocIR XML string.
    
    Returns XML fragment for <docir:table_content> element.
    """
    lines = []
    lines.append(f'<docir:table_content rows="{table_data.num_rows}" cols="{table_data.num_cols}">')
    
    # Table style
    lines.append('  <docir:table_style border_visible="true" border_color="#000000"' + 
                 (' header_row="true"' if table_data.has_header else '') + '/>')
    
    # Row groups
    if table_data.has_header:
        lines.append('  <docir:row_group type="header">')
        header_row = table_data.rows[0]
        lines.append('    <docir:row>')
        for cell in header_row:
            span_attrs = ""
            if cell.col_span > 1:
                span_attrs += f' col_span="{cell.col_span}"'
            if cell.row_span > 1:
                span_attrs += f' row_span="{cell.row_span}"'
            lines.append(f'      <docir:cell{span_attrs}>')
            lines.append('        <docir:text_content>')
            lines.append('          <docir:paragraph>')
            escaped = cell.content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            lines.append(f'            <docir:run>{escaped}</docir:run>')
            lines.append('          </docir:paragraph>')
            lines.append('        </docir:text_content>')
            lines.append('      </docir:cell>')
        lines.append('    </docir:row>')
        lines.append('  </docir:row_group>')
        
        # Body rows
        if len(table_data.rows) > 1:
            lines.append('  <docir:row_group type="body">')
            for row in table_data.rows[1:]:
                lines.append('    <docir:row>')
                for cell in row:
                    span_attrs = ""
                    if cell.col_span > 1:
                        span_attrs += f' col_span="{cell.col_span}"'
                    if cell.row_span > 1:
                        span_attrs += f' row_span="{cell.row_span}"'
                    lines.append(f'      <docir:cell{span_attrs}>')
                    lines.append('        <docir:text_content>')
                    lines.append('          <docir:paragraph>')
                    escaped = cell.content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    lines.append(f'            <docir:run>{escaped}</docir:run>')
                    lines.append('          </docir:paragraph>')
                    lines.append('        </docir:text_content>')
                    lines.append('      </docir:cell>')
                lines.append('    </docir:row>')
            lines.append('  </docir:row_group>')
    else:
        # All body rows
        lines.append('  <docir:row_group type="body">')
        for row in table_data.rows:
            lines.append('    <docir:row>')
            for cell in row:
                span_attrs = ""
                if cell.col_span > 1:
                    span_attrs += f' col_span="{cell.col_span}"'
                if cell.row_span > 1:
                    span_attrs += f' row_span="{cell.row_span}"'
                lines.append(f'      <docir:cell{span_attrs}>')
                lines.append('        <docir:text_content>')
                lines.append('          <docir:paragraph>')
                escaped = cell.content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                lines.append(f'            <docir:run>{escaped}</docir:run>')
                lines.append('          </docir:paragraph>')
                lines.append('        </docir:text_content>')
                lines.append('      </docir:cell>')
            lines.append('    </docir:row>')
        lines.append('  </docir:row_group>')
    
    lines.append('</docir:table_content>')
    return "\n".join(lines)


# ============================================================
# Tests
# ============================================================

def test_markdown_table():
    """Test markdown table parsing."""
    content = """| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |"""
    
    result = parse_table_content(content)
    assert result.num_rows == 3, f"Expected 3 rows, got {result.num_rows}"
    assert result.num_cols == 3, f"Expected 3 cols, got {result.num_cols}"
    assert result.has_header == True, "Expected header row"
    assert result.rows[0][0].content == "Header 1"
    assert result.rows[1][0].content == "Cell 1"
    print("✓ Markdown table parsing works")


def test_empty_content():
    """Test empty content handling."""
    result = parse_table_content(None)
    assert result.num_rows == 1
    assert result.num_cols == 1
    assert result.rows[0][0].content == "[Empty table]"
    print("✓ Empty content handling works")


def test_code_fence_cleanup():
    """Test that code fences are removed."""
    content = """```markdown
| A | B |
|---|---|
| 1 | 2 |
```"""
    
    result = parse_table_content(content)
    assert result.num_rows == 2, f"Expected 2 rows, got {result.num_rows}"
    assert result.rows[0][0].content == "A"
    print("✓ Code fence cleanup works")


def test_xml_generation():
    """Test XML generation from TableData."""
    content = """| Name | Value |
|------|-------|
| Foo  | 123   |
| Bar  | 456   |"""
    
    table_data = parse_table_content(content)
    xml = table_data_to_docir_xml(table_data)
    
    assert 'rows="3"' in xml
    assert 'cols="2"' in xml
    assert 'header_row="true"' in xml
    assert '<docir:row_group type="header">' in xml
    assert '<docir:row_group type="body">' in xml
    assert 'Foo' in xml
    assert '456' in xml
    print("✓ XML generation works")
    print("\nGenerated XML:")
    print(xml)


if __name__ == "__main__":
    test_markdown_table()
    test_empty_content()
    test_code_fence_cleanup()
    test_xml_generation()
    print("\n✅ All table parser tests passed!")
