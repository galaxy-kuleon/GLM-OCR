#!/usr/bin/env python3
"""
Create a multi-page test PDF for DocIR pipeline testing.
"""

import pymupdf
from pathlib import Path


def create_multipage_test_pdf(output_path: Path, num_pages: int = 3):
    """
    Create a multi-page test PDF with various content types.
    
    Args:
        output_path: Path to output PDF
        num_pages: Number of pages to create
    """
    doc = pymupdf.open()
    
    for page_idx in range(num_pages):
        # Create A4 page
        page = doc.new_page(width=595, height=842)
        
        # Add title
        title = f"Multi-Page Test Document - Page {page_idx + 1}"
        page.insert_text(
            (72, 72),
            title,
            fontsize=18,
            fontname="helv",
            color=(0, 0, 0)
        )
        
        # Add body text
        body_text = f"""This is page {page_idx + 1} of the multi-page test document.
It contains various content types to test the DocIR pipeline.

Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor 
incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis 
nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.

Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu 
fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in 
culpa qui officia deserunt mollit anim id est laborum.
"""
        
        page.insert_text(
            (72, 120),
            body_text,
            fontsize=11,
            fontname="helv",
            color=(0, 0, 0)
        )
        
        # Add a table on page 2
        if page_idx == 1:
            table_y = 400
            page.insert_text(
                (72, table_y),
                "Sample Table:",
                fontsize=12,
                fontname="helv",
                color=(0, 0, 0)
            )
            
            # Draw simple table
            table_x = 72
            table_y_start = table_y + 20
            col_width = 150
            row_height = 25
            
            # Header row
            for col in range(3):
                x = table_x + col * col_width
                y = table_y_start
                rect = pymupdf.Rect(x, y, x + col_width, y + row_height)
                page.draw_rect(rect, color=(0, 0, 0), width=1)
                page.insert_text(
                    (x + 10, y + 17),
                    f"Header {col + 1}",
                    fontsize=10,
                    fontname="helv",
                    color=(0, 0, 0)
                )
            
            # Data rows
            for row in range(3):
                for col in range(3):
                    x = table_x + col * col_width
                    y = table_y_start + (row + 1) * row_height
                    rect = pymupdf.Rect(x, y, x + col_width, y + row_height)
                    page.draw_rect(rect, color=(0, 0, 0), width=1)
                    page.insert_text(
                        (x + 10, y + 17),
                        f"Row {row + 1}, Col {col + 1}",
                        fontsize=10,
                        fontname="helv",
                        color=(0, 0, 0)
                    )
        
        # Add page number
        page.insert_text(
            (523, 820),
            f"Page {page_idx + 1}",
            fontsize=9,
            fontname="helv",
            color=(0.5, 0.5, 0.5)
        )
    
    doc.save(str(output_path))
    doc.close()
    
    print(f"✓ Created multi-page test PDF: {output_path}")
    print(f"  Pages: {num_pages}")


if __name__ == "__main__":
    import sys
    
    output_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("/tmp/multipage-test.pdf")
    num_pages = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    
    create_multipage_test_pdf(output_path, num_pages)
