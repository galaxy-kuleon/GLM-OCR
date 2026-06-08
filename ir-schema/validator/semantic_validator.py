#!/usr/bin/env python3
"""
DocIR Semantic Validator v0.1.0

Validates DocIR XML documents beyond what XSD schema can check:
- Bounding boxes within page boundaries
- Table row/col counts match actual cell counts
- Region order is monotonically increasing per page
- Image references resolve to existing assets
- Confidence values in valid range [0, 1]
- Cross-page hint references are valid
- Provenance data completeness

Usage:
    python semantic_validator.py <docir.xml> [--xsd <schema.xsd>] [--strict]

Exit codes:
    0 = All checks passed
    1 = Validation errors found
    2 = File/parse errors
"""

import sys
import argparse
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional, Set
from lxml import etree

NS = {"docir": "urn:docir:v0.1"}


@dataclass
class ValidationIssue:
    level: str  # "error" | "warning" | "info"
    rule: str
    message: str
    location: str = ""

    def __str__(self):
        loc = f" @ {self.location}" if self.location else ""
        return f"[{self.level.upper()}] ({self.rule}) {self.message}{loc}"


@dataclass
class ValidationResult:
    issues: List[ValidationIssue] = field(default_factory=list)
    
    @property
    def errors(self) -> List[ValidationIssue]:
        return [i for i in self.issues if i.level == "error"]
    
    @property
    def warnings(self) -> List[ValidationIssue]:
        return [i for i in self.issues if i.level == "warning"]
    
    @property
    def passed(self) -> bool:
        return len(self.errors) == 0
    
    def add(self, level: str, rule: str, message: str, location: str = ""):
        self.issues.append(ValidationIssue(level, rule, message, location))
    
    def summary(self) -> str:
        e = len(self.errors)
        w = len(self.warnings)
        status = "PASS" if self.passed else "FAIL"
        return f"{status}: {e} error(s), {w} warning(s), {len(self.issues)} total issue(s)"


def validate_xsd(xml_path: Path, xsd_path: Path, result: ValidationResult):
    """Validate XML against XSD schema."""
    try:
        schema_doc = etree.parse(str(xsd_path))
        schema = etree.XMLSchema(schema_doc)
        doc = etree.parse(str(xml_path))
        if not schema.validate(doc):
            for error in schema.error_log:
                result.add("error", "xsd", str(error), f"line {error.line}")
        else:
            result.add("info", "xsd", "XSD schema validation passed")
    except etree.XMLSchemaParseError as e:
        result.add("error", "xsd", f"Schema parse error: {e}")
    except etree.XMLSyntaxError as e:
        result.add("error", "xml", f"XML syntax error: {e}")


def validate_bbox_in_page(tree: etree._ElementTree, result: ValidationResult):
    """Check that all bounding boxes are within page boundaries."""
    for page in tree.findall(".//docir:page", NS):
        page_idx = page.get("index", "?")
        
        # Get page dimensions
        page_size = page.find("docir:page_size", NS)
        if page_size is None:
            # Try document default
            page_size = tree.find(".//docir:default_page_size", NS)
        if page_size is None:
            result.add("warning", "bbox-page", f"Page {page_idx}: no page_size defined, cannot validate bbox")
            continue
        
        pw = float(page_size.get("width_pt", 0))
        ph = float(page_size.get("height_pt", 0))
        
        for region in page.findall(".//docir:region", NS):
            rid = region.get("id", "?")
            bbox = region.find("docir:bbox", NS)
            if bbox is None:
                result.add("error", "bbox-missing", f"Page {page_idx}, region {rid}: missing bbox")
                continue
            
            x = float(bbox.get("x", 0))
            y = float(bbox.get("y", 0))
            w = float(bbox.get("width", 0))
            h = float(bbox.get("height", 0))
            
            # Check non-negative dimensions
            if w <= 0 or h <= 0:
                result.add("error", "bbox-dimension", f"Page {page_idx}, region {rid}: non-positive dimensions ({w}×{h})")
            
            # Check within page bounds (with 5% tolerance for rounding)
            tolerance = 0.05
            if x < -pw * tolerance or y < -ph * tolerance:
                result.add("warning", "bbox-origin", f"Page {page_idx}, region {rid}: bbox origin ({x},{y}) outside page")
            if x + w > pw * (1 + tolerance) or y + h > ph * (1 + tolerance):
                result.add("warning", "bbox-overflow", f"Page {page_idx}, region {rid}: bbox extends beyond page ({x+w:.1f} > {pw} or {y+h:.1f} > {ph})")


def validate_table_structure(tree: etree._ElementTree, result: ValidationResult):
    """Check table row/col counts match actual cell counts."""
    for table in tree.findall(".//docir:table_content", NS):
        declared_rows = int(table.get("rows", 0))
        declared_cols = int(table.get("cols", 0))
        
        # Count actual rows (sum across all row_groups)
        actual_rows = 0
        max_cols = 0
        
        for row_group in table.findall("docir:row_group", NS):
            for row in row_group.findall("docir:row", NS):
                actual_rows += 1
                # Count cells with colspan
                row_col_count = 0
                for cell in row.findall("docir:cell", NS):
                    col_span = int(cell.get("col_span", 1))
                    row_col_count += col_span
                max_cols = max(max_cols, row_col_count)
        
        if declared_rows != actual_rows:
            result.add("error", "table-rows", 
                       f"Table declares {declared_rows} rows but has {actual_rows} actual rows")
        if declared_cols != max_cols and max_cols > 0:
            result.add("warning", "table-cols",
                       f"Table declares {declared_cols} cols but max actual is {max_cols}")
        
        # Check for nested tables
        for nested in table.findall(".//docir:table_content", NS):
            if nested is not table:
                result.add("info", "table-nested", "Nested table detected — verify DOCX generator handles this")


def validate_region_order(tree: etree._ElementTree, result: ValidationResult):
    """Check that region order is monotonically increasing per page."""
    for page in tree.findall(".//docir:page", NS):
        page_idx = page.get("index", "?")
        orders = []
        
        for region in page.findall(".//docir:region", NS):
            rid = region.get("id", "?")
            order = region.get("order")
            if order is None:
                result.add("error", "region-order-missing", f"Page {page_idx}, region {rid}: missing order attribute")
                continue
            orders.append((int(order), rid))
        
        # Check monotonically increasing
        for i in range(1, len(orders)):
            if orders[i][0] <= orders[i-1][0]:
                result.add("error", "region-order",
                           f"Page {page_idx}: region {orders[i][1]} (order={orders[i][0]}) "
                           f"not after {orders[i-1][1]} (order={orders[i-1][0]})")
        
        # Check starts from 0
        if orders and orders[0][0] != 0:
            result.add("warning", "region-order-start",
                       f"Page {page_idx}: first region order is {orders[0][0]}, expected 0")


def validate_image_references(tree: etree._ElementTree, result: ValidationResult):
    """Check that image references resolve to existing assets."""
    # Collect asset IDs
    asset_ids: Set[str] = set()
    for asset in tree.findall(".//docir:asset", NS):
        aid = asset.get("id")
        if aid:
            asset_ids.add(aid)
    
    # Check image references
    for img_ref in tree.findall(".//docir:image_reference", NS):
        ref_id = img_ref.get("asset_id")
        if ref_id and ref_id not in asset_ids:
            result.add("error", "image-ref",
                       f"Image reference '{ref_id}' not found in asset registry")
    
    # Check for orphaned assets
    referenced_ids: Set[str] = set()
    for img_ref in tree.findall(".//docir:image_reference", NS):
        aid = img_ref.get("asset_id")
        if aid:
            referenced_ids.add(aid)
    
    for aid in asset_ids:
        if aid not in referenced_ids:
            result.add("warning", "asset-orphan",
                       f"Asset '{aid}' defined but not referenced by any region")


def validate_confidence(tree: etree._ElementTree, result: ValidationResult):
    """Check confidence values are in valid range [0, 1]."""
    for prov in tree.findall(".//docir:provenance", NS):
        conf_elem = prov.find("docir:confidence", NS)
        if conf_elem is not None and conf_elem.text:
            try:
                conf = float(conf_elem.text)
                if conf < 0 or conf > 1:
                    region = prov.getparent()
                    rid = region.get("id", "?") if region is not None else "?"
                    result.add("error", "confidence-range",
                               f"Region {rid}: confidence {conf} outside [0, 1]")
                elif conf < 0.5:
                    region = prov.getparent()
                    rid = region.get("id", "?") if region is not None else "?"
                    result.add("warning", "confidence-low",
                               f"Region {rid}: low confidence {conf:.2f} (< 0.5)")
            except ValueError:
                result.add("error", "confidence-parse",
                           f"Cannot parse confidence value: '{conf_elem.text}'")


def validate_cross_page_hints(tree: etree._ElementTree, result: ValidationResult):
    """Check cross-page hint references are valid region IDs."""
    # Collect all region IDs
    all_region_ids: Set[str] = set()
    for region in tree.findall(".//docir:region", NS):
        rid = region.get("id")
        if rid:
            all_region_ids.add(rid)
    
    # Check hint references
    for hint in tree.findall(".//docir:hint", NS):
        from_r = hint.get("from_region")
        to_r = hint.get("to_region")
        
        if from_r and from_r not in all_region_ids:
            result.add("error", "cross-page-from",
                       f"Cross-page hint references non-existent region '{from_r}'")
        if to_r and to_r not in all_region_ids:
            result.add("error", "cross-page-to",
                       f"Cross-page hint references non-existent region '{to_r}'")
    
    # Check merge_hint references
    for mh in tree.findall(".//docir:merge_hint", NS):
        linked = mh.get("linked_region")
        if linked and linked not in all_region_ids:
            parent = mh.getparent()
            rid = parent.get("id", "?") if parent is not None else "?"
            result.add("error", "merge-hint-ref",
                       f"Region {rid}: merge_hint references non-existent region '{linked}'")


def validate_provenance_completeness(tree: etree._ElementTree, result: ValidationResult):
    """Check that all regions have provenance data."""
    for page in tree.findall(".//docir:page", NS):
        page_idx = page.get("index", "?")
        for region in page.findall(".//docir:region", NS):
            rid = region.get("id", "?")
            prov = region.find("docir:provenance", NS)
            
            if prov is None:
                result.add("warning", "provenance-missing",
                           f"Page {page_idx}, region {rid}: no provenance data")
                continue
            
            source = prov.find("docir:source", NS)
            if source is None or not source.text:
                result.add("warning", "provenance-source",
                           f"Page {page_idx}, region {rid}: provenance missing source")
            
            conf = prov.find("docir:confidence", NS)
            if conf is None:
                result.add("info", "provenance-confidence",
                           f"Page {page_idx}, region {rid}: provenance missing confidence")


def validate_style_evidence(tree: etree._ElementTree, result: ValidationResult):
    """Check that runs with computed styles also have evidence."""
    for run in tree.findall(".//docir:run", NS):
        has_computed = any([
            run.get("font_size_pt"),
            run.get("color"),
            run.get("font_name"),
        ])
        has_evidence = any([
            run.get("evidence_pixel_height"),
            run.get("evidence_color_sample"),
            run.get("evidence_confidence"),
        ])
        
        if has_computed and not has_evidence:
            # Find parent region for context
            region = run.getparent()
            while region is not None and region.tag != f"{{{NS['docir']}}}region":
                region = region.getparent()
            rid = region.get("id", "?") if region is not None else "?"
            
            result.add("warning", "style-no-evidence",
                       f"Region {rid}: run has computed style but no evidence attributes")


def run_all_validations(xml_path: Path, xsd_path: Optional[Path] = None, strict: bool = False) -> ValidationResult:
    """Run all validation checks on a DocIR XML file."""
    result = ValidationResult()
    
    # Parse XML
    try:
        tree = etree.parse(str(xml_path))
    except etree.XMLSyntaxError as e:
        result.add("error", "xml-parse", f"Cannot parse XML: {e}")
        return result
    
    # XSD validation (if schema provided)
    if xsd_path and xsd_path.exists():
        validate_xsd(xml_path, xsd_path, result)
    
    # Semantic validations
    validate_bbox_in_page(tree, result)
    validate_table_structure(tree, result)
    validate_region_order(tree, result)
    validate_image_references(tree, result)
    validate_confidence(tree, result)
    validate_cross_page_hints(tree, result)
    validate_provenance_completeness(tree, result)
    validate_style_evidence(tree, result)
    
    # In strict mode, warnings become errors
    if strict:
        for issue in result.issues:
            if issue.level == "warning":
                issue.level = "error"
    
    return result


def main():
    parser = argparse.ArgumentParser(description="DocIR Semantic Validator")
    parser.add_argument("xml_file", type=Path, help="DocIR XML file to validate")
    parser.add_argument("--xsd", type=Path, default=None, help="XSD schema file (optional)")
    parser.add_argument("--strict", action="store_true", help="Treat warnings as errors")
    parser.add_argument("--verbose", "-v", action="store_true", help="Show all issues including info")
    args = parser.parse_args()
    
    if not args.xml_file.exists():
        print(f"Error: File not found: {args.xml_file}", file=sys.stderr)
        sys.exit(2)
    
    result = run_all_validations(args.xml_file, args.xsd, args.strict)
    
    # Print results
    for issue in result.issues:
        if issue.level == "info" and not args.verbose:
            continue
        print(issue)
    
    print(f"\n{'='*60}")
    print(result.summary())
    print(f"{'='*60}")
    
    sys.exit(0 if result.passed else 1)


if __name__ == "__main__":
    main()
