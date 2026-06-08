#!/usr/bin/env python3
"""
End-to-End DocIR Pipeline

Chains all components:
  PDF → GLM-OCR → DocIR XML → Style Extraction → DOCX

Usage:
    python run_pipeline.py input.pdf -o output.docx
    
Requirements:
    - GLM-OCR installed and configured (Ollama with glm-ocr model)
    - VLM service running (qwen3.6-35b-a3b-q7 @ localhost:11234)
    - Python packages: pymupdf, lxml, requests, python-docx
"""

import sys
import subprocess
import tempfile
from pathlib import Path


def run_command(cmd, description):
    """Run a command and check for errors."""
    print(f"\n{'='*60}")
    print(f"→ {description}")
    print(f"{'='*60}")
    print(f"Command: {' '.join(str(c) for c in cmd)}")
    
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode != 0:
        print(f"\n✗ Error: {description} failed with exit code {result.returncode}")
        sys.exit(1)
    
    print(f"✓ {description} completed successfully")


def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description='End-to-end DocIR pipeline: PDF → DOCX'
    )
    parser.add_argument(
        'input_pdf',
        type=Path,
        help='Input PDF file'
    )
    parser.add_argument(
        '-o', '--output',
        type=Path,
        required=True,
        help='Output DOCX file'
    )
    parser.add_argument(
        '--work-dir',
        type=Path,
        help='Working directory for intermediate files (default: temp dir)'
    )
    parser.add_argument(
        '--skip-ocr',
        action='store_true',
        help='Skip GLM-OCR step (use existing model JSON)'
    )
    parser.add_argument(
        '--skip-style',
        action='store_true',
        help='Skip style extraction step'
    )
    parser.add_argument(
        '--glmocr-output',
        type=Path,
        help='GLM-OCR output directory (for --skip-ocr)'
    )
    parser.add_argument(
        '--dpi',
        type=int,
        default=200,
        help='DPI for style extraction cropping (default: 200)'
    )
    parser.add_argument(
        '--positioned',
        action='store_true',
        help='Use absolute positioning in DOCX (places elements at PDF coordinates)'
    )
    parser.add_argument(
        '--table-styles',
        action='store_true',
        help='Also extract table styles (borders, headers) via VLM'
    )
    parser.add_argument(
        '--digital',
        action='store_true',
        help='Use digital PDF fast path (exact fonts from PDF, no VLM). '
             'Automatically detects digital vs scanned pages.'
    )
    parser.add_argument(
        '--parallel',
        type=int,
        default=1,
        help='Number of parallel VLM requests for style extraction (default: 1)'
    )
    
    args = parser.parse_args()
    
    if not args.input_pdf.exists():
        print(f"Error: Input PDF not found: {args.input_pdf}", file=sys.stderr)
        sys.exit(1)
    
    # Get script directory
    script_dir = Path(__file__).parent
    
    # Create working directory
    if args.work_dir:
        work_dir = args.work_dir
        work_dir.mkdir(parents=True, exist_ok=True)
    else:
        work_dir = Path(tempfile.mkdtemp(prefix='docir-pipeline-'))
    
    print(f"Working directory: {work_dir}")
    
    # Step 1: GLM-OCR
    if not args.skip_ocr:
        glmocr_output = work_dir / 'glmocr-output'
        
        # Check if uv is available
        try:
            subprocess.run(['uv', '--version'], capture_output=True, check=True)
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("Error: 'uv' command not found. Please install uv first.", file=sys.stderr)
            sys.exit(1)
        
        # Run GLM-OCR pipeline
        glmocr_cmd = [
            'uv', 'run', '--extra', 'layout',
            'glmocr', 'parse',
            str(args.input_pdf),
            '--mode', 'selfhosted',
            '--layout-device', 'cpu',
            '--set', 'pipeline.ocr_api.api_mode', 'ollama_generate',
            '--set', 'pipeline.ocr_api.api_path', '/api/generate',
            '--set', 'pipeline.ocr_api.api_host', 'localhost',
            '--set', 'pipeline.ocr_api.api_port', '11434',
            '--set', 'pipeline.ocr_api.model', 'glm-ocr:latest',
            '--set', 'pipeline.layout.device', 'cpu',
            '--output', str(glmocr_output),
            '--log-level', 'INFO'
        ]
        
        run_command(glmocr_cmd, "Step 1: GLM-OCR Pipeline")
        
        # Find the model JSON
        pdf_stem = args.input_pdf.stem
        model_json_candidates = list(glmocr_output.rglob(f'*{pdf_stem}*model.json'))
        if not model_json_candidates:
            print(f"Error: Could not find model JSON in {glmocr_output}", file=sys.stderr)
            sys.exit(1)
        
        model_json = model_json_candidates[0]
        print(f"✓ Found model JSON: {model_json}")
    else:
        if not args.glmocr_output:
            print("Error: --glmocr-output required when using --skip-ocr", file=sys.stderr)
            sys.exit(1)
        
        pdf_stem = args.input_pdf.stem
        model_json_candidates = list(args.glmocr_output.rglob(f'*{pdf_stem}*model.json'))
        if not model_json_candidates:
            print(f"Error: Could not find model JSON in {args.glmocr_output}", file=sys.stderr)
            sys.exit(1)
        
        model_json = model_json_candidates[0]
        print(f"✓ Using existing model JSON: {model_json}")
    
    # Step 2: IR Builder
    docir_xml = work_dir / f'{args.input_pdf.stem}.docir.xml'
    
    ir_builder_cmd = [
        sys.executable,
        str(script_dir / 'builder' / 'ir_builder.py'),
        str(model_json),
        str(args.input_pdf),
        '-o', str(docir_xml),
        '--title', args.input_pdf.stem
    ]
    
    run_command(ir_builder_cmd, "Step 2: IR Builder")
    
    # Step 3: Style Extraction
    docir_styled = work_dir / f'{args.input_pdf.stem}-styled.docir.xml'
    
    if not args.skip_style:
        if args.digital:
            # Digital PDF fast path: exact fonts from PDF, no VLM
            style_extractor_cmd = [
                sys.executable,
                str(script_dir / 'style_extractor' / 'digital_extractor.py'),
                str(docir_xml),
                str(args.input_pdf),
                '-o', str(docir_styled),
                '--dpi', str(args.dpi)
            ]
            run_command(style_extractor_cmd, "Step 3: Style Extraction (digital fast path)")
        else:
            # VLM-based style extraction (for scanned PDFs)
            style_extractor_cmd = [
                sys.executable,
                str(script_dir / 'style_extractor' / 'style_extractor.py'),
                str(docir_xml),
                str(args.input_pdf),
                '-o', str(docir_styled),
                '--dpi', str(args.dpi)
            ]
            
            if args.table_styles:
                style_extractor_cmd.append('--table-styles')
            
            if args.parallel > 1:
                style_extractor_cmd.extend(['--parallel', str(args.parallel)])
            
            run_command(style_extractor_cmd, "Step 3: Style Extraction (VLM)")
        
        final_docir = docir_styled
    else:
        print("\n⊘ Skipping style extraction")
        final_docir = docir_xml
    
    # Step 4: DOCX Generator
    docx_generator_cmd = [
        sys.executable,
        str(script_dir / 'docx_generator' / 'docx_generator.py'),
        str(final_docir),
        '-o', str(args.output),
        '--base-dir', str(work_dir)
    ]
    
    if args.positioned:
        docx_generator_cmd.append('--positioned')
    
    run_command(docx_generator_cmd, "Step 4: DOCX Generator")
    
    # Summary
    print(f"\n{'='*60}")
    print("✓ Pipeline completed successfully!")
    print(f"{'='*60}")
    print(f"Input PDF:    {args.input_pdf}")
    print(f"Output DOCX:  {args.output}")
    print(f"Working dir:  {work_dir}")
    print(f"\nIntermediate files:")
    print(f"  Model JSON:   {model_json}")
    print(f"  DocIR XML:    {docir_xml}")
    if not args.skip_style:
        print(f"  Styled XML:   {docir_styled}")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
