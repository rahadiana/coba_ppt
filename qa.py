#!/usr/bin/env python3
"""
qa.py — Visual QA untuk PPTX
==============================
Render .pptx ke PDF → images untuk inspeksi visual.
Pakai subagent untuk inspeksi setiap slide.

Dependencies:
    pip install markitdown[pptx]
    sudo apt install -y libreoffice poppler-utils

Usage:
    python qa.py input.pptx                     # render + inspect
    python qa.py input.pptx --inspect           # render + subagent inspect
    python qa.py input.pptx --render-only       # hanya render ke images
    python qa.py input.pptx --text-only         # hanya text extraction
"""

import subprocess, sys, os, json, tempfile, shutil, glob, re
from pathlib import Path


# ═══════════════════════════════════════════════════════════
# TEXT EXTRACTION — via markitdown
# ═══════════════════════════════════════════════════════════

def extract_text(pptx_path):
    """Extract text content from PPTX using markitdown."""
    try:
        result = subprocess.run(
            [sys.executable, "-m", "markitdown", pptx_path],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0:
            return result.stdout
        else:
            return f"[markitdown error] {result.stderr}"
    except FileNotFoundError:
        return "[markitdown not installed — pip install markitdown[pptx]]"
    except Exception as e:
        return f"[error extracting text] {e}"


# ═══════════════════════════════════════════════════════════
# RENDER — PPTX → PDF → JPG images
# ═══════════════════════════════════════════════════════════

def render_to_images(pptx_path, output_dir=None, dpi=150):
    """
    Convert PPTX to individual slide images.
    
    Returns:
        list of (slide_num, image_path)
    """
    pptx_path = os.path.abspath(pptx_path)
    stem = Path(pptx_path).stem
    
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix=f"qa_{stem}_")
    else:
        os.makedirs(output_dir, exist_ok=True)
    
    # Step 1: PPTX → PDF via LibreOffice
    pdf_path = os.path.join(output_dir, f"{stem}.pdf")
    
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("❌ LibreOffice tidak ditemukan. Install:")
        print("   sudo apt install -y libreoffice")
        return []
    
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", output_dir, pptx_path],
            check=True, capture_output=True, text=True, timeout=60
        )
    except subprocess.CalledProcessError as e:
        print(f"❌ LibreOffice gagal: {e.stderr}")
        return []
    
    if not os.path.exists(pdf_path):
        # LibreOffice mungkin bikin nama file beda
        pdfs = glob.glob(os.path.join(output_dir, "*.pdf"))
        if pdfs:
            pdf_path = pdfs[0]
        else:
            print("❌ PDF tidak ditemukan setelah konversi")
            return []
    
    print(f"  📄 PDF: {pdf_path}")
    
    # Step 2: PDF → JPG images via pdftoppm
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        print("❌ pdftoppm tidak ditemukan. Install:")
        print("   sudo apt install -y poppler-utils")
        return []
    
    output_pattern = os.path.join(output_dir, f"slide")
    
    try:
        subprocess.run(
            [pdftoppm, "-jpeg", "-r", str(dpi), pdf_path, output_pattern],
            check=True, capture_output=True, text=True, timeout=60
        )
    except subprocess.CalledProcessError as e:
        print(f"❌ pdftoppm gagal: {e.stderr}")
        return []
    
    # Collect slide images
    images = sorted(glob.glob(os.path.join(output_dir, "slide-*.jpg")))
    slide_images = []
    for img_path in images:
        match = re.search(r'slide-(\d+)', os.path.basename(img_path))
        slide_num = int(match.group(1)) if match else 0
        slide_images.append((slide_num, img_path))
    
    print(f"  🖼️  {len(slide_images)} slide images → {output_dir}/")
    return slide_images


# ═══════════════════════════════════════════════════════════
# INSPECT PROMPT — untuk subagent
# ═══════════════════════════════════════════════════════════

INSPECT_PROMPT_TEMPLATE = """Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Low-contrast text (e.g., light gray text on cream-colored background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without a contrasting circle)
- Text boxes too narrow causing excessive wrapping
- Leftover placeholder content

For each slide, list issues or areas of concern, even if minor.

Read and analyze these images:
{image_list}

Report ALL issues found, including minor ones. Format as a markdown list per slide."""


def build_inspect_prompt(slide_images):
    """Build the subagent inspection prompt."""
    lines = []
    for num, path in slide_images:
        lines.append(f"{num}. {path}")
    return INSPECT_PROMPT_TEMPLATE.format(image_list='\n'.join(lines))


# ═══════════════════════════════════════════════════════════
# CHECK — text-only quick check
# ═══════════════════════════════════════════════════════════

def check_placeholder(text):
    """Check for leftover placeholder text."""
    issues = []
    patterns = [
        (r'xxxx', 'Placeholder "xxxx"'),
        (r'lorem|ipsum', 'Lorem ipsum text'),
        (r'this\s+page.*layout', 'Generic layout text'),
        (r'click\s+to\s+add', 'Click-to-add placeholder'),
    ]
    for pattern, label in patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for m in matches:
            line_start = max(0, m.start() - 40)
            line_end = min(len(text), m.end() + 40)
            context = text[line_start:line_end].replace('\n', ' ')
            issues.append(f"  ⚠️  {label}: ...{context}...")
    return issues


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="QA: render PPTX → images untuk inspeksi visual")
    parser.add_argument("pptx", help="Path ke file .pptx")
    parser.add_argument("--render-only", action="store_true",
                       help="Hanya render ke images, tanpa text check")
    parser.add_argument("--text-only", action="store_true",
                       help="Hanya text extraction + placeholder check")
    parser.add_argument("--inspect", action="store_true",
                       help="Render + generate subagent inspect prompt")
    parser.add_argument("--output-dir", help="Output directory untuk images")
    parser.add_argument("--dpi", type=int, default=150,
                       help="DPI untuk render (default: 150)")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.pptx):
        print(f"❌ File tidak ditemukan: {args.pptx}")
        sys.exit(1)
    
    pptx_path = args.pptx
    stem = Path(pptx_path).stem
    
    print(f"📋 QA Report: {pptx_path}")
    print("=" * 60)
    
    # Text extraction
    text = extract_text(pptx_path)
    if not args.render_only:
        print(f"\n📝 Text Content ({len(text)} chars):")
        print("-" * 40)
        # Print first 2000 chars
        print(text[:2000])
        if len(text) > 2000:
            print(f"... ({len(text) - 2000} more chars)")
        
        # Placeholder check
        issues = check_placeholder(text)
        if issues:
            print(f"\n⚠️  Placeholder Issues:")
            for issue in issues:
                print(issue)
        else:
            print("\n✅ No placeholder text found")
    
    # Render
    if not args.text_only:
        print(f"\n🎨 Rendering to images (DPI={args.dpi}):")
        images = render_to_images(pptx_path, args.output_dir, args.dpi)
        
        if images:
            print(f"\n📸 Slide images:")
            for num, path in images:
                print(f"  Slide {num:02d}: {path}")
            
            if args.inspect:
                print(f"\n🔍 Subagent Inspect Prompt:")
                print("-" * 40)
                print(build_inspect_prompt(images))
    
    print("\n✅ QA selesai")


if __name__ == "__main__":
    main()
