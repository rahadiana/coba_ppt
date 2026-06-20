#!/usr/bin/env python3
"""
buat_ppt_generik.py — GENERIC PPT GENERATOR
============================================
Entry point untuk generate PPT dari content terpisah.
Engine dan content dipisah — ganti konten tanpa edit engine.

CARA PAKAI:
    # Pakai content default (content_perwal_51_2024.py):
    python3 buat_ppt_generik.py

    # Pakai content kustom:
    CONTENT_MODULE=content_kustom python3 buat_ppt_generik.py

    # Atau dari script lain:
    from ppt_engine import Engine
    from content_perwal_51_2024 import PRESENTATION, SLIDES
    engine = Engine()
    engine.build(SLIDES, source_text=PRESENTATION['source'],
                 output_path=PRESENTATION['output'])

STRUKTUR FILE:
    ppt_engine.py               ← REUSABLE (tidak perlu diubah)
    content_perwal_51_2024.py   ← CONTENT (ganti untuk PPT lain)
    buat_ppt_generik.py         ← ENTRY POINT (file ini)
"""

import os, sys, importlib


def main():
    # Tentukan module content
    content_module_name = os.environ.get("CONTENT_MODULE", "content_perwal_51_2024")
    
    try:
        content = importlib.import_module(content_module_name)
    except ImportError:
        print(f"❌ Content module '{content_module_name}' tidak ditemukan.")
        print("   Buat file content atau set CONTENT_MODULE=namafile_tanpa_.py")
        sys.exit(1)
    
    from ppt_engine import Engine
    
    # Validasi content
    pres = getattr(content, "PRESENTATION", {})
    slides = getattr(content, "SLIDES", [])
    
    if not slides:
        print("❌ SLIDES kosong atau tidak ditemukan di content module.")
        sys.exit(1)
    
    title = pres.get("title", "Presentation")
    source = pres.get("source", "Sumber: ...")
    output = pres.get("output", "output.pptx")
    
    print(f"📋 Judul: {title}")
    print(f"📄 Slide : {len(slides)} definisi")
    print(f"🔧 Engine: ppt_engine.py")
    print(f"📁 Output: {output}")
    print()
    
    # Build
    engine = Engine()
    prs = engine.build(slides, source_text=source, output_path=output)
    
    print(f"✅ BERHASIL: {output}")
    print(f"   {len(prs.slides)} slide, {os.path.getsize(output):,} bytes")
    
    # Ringkasan slide
    types = {}
    for s in slides:
        t = s.get("type", "?")
        types[t] = types.get(t, 0) + 1
    print(f"\n📊 Komposisi slide:")
    for t, n in sorted(types.items()):
        print(f"   • {t}: {n}")
    
    abs_path = os.path.abspath(output)
    print(f"\n📍 Path: {abs_path}")


if __name__ == "__main__":
    main()
