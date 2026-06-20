#!/usr/bin/env python3
"""
buat_ppt_generik.py — GENERIC PPT GENERATOR
============================================
Entry point untuk generate PPT dari content module.
LLM bisa generate langsung via python3 -c tanpa file ini.

CARA PAKAI:
    # Via content module (namai *_tmp.py biar di-ignore git):
    CONTENT_MODULE=content_xxx_tmp python3 buat_ppt_generik.py

    # Langsung dari CLI (LLM workflow — tanpa file content):
    python3 -c "
    from ppt_engine import Engine
    SLIDES = [{'type':'cover','data':{...}}]
    Engine().build(SLIDES, source_text='Sumber: ...', output_path='output.pptx')
    "

STRUKTUR FILE:
    ppt_engine.py               ← ENGINE REUSABLE
    buat_ppt_generik.py         ← ENTRY POINT (file ini)
    *_tmp.py                    ← CONTENT MODULE (di-ignore git via .gitignore)

NAMA FILE:
    Content module harus pakai suffix *_tmp.py biar tidak ikut ter-track git.
    Contoh: content_regulasi_tmp.py, content_trigonometri_tmp.py
"""

import os, sys, importlib


def main():
    content_module_name = os.environ.get("CONTENT_MODULE")
    
    if not content_module_name:
        print("=" * 60)
        print("📋 PPT GENERATOR — Panduan Cepat")
        print("=" * 60)
        print()
        print("Content module tidak ditentukan.")
        print()
        print("🔥 LLM Workflow — langsung dari CLI:")
        print()
        print("  python3 -c \"from ppt_engine import Engine;\"")
        print("  python3 -c \"Engine().build(SLIDES, ...)\"")
        print()
        print("📦 Content module — jika sudah punya file *_tmp.py:")
        print()
        print("  CONTENT_MODULE=content_xxx_tmp python3 buat_ppt_generik.py")
        print()
        print("📌 Convention: namai file *_tmp.py biar di-ignore git")
        print()
        print("📖 Baca panduan lengkap: cat agent.md")
        print("=" * 60)
        sys.exit(0)
    
    try:
        content = importlib.import_module(content_module_name)
    except ImportError:
        print(f"❌ Content module '{content_module_name}' tidak ditemukan.")
        print(f"   Pastikan file {content_module_name}.py ada.")
        print(f"   📌 Convention: pakai suffix _tmp.py (e.g. content_xxx_tmp.py)")
        sys.exit(1)
    
    from ppt_engine import Engine
    
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
    
    engine = Engine()
    prs = engine.build(slides, source_text=source, output_path=output)
    
    print(f"✅ BERHASIL: {output}")
    print(f"   {len(prs.slides)} slide, {os.path.getsize(output):,} bytes")
    
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
