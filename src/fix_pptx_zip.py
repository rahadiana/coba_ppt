#!/usr/bin/env python3
"""
fix_pptx_zip.py — Post-processing fix untuk PptxGenJS output

Masalah:
  PptxGenJS menulis [Content_Types].xml sebagai entry ZIP ke-19, bukan
  pertama. Microsoft PowerPoint strict terhadap ECMA-376 OPC spec §9.1.2.1:
  "The content types stream shall be the first stream in the package."

Akibat:
  - ❌ Microsoft PowerPoint: file corrupt, tidak bisa dibuka
  - ✅ LibreOffice Impress: render sempurna
  - ✅ python-pptx / lxml: parse sukses

Fix:
  Baca semua entry ZIP → simpan content → tulis ulang dengan
  [Content_Types].xml sebagai entry ZIP pertama.

Penggunaan:
  python3 fix_pptx_zip.py <input.pptx> [output.pptx]
  
  Jika output tidak disebutkan, file input akan di-overwrite in-place.
"""

import sys
import os
import zipfile
import io
import shutil


def fix_pptx_zip(input_path: str, output_path: str = None) -> bool:
    """
    Perbaiki urutan ZIP entry agar [Content_Types].xml menjadi yang pertama.
    
    Args:
        input_path: Path ke file PPTX yang bermasalah
        output_path: Path output (None = overwrite input)
    
    Returns:
        True jika sukses, False jika tidak ada perubahan diperlukan
    """
    if output_path is None:
        output_path = input_path
    
    # Baca semua entry dari file asli
    entries = []
    with zipfile.ZipFile(input_path, 'r') as zin:
        for name in zin.namelist():
            info = zin.getinfo(name)
            entries.append((name, info.compress_type, zin.read(name)))
    
    # Cari posisi [Content_Types].xml
    ct_names = [i for i, (name, _, _) in enumerate(entries) if name == '[Content_Types].xml']
    
    if not ct_names:
        print("❌ [Content_Types].xml tidak ditemukan dalam ZIP!")
        return False
    
    ct_idx = ct_names[0]
    
    if ct_idx == 0:
        print("✅ [Content_Types].xml sudah di posisi pertama. Tidak perlu perbaikan.")
        return False
    
    print(f"📦 [Content_Types].xml saat ini di posisi ZIP ke-{ct_idx} dari {len(entries)}")
    
    # Extract [Content_Types].xml data
    ct_data = entries[ct_idx][2]
    
    # Tulis ulang ZIP dengan urutan baru
    tmp_path = output_path + '.fix_tmp'
    
    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        # 1. Tulis [Content_Types].xml sebagai entry PERTAMA
        zout.writestr('[Content_Types].xml', ct_data)
        
        # 2. Tulis semua entry lain (skip folder entries, skip [Content_Types].xml)
        for name, ctype, data in entries:
            if name == '[Content_Types].xml':
                continue
            if name.endswith('/'):
                continue  # skip folder entries, ZIP_DEFLATED won't store them anyway
            zout.writestr(name, data)
    
    # Overwrite output
    shutil.move(tmp_path, output_path)
    
    # Verifikasi
    with zipfile.ZipFile(output_path, 'r') as z:
        names = z.namelist()
        new_ct_idx = names.index('[Content_Types].xml')
        
        if new_ct_idx != 0:
            print(f"❌ GAGAL! [Content_Types].xml masih di posisi {new_ct_idx}")
            return False
    
    orig_size = os.path.getsize(input_path)
    new_size = os.path.getsize(output_path)
    
    print(f"✅ [Content_Types].xml sekarang di posisi 0 (PERTAMA)")
    print(f"   Ukuran: {orig_size//1024}KB → {new_size//1024}KB")
    
    return True


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 fix_pptx_zip.py <input.pptx> [output.pptx]")
        print("       Jika output tidak disebutkan, file akan di-overwrite.")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(input_path):
        print(f"❌ File tidak ditemukan: {input_path}")
        sys.exit(1)
    
    fix_pptx_zip(input_path, output_path)


if __name__ == '__main__':
    main()
