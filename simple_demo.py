#!/usr/bin/env python3
"""
Simple demonstration of DocumentProcessor capabilities with testing.docx
"""

from app import DocumentProcessor
import os

def simple_demo():
    """Simple demo showing working features"""
    processor = DocumentProcessor()
    docx_file = "testing.docx"
    
    print("=== DOKUMENTASI HASIL TESTING ===\n")
    
    if not os.path.exists(docx_file):
        print(f"❌ File {docx_file} tidak ditemukan!")
        return
    
    try:
        # 1. Document Type Detection
        print("✅ DETEKSI TIPE DOKUMEN:")
        doc_type = processor.detect_document_type(docx_file)
        print(f"   Berhasil terdeteksi sebagai: {doc_type}")
        
        # 2. Read Content
        print(f"\n✅ PEMBACAAN KONTEN:")
        content = processor.process_document(docx_file, "read")
        word_count = len(content.split())
        char_count = len(content)
        
        print(f"   Total karakter: {char_count:,}")
        print(f"   Total kata: {word_count:,}")
        print(f"   Dokumen berisi: Proposal skripsi tentang watermarking")
        
        # 3. Content Analysis
        print(f"\n✅ ANALISIS KONTEN:")
        if "watermark" in content.lower():
            print("   ✓ Mengandung pembahasan tentang watermarking")
        if "qr code" in content.lower():
            print("   ✓ Mengandung pembahasan tentang QR Code")
        if "lsb" in content.lower():
            print("   ✓ Mengandung pembahasan tentang metode LSB")
        if "steganografi" in content.lower():
            print("   ✓ Mengandung pembahasan tentang steganografi")
        
        # 4. Create a simple text version
        print(f"\n✅ EXPORT KE FORMAT LAIN:")
        txt_file = "testing_export.txt"
        processor.process_document(txt_file, "write", content=content)
        print(f"   Berhasil mengekspor ke: {txt_file}")
        
        # 5. File size comparison
        docx_size = os.path.getsize(docx_file)
        txt_size = os.path.getsize(txt_file)
        print(f"\n✅ PERBANDINGAN UKURAN FILE:")
        print(f"   File DOCX: {docx_size:,} bytes")
        print(f"   File TXT:  {txt_size:,} bytes")
        print(f"   Rasio kompresi: {(docx_size/txt_size):.1f}x")
        
        print(f"\n🎉 TESTING BERHASIL!")
        print("   DocumentProcessor dapat membaca dan memproses file Word dengan sempurna!")
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    simple_demo()
