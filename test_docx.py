#!/usr/bin/env python3
"""
Test script for testing.docx file using DocumentProcessor
"""

from app import DocumentProcessor
import os

def test_docx_file():
    """Test the existing testing.docx file"""
    processor = DocumentProcessor()
    docx_file = "testing.docx"
    
    print("=== TESTING DOCX FILE ===\n")
    
    # Check if file exists
    if not os.path.exists(docx_file):
        print(f"❌ File {docx_file} tidak ditemukan!")
        return
    
    print(f"✅ File {docx_file} ditemukan!")
    
    try:
        # 1. Detect document type
        print("\n1. DETEKSI TIPE DOKUMEN:")
        doc_type = processor.detect_document_type(docx_file)
        print(f"Tipe dokumen: {doc_type}")
        
        # 2. Read content
        print("\n2. MEMBACA KONTEN DOKUMEN:")
        content = processor.process_document(docx_file, "read")
        print("Konten dokumen:")
        print("-" * 50)
        print(content)
        print("-" * 50)
        
        # 3. Create backup and test append
        backup_file = "testing_backup.docx"
        print(f"\n3. MEMBUAT BACKUP DAN TEST APPEND:")
        
        # Copy content to backup
        processor.process_document(backup_file, "write", content=content)
        print(f"✅ Backup dibuat: {backup_file}")
        
        # Append new content
        new_content = "\n\nTeks tambahan dari DocumentProcessor pada tanggal 31 Juli 2025."
        result = processor.process_document(backup_file, "append", content=new_content)
        print(f"✅ {result}")
        
        # Read backup to verify
        backup_content = processor.process_document(backup_file, "read")
        print("\nKonten backup setelah append:")
        print("-" * 50)
        print(backup_content)
        print("-" * 50)
        
        # 4. Test replace functionality
        print("\n4. TEST REPLACE FUNCTIONALITY:")
        if "test" in content.lower() or "testing" in content.lower():
            result = processor.process_document(backup_file, "replace", 
                                              old_text="testing", 
                                              new_text="TESTING (MODIFIED)")
            print(f"✅ {result}")
            
            # Read to verify replace
            modified_content = processor.process_document(backup_file, "read")
            print("\nKonten setelah replace:")
            print("-" * 50)
            print(modified_content)
            print("-" * 50)
        else:
            print("❌ Kata 'testing' tidak ditemukan untuk test replace")
        
        print(f"\n✅ Testing selesai! File backup tersimpan sebagai {backup_file}")
        
    except ImportError as e:
        print(f"❌ Error: Library python-docx belum terinstall!")
        print("Jalankan: pip install python-docx")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    test_docx_file()
