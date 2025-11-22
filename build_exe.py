#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AFT Sizing Automation - EXE Build Script
PyInstaller ile Windows executable oluşturur
"""

import subprocess
import sys
import os

def build_exe():
    """EXE dosyası oluştur"""
    print("=" * 60)
    print("AFT Sizing Automation - EXE Build")
    print("=" * 60)
    
    # PyInstaller'ı kontrol et
    try:
        import PyInstaller
        print("✅ PyInstaller yüklü")
    except ImportError:
        print("❌ PyInstaller bulunamadı!")
        print("Yüklemek için: pip install pyinstaller")
        return False
    
    # PyInstaller komutu
    cmd = [
        "pyinstaller",
        "--name=AFT_Sizing_Automation",
        "--onefile",  # Tek dosya olarak
        "--windowed",  # Konsol penceresi gösterme (GUI için)
        "--icon=NONE",  # Icon yok (kod içinde oluşturuluyor)
        # Windows'ta ; kullan, macOS/Linux'ta : kullan
        "--add-data=Logs" + (";Logs" if sys.platform == "win32" else ":Logs"),
        "--add-data=Results" + (";Results" if sys.platform == "win32" else ":Results"),
        "--hidden-import=customtkinter",
        "--hidden-import=openpyxl",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=tkinterdnd2",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.ImageDraw",
        "--hidden-import=PIL.ImageFont",
        "--hidden-import=PIL.ImageTk",
        "--collect-all=customtkinter",  # customtkinter'ın tüm dosyalarını topla
        "--collect-all=openpyxl",  # openpyxl'ın tüm dosyalarını topla
        "--noconfirm",  # Onay isteme
        "--clean",  # Önceki build'i temizle
        "s2dgui4.py"
    ]
    
    print("\n🔨 EXE oluşturuluyor...")
    print("Bu işlem birkaç dakika sürebilir...\n")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ EXE başarıyla oluşturuldu!")
        print(f"\n📦 Dosya konumu: dist/AFT_Sizing_Automation.exe")
        print("\n💡 Notlar:")
        print("   - EXE dosyasını Windows'ta çalıştırabilirsiniz")
        print("   - İlk çalıştırmada Windows Defender uyarısı çıkabilir (normal)")
        print("   - Logs ve Results klasörleri otomatik oluşturulacak")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Hata: {e}")
        print(f"\nHata çıktısı:\n{e.stderr}")
        return False
    except Exception as e:
        print(f"❌ Beklenmeyen hata: {e}")
        return False

if __name__ == "__main__":
    success = build_exe()
    sys.exit(0 if success else 1)

