#!/bin/bash
# macOS'tan Windows VM'de EXE oluşturma scripti
# Parallels, VMware veya VirtualBox kullanılabilir

echo "=========================================="
echo "Windows EXE Build - VM Yöntemi"
echo "=========================================="
echo ""
echo "Bu script Windows VM'inizde çalıştırılmalıdır."
echo ""
echo "Adımlar:"
echo "1. Windows VM'inizi başlatın"
echo "2. Bu projeyi VM'e kopyalayın"
echo "3. VM'de şu komutları çalıştırın:"
echo ""
echo "   pip install -r requirements.txt"
echo "   pip install pyinstaller"
echo "   python build_exe.py"
echo ""
echo "VEYA"
echo ""
echo "   pyinstaller build_exe.spec"
echo ""
echo "=========================================="

