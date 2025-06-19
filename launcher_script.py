#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word轉PowerPoint工具 v5.0 - 統一啟動器
支援啟動 Streamlit 網頁版或桌面版
"""

import sys
import os
import platform
import subprocess
import argparse
from pathlib import Path

def check_file_exists(filename):
    """檢查檔案是否存在"""
    if not os.path.exists(filename):
        print(f"❌ 找不到檔案: {filename}")
        return False
    return True

def check_dependencies():
    """檢查相依套件是否已安裝"""
    required_packages = {
        'docx': 'python-docx',
        'pptx': 'python-pptx', 
        'PIL': 'pillow',
        'streamlit': 'streamlit',
        'PySide6': 'PySide6'
    }
    
    missing_packages = []
    available_packages = []
    
    for module, package in required_packages.items():
        try:
            __import__(module)
            available_packages.append(f"✅ {package}")
        except ImportError:
            missing_packages.append(package)
            available_packages.append(f"❌ {package}")
    
    print("📦 套件狀態檢查:")
    for status in available_packages:
        print(f"  {status}")
    
    return missing_packages

def install_dependencies(missing_packages):
    """安裝缺少的套件"""
    if not missing_packages:
        return True
    
    print(f"\n🔧 發現缺少的套件: {', '.join(missing_packages)}")
    response = input("是否要自動安裝這些套件？ (y/n): ").lower().strip()
    
    if response in ['y', 'yes', '是']:
        try:
            cmd = [sys.executable, '-m', 'pip', 'install'] + missing_packages
            print(f"執行命令: {' '.join(cmd)}")
            
            process = subprocess.run(cmd, check=True, capture_output=True, text=True)
            print("✅ 套件安裝成功!")
            return True
            
        except subprocess.CalledProcessError as e:
            print(f"❌ 套件安裝失敗: {e}")
            print("請手動執行以下命令:")
            print(f"pip install {' '.join(missing_packages)}")
            return False
    else:
        print("請手動安裝缺少的套件:")
        print(f"pip install {' '.join(missing_packages)}")
        return False

def launch_streamlit():
    """啟動 Streamlit 版本"""
    print("🌐 啟動 Streamlit 網頁版...")
    
    if not check_file_exists('streamlit_app.py'):
        return False
    
    if not check_file_exists('word_to_pptx_core.py'):
        print("❌ 找不到核心模組 word_to_pptx_core.py")
        return False
    
    try:
        # 檢查 streamlit 是否可用
        import streamlit
        
        # 啟動 streamlit
        cmd = [sys.executable, '-m', 'streamlit', 'run', 'streamlit_app.py']
        print(f"執行命令: {' '.join(cmd)}")
        
        subprocess.run(cmd)
        return True
        
    except ImportError:
        print("❌ Streamlit 未安裝，請執行: pip install streamlit")
        return False
    except Exception as e:
        print(f"❌ 啟動 Streamlit 失敗: {e}")
        return False

def launch_desktop():
    """啟動桌面版本"""
    print("🖥️ 啟動桌面版...")
    
    if not check_file_exists('standalone_app.py'):
        return False
    
    if not check_file_exists('word_to_pptx_core.py'):
        print("❌ 找不到核心模組 word_to_pptx_core.py")
        return False
    
    try:
        # 檢查 PySide6 是否可用
        import PySide6
        
        # 啟動桌面應用
        cmd = [sys.executable, 'standalone_app.py']
        print(f"執行命令: {' '.join(cmd)}")
        
        subprocess.run(cmd)
        return True
        
    except ImportError:
        print("❌ PySide6 未安裝，請執行: pip install PySide6")
        return False
    except Exception as e:
        print(f"❌ 啟動桌面版失敗: {e}")
        return False

def show_system_info():
    """顯示系統資訊"""
    print("🔍 系統資訊:")
    print(f"  🖥️  作業系統: {platform.system()} {platform.release()}")
    print(f"  🐍 Python版本: {platform.python_version()}")
    print(f"  📁 當前目錄: {os.getcwd()}")
    
    # 檢查檔案
    files = ['word_to_pptx_core.py', 'streamlit_app.py', 'standalone_app.py']
    print("\n📄 檔案檢查:")
    for file in files:
        status = "✅" if os.path.exists(file) else "❌"
        print(f"  {status} {file}")

def show_help():
    """顯示幫助資訊"""
    help_text = """
🚀 Word轉PowerPoint工具 v5.0 - 啟動器

用法:
  python launcher.py [選項]

選項:
  -w, --web       啟動 Streamlit 網頁版
  -d, --desktop   啟動 PySide6 桌面版
  -c, --check     檢查系統環境和相依套件
  -i, --install   檢查並安裝缺少的套件
  -h, --help      顯示此幫助資訊

範例:
  python launcher.py --web      # 啟動網頁版
  python launcher.py --desktop  # 啟動桌面版
  python launcher.py --check    # 檢查環境

功能特色:
  🎯 智慧識別中文章節標題
  📐 自動分頁與溢出檢測
  🖼️ 1080p 高品質預覽
  🌈 漸層背景美化
  🌍 跨平台支援 (Windows/macOS/Linux)

支援格式:
  📄 Word: .docx, .doc
  📊 PowerPoint: .pptx, .ppt
  🖼️ 預覽: .jpg (1920×1080)
"""
    print(help_text)

def interactive_menu():
    """互動式選單"""
    print("\n" + "="*60)
    print("🚀 Word轉PowerPoint工具 v5.0")
    print("="*60)
    
    while True:
        print("\n請選擇啟動模式:")
        print("1. 🌐 Streamlit 網頁版")
        print("2. 🖥️  PySide6 桌面版") 
        print("3. 🔍 檢查系統環境")
        print("4. 📦 安裝相依套件")
        print("5. ❓ 顯示幫助")
        print("0. 🚪 退出")
        
        try:
            choice = input("\n請輸入選項 (0-5): ").strip()
            
            if choice == '1':
                if launch_streamlit():
                    break
            elif choice == '2':
                if launch_desktop():
                    break
            elif choice == '3':
                show_system_info()
                missing = check_dependencies()
                if missing:
                    print(f"\n缺少套件: {', '.join(missing)}")
                else:
                    print("\n✅ 所有套件都已安裝!")
            elif choice == '4':
                missing = check_dependencies()
                if missing:
                    install_dependencies(missing)
                else:
                    print("✅ 所有套件都已安裝!")
            elif choice == '5':
                show_help()
            elif choice == '0':
                print("👋 再見!")
                break
            else:
                print("❌ 無效選項，請重新選擇")
                
        except KeyboardInterrupt:
            print("\n\n👋 再見!")
            break
        except Exception as e:
            print(f"❌ 發生錯誤: {e}")

def main():
    """主函式"""
    parser = argparse.ArgumentParser(
        description='Word轉PowerPoint工具 v5.0 啟動器',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('-w', '--web', action='store_true',
                       help='啟動 Streamlit 網頁版')
    parser.add_argument('-d', '--desktop', action='store_true',
                       help='啟動 PySide6 桌面版')
    parser.add_argument('-c', '--check', action='store_true',
                       help='檢查系統環境和相依套件')
    parser.add_argument('-i', '--install', action='store_true',
                       help='檢查並安裝缺少的套件')
    
    args = parser.parse_args()
    
    # 顯示歡迎資訊
    print("🚀 Word轉PowerPoint工具 v5.0 啟動器")
    print(f"🖥️  系統: {platform.system()}")
    print(f"🐍 Python: {platform.python_version()}")
    print()
    
    # 處理命令行參數
    if args.web:
        launch_streamlit()
    elif args.desktop:
        launch_desktop()
    elif args.check:
        show_system_info()
        missing = check_dependencies()
        if missing:
            print(f"\n缺少套件: {', '.join(missing)}")
            print("執行 'python launcher.py --install' 來安裝")
        else:
            print("\n✅ 環境檢查完成，所有套件都已安裝!")
    elif args.install:
        missing = check_dependencies()
        if missing:
            install_dependencies(missing)
        else:
            print("✅ 所有套件都已安裝!")
    else:
        # 沒有參數時顯示互動式選單
        interactive_menu()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 程式已中斷")
    except Exception as e:
        print(f"\n❌ 發生未預期的錯誤: {e}")
        import traceback
        traceback.print_exc()