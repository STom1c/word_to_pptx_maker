#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit 雲端部署測試腳本
在部署前驗證所有功能是否正常運作
"""

import os
import sys
import platform
import importlib
import subprocess
from pathlib import Path

def print_header(title):
    """列印標題"""
    print("\n" + "="*60)
    print(f"📋 {title}")
    print("="*60)

def print_result(test_name, success, details=""):
    """列印測試結果"""
    icon = "✅" if success else "❌"
    print(f"{icon} {test_name}")
    if details:
        print(f"   📝 {details}")

def test_python_version():
    """測試 Python 版本"""
    print_header("Python 版本檢查")
    
    version = sys.version_info
    required_major, required_minor = 3, 7
    
    current_version = f"{version.major}.{version.minor}.{version.micro}"
    success = version.major >= required_major and version.minor >= required_minor
    
    print_result(
        f"Python 版本: {current_version}",
        success,
        f"需要 Python {required_major}.{required_minor}+" if not success else "版本符合要求"
    )
    
    return success

def test_required_files():
    """測試必要檔案"""
    print_header("必要檔案檢查")
    
    required_files = [
        "streamlit_app.py",
        "word_to_pptx_core.py", 
        "requirements.txt"
    ]
    
    optional_files = [
        ".streamlit/config.toml",
        "README.md",
        ".gitignore"
    ]
    
    all_success = True
    
    # 檢查必要檔案
    for file in required_files:
        exists = os.path.exists(file)
        print_result(f"必要檔案: {file}", exists)
        if not exists:
            all_success = False
    
    # 檢查可選檔案
    for file in optional_files:
        exists = os.path.exists(file)
        print_result(f"可選檔案: {file}", exists, "建議添加" if not exists else "")
    
    return all_success

def test_dependencies():
    """測試相依套件"""
    print_header("相依套件檢查")
    
    required_packages = {
        'streamlit': 'streamlit',
        'docx': 'python-docx',
        'pptx': 'python-pptx',
        'PIL': 'Pillow'
    }
    
    all_success = True
    
    for module, package in required_packages.items():
        try:
            importlib.import_module(module)
            print_result(f"套件: {package}", True)
        except ImportError:
            print_result(f"套件: {package}", False, f"請執行: pip install {package}")
            all_success = False
    
    return all_success

def test_core_module():
    """測試核心模組"""
    print_header("核心模組檢查")
    
    try:
        # 嘗試導入核心模組
        sys.path.insert(0, os.getcwd())
        import word_to_pptx_core
        
        # 檢查核心類別
        required_classes = [
            'WordToPPTXConverter',
            'WordDocumentAnalyzer', 
            'ContentToSlideMapper',
            'PPTXImageExporter'
        ]
        
        all_success = True
        
        for class_name in required_classes:
            if hasattr(word_to_pptx_core, class_name):
                print_result(f"核心類別: {class_name}", True)
            else:
                print_result(f"核心類別: {class_name}", False)
                all_success = False
        
        # 測試基本功能
        try:
            converter = word_to_pptx_core.WordToPPTXConverter()
            print_result("核心模組初始化", True)
        except Exception as e:
            print_result("核心模組初始化", False, str(e))
            all_success = False
        
        return all_success
        
    except ImportError as e:
        print_result("核心模組導入", False, str(e))
        return False

def test_streamlit_app():
    """測試 Streamlit 應用程式"""
    print_header("Streamlit 應用程式檢查")
    
    try:
        # 檢查 streamlit_app.py 語法
        with open('streamlit_app.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        compile(content, 'streamlit_app.py', 'exec')
        print_result("Streamlit 應用程式語法", True)
        
        # 嘗試導入應用程式模組
        import streamlit_app
        print_result("Streamlit 應用程式導入", True)
        
        return True
        
    except SyntaxError as e:
        print_result("Streamlit 應用程式語法", False, f"語法錯誤: {e}")
        return False
    except Exception as e:
        print_result("Streamlit 應用程式導入", False, str(e))
        return False

def test_requirements_file():
    """測試 requirements.txt"""
    print_header("Requirements 檔案檢查")
    
    try:
        with open('requirements.txt', 'r', encoding='utf-8') as f:
            requirements = f.read().strip().split('\n')
        
        # 過濾空行和註解
        packages = [req.strip() for req in requirements 
                   if req.strip() and not req.strip().startswith('#')]
        
        print_result(f"Requirements 檔案", True, f"找到 {len(packages)} 個套件")
        
        # 檢查核心套件
        core_packages = ['streamlit', 'python-docx', 'python-pptx', 'Pillow']
        missing_core = []
        
        for core_pkg in core_packages:
            found = any(core_pkg.lower() in pkg.lower() for pkg in packages)
            if not found:
                missing_core.append(core_pkg)
        
        if missing_core:
            print_result("核心套件完整性", False, f"缺少: {', '.join(missing_core)}")
            return False
        else:
            print_result("核心套件完整性", True)
            return True
        
    except FileNotFoundError:
        print_result("Requirements 檔案", False, "檔案不存在")
        return False

def test_streamlit_config():
    """測試 Streamlit 配置"""
    print_header("Streamlit 配置檢查")
    
    config_path = ".streamlit/config.toml"
    
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_content = f.read()
            
            # 檢查基本配置項
            required_sections = ['server', 'browser', 'theme']
            all_found = True
            
            for section in required_sections:
                if f'[{section}]' in config_content:
                    print_result(f"配置段落: [{section}]", True)
                else:
                    print_result(f"配置段落: [{section}]", False, "建議添加")
                    all_found = False
            
            return True
            
        except Exception as e:
            print_result("Streamlit 配置檔案", False, str(e))
            return False
    else:
        print_result("Streamlit 配置檔案", False, "檔案不存在，建議創建")
        return False

def test_local_streamlit():
    """測試本地 Streamlit 啟動"""
    print_header("本地 Streamlit 測試")
    
    try:
        # 檢查 streamlit 命令是否可用
        result = subprocess.run(['streamlit', '--version'], 
                              capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            print_result("Streamlit CLI", True, f"版本: {result.stdout.strip()}")
            
            # 提示本地測試
            print("\n💡 本地測試建議:")
            print("   執行: streamlit run streamlit_app.py")
            print("   然後在瀏覽器中測試所有功能")
            
            return True
        else:
            print_result("Streamlit CLI", False, "無法執行 streamlit 命令")
            return False
            
    except subprocess.TimeoutExpired:
        print_result("Streamlit CLI", False, "命令執行超時")
        return False
    except FileNotFoundError:
        print_result("Streamlit CLI", False, "未安裝 streamlit")
        return False

def generate_deployment_checklist():
    """生成部署檢查清單"""
    print_header("部署檢查清單")
    
    checklist = [
        "✅ Python 版本 >= 3.7",
        "✅ 所有必要檔案存在",
        "✅ 相依套件已安裝",
        "✅ 核心模組正常運作",
        "✅ Streamlit 應用程式語法正確",
        "✅ requirements.txt 完整",
        "✅ Streamlit 配置檔案 (可選)",
        "✅ 本地測試通過"
    ]
    
    print("📋 部署前確認清單:")
    for item in checklist:
        print(f"   {item}")
    
    print("\n🚀 部署平台選擇:")
    print("   • Streamlit Cloud (推薦): share.streamlit.io")
    print("   • Heroku: heroku.com")  
    print("   • Railway: railway.app")
    print("   • Google Cloud Run: cloud.google.com")
    
    print("\n📚 詳細部署指南:")
    print("   查看 'Streamlit 雲端部署指南.md' 檔案")

def main():
    """主函式"""
    print("🌐 Streamlit 雲端部署測試")
    print(f"🖥️  系統: {platform.system()} {platform.release()}")
    print(f"🐍 Python: {platform.python_version()}")
    
    # 執行所有測試
    tests = [
        ("Python 版本", test_python_version),
        ("必要檔案", test_required_files),
        ("相依套件", test_dependencies),
        ("核心模組", test_core_module),
        ("Streamlit 應用程式", test_streamlit_app),
        ("Requirements 檔案", test_requirements_file),
        ("Streamlit 配置", test_streamlit_config),
        ("本地 Streamlit", test_local_streamlit)
    ]
    
    results = []
    
    for test_name, test_func in tests:
        try:
            success = test_func()
            results.append((test_name, success))
        except Exception as e:
            print_result(f"測試 {test_name}", False, f"測試失敗: {e}")
            results.append((test_name, False))
    
    # 總結
    print_header("測試總結")
    
    passed = sum(1 for _, success in results if success)
    total = len(results)
    
    print(f"📊 測試結果: {passed}/{total} 通過")
    
    if passed == total:
        print("🎉 恭喜！所有測試都通過了，可以開始部署！")
    else:
        print("⚠️  有部分測試未通過，請修復後再次測試")
        
        failed_tests = [name for name, success in results if not success]
        print(f"❌ 未通過的測試: {', '.join(failed_tests)}")
    
    # 生成檢查清單
    generate_deployment_checklist()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 測試已中斷")
    except Exception as e:
        print(f"\n❌ 測試過程中發生錯誤: {e}")
        import traceback
        traceback.print_exc()