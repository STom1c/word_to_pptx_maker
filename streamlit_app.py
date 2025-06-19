#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit Word轉PowerPoint應用程式
基於核心模組的Web界面 - 雲端部署版
"""

import streamlit as st
import os
import sys
import tempfile
import zipfile
import shutil
import logging
from io import BytesIO
import platform
from typing import List
from pathlib import Path

# 設定日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 確保當前目錄在 Python 路徑中 (用於雲端部署)
current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

# 導入核心模組 - 增強錯誤處理
CORE_AVAILABLE = False
try:
    # 嘗試導入核心模組
    import word_to_pptx_core
    from word_to_pptx_core import (
        WordToPPTXConverter, 
        check_dependencies, 
        get_dependency_status,
        ConversionResult
    )
    CORE_AVAILABLE = True
    logger.info("✅ 核心模組載入成功")
except ImportError as e:
    logger.error(f"❌ 核心模組載入失敗: {e}")
    st.error("❌ 核心模組載入失敗，請檢查 word_to_pptx_core.py 檔案")
except Exception as e:
    logger.error(f"❌ 核心模組初始化失敗: {e}")
    st.error(f"❌ 核心模組初始化失敗: {str(e)}")

# 頁面配置
st.set_page_config(
    page_title="Word轉PowerPoint工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def initialize_session_state():
    """初始化session state"""
    if 'conversion_result' not in st.session_state:
        st.session_state.conversion_result = None
    if 'word_file' not in st.session_state:
        st.session_state.word_file = None
    if 'template_file' not in st.session_state:
        st.session_state.template_file = None
    if 'preview_images' not in st.session_state:
        st.session_state.preview_images = []
    if 'temp_dirs' not in st.session_state:
        st.session_state.temp_dirs = []  # 追蹤暫存目錄用於清理
    if 'save_preview_to_disk' not in st.session_state:
        st.session_state.save_preview_to_disk = False  # 預設不保存到磁碟

def cleanup_temp_files():
    """清理暫存檔案 - 雲端部署安全"""
    if 'temp_dirs' in st.session_state:
        for temp_dir in st.session_state.temp_dirs:
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    logger.info(f"已清理暫存目錄: {temp_dir}")
            except Exception as e:
                logger.warning(f"清理暫存目錄失敗: {e}")
        st.session_state.temp_dirs = []

def check_system_info():
    """檢查系統資訊 - 雲端部署版"""
    with st.sidebar:
        with st.expander("🔍 系統資訊", expanded=False):
            st.write(f"**作業系統:** {platform.system()} {platform.release()}")
            st.write(f"**Python版本:** {platform.python_version()}")
            
            # 檢查是否在雲端環境
            is_cloud = os.environ.get('STREAMLIT_SHARING_MODE') or \
                      os.environ.get('STREAMLIT_CLOUD') or \
                      'streamlit.io' in os.environ.get('HOSTNAME', '')
            
            if is_cloud:
                st.info("🌐 運行在 Streamlit Cloud 環境")
            
            # 檢查相依性
            if CORE_AVAILABLE:
                try:
                    dep_status = get_dependency_status()
                    st.write("**套件狀態:**")
                    for pkg, status in dep_status.items():
                        icon = "✅" if status else "❌"
                        st.write(f"{icon} {pkg}")
                    
                    if all(dep_status.values()):
                        st.success("所有相依套件已安裝")
                    else:
                        missing = [pkg for pkg, status in dep_status.items() if not status]
                        st.error(f"缺少套件: {', '.join(missing)}")
                        if is_cloud:
                            st.info("💡 雲端部署請檢查 requirements.txt")
                        else:
                            st.code(f"pip install {' '.join(missing)}")
                except Exception as e:
                    st.error(f"檢查相依性時發生錯誤: {e}")

def render_header():
    """渲染頁首"""
    st.markdown("""
    <div style="background: linear-gradient(90deg, #3498db, #e74c3c); padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        <h1 style="color: white; text-align: center; margin: 0;">
            🚀 Word轉PowerPoint工具 v5.0
        </h1>
        <p style="color: white; text-align: center; margin: 0.5rem 0 0 0;">
            智慧章節識別 | 自動分頁 | 1080p高品質預覽 | 跨平台支援
        </p>
    </div>
    """, unsafe_allow_html=True)

def render_features():
    """渲染功能特色"""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div style="background: #f8f9fa; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3>🎯 智慧識別</h3>
            <p>自動識別中文章節標題</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div style="background: #f8f9fa; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3>📐 自動分頁</h3>
            <p>智慧文字溢出檢測</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div style="background: #f8f9fa; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3>🖼️ 高品質預覽</h3>
            <p>1080p 解析度圖片</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown("""
        <div style="background: #f8f9fa; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3>🌍 跨平台</h3>
            <p>支援 Mac/Windows/Linux</p>
        </div>
        """, unsafe_allow_html=True)

def render_file_upload():
    """渲染檔案上傳區域"""
    st.markdown("### 📁 步驟1: 上傳檔案")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📄 Word文件")
        word_file = st.file_uploader(
            "選擇Word文件",
            type=['docx', 'doc'],
            key="word_uploader",
            help="支援 .docx 和 .doc 格式"
        )
        
        if word_file:
            st.session_state.word_file = word_file
            st.success(f"✅ 已選擇: {word_file.name}")
            st.info(f"檔案大小: {len(word_file.getvalue()) / 1024:.1f} KB")
    
    with col2:
        st.markdown("#### 📊 PowerPoint範本")
        template_file = st.file_uploader(
            "選擇PowerPoint範本",
            type=['pptx', 'ppt'],
            key="template_uploader",
            help="支援 .pptx 和 .ppt 格式"
        )
        
        if template_file:
            st.session_state.template_file = template_file
            st.success(f"✅ 已選擇: {template_file.name}")
            st.info(f"檔案大小: {len(template_file.getvalue()) / 1024:.1f} KB")
    
    # 一鍵清除功能
    st.markdown("#### 🗑️ 檔案管理")
    col_clear1, col_clear2, col_clear3 = st.columns(3)
    
    with col_clear1:
        if st.button("清除Word文件", type="secondary", use_container_width=True, key="clear_word_btn"):
            st.session_state.word_file = None
            st.success("✅ Word文件已清除")
            st.experimental_rerun()
    
    with col_clear2:
        if st.button("清除範本文件", type="secondary", use_container_width=True, key="clear_template_btn"):
            st.session_state.template_file = None
            st.success("✅ 範本文件已清除")
            st.experimental_rerun()
    
    with col_clear3:
        if st.button("🗑️ 清除所有檔案", type="secondary", use_container_width=True, key="clear_all_files_btn"):
            st.session_state.word_file = None
            st.session_state.template_file = None
            st.session_state.conversion_result = None
            st.session_state.preview_images = []
            cleanup_temp_files()
            st.success("✅ 所有檔案已清除")
            st.experimental_rerun()

def render_conversion_settings():
    """渲染轉換設定"""
    st.markdown("### ⚙️ 步驟2: 轉換設定")
    
    col1, col2 = st.columns(2)
    
    with col1:
        output_filename = st.text_input(
            "輸出檔案名稱",
            value="presentation.pptx",
            help="設定輸出的PowerPoint檔案名稱"
        )
    
    with col2:
        generate_preview = st.checkbox(
            "生成預覽圖片",
            value=True,
            help="生成1080p高品質投影片預覽圖片（總是顯示預覽）"
        )
    
    # 預覽圖片保存設定
    st.markdown("#### 🖼️ 預覽圖片設定")
    col_preview1, col_preview2 = st.columns(2)
    
    with col_preview1:
        save_preview_to_disk = st.checkbox(
            "💾 保存預覽圖片到本地",
            value=st.session_state.save_preview_to_disk,
            help="是否將預覽圖片保存到本地磁碟（預設關閉，不影響線上預覽功能）"
        )
        st.session_state.save_preview_to_disk = save_preview_to_disk
    
    with col_preview2:
        if save_preview_to_disk:
            st.info("🔸 預覽圖片將保存為ZIP檔案供下載")
        else:
            st.info("🔸 僅在記憶體中生成預覽，不保存到磁碟")
    
    # 進階設定
    with st.expander("🔧 進階設定", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            max_content_length = st.slider(
                "單張投影片最大內容長度",
                min_value=100,
                max_value=500,
                value=220,
                help="超過此長度將自動分頁"
            )
        
        with col2:
            max_content_items = st.slider(
                "單張投影片最大內容項目數",
                min_value=2,
                max_value=8,
                value=4,
                help="超過此項目數將自動分頁"
            )
    
    return output_filename, generate_preview, max_content_length, max_content_items

def perform_conversion(output_filename: str, generate_preview: bool):
    """執行轉換 - 雲端部署優化版"""
    if not st.session_state.word_file or not st.session_state.template_file:
        st.error("❌ 請先上傳Word文件和PowerPoint範本")
        return
    
    if not CORE_AVAILABLE:
        st.error("❌ 核心模組不可用")
        return
    
    # 檢查相依性
    try:
        if not check_dependencies():
            st.error("❌ 相依套件不完整，請檢查 requirements.txt")
            return
    except Exception as e:
        st.error(f"❌ 檢查相依性時發生錯誤: {e}")
        return
    
    # 清理之前的暫存檔案
    cleanup_temp_files()
    
    # 顯示進度
    progress_container = st.container()
    with progress_container:
        progress_bar = st.progress(0)
        status_text = st.empty()
    
    try:
        # 初始化轉換器
        status_text.text("🔄 初始化轉換器...")
        progress_bar.progress(10)
        logger.info("開始轉換過程")
        
        converter = WordToPPTXConverter()
        
        # 準備檔案內容
        status_text.text("📄 讀取檔案內容...")
        progress_bar.progress(30)
        
        word_content = st.session_state.word_file.getvalue()
        template_content = st.session_state.template_file.getvalue()
        
        # 檔案大小檢查 (雲端部署限制)
        max_file_size = 50 * 1024 * 1024  # 50MB
        if len(word_content) > max_file_size:
            st.error(f"❌ Word檔案過大 ({len(word_content)/1024/1024:.1f}MB)，請使用小於50MB的檔案")
            return
        
        if len(template_content) > max_file_size:
            st.error(f"❌ 範本檔案過大 ({len(template_content)/1024/1024:.1f}MB)，請使用小於50MB的檔案")
            return
        
        # 執行轉換 - 根據設定決定是否保存預覽圖片到磁碟
        status_text.text("🔄 正在轉換...")
        progress_bar.progress(50)
        
        # 決定預覽模式：總是生成預覽，但根據設定決定是否保存到磁碟
        save_to_disk = st.session_state.save_preview_to_disk if generate_preview else False
        
        result = converter.convert(
            word_file_content=word_content,
            template_file_content=template_content,
            output_path=None,  # Streamlit模式不需要儲存檔案
            generate_preview=generate_preview,
            save_preview_to_disk=save_to_disk  # 傳遞保存設定
        )
        
        progress_bar.progress(80)
        
        if result.success:
            status_text.text("✅ 轉換完成！")
            progress_bar.progress(100)
            logger.info("轉換成功完成")
            
            # 儲存結果到session state
            st.session_state.conversion_result = result
            
            # 生成下載用的pptx檔案
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dirs.append(temp_dir)  # 記錄暫存目錄
            output_path = os.path.join(temp_dir, output_filename)
            
            # 重新執行轉換以生成檔案
            status_text.text("📁 準備下載檔案...")
            file_result = converter.convert(
                word_file_content=word_content,
                template_file_content=template_content,
                output_path=output_path,
                generate_preview=False  # 避免重複生成預覽
            )
            
            if file_result.success:
                # 讀取生成的檔案
                with open(output_path, 'rb') as f:
                    pptx_data = f.read()
                
                # 提供下載
                st.download_button(
                    label="📥 下載PowerPoint檔案",
                    data=pptx_data,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_pptx"
                )
                
                st.success(f"🎉 轉換成功！共生成 {result.slides_count} 張投影片")
                logger.info(f"轉換完成，生成 {result.slides_count} 張投影片")
                
                # 預覽圖片處理
                if generate_preview and result.preview_images:
                    st.session_state.preview_images = result.preview_images
                    
                    # 如果啟用了保存到磁碟，提供預覽圖片下載
                    if st.session_state.save_preview_to_disk:
                        render_preview_download()
                        st.info("💾 預覽圖片已保存，可下載ZIP檔案")
                    else:
                        st.info("🖼️ 預覽圖片僅在記憶體中，可在下方查看")
                elif generate_preview:
                    st.warning("⚠️ 預覽圖片生成失敗，但轉換成功")
            else:
                st.error(f"❌ 生成下載檔案失敗: {file_result.error_message}")
                logger.error(f"生成下載檔案失敗: {file_result.error_message}")
            
        else:
            status_text.text("❌ 轉換失敗")
            st.error(f"轉換失敗: {result.error_message}")
            logger.error(f"轉換失敗: {result.error_message}")
            
    except Exception as e:
        status_text.text("❌ 發生錯誤")
        error_msg = f"轉換過程中發生錯誤: {str(e)}"
        st.error(error_msg)
        logger.error(error_msg, exc_info=True)
    finally:
        # 隱藏進度條
        with progress_container:
            progress_bar.empty()
            status_text.empty()

def render_preview_download():
    """渲染預覽圖片下載"""
    if not st.session_state.preview_images:
        return
    
    # 創建預覽圖片zip檔案
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for i, image_path in enumerate(st.session_state.preview_images, 1):
            if os.path.exists(image_path):
                with open(image_path, 'rb') as img_file:
                    zip_file.writestr(f"slide_{i:02d}.jpg", img_file.read())
    
    zip_buffer.seek(0)
    
    st.download_button(
        label="🖼️ 下載預覽圖片 (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="slides_preview.zip",
        mime="application/zip",
        key="download_preview_zip"
    )

def render_preview():
    """渲染預覽區域"""
    if not st.session_state.preview_images:
        return
    
    st.markdown("### 🖼️ 投影片預覽")
    
    # 顯示預覽資訊
    st.info(f"📊 共生成 {len(st.session_state.preview_images)} 張 1080p 高品質預覽圖片")
    
    # 圖片展示選項
    col1, col2, col3 = st.columns([1, 1, 2])
    
    with col1:
        view_mode = st.selectbox(
            "預覽模式",
            ["網格檢視", "單張檢視"],
            key="view_mode"
        )
    
    with col2:
        if view_mode == "網格檢視":
            columns = st.selectbox("每行圖片數", [1, 2, 3], index=1, key="grid_columns")
        else:
            slide_index = st.selectbox(
                "選擇投影片",
                range(1, len(st.session_state.preview_images) + 1),
                key="slide_index"
            )
    
    # 顯示預覽圖片
    if view_mode == "網格檢視":
        # 網格檢視
        for i in range(0, len(st.session_state.preview_images), columns):
            cols = st.columns(columns)
            for j in range(columns):
                if i + j < len(st.session_state.preview_images):
                    image_path = st.session_state.preview_images[i + j]
                    if os.path.exists(image_path):
                        with cols[j]:
                            st.image(
                                image_path,
                                caption=f"投影片 {i + j + 1}",
                                use_column_width=True
                            )
    else:
        # 單張檢視
        image_path = st.session_state.preview_images[slide_index - 1]
        if os.path.exists(image_path):
            st.image(
                image_path,
                caption=f"投影片 {slide_index}",
                use_column_width=True
            )
            
            # 顯示圖片資訊
            try:
                from PIL import Image
                img = Image.open(image_path)
                file_size = os.path.getsize(image_path)
                st.caption(f"尺寸: {img.width}×{img.height} | 大小: {file_size/1024:.1f} KB")
            except:
                pass

def render_usage_guide():
    """渲染使用指南"""
    with st.sidebar:
        with st.expander("📖 使用指南", expanded=False):
            st.markdown("""
            **使用步驟:**
            1. 上傳Word文件 (.docx/.doc)
            2. 上傳PowerPoint範本 (.pptx/.ppt)
            3. 設定輸出檔案名稱
            4. 點擊開始轉換
            5. 下載生成的檔案和預覽圖片
            
            **功能特色:**
            - 🎯 智慧識別中文章節標題
            - 📐 自動分頁與溢出檢測
            - 🖼️ 1080p 高品質預覽
            - 🌈 漸層背景美化
            - 🔤 跨平台字體優化
            
            **支援格式:**
            - Word: .docx, .doc
            - PowerPoint: .pptx, .ppt
            - 預覽: .jpg (1920×1080)
            """)

def main():
    """主函式 - 雲端部署版"""
    try:
        # 初始化
        initialize_session_state()
        
        # 檢查核心模組
        if not CORE_AVAILABLE:
            st.error("❌ 核心模組載入失敗")
            st.info("💡 如果您在雲端部署，請確保 word_to_pptx_core.py 檔案已正確上傳")
            st.stop()
        
        # 頁面清理按鈕 (在側邊欄)
        with st.sidebar:
            if st.button("🗑️ 清理暫存檔案", help="清理伺服器上的暫存檔案", key="cleanup_temp_btn"):
                cleanup_temp_files()
                st.success("✅ 暫存檔案已清理")
        
        # 渲染頁面
        render_header()
        render_features()
        check_system_info()
        render_usage_guide()
        
        # 主要內容區域
        col1, col2 = st.columns([2, 1])
        
        with col1:
            # 檔案上傳
            render_file_upload()
            
            # 轉換設定
            output_filename, generate_preview, max_content_length, max_content_items = render_conversion_settings()
            
            # 轉換按鈕
            st.markdown("### 🚀 步驟3: 開始轉換")
            
            col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
            with col_btn2:
                if st.button(
                    "🚀 開始轉換為PowerPoint",
                    type="primary",
                    use_container_width=True,
                    disabled=not (st.session_state.word_file and st.session_state.template_file),
                    key="start_conversion_btn"
                ):
                    perform_conversion(output_filename, generate_preview)
        
        with col2:
            # 快速操作
            st.markdown("### ⚡ 快速操作")
            
            if st.button("🗑️ 清除所有檔案", type="secondary", use_container_width=True):
                st.session_state.word_file = None
                st.session_state.template_file = None
                st.session_state.conversion_result = None
                st.session_state.preview_images = []
                cleanup_temp_files()
                st.experimental_rerun()
            
            # 範例檔案下載（如果有的話）
            st.markdown("### 📝 範例檔案")
            st.markdown("""
            **範例Word文件結構:**
            ```
            一、第一章標題
            內容描述...
            
            二、第二章標題
            (一) 第一節
            詳細內容...
            
            (二) 第二節
            更多內容...
            ```
            """)
            
            # 雲端部署資訊
            if os.environ.get('STREAMLIT_SHARING_MODE') or os.environ.get('STREAMLIT_CLOUD'):
                st.markdown("### 🌐 雲端部署")
                st.info("此應用程式運行在 Streamlit Cloud 上")
        
        # 預覽區域（全寬）
        if st.session_state.preview_images:
            st.markdown("---")
            render_preview()
        
        # 頁腳
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**🚀 Word轉PowerPoint工具 v5.0**")
        
        with col2:
            st.markdown(f"**🖥️ 運行平台:** {platform.system()}")
        
        with col3:
            st.markdown("**🌐 Streamlit 雲端版**")
    
    except Exception as e:
        st.error(f"❌ 應用程式發生錯誤: {str(e)}")
        logger.error(f"應用程式錯誤: {e}", exc_info=True)
        
        # 提供重啟選項
        if st.button("🔄 重新載入應用程式", key="reload_app_btn"):
            st.experimental_rerun()

# 確保暫存檔案在應用程式結束時被清理
import atexit
atexit.register(lambda: cleanup_temp_files() if 'st' in globals() else None)

if __name__ == "__main__":
    main()