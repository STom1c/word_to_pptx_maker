#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word轉PowerPoint獨立應用程式
支援 macOS 和 Windows 的桌面版本
基於 PySide6 和核心模組
"""

import sys
import os
import platform
import json
import subprocess
from pathlib import Path

# PySide6 導入
try:
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
        QPushButton, QLabel, QProgressBar, QFileDialog,
        QWidget, QFrame, QScrollArea, QGroupBox,
        QMessageBox, QSplitter, QTextEdit, QCheckBox
    )
    from PySide6.QtCore import Qt, QThread, Signal
    from PySide6.QtGui import QFont, QPixmap, QDragEnterEvent, QDropEvent
    PYSIDE6_AVAILABLE = True
except ImportError:
    PYSIDE6_AVAILABLE = False
    print("❌ PySide6 不可用，請安裝: pip install PySide6")

# 核心模組導入
try:
    from word_to_pptx_core import (
        WordToPPTXConverter, 
        check_dependencies, 
        get_dependency_status,
        ConversionResult
    )
    CORE_AVAILABLE = True
except ImportError:
    CORE_AVAILABLE = False
    print("❌ 核心模組不可用，請確保 word_to_pptx_core.py 在同一目錄下")

class ConfigManager:
    """設定管理器 - 記憶上次使用的檔案路徑"""
    
    def __init__(self):
        self.config_file = os.path.join(os.path.expanduser("~"), ".word_to_pptx_config.json")
        self.config = self.load_config()
    
    def load_config(self) -> dict:
        """載入設定檔"""
        default_config = {
            "last_word_path": "",
            "last_template_path": "",
            "last_output_dir": "",
            "window_geometry": None
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 驗證檔案是否仍然存在
                    if config.get("last_word_path") and not os.path.exists(config["last_word_path"]):
                        config["last_word_path"] = ""
                    if config.get("last_template_path") and not os.path.exists(config["last_template_path"]):
                        config["last_template_path"] = ""
                    return {**default_config, **config}
        except Exception as e:
            print(f"載入設定檔時發生錯誤: {e}")
        
        return default_config
    
    def save_config(self):
        """儲存設定檔"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"儲存設定檔時發生錯誤: {e}")
    
    def set_last_word_path(self, path: str):
        self.config["last_word_path"] = path
        self.save_config()
    
    def set_last_template_path(self, path: str):
        self.config["last_template_path"] = path
        self.save_config()
    
    def set_last_output_dir(self, dir_path: str):
        self.config["last_output_dir"] = dir_path
        self.save_config()
    
    def get_last_word_path(self) -> str:
        return self.config.get("last_word_path", "")
    
    def get_last_template_path(self) -> str:
        return self.config.get("last_template_path", "")
    
    def get_last_output_dir(self) -> str:
        return self.config.get("last_output_dir", "")

class ConversionWorker(QThread):
    """轉換工作執行緒"""
    
    progress_updated = Signal(int)
    status_updated = Signal(str)
    finished_successfully = Signal(ConversionResult)
    error_occurred = Signal(str)
    
    def __init__(self, word_path: str, template_path: str, output_path: str, save_preview_to_disk: bool = True):
        super().__init__()
        self.word_path = word_path
        self.template_path = template_path
        self.output_path = output_path
        self.save_preview_to_disk = save_preview_to_disk
    
    def run(self):
        """執行轉換"""
        try:
            self.status_updated.emit("正在初始化轉換器...")
            self.progress_updated.emit(10)
            
            converter = WordToPPTXConverter()
            
            self.status_updated.emit("正在轉換文件...")
            self.progress_updated.emit(50)
            
            result = converter.convert(
                word_file_path=self.word_path,
                template_file_path=self.template_path,
                output_path=self.output_path,
                generate_preview=True,
                save_preview_to_disk=self.save_preview_to_disk
            )
            
            self.progress_updated.emit(100)
            
            if result.success:
                self.status_updated.emit("轉換完成！")
                self.finished_successfully.emit(result)
            else:
                self.error_occurred.emit(result.error_message)
                
        except Exception as e:
            import traceback
            error_details = f"{str(e)}\n\n詳細錯誤資訊:\n{traceback.format_exc()}"
            self.error_occurred.emit(error_details)

class DropArea(QFrame):
    """拖放區域元件"""
    
    file_dropped = Signal(str)
    
    def __init__(self, file_type: str, parent=None):
        super().__init__(parent)
        self.file_type = file_type
        self.setup_ui()
        
    def setup_ui(self):
        """設定UI"""
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.StyledPanel)
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #3498db;
                border-radius: 8px;
                background-color: #f8f9fa;
                padding: 15px;
                min-height: 120px;
                max-height: 160px;
            }
            DropArea:hover {
                background-color: #e3f2fd;
                border-color: #2196f3;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 圖示
        icon_label = QLabel("📄" if self.file_type == "word" else "📊")
        icon_label.setFont(QFont("Arial", 36))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setMaximumHeight(45)
        
        # 文字說明
        if self.file_type == "word":
            text = "拖放WORD檔案到此處\n或點擊選擇檔案"
        else:
            text = "拖放POWERPOINT範本到此處\n或點擊選擇檔案"
        text_label = QLabel(text)
        text_label.setFont(QFont("Microsoft JhengHei", 10))
        text_label.setAlignment(Qt.AlignCenter)
        text_label.setStyleSheet("color: #666; margin: 5px; line-height: 1.4;")
        text_label.setWordWrap(True)
        text_label.setMaximumHeight(40)
        
        layout.addWidget(icon_label)
        layout.addWidget(text_label)
        layout.addStretch()
        
        self.setMinimumHeight(120)
        self.setMaximumHeight(160)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖拽進入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                DropArea {
                    border: 2px solid #4caf50;
                    background-color: #e8f5e8;
                }
            """)
    
    def dragLeaveEvent(self, event):
        """拖拽離開事件"""
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #3498db;
                background-color: #f8f9fa;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        """放置事件"""
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            self.file_dropped.emit(files[0])
        
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #3498db;
                background-color: #f8f9fa;
            }
        """)
    
    def mousePressEvent(self, event):
        """滑鼠點擊選擇檔案"""
        if self.file_type == "word":
            file_filter = "Word Documents (*.docx *.doc)"
            dialog_title = "選擇WORD檔案"
        else:
            file_filter = "PowerPoint Files (*.pptx *.ppt)"
            dialog_title = "選擇POWERPOINT檔案"
            
        file_path, _ = QFileDialog.getOpenFileName(
            self, dialog_title, "", file_filter
        )
        
        if file_path:
            self.file_dropped.emit(file_path)

class PreviewWidget(QScrollArea):
    """預覽元件"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        """設定UI"""
        self.setWidgetResizable(True)
        self.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ddd;
                border-radius: 8px;
                background-color: #f8f9fa;
            }
        """)
        
        content = QWidget()
        self.content_layout = QVBoxLayout(content)
        self.content_layout.setSpacing(15)
        self.content_layout.setContentsMargins(15, 15, 15, 15)
        
        self.setWidget(content)
    
    def update_preview(self, image_paths: list):
        """更新預覽"""
        try:
            self.clear_preview()
            
            if not image_paths:
                error_label = QLabel("❌ 無法載入預覽圖片")
                error_label.setStyleSheet("color: #e74c3c; padding: 20px; text-align: center;")
                error_label.setAlignment(Qt.AlignCenter)
                self.content_layout.addWidget(error_label)
                return
            
            # 顯示預覽圖片資訊
            info_label = QLabel(f"🖼️ 高品質預覽圖片 | 共 {len(image_paths)} 張 1080p 圖片")
            info_label.setFont(QFont("Microsoft JhengHei", 10, QFont.Weight.Bold))
            info_label.setStyleSheet("color: #27ae60; padding: 10px; background: #f0f8f0; border-radius: 5px; margin-bottom: 10px;")
            info_label.setAlignment(Qt.AlignCenter)
            self.content_layout.addWidget(info_label)
            
            success_count = 0
            for i, image_path in enumerate(image_paths):
                if os.path.exists(image_path):
                    try:
                        preview_item = self.create_image_preview(image_path, i + 1)
                        self.content_layout.addWidget(preview_item)
                        success_count += 1
                        QApplication.processEvents()
                    except Exception as e:
                        print(f"建立預覽項目失敗: {e}")
                        continue
            
            if success_count > 0:
                result_label = QLabel(f"✅ 成功載入 {success_count} 張預覽圖片")
                result_label.setStyleSheet("color: #27ae60; padding: 5px; font-weight: bold;")
                result_label.setAlignment(Qt.AlignCenter)
                self.content_layout.addWidget(result_label)
            
        except Exception as e:
            self.clear_preview()
            error_label = QLabel(f"❌ 預覽載入失敗: {str(e)}")
            error_label.setStyleSheet("color: #e74c3c; padding: 20px; background: #fdf2f2; border-radius: 5px;")
            error_label.setWordWrap(True)
            self.content_layout.addWidget(error_label)
    
    def clear_preview(self):
        """清除預覽"""
        while self.content_layout.count():
            child = self.content_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
    
    def create_image_preview(self, image_path: str, slide_number: int) -> QWidget:
        """建立圖片預覽"""
        frame = QFrame()
        frame.setFrameStyle(QFrame.Box)
        frame.setStyleSheet("""
            QFrame {
                border: 2px solid #3498db;
                border-radius: 10px;
                background-color: white;
                margin: 5px;
                padding: 10px;
            }
        """)
        
        layout = QVBoxLayout(frame)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)
        
        # 標題欄
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #3498db;
                border-radius: 5px;
                padding: 5px;
                border: none;
            }
        """)
        
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(8, 4, 8, 4)
        
        number_label = QLabel(f"投影片 {slide_number}")
        number_label.setFont(QFont("Microsoft JhengHei", 11, QFont.Weight.Bold))
        number_label.setStyleSheet("color: white;")
        
        engine_label = QLabel("核心引擎")
        engine_label.setFont(QFont("Microsoft JhengHei", 8))
        engine_label.setStyleSheet("color: #f1c40f; background: rgba(255,255,255,20); padding: 2px 6px; border-radius: 3px;")
        
        header_layout.addWidget(number_label, 1)
        header_layout.addWidget(engine_label)
        
        layout.addWidget(header_frame)
        
        # 圖片容器
        image_container = QFrame()
        image_container.setStyleSheet("""
            QFrame {
                border: 1px solid #34495e;
                border-radius: 5px;
                background-color: white;
                padding: 5px;
            }
        """)
        
        image_layout = QVBoxLayout(image_container)
        image_layout.setContentsMargins(5, 5, 5, 5)
        
        try:
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                # 保持原始比例，但限制最大尺寸
                target_width = 900
                target_height = 507  # 16:9 比例
                
                scaled_pixmap = pixmap.scaled(
                    target_width, target_height, 
                    Qt.KeepAspectRatio, 
                    Qt.SmoothTransformation
                )
                
                image_display = QLabel()
                image_display.setPixmap(scaled_pixmap)
                image_display.setAlignment(Qt.AlignCenter)
                image_display.setStyleSheet("border: none; background: white;")
                
                image_layout.addWidget(image_display)
                
                # 檔案資訊
                file_size = os.path.getsize(image_path) if os.path.exists(image_path) else 0
                relative_path = os.path.basename(image_path)
                image_info = QLabel(f"尺寸: {pixmap.width()}×{pixmap.height()} (1080p) | 大小: {file_size//1024}KB | 檔案: {relative_path}")
                image_info.setFont(QFont("Microsoft JhengHei", 8))
                image_info.setStyleSheet("color: #7f8c8d; margin-top: 5px;")
                image_info.setAlignment(Qt.AlignCenter)
                image_layout.addWidget(image_info)
                
            else:
                error_display = QLabel("圖片載入失敗")
                error_display.setAlignment(Qt.AlignCenter)
                error_display.setStyleSheet("color: #e74c3c; padding: 40px; font-size: 14px;")
                image_layout.addWidget(error_display)
                
        except Exception as e:
            print(f"圖片顯示錯誤: {e}")
            error_display = QLabel(f"圖片顯示錯誤\n{str(e)[:50]}")
            error_display.setAlignment(Qt.AlignCenter)
            error_display.setStyleSheet("color: #e74c3c; padding: 40px; font-size: 12px;")
            image_layout.addWidget(error_display)
        
        layout.addWidget(image_container)
        
        frame.setMaximumHeight(650)
        frame.setMinimumHeight(550)
        return frame

class MainWindow(QMainWindow):
    """主視窗"""
    
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        
        self.word_path = ""
        self.template_path = ""
        self.output_path = ""
        self.worker = None
        self.save_preview_to_disk = False  # 預設不保存預覽圖片到磁碟
        
        self.setup_ui()
        self.setup_connections()
        self.load_last_used_paths()
        self.check_ready_to_convert()
        
    def setup_ui(self):
        """設定UI"""
        self.setWindowTitle(f"Word轉PowerPoint工具 v5.0 - 桌面版")
        self.setGeometry(100, 100, 1400, 900)
        
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #bdc3c7;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 8px;
                font-size: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px 0 4px;
            }
        """)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        left_panel = self.create_left_panel()
        left_panel.setMaximumWidth(400)
        
        right_panel = self.create_right_panel()
        
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([300, 1000])
        
        main_layout.addWidget(splitter)
        
    def create_left_panel(self) -> QWidget:
        """建立左側面板"""
        panel = QWidget()
        panel.setMinimumWidth(300)
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # 標題
        title = QLabel(f"🚀 Word轉PowerPoint工具 ({platform.system()})")
        title.setFont(QFont("Microsoft JhengHei", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #2c3e50; margin: 10px;")
        title.setMaximumHeight(40)
        layout.addWidget(title)
        
        # 功能說明
        features = QLabel("✨ v5.0: 核心模組 | 1080p高解析度 | 智慧分頁 | 跨平台支援")
        features.setFont(QFont("Microsoft JhengHei", 9))
        features.setAlignment(Qt.AlignCenter)
        features.setStyleSheet("color: #27ae60; margin-bottom: 5px;")
        features.setMaximumHeight(25)
        layout.addWidget(features)
        
        # 檔案選擇區域
        file_group = QGroupBox("1. 選擇檔案")
        file_group.setMaximumHeight(420)
        file_layout = QVBoxLayout(file_group)
        file_layout.setSpacing(8)
        file_layout.setContentsMargins(10, 20, 10, 10)
        
        # Word檔案拖放區
        word_label = QLabel("Word文件:")
        word_label.setFont(QFont("Microsoft JhengHei", 11, QFont.Weight.Bold))
        word_label.setStyleSheet("color: #2c3e50; margin-bottom: 3px;")
        word_label.setMaximumHeight(20)
        
        self.word_drop = DropArea("word")
        self.word_drop.setMaximumHeight(140)
        
        self.word_status = QLabel("未選擇Word檔案")
        self.word_status.setStyleSheet("color: #e74c3c; margin: 5px; font-size: 10px;")
        self.word_status.setMaximumHeight(28)
        
        file_layout.addWidget(word_label)
        file_layout.addWidget(self.word_drop)
        file_layout.addWidget(self.word_status)
        
        # 分隔線
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #bdc3c7;")
        line.setMaximumHeight(5)
        file_layout.addWidget(line)
        
        # PowerPoint範本拖放區
        template_label = QLabel("PowerPoint範本:")
        template_label.setFont(QFont("Microsoft JhengHei", 11, QFont.Weight.Bold))
        template_label.setStyleSheet("color: #2c3e50; margin-bottom: 3px;")
        template_label.setMaximumHeight(20)
        
        self.template_drop = DropArea("pptx")
        self.template_drop.setMaximumHeight(140)
        
        self.template_status = QLabel("未選擇PowerPoint範本")
        self.template_status.setStyleSheet("color: #e74c3c; margin: 5px; font-size: 10px;")
        self.template_status.setMaximumHeight(28)
        
        file_layout.addWidget(template_label)
        file_layout.addWidget(self.template_drop)
        file_layout.addWidget(self.template_status)
        
        layout.addWidget(file_group)
        
        # 檔案管理區域
        management_group = QGroupBox("檔案管理")
        management_group.setMaximumHeight(120)
        management_layout = QVBoxLayout(management_group)
        management_layout.setSpacing(8)
        management_layout.setContentsMargins(10, 15, 10, 10)
        
        # 清除按鈕
        clear_layout = QHBoxLayout()
        clear_layout.setSpacing(8)
        
        clear_word_btn = QPushButton("清除Word")
        clear_word_btn.setMaximumHeight(30)
        clear_word_btn.clicked.connect(self.clear_word_file)
        
        clear_template_btn = QPushButton("清除範本")
        clear_template_btn.setMaximumHeight(30)
        clear_template_btn.clicked.connect(self.clear_template_file)
        
        clear_all_btn = QPushButton("🗑️ 清除全部")
        clear_all_btn.setMaximumHeight(30)
        clear_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        clear_all_btn.clicked.connect(self.clear_all_files)
        
        clear_layout.addWidget(clear_word_btn)
        clear_layout.addWidget(clear_template_btn)
        clear_layout.addWidget(clear_all_btn)
        
        management_layout.addLayout(clear_layout)
        
        layout.addWidget(management_group)
        
        # 轉換設定
        settings_group = QGroupBox("2. 轉換設定")
        settings_group.setMaximumHeight(220)
        settings_layout = QVBoxLayout(settings_group)
        settings_layout.setSpacing(8)
        settings_layout.setContentsMargins(10, 15, 10, 10)
        
        # 設定說明
        settings_desc = QLabel("• 智慧章節識別與自動分頁\n• 完全清除範本原始內容\n• 文字溢出自動檢測\n• 高品質 1080p 預覽圖片\n• 跨平台字體優化")
        settings_desc.setFont(QFont("Microsoft JhengHei", 9))
        settings_desc.setStyleSheet("color: #27ae60; margin-bottom: 5px;")
        settings_desc.setMaximumHeight(100)
        settings_layout.addWidget(settings_desc)
        
        # 預覽圖片設定
        preview_checkbox = QCheckBox("💾 保存預覽圖片到本地磁碟")
        preview_checkbox.setFont(QFont("Microsoft JhengHei", 9))
        preview_checkbox.setStyleSheet("color: #2c3e50; margin: 5px;")
        preview_checkbox.setChecked(self.save_preview_to_disk)
        preview_checkbox.stateChanged.connect(self.on_preview_save_changed)
        settings_layout.addWidget(preview_checkbox)
        
        preview_note = QLabel("註：預覽功能總是啟用，此選項僅控制是否永久保存到磁碟")
        preview_note.setFont(QFont("Microsoft JhengHei", 8))
        preview_note.setStyleSheet("color: #7f8c8d; margin-left: 20px;")
        preview_note.setWordWrap(True)
        settings_layout.addWidget(preview_note)
        
        # 輸出路徑
        output_layout = QHBoxLayout()
        output_layout.setSpacing(8)
        
        self.output_label = QLabel("將自動設定輸出位置...")
        self.output_label.setStyleSheet("""
            QLabel {
                border: 1px solid #ccc; 
                padding: 8px; 
                background: white; 
                color: #666;
                border-radius: 4px;
                font-size: 10px;
            }
        """)
        self.output_label.setMaximumHeight(35)
        self.output_label.setWordWrap(True)
        
        output_btn = QPushButton("選擇")
        output_btn.setMaximumHeight(35)
        output_btn.setMaximumWidth(80)
        output_btn.clicked.connect(self.select_output_path)
        
        output_layout.addWidget(self.output_label, 1)
        output_layout.addWidget(output_btn)
        
        settings_layout.addLayout(output_layout)
        
        layout.addWidget(settings_group)
        
        # 轉換按鈕
        self.convert_btn = QPushButton("🚀 開始轉換為PowerPoint")
        self.convert_btn.setEnabled(False)
        self.convert_btn.setFont(QFont("Microsoft JhengHei", 13, QFont.Weight.Bold))
        self.convert_btn.setMaximumHeight(50)
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)
        
        # 進度條
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(20)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 4px;
                text-align: center;
                font-weight: bold;
                font-size: 10px;
            }
            QProgressBar::chunk {
                background-color: #27ae60;
                border-radius: 2px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
        # 狀態標籤
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #27ae60; font-weight: bold; font-size: 11px; margin: 5px;")
        self.status_label.setMaximumHeight(25)
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        
        return panel
    
    def create_right_panel(self) -> QWidget:
        """建立右側預覽面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(5)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # 預覽標題
        preview_title = QLabel("📋 投影片預覽 (核心引擎)")
        preview_title.setFont(QFont("Microsoft JhengHei", 16, QFont.Weight.Bold))
        preview_title.setStyleSheet("color: #2c3e50; margin: 10px;")
        layout.addWidget(preview_title)
        
        # 預覽說明
        preview_desc = QLabel("🎯 核心渲染引擎 | 📐 智慧分頁檢測 | 🖼️ 1080p 高品質預覽 | ✨ 漸層美化")
        preview_desc.setFont(QFont("Microsoft JhengHei", 10))
        preview_desc.setStyleSheet("""
            color: #27ae60; 
            background: #f0f8f0; 
            padding: 8px; 
            border-radius: 5px; 
            border: 1px solid #27ae60;
            margin: 5px 10px;
        """)
        preview_desc.setWordWrap(True)
        layout.addWidget(preview_desc)
        
        # 預覽區域
        self.preview_widget = PreviewWidget()
        layout.addWidget(self.preview_widget)
        
        return panel
    
    def setup_connections(self):
        """設定信號連接"""
        self.word_drop.file_dropped.connect(self.on_word_file_selected)
        self.template_drop.file_dropped.connect(self.on_template_file_selected)
    
    def load_last_used_paths(self):
        """載入上次使用的檔案路徑"""
        try:
            last_word_path = self.config_manager.get_last_word_path()
            last_template_path = self.config_manager.get_last_template_path()
            
            if last_word_path and os.path.exists(last_word_path):
                self.on_word_file_selected(last_word_path)
                print(f"自動載入上次的Word檔案: {last_word_path}")
            
            if last_template_path and os.path.exists(last_template_path):
                self.on_template_file_selected(last_template_path)
                print(f"自動載入上次的範本檔案: {last_template_path}")
                
        except Exception as e:
            print(f"載入上次使用路徑時發生錯誤: {e}")
    
    def on_word_file_selected(self, file_path: str):
        """Word檔案選擇處理"""
        if file_path.lower().endswith(('.docx', '.doc')):
            self.word_path = file_path
            filename = os.path.basename(file_path)
            self.word_status.setText(f"✅ {filename}")
            self.word_status.setStyleSheet("color: #27ae60; margin: 5px; font-weight: bold;")
            
            self.config_manager.set_last_word_path(file_path)
            self.auto_set_output_path(file_path)
        else:
            QMessageBox.warning(self, "檔案格式錯誤", "請選擇Word文件(.docx或.doc)")
            return
            
        self.check_ready_to_convert()
    
    def on_template_file_selected(self, file_path: str):
        """範本檔案選擇處理"""
        if file_path.lower().endswith(('.pptx', '.ppt')):
            self.template_path = file_path
            filename = os.path.basename(file_path)
            self.template_status.setText(f"✅ {filename}")
            self.template_status.setStyleSheet("color: #27ae60; margin: 5px; font-weight: bold;")
            
            self.config_manager.set_last_template_path(file_path)
        else:
            QMessageBox.warning(self, "檔案格式錯誤", "請選擇PowerPoint檔案(.pptx或.ppt)")
            return
            
        self.check_ready_to_convert()
    
    def auto_set_output_path(self, word_path: str):
        """根據Word檔案路徑自動設定輸出路徑"""
        try:
            file_dir = os.path.dirname(word_path)
            file_name_without_ext = os.path.splitext(os.path.basename(word_path))[0]
            
            output_path = os.path.join(file_dir, f"{file_name_without_ext}.pptx")
            
            counter = 1
            original_output_path = output_path
            while os.path.exists(output_path):
                base_name = os.path.splitext(original_output_path)[0]
                output_path = f"{base_name}_{counter}.pptx"
                counter += 1
            
            self.output_path = output_path
            self.output_label.setText(f"📁 {os.path.basename(output_path)}")
            self.output_label.setStyleSheet("border: 1px solid #27ae60; padding: 8px; background: #f8fff8; color: #27ae60;")
            
            self.config_manager.set_last_output_dir(file_dir)
            
        except Exception as e:
            print(f"自動設定輸出路徑時發生錯誤: {e}")
    
    def select_output_path(self):
        """選擇輸出路徑"""
        default_dir = self.config_manager.get_last_output_dir()
        default_name = "presentation.pptx"
        
        if self.word_path:
            default_dir = os.path.dirname(self.word_path)
            word_name = os.path.splitext(os.path.basename(self.word_path))[0]
            default_name = f"{word_name}.pptx"
        
        default_path = os.path.join(default_dir, default_name) if default_dir else default_name
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "儲存PowerPoint檔案", default_path, "PowerPoint Files (*.pptx)"
        )
        
        if file_path:
            if not file_path.endswith('.pptx'):
                file_path += '.pptx'
            self.output_path = file_path
            self.output_label.setText(f"📁 {os.path.basename(file_path)}")
            self.output_label.setStyleSheet("border: 1px solid #27ae60; padding: 8px; background: #f8fff8; color: #27ae60;")
            
            self.config_manager.set_last_output_dir(os.path.dirname(file_path))
            
            self.check_ready_to_convert()
    
    def clear_word_file(self):
        """清除Word檔案"""
        self.word_path = ""
        self.word_status.setText("未選擇Word檔案")
        self.word_status.setStyleSheet("color: #e74c3c; margin: 5px; font-size: 10px;")
        self.check_ready_to_convert()
        QMessageBox.information(self, "清除完成", "Word檔案已清除")
    
    def clear_template_file(self):
        """清除範本檔案"""
        self.template_path = ""
        self.template_status.setText("未選擇PowerPoint範本")
        self.template_status.setStyleSheet("color: #e74c3c; margin: 5px; font-size: 10px;")
        self.check_ready_to_convert()
        QMessageBox.information(self, "清除完成", "範本檔案已清除")
    
    def clear_all_files(self):
        """清除所有檔案"""
        reply = QMessageBox.question(
            self, "確認清除", 
            "確定要清除所有檔案路徑嗎？\n這將清除Word文件、範本文件和輸出路徑。",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.word_path = ""
            self.template_path = ""
            self.output_path = ""
            
            self.word_status.setText("未選擇Word檔案")
            self.word_status.setStyleSheet("color: #e74c3c; margin: 5px; font-size: 10px;")
            
            self.template_status.setText("未選擇PowerPoint範本")
            self.template_status.setStyleSheet("color: #e74c3c; margin: 5px; font-size: 10px;")
            
            self.output_label.setText("將自動設定輸出位置...")
            self.output_label.setStyleSheet("""
                QLabel {
                    border: 1px solid #ccc; 
                    padding: 8px; 
                    background: white; 
                    color: #666;
                    border-radius: 4px;
                    font-size: 10px;
                }
            """)
            
            # 清除預覽
            self.preview_widget.clear_preview()
            
            self.check_ready_to_convert()
            QMessageBox.information(self, "清除完成", "所有檔案路徑已清除")
    
    def on_preview_save_changed(self, state):
        """預覽圖片保存設定變更"""
        self.save_preview_to_disk = state == 2  # 2 表示選中
        
        if self.save_preview_to_disk:
            QMessageBox.information(
                self, "設定變更", 
                "✅ 已啟用預覽圖片保存到本地磁碟\n\n預覽圖片將永久保存在輸出檔案同目錄的資料夾中。"
            )
        else:
            QMessageBox.information(
                self, "設定變更", 
                "❌ 已禁用預覽圖片保存到本地磁碟\n\n預覽圖片僅在記憶體中生成，關閉程式後不會保留。\n注意：預覽功能仍然正常運作。"
            )
    
    def check_ready_to_convert(self):
        """檢查是否準備好轉換"""
        word_ready = bool(self.word_path and self.word_path.strip())
        template_ready = bool(self.template_path and self.template_path.strip())
        output_ready = bool(self.output_path and self.output_path.strip())
        
        ready = word_ready and template_ready and output_ready
        
        self.convert_btn.setEnabled(ready)
        
        if ready:
            self.convert_btn.setText("🚀 開始轉換為PowerPoint")
            self.convert_btn.setStyleSheet("""
                QPushButton {
                    background-color: #27ae60;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 5px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #229954;
                }
            """)
        else:
            missing_items = []
            if not word_ready:
                missing_items.append("Word文件")
            if not template_ready:
                missing_items.append("PowerPoint範本")
            if not output_ready:
                missing_items.append("輸出路徑")
            
            self.convert_btn.setText(f"請選擇: {', '.join(missing_items)}")
            self.convert_btn.setStyleSheet("""
                QPushButton {
                    background-color: #bdc3c7;
                    color: #7f8c8d;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 5px;
                    font-size: 14px;
                    font-weight: bold;
                }
            """)
    
    def start_conversion(self):
        """開始轉換"""
        if not all([self.word_path, self.template_path, self.output_path]):
            QMessageBox.warning(
                self, "準備不完整", 
                "請確保已選擇Word文件、PowerPoint範本和輸出路徑"
            )
            return
        
        if not os.path.exists(self.word_path):
            QMessageBox.critical(self, "檔案不存在", f"Word文件不存在：\n{self.word_path}")
            return
            
        if not os.path.exists(self.template_path):
            QMessageBox.critical(self, "檔案不存在", f"PowerPoint範本不存在：\n{self.template_path}")
            return
        
        self.convert_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        status_msg = f"正在準備轉換...\n預覽圖片保存: {'啟用' if self.save_preview_to_disk else '禁用'}"
        self.status_label.setText(status_msg)
        
        self.worker = ConversionWorker(
            self.word_path, 
            self.template_path, 
            self.output_path,
            self.save_preview_to_disk
        )
        
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.status_updated.connect(self.status_label.setText)
        self.worker.finished_successfully.connect(self.on_conversion_finished)
        self.worker.error_occurred.connect(self.on_conversion_error)
        
        self.worker.start()
    
    def on_conversion_finished(self, result: ConversionResult):
        """轉換完成處理"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("✅ 轉換完成！")
        self.status_label.setStyleSheet("color: #27ae60; font-weight: bold;")
        
        self.check_ready_to_convert()
        
        # 更新預覽（總是顯示預覽，不管是否保存到磁碟）
        if result.preview_images:
            self.preview_widget.update_preview(result.preview_images)
        
        # 顯示完成訊息
        output_name = os.path.splitext(os.path.basename(result.output_path))[0]
        preview_folder = f"{output_name}_預覽圖片"
        
        # 根據預覽圖片保存設定顯示不同訊息
        if self.save_preview_to_disk and result.preview_images:
            preview_status = f"📁 預覽圖片資料夾: {preview_folder}\n💾 預覽圖片已永久保存到本地磁碟"
        elif result.preview_images:
            preview_status = "🖼️ 預覽圖片已生成（僅記憶體中，未保存到磁碟）"
        else:
            preview_status = "⚠️ 預覽圖片生成失敗"
        
        message = f"""🎉 PowerPoint轉換完成！

📊 已建立檔案: {os.path.basename(result.output_path)}
{preview_status}
📊 投影片數量: {result.slides_count}

✨ v5.0 增強功能:
• 智慧章節識別與自動分頁
• 完全清除範本原始內容
• 文字溢出檢測與處理
• 1080p 高解析度預覽圖片
• 漸層背景與美化效果
• 跨平台字體優化
• 可選的預覽圖片保存功能

右側顯示投影片預覽。
{'預覽圖片已永久保存在輸出檔案同目錄下的資料夾中。' if self.save_preview_to_disk else '預覽圖片僅在記憶體中，關閉程式後不會保留。'}

是否立即開啟PowerPoint檔案？"""

        reply = QMessageBox.question(
            self, "轉換完成", message,
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.open_file(result.output_path)
    
    def open_file(self, file_path: str):
        """跨平台開啟檔案"""
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":
                subprocess.run(["open", file_path])
            else:
                subprocess.run(["xdg-open", file_path])
        except Exception as e:
            QMessageBox.information(
                self, "提示", 
                f"檔案已儲存到：{file_path}\n請手動開啟檔案。"
            )
    
    def on_conversion_error(self, error_message: str):
        """轉換錯誤處理"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("❌ 轉換失敗")
        self.status_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
        
        self.check_ready_to_convert()
        
        detailed_message = f"轉換過程中發生錯誤：\n{error_message}\n\n"
        detailed_message += f"系統資訊：{platform.system()}\n\n"
        detailed_message += "可能的解決方法：\n"
        detailed_message += "1. 確保Word文件不是受保護的\n"
        detailed_message += "2. 確保PowerPoint範本檔案完整\n"
        detailed_message += "3. 檢查檔案路徑中是否包含特殊字元\n"
        detailed_message += "4. 嘗試關閉正在使用這些檔案的其他程式\n"
        detailed_message += "5. 確認Word文件中包含可識別的章節標題\n"
        
        QMessageBox.critical(self, "轉換錯誤", detailed_message)

def main():
    """主函式"""
    if not PYSIDE6_AVAILABLE:
        print("❌ PySide6 不可用，請安裝: pip install PySide6")
        return
        
    if not CORE_AVAILABLE:
        print("❌ 核心模組不可用，請確保 word_to_pptx_core.py 在同一目錄下")
        return
    
    app = QApplication(sys.argv)
    
    app.setApplicationName("Word轉PowerPoint工具")
    app.setApplicationVersion("5.0")
    app.setOrganizationName("智慧辦公工具")
    
    # 設定高DPI支援
    app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    print("=" * 80)
    print("🚀 Word轉PowerPoint工具 v5.0 (桌面版) 🚀")
    print("=" * 80)
    print(f"🖥️  運行系統: {platform.system()}")
    print("📱 UI框架: PySide6")
    print("🔧 核心引擎: word_to_pptx_core")
    print("")
    
    # 檢查相依性
    if CORE_AVAILABLE:
        dep_status = get_dependency_status()
        print("🛠️  相依性檢查:")
        for pkg, status in dep_status.items():
            icon = "✅" if status else "❌"
            print(f"  {icon} {pkg}")
        
        if not all(dep_status.values()):
            missing = [pkg for pkg, status in dep_status.items() if not status]
            print(f"\n❌ 缺少套件: {', '.join(missing)}")
            print(f"請執行: pip install {' '.join(missing)}")
            sys.exit(1)
    
    print("\n🚀 啟動桌面應用程式...")
    print("=" * 80)
    
    try:
        main()
    except Exception as e:
        print(f"❌ 程式執行錯誤: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)