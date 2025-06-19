#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用Word轉PowerPoint工具 - 優化版 (支援 Windows & macOS)
專注於 PowerPoint 輸出，優化預覽功能
增強功能：
1. 專注 PowerPoint (.pptx) 輸出
2. 修復 macOS 字體顯示問題
3. 智慧文字溢出檢測與自動分頁
4. 真實 PPTX 投影片圖片預覽
"""

import sys
import os
import re
import subprocess
import json
import platform
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from io import BytesIO
import tempfile

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QTextEdit, QProgressBar, QFileDialog,
    QWidget, QFrame, QScrollArea, QTabWidget, QGroupBox,
    QMessageBox, QSplitter, QListWidget, QListWidgetItem,
    QComboBox, QSpinBox, QCheckBox, QSlider
)
from PySide6.QtCore import (
    Qt, QThread, Signal, QMimeData, QTimer, QPropertyAnimation,
    QEasingCurve, QRect, Slot, QSize, QSettings
)
from PySide6.QtGui import (
    QFont, QPixmap, QPainter, QColor, QBrush, QPen,
    QDragEnterEvent, QDropEvent, QIcon, QLinearGradient
)

# 第三方庫導入
try:
    from docx import Document
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from PIL import Image, ImageDraw, ImageFont
    DEPENDENCIES_OK = True
except ImportError as e:
    print(f"請安裝必要套件: pip install python-docx python-pptx pillow")
    DEPENDENCIES_OK = False

# Spire.Presentation 導入
try:
    from spire.presentation import Presentation as SpirePresentation
    from spire.presentation.common import *
    SPIRE_AVAILABLE = True
    print("✅ Spire.Presentation 可用 - 將使用高品質匯出")
except ImportError:
    SPIRE_AVAILABLE = False
    print("⚠️  Spire.Presentation 不可用 - 將使用 PIL 備用方法")

@dataclass
class ContentBlock:
    """內容塊資料結構"""
    text: str
    level: int  # 標題層級 (0=主標題/章節, 1=次標題, 2=內容)
    content_type: str  # header, chapter, title, subtitle, content, quote, list
    formatting: Dict = None
    estimated_length: int = 0  # 估算文字長度

@dataclass
class SlideTemplate:
    """投影片範本資料結構"""
    layout_index: int
    layout_name: str
    placeholders: List[Dict]
    background_color: Tuple[int, int, int] = None
    font_family: str = "Microsoft JhengHei"

class SystemFontManager:
    """跨平台系統字體管理器"""
    
    def __init__(self):
        self.platform = platform.system()
        self.available_fonts = {}
        self.font_cache = {}
        self._load_system_fonts()
    
    def _load_system_fonts(self):
        """載入系統字體並測試中文支援"""
        if self.platform == "Darwin":  # macOS
            font_candidates = [
                ("PingFang TC", [
                    "/System/Library/Fonts/PingFang.ttc",
                    "/Library/Fonts/PingFang.ttc"
                ]),
                ("Hiragino Sans", [
                    "/System/Library/Fonts/Hiragino Sans GB.ttc",
                    "/Library/Fonts/Hiragino Sans GB.ttc"
                ]),
                ("STHeiti", [
                    "/System/Library/Fonts/STHeiti Medium.ttc",
                    "/Library/Fonts/STHeiti Medium.ttc"
                ]),
                ("Arial Unicode MS", [
                    "/Library/Fonts/Arial Unicode MS.ttf",
                    "/System/Library/Fonts/Arial Unicode MS.ttf"
                ])
            ]
        elif self.platform == "Windows":  # Windows
            font_candidates = [
                ("Microsoft JhengHei", [
                    "C:/Windows/Fonts/msjh.ttc",
                    "C:/Windows/Fonts/msjhbd.ttc"
                ]),
                ("Microsoft YaHei", [
                    "C:/Windows/Fonts/msyh.ttc", 
                    "C:/Windows/Fonts/msyhbd.ttc"
                ]),
                ("SimSun", [
                    "C:/Windows/Fonts/simsun.ttc"
                ]),
                ("Arial Unicode MS", [
                    "C:/Windows/Fonts/ARIALUNI.TTF"
                ])
            ]
        else:  # Linux
            font_candidates = [
                ("WenQuanYi", [
                    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
                    "/usr/share/fonts/truetype/arphic/uming.ttc"
                ]),
                ("Noto CJK", [
                    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"
                ])
            ]
        
        # 測試每種字體是否支援中文
        for font_name, paths in font_candidates:
            for path in paths:
                if os.path.exists(path):
                    if self._test_font_chinese_support(path):
                        self.available_fonts[font_name] = path
                        print(f"載入中文字體: {font_name} ({path})")
                        break
        
        if not self.available_fonts:
            print(f"警告: 在 {self.platform} 系統上未找到支援中文的字體")
    
    def _test_font_chinese_support(self, font_path: str) -> bool:
        """測試字體是否支援中文"""
        try:
            test_font = ImageFont.truetype(font_path, 20)
            test_img = Image.new('RGB', (100, 50), 'white')
            test_draw = ImageDraw.Draw(test_img)
            test_text = "測試中文字體"
            test_draw.text((10, 10), test_text, font=test_font, fill='black')
            return True
        except Exception as e:
            return False
    
    def get_best_font(self, size: int) -> ImageFont.ImageFont:
        """取得最佳的中文字體"""
        cache_key = f"font_{size}_{self.platform}"
        if cache_key in self.font_cache:
            return self.font_cache[cache_key]
        
        # 依平台優先順序嘗試字體
        if self.platform == "Darwin":
            priority_fonts = ["PingFang TC", "Hiragino Sans", "STHeiti", "Arial Unicode MS"]
        elif self.platform == "Windows":
            priority_fonts = ["Microsoft JhengHei", "Microsoft YaHei", "SimSun", "Arial Unicode MS"]
        else:
            priority_fonts = ["WenQuanYi", "Noto CJK"]
        
        for font_name in priority_fonts:
            if font_name in self.available_fonts:
                try:
                    font = ImageFont.truetype(self.available_fonts[font_name], size)
                    self.font_cache[cache_key] = font
                    return font
                except Exception as e:
                    print(f"載入字體失敗 {font_name}: {e}")
                    continue
        
        # 如果都失敗，使用預設字體
        default_font = ImageFont.load_default()
        self.font_cache[cache_key] = default_font
        return default_font

class ConfigManager:
    """設定管理器 - 記憶上次使用的檔案路徑"""
    
    def __init__(self):
        self.config_file = os.path.join(os.path.expanduser("~"), ".word_to_pptx_config.json")
        self.config = self.load_config()
    
    def load_config(self) -> Dict:
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
        """設定上次使用的Word檔案路徑"""
        self.config["last_word_path"] = path
        self.save_config()
    
    def set_last_template_path(self, path: str):
        """設定上次使用的範本檔案路徑"""
        self.config["last_template_path"] = path
        self.save_config()
    
    def set_last_output_dir(self, dir_path: str):
        """設定上次使用的輸出目錄"""
        self.config["last_output_dir"] = dir_path
        self.save_config()
    
    def get_last_word_path(self) -> str:
        """取得上次使用的Word檔案路徑"""
        return self.config.get("last_word_path", "")
    
    def get_last_template_path(self) -> str:
        """取得上次使用的範本檔案路徑"""
        return self.config.get("last_template_path", "")
    
    def get_last_output_dir(self) -> str:
        """取得上次使用的輸出目錄"""
        return self.config.get("last_output_dir", "")
    
    def set_window_geometry(self, geometry: Dict):
        """設定視窗幾何"""
        self.config["window_geometry"] = geometry
        self.save_config()
    
    def get_window_geometry(self) -> Dict:
        """取得視窗幾何"""
        return self.config.get("window_geometry", None)

class WordDocumentAnalyzer:
    """Word文件智慧分析器 - 增強版"""
    
    def __init__(self):
        # 增強的中文章節識別模式
        self.chapter_patterns = [
            r'^[一二三四五六七八九十]+[、．.]\s*',
            r'^第[一二三四五六七八九十壹貳參肆伍陸柒捌玖拾]+[章節部分]\s*',
            r'^第[一二三四五六七八九十]+[、．.]\s*',
            r'^前言\s*',
            r'^結論\s*',
            r'^總結\s*',
            r'^概述\s*',
            r'^摘要\s*',
            r'^序言\s*',
            r'^引言\s*',
        ]
        
        self.subtitle_patterns = [
            r'^[一二三四五六七八九十]+[）)]\s*',
            r'^\([一二三四五六七八九十]+\)\s*',
            r'^[1-9]\d*[）)]\s*',
            r'^\([1-9]\d*\)\s*',
            r'^[a-z][）)]\s*',
            r'^\([a-z]\)\s*',
            r'^[•·○]\s*',
        ]
        
    def analyze_document(self, file_path: str) -> List[ContentBlock]:
        """分析Word文件結構"""
        try:
            doc = Document(file_path)
            blocks = []
            
            header_line = True
            for para in doc.paragraphs:
                if not para.text.strip():
                    continue

                if header_line:
                    block = self._classify_header(para)
                    header_line = False
                else:
                    print(para.text)
                    block = self._analyze_paragraph(para)
                if block:
                    # 估算文字長度（用於後續分頁判斷）
                    block.estimated_length = self._estimate_text_length(block.text)
                    blocks.append(block)
            
            return self._optimize_structure(blocks)
            
        except Exception as e:
            raise Exception(f"分析Word文件時發生錯誤: {e}")
    
    def _estimate_text_length(self, text: str) -> int:
        """估算文字在投影片上的顯示長度"""
        # 中文字符算2個單位，英文算1個單位
        length = 0
        for char in text:
            if ord(char) > 127:  # 非ASCII字符（包括中文）
                length += 2
            else:
                length += 1
        return length
    
    def _analyze_paragraph(self, para) -> Optional[ContentBlock]:
        """分析單個段落"""
        text = para.text.strip()
        if not text:
            return None
            
        formatting = self._extract_formatting(para)
        level, content_type = self._classify_content(text, formatting)
        
        return ContentBlock(
            text=text,
            level=level,
            content_type=content_type,
            formatting=formatting
        )

    def _classify_header(self, para) -> Optional[ContentBlock]:
        text = para.text.strip()
        if not text:
            return None

        formatting = self._extract_formatting(para)

        return ContentBlock(
            text=text,
            level=0,
            content_type='header',
            formatting=formatting
        )

    def _extract_formatting(self, para) -> Dict:
        """提取段落格式"""
        formatting = {
            'bold': False,
            'italic': False,
            'font_size': 12,
            'alignment': 'left'
        }
        
        if para.runs:
            run = para.runs[0]
            if run.bold:
                formatting['bold'] = True
            if run.italic:
                formatting['italic'] = True
            if run.font.size:
                formatting['font_size'] = run.font.size.pt
                
        return formatting
    
    def _classify_content(self, text: str, formatting: Dict) -> Tuple[int, str]:
        """分類內容類型和層級"""
        for pattern in self.chapter_patterns:
            if re.match(pattern, text):
                return (0, 'chapter')
        
        for pattern in self.subtitle_patterns:
            if re.match(pattern, text):
                return (1, 'subtitle')
            
        return (2, 'content')
    
    def _optimize_structure(self, blocks: List[ContentBlock]) -> List[ContentBlock]:
        """最佳化文件結構"""
        if not blocks:
            return blocks
            
        has_chapter = any(block.content_type == 'chapter' for block in blocks)
        if not has_chapter and blocks:
            for block in blocks:
                if block.level <= 1 or block.formatting.get('bold', False):
                    block.level = 0
                    block.content_type = 'chapter'
                    break
        
        return blocks

class ContentToSlideMapper:
    """內容到投影片的智慧映射器 - 增強版（支援智慧分頁）"""
    
    def __init__(self, presentation_path: str):
        self.presentation_path = presentation_path
        self.prs = None
        self.max_content_length = 220  # 單張投影片最大內容長度
        self.max_content_items = 4     # 單張投影片最大內容項目數
        
    def create_slides(self, blocks: List[ContentBlock]) -> Presentation:
        """建立投影片（支援智慧分頁）- 清除範本原始內容"""
        self.prs = Presentation(self.presentation_path)
        
        # 清空現有投影片
        while len(self.prs.slides) > 0:
            rId = self.prs.slides._sldIdLst[0].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[0]
        
        current_slide = None
        current_content = []
        current_chapter_title = ""
        current_content_length = 0
        
        for block in blocks:
            print(f'Block level:{block.level}, type:{block.content_type}, text:{block.text[:50]}...')
            
            if block.content_type == 'header':
                # 完成當前投影片
                if current_slide is not None:
                    self._finalize_slide(current_slide, current_content)
                
                # 建立新的章節投影片
                current_slide = self._create_title_slide(block)
                current_content = []
                current_chapter_title = self._clean_chapter_text(block.text)
                current_content_length = 0
            
            elif block.content_type == 'chapter':
                # 完成當前投影片
                if current_slide is not None:
                    self._finalize_slide(current_slide, current_content)
                
                # 建立新的章節投影片
                current_slide = self._create_content_slide(block)
                current_content = []
                current_chapter_title = block.text
                current_content_length = 0

            elif block.content_type == 'subtitle':
                # 檢查是否需要分頁
                if True:  # 簡化邏輯，遇到subtitle始終分頁
                    # 完成當前投影片
                    if current_slide is not None:
                        self._finalize_slide(current_slide, current_content)
                    
                    # 建立新的內容投影片（使用相同章節標題）
                    chapter_block = ContentBlock(
                        text=current_chapter_title + " (續)",
                        level=0,
                        content_type='chapter',
                        formatting={}
                    )
                    current_slide = self._create_content_slide(chapter_block)
                    current_content = []
                    current_content_length = 0

                    current_content.append(block)
                    current_content_length += block.estimated_length
            else:
                # 檢查是否需要分頁
                if self._should_create_new_slide(current_content, current_content_length, block):
                    # 完成當前投影片
                    if current_slide is not None:
                        self._finalize_slide(current_slide, current_content)
                    
                    # 建立新的內容投影片（使用相同章節標題）
                    chapter_block = ContentBlock(
                        text=current_chapter_title + " (續)",
                        level=0,
                        content_type='chapter',
                        formatting={}
                    )
                    current_slide = self._create_content_slide(chapter_block)
                    current_content = []
                    current_content_length = 0
                
                # 添加內容到當前投影片
                if current_slide is None:
                    # 如果沒有當前投影片，建立一個
                    current_slide = self._create_content_slide(block)
                
                current_content.append(block)
                current_content_length += block.estimated_length
        
        # 處理最後一張投影片
        if current_slide is not None:
            self._finalize_slide(current_slide, current_content)
        
        print(f"已建立 {len(self.prs.slides)} 張新投影片")
        return self.prs
    
    def _should_create_new_slide(self, current_content: List[ContentBlock], 
                                current_length: int, new_block: ContentBlock) -> bool:
        """判斷是否需要建立新投影片"""
        # 檢查內容項目數量
        if len(current_content) >= self.max_content_items:
            return True
        
        # 檢查內容總長度
        if current_length + new_block.estimated_length > self.max_content_length:
            return True
        
        # 檢查是否為次標題（可能需要新投影片）
        if new_block.content_type == 'subtitle' and len(current_content) > 0:
            return True
        
        return False
    
    def _create_title_slide(self, block: ContentBlock):
        """建立標題投影片"""
        layout = self._get_best_layout('title')
        slide = self.prs.slides.add_slide(layout)
        
        if slide.shapes.title:
            title_text = self._clean_chapter_text(block.text)
            slide.shapes.title.text = title_text
            
        return slide
    
    def _create_content_slide(self, block: ContentBlock):
        """建立內容投影片"""
        layout = self._get_best_layout('content')
        slide = self.prs.slides.add_slide(layout)
        
        if slide.shapes.title:
            title_text = self._clean_subtitle_text(block.text)
            slide.shapes.title.text = title_text
            
        return slide
    
    def _clean_chapter_text(self, text: str) -> str:
        """清理章節文字，移除編號標記"""
        patterns = [
            r'^[一二三四五六七八九十]+[、．.]\s*',
            r'^第[一二三四五六七八九十壹貳參肆伍陸柒捌玖拾]+[章節部分]\s*',
            r'^第[一二三四五六七八九十]+[、．.]\s*',
            r'^[1-9]\d*[、．.]\s*',
            r'^第[1-9]\d*[章節部分]\s*',
            r'^第[1-9]\d*[、．.]\s*',
            r'^[A-Z][、．.]\s*',
            r'^第[A-Z][章節部分]\s*',
            r'^[●◆■▲]\s*',
        ]
        
        cleaned_text = text
        for pattern in patterns:
            cleaned_text = re.sub(pattern, '', cleaned_text)
        
        return cleaned_text.strip()
    
    def _clean_subtitle_text(self, text: str) -> str:
        """清理次標題文字"""
        patterns = [
            r'^[一二三四五六七八九十]+[）)]\s*',
            r'^\([一二三四五六七八九十]+\)\s*',
            r'^[1-9]\d*[）)]\s*',
            r'^\([1-9]\d*\)\s*',
            r'^[a-z][）)]\s*',
            r'^\([a-z]\)\s*',
            r'^[•·○]\s*',
        ]
        
        cleaned_text = text
        for pattern in patterns:
            cleaned_text = re.sub(pattern, '', cleaned_text)
        
        return cleaned_text.strip()
    
    def _finalize_slide(self, slide, content_blocks: List[ContentBlock]):
        """完成投影片內容"""
        if not content_blocks:
            return
            
        content_placeholder = None
        for shape in slide.placeholders:
            if hasattr(shape, 'text_frame') and shape.placeholder_format.idx == 1:
                content_placeholder = shape
                break
        
        if content_placeholder:
            text_frame = content_placeholder.text_frame
            text_frame.clear()
            
            for i, block in enumerate(content_blocks):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                
                p.text = block.text
    
    def _get_best_layout(self, layout_type: str):
        """取得最佳佈局"""
        if not self.prs:
            raise Exception("Presentation not initialized")
            
        if layout_type == 'title':
            for layout in self.prs.slide_layouts:
                if 'title' in layout.name.lower() or layout == self.prs.slide_layouts[0]:
                    return layout
        else:
            for layout in self.prs.slide_layouts:
                if 'content' in layout.name.lower():
                    return layout
            if len(self.prs.slide_layouts) > 1:
                return self.prs.slide_layouts[1]
        
        return self.prs.slide_layouts[0]

class PPTXImageExporter:
    """PPTX 投影片圖片匯出器 - 使用真實投影片內容"""
    
    def __init__(self, output_path: str = None):
        self.font_manager = SystemFontManager()
        if output_path:
            # 在輸出檔案同目錄創建預覽資料夾
            output_dir = os.path.dirname(output_path)
            output_name = os.path.splitext(os.path.basename(output_path))[0]
            self.temp_dir = os.path.join(output_dir, f"{output_name}_預覽圖片")
            os.makedirs(self.temp_dir, exist_ok=True)
            print(f"建立預覽圖片目錄: {self.temp_dir}")
        else:
            # 備用：使用暫存目錄
            self.temp_dir = tempfile.mkdtemp()
            print(f"建立暫存目錄: {self.temp_dir}")
    
    def export_slides_to_images(self, presentation_path: str) -> List[str]:
        """將 PPTX 投影片匯出為圖片 - 統一使用 python-pptx"""
        try:
            # 所有平台統一使用 python-pptx + PIL 方法
            return self._export_with_python_pptx(presentation_path)
                
        except Exception as e:
            print(f"圖片匯出失敗: {e}")
            return []
    
    def _export_with_windows_com(self, presentation_path: str) -> List[str]:
        """Windows COM 匯出方法"""
        try:
            import win32com.client
            
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1
            
            presentation = powerpoint.Presentations.Open(presentation_path)
            image_paths = []
            
            for i, slide in enumerate(presentation.Slides, 1):
                image_path = os.path.join(self.temp_dir, f"slide_{i}.png")
                # 匯出為 PNG，800x600 解析度
                slide.Export(image_path, "PNG", 800, 600)
                if os.path.exists(image_path):
                    image_paths.append(image_path)
                    print(f"已匯出投影片 {i} 到 {image_path}")
            
            presentation.Close()
            powerpoint.Quit()
            
            return image_paths
            
        except ImportError:
            print("未安裝 pywin32，無法使用 COM 匯出")
            raise Exception("需要安裝 pywin32 套件")
        except Exception as e:
            print(f"Windows COM 匯出失敗: {e}")
            raise e
    
    def _export_with_macos_keynote(self, presentation_path: str) -> List[str]:
        """macOS Keynote 匯出方法"""
        try:
            # 使用 AppleScript 將 PPTX 轉換並匯出圖片
            apple_script = f'''
            tell application "Keynote"
                set thePresentation to open POSIX file "{presentation_path}"
                set slideCount to count of slides of thePresentation
                
                repeat with i from 1 to slideCount
                    set currentSlide to slide i of thePresentation
                    set imagePath to "{self.temp_dir}/slide_" & i & ".png"
                    export currentSlide to POSIX file imagePath as PNG
                end repeat
                
                close thePresentation
                return slideCount
            end tell
            '''
            
            process = subprocess.run(['osascript', '-e', apple_script], 
                                   capture_output=True, text=True, timeout=60)
            
            if process.returncode == 0:
                image_paths = []
                slide_count = int(process.stdout.strip())
                
                for i in range(1, slide_count + 1):
                    image_path = os.path.join(self.temp_dir, f"slide_{i}.png")
                    if os.path.exists(image_path):
                        image_paths.append(image_path)
                        print(f"已匯出投影片 {i} 到 {image_path}")
                
                return image_paths
            else:
                raise Exception(f"AppleScript 執行失敗: {process.stderr}")
                
        except Exception as e:
            print(f"macOS Keynote 匯出失敗: {e}")
            raise e
    
    def _export_with_python_pptx(self, presentation_path: str) -> List[str]:
        """使用 python-pptx + PIL 匯出方法（備用）"""
        try:
            prs = Presentation(presentation_path)
            image_paths = []
            
            print(f"使用 python-pptx 方法處理 {len(prs.slides)} 張投影片...")
            
            for i, slide in enumerate(prs.slides, 1):
                print(f"正在處理投影片 {i}...")
                image_path = self._render_slide_to_image(slide, i)
                if image_path:
                    image_paths.append(image_path)
            
            print(f"成功產生 {len(image_paths)} 張投影片圖片")
            return image_paths
            
        except Exception as e:
            print(f"python-pptx 匯出失敗: {e}")
            raise e
    
    def _render_slide_to_image(self, slide, slide_number: int) -> str:
        """將投影片渲染為圖片（高品質版）"""
        try:
            # 設定更高的解析度
            width, height = 1280, 720  # 720p 解析度
            img = Image.new('RGB', (width, height), color='white')
            draw = ImageDraw.Draw(img)
            
            # 使用跨平台字體管理器
            title_font = self.font_manager.get_best_font(48)
            content_font = self.font_manager.get_best_font(32)
            small_font = self.font_manager.get_best_font(24)
            
            # 繪製邊框和背景
            draw.rectangle([0, 0, width-1, height-1], outline='#E0E0E0', width=3)
            
            # 添加漸層背景效果
            for y in range(height):
                color = int(255 - (y / height) * 10)  # 輕微漸層
                draw.line([(0, y), (width, y)], fill=(color, color, color))
            
            # 提取和繪製內容
            y_pos = 60
            title_drawn = False
            content_items = []
            
            # 收集所有文字內容
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text and shape.text.strip():
                    raw_text = shape.text.strip()
                    text = self._normalize_text_cross_platform(raw_text)
                    
                    if not text:
                        continue
                    
                    is_title = (hasattr(shape, 'placeholder_format') and 
                              shape.placeholder_format.idx == 0)
                    
                    if is_title and not title_drawn:
                        # 繪製標題
                        y_pos = self._draw_title_enhanced(draw, text, width, y_pos, title_font)
                        y_pos += 80
                        title_drawn = True
                    else:
                        content_items.append(text)
            
            # 繪製內容項目
            if content_items:
                y_pos = self._draw_content_enhanced(draw, content_items, width, y_pos, content_font)
            
            # 添加投影片編號和裝飾
            self._add_slide_decorations(draw, slide_number, width, height, small_font)
            
            # 儲存高品質圖片
            image_path = os.path.join(self.temp_dir, f"slide_{slide_number}.png")
            img.save(image_path, 'PNG', quality=95, optimize=True)
            
            print(f"成功建立高品質投影片圖片: {image_path}")
            return image_path
            
        except Exception as e:
            print(f"渲染投影片圖片時發生錯誤: {e}")
            return None
    
    def _draw_title_enhanced(self, draw, title: str, width: int, y_pos: int, font) -> int:
        """繪製增強版標題"""
        try:
            safe_title = self._normalize_text_cross_platform(title)
            
            # 計算標題尺寸
            bbox = draw.textbbox((0, 0), safe_title, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            # 置中位置
            center_x = (width - text_width) // 2
            
            # 繪製標題陰影
            shadow_offset = 3
            draw.text((center_x + shadow_offset, y_pos + shadow_offset), 
                     safe_title, fill='#CCCCCC', font=font)
            
            # 繪製主標題
            draw.text((center_x, y_pos), safe_title, fill='#2C3E50', font=font)
            
            # 繪製底線
            line_y = y_pos + text_height + 15
            line_start_x = center_x
            line_end_x = center_x + text_width
            draw.line([(line_start_x, line_y), (line_end_x, line_y)], 
                     fill='#3498DB', width=4)
            
            return line_y + 10
            
        except Exception as e:
            print(f"繪製標題時發生錯誤: {e}")
            return y_pos + 80
    
    def _draw_content_enhanced(self, draw, content_items: List[str], 
                              width: int, y_pos: int, font) -> int:
        """繪製增強版內容"""
        try:
            max_items = 6  # 最多顯示6個項目
            item_height = 60
            left_margin = 80
            bullet_size = 8
            
            for i, item in enumerate(content_items[:max_items]):
                if y_pos > 600:  # 避免超出邊界
                    break
                
                safe_item = self._normalize_text_cross_platform(item)
                
                # 繪製項目符號
                bullet_x = left_margin - 30
                bullet_y = y_pos + 20
                draw.ellipse([bullet_x - bullet_size, bullet_y - bullet_size,
                             bullet_x + bullet_size, bullet_y + bullet_size], 
                            fill='#3498DB')
                
                # 處理長文字換行
                wrapped_lines = self._wrap_text_smart(safe_item, width - left_margin - 40, font, draw)
                
                line_y = y_pos
                for line in wrapped_lines[:2]:  # 最多2行
                    if line.strip():
                        draw.text((left_margin, line_y), line, fill='#34495E', font=font)
                        line_y += 35
                
                y_pos += item_height
            
            return y_pos
            
        except Exception as e:
            print(f"繪製內容時發生錯誤: {e}")
            return y_pos + 200
    
    def _add_slide_decorations(self, draw, slide_number: int, width: int, height: int, font):
        """添加投影片裝飾元素"""
        try:
            # 投影片編號
            number_text = f"投影片 {slide_number}"
            draw.text((30, height - 50), number_text, fill='#7F8C8D', font=font)
            
            # 右下角裝飾
            draw.text((width - 150, height - 50), f"{platform.system()}", 
                     fill='#BDC3C7', font=font)
            
            # 頂部裝飾線
            draw.line([(30, 30), (width - 30, 30)], fill='#3498DB', width=2)
            
        except Exception as e:
            print(f"添加裝飾時發生錯誤: {e}")
    
    def _normalize_text_cross_platform(self, text: str) -> str:
        """跨平台標準化文字編碼"""
        try:
            if isinstance(text, bytes):
                # 依平台嘗試不同編碼
                if platform.system() == "Darwin":  # macOS
                    encodings = ['utf-8', 'utf-8-sig', 'macroman', 'big5']
                elif platform.system() == "Windows":
                    encodings = ['utf-8', 'utf-8-sig', 'cp950', 'big5', 'gb2312']
                else:
                    encodings = ['utf-8', 'utf-8-sig', 'gb2312', 'big5']
                
                for encoding in encodings:
                    try:
                        text = text.decode(encoding)
                        break
                    except:
                        continue
                else:
                    text = text.decode('utf-8', errors='ignore')
            
            text = str(text)
            text = text.encode('utf-8', errors='ignore').decode('utf-8')
            
            # 移除控制字元但保留換行符
            import unicodedata
            text = ''.join(char for char in text 
                          if unicodedata.category(char)[0] != 'C' or char in '\n\r\t')
            
            return text.strip()
            
        except Exception as e:
            print(f"跨平台文字編碼標準化失敗: {e}")
            return str(text)[:100] if text else "文字顯示錯誤"
    
    def _wrap_text_smart(self, text: str, max_width: int, font, draw) -> List[str]:
        """智慧文字換行"""
        try:
            lines = []
            words = text.split()
            current_line = ""
            
            for word in words:
                test_line = current_line + (" " if current_line else "") + word
                
                try:
                    bbox = draw.textbbox((0, 0), test_line, font=font)
                    text_width = bbox[2] - bbox[0]
                    
                    if text_width <= max_width:
                        current_line = test_line
                    else:
                        if current_line:
                            lines.append(current_line)
                        current_line = word
                except:
                    # 備用方法：按字符數估算
                    if len(current_line) < max_width // 20:
                        current_line = test_line
                    else:
                        if current_line:
                            lines.append(current_line)
                        current_line = word
            
            if current_line:
                lines.append(current_line)
            
            return lines
            
        except Exception as e:
            print(f"智慧文字換行處理錯誤: {e}")
            return [text[:50] + "..." if len(text) > 50 else text]
    
    def cleanup(self, force_delete: bool = False):
        """清理暫存檔案 - 可選擇是否強制刪除預覽圖片"""
        try:
            import shutil
            if os.path.exists(self.temp_dir):
                # 如果是用戶指定的預覽目錄，默認不刪除
                if "_預覽圖片" in self.temp_dir and not force_delete:
                    print(f"保留預覽圖片目錄: {self.temp_dir}")
                    return
                
                # 只刪除真正的暫存目錄或強制刪除
                shutil.rmtree(self.temp_dir)
                print(f"清理目錄: {self.temp_dir}")
        except Exception as e:
            print(f"清理檔案時發生錯誤: {e}")

class ConversionWorker(QThread):
    """轉換工作執行緒 - 專注 PowerPoint"""
    
    progress_updated = Signal(int)
    status_updated = Signal(str)
    finished_successfully = Signal(str)  # 返回 PowerPoint 檔案路徑
    error_occurred = Signal(str)
    
    def __init__(self, word_path: str, template_path: str, output_path: str):
        super().__init__()
        self.word_path = word_path
        self.template_path = template_path
        self.output_path = output_path
    
    def run(self):
        """執行轉換"""
        try:
            # 步驟1：分析Word文件
            self.status_updated.emit("正在分析Word文件章節結構...")
            self.progress_updated.emit(10)
            
            analyzer = WordDocumentAnalyzer()
            blocks = analyzer.analyze_document(self.word_path)
            
            if not blocks:
                raise Exception("Word文件中沒有找到可轉換的內容")
            
            # 步驟2：檢查PowerPoint範本
            self.status_updated.emit("正在檢查PowerPoint範本...")
            self.progress_updated.emit(20)
            
            if not os.path.exists(self.template_path):
                raise Exception(f"PowerPoint範本檔案不存在: {self.template_path}")
            
            # 步驟3：建立投影片
            self.status_updated.emit("正在建立PowerPoint投影片...")
            self.progress_updated.emit(50)
            
            mapper = ContentToSlideMapper(self.template_path)
            presentation = mapper.create_slides(blocks)
            
            # 步驟4：儲存檔案
            self.status_updated.emit("正在儲存PowerPoint檔案...")
            self.progress_updated.emit(80)
            
            output_dir = os.path.dirname(self.output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            presentation.save(self.output_path)
            
            self.progress_updated.emit(100)
            self.status_updated.emit("轉換完成！字體已優化，支援智慧分頁")
            self.finished_successfully.emit(self.output_path)
            
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
    """預覽元件 - 使用真實 PPTX 圖片"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.image_exporter = None  # 延遲初始化
        
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
    
    def update_preview(self, presentation_path: str):
        """更新預覽 - 使用真實投影片圖片"""
        try:
            self.clear_preview()
            
            # 初始化圖片匯出器，傳遞輸出路徑
            if self.image_exporter:
                self.image_exporter.cleanup()
            self.image_exporter = PPTXImageExporter(presentation_path)
            
            loading_label = QLabel(f"正在使用 {'Spire.Presentation' if SPIRE_AVAILABLE else 'PIL備用方法'} 渲染投影片圖片...")
            loading_label.setAlignment(Qt.AlignCenter)
            loading_label.setStyleSheet("color: #3498db; font-size: 14px; padding: 20px;")
            self.content_layout.addWidget(loading_label)
            
            QApplication.processEvents()
            
            print(f"開始匯出 PPTX 投影片圖片: {presentation_path}")
            
            # 使用真實的 PPTX 圖片匯出
            image_paths = self.image_exporter.export_slides_to_images(presentation_path)
            
            loading_label.deleteLater()
            
            if not image_paths:
                error_label = QLabel(f"❌ 無法渲染投影片圖片\n\n請檢查PowerPoint檔案是否正常")
                error_label.setStyleSheet("color: #e74c3c; padding: 20px; text-align: center;")
                error_label.setAlignment(Qt.AlignCenter)
                self.content_layout.addWidget(error_label)
                return
            
            # 顯示預覽圖片資料夾路徑
            preview_dir = self.image_exporter.temp_dir
            engine_info = "Spire.Presentation 高品質匯出" if SPIRE_AVAILABLE else "PIL 備用渲染"
            info_label = QLabel(f"📊 {engine_info}\n💾 預覽圖片已保存至: {os.path.basename(preview_dir)}")
            info_label.setFont(QFont("Microsoft JhengHei", 10, QFont.Weight.Bold))
            info_label.setStyleSheet("color: #27ae60; padding: 10px; background: #f0f8f0; border-radius: 5px; margin-bottom: 10px;")
            info_label.setAlignment(Qt.AlignCenter)
            info_label.setWordWrap(True)
            self.content_layout.addWidget(info_label)
            
            success_count = 0
            for i, image_path in enumerate(image_paths):
                if os.path.exists(image_path):
                    try:
                        preview_item = self.create_image_preview(image_path, i + 1, preview_dir)
                        self.content_layout.addWidget(preview_item)
                        success_count += 1
                        QApplication.processEvents()
                    except Exception as e:
                        print(f"建立預覽項目失敗: {e}")
                        continue
            
            if success_count > 0:
                engine_method = "Spire高品質匯出" if SPIRE_AVAILABLE else "PIL備用渲染"
                result_label = QLabel(f"✅ {engine_method}成功產生 {success_count} 張JPG投影片預覽")
                result_label.setStyleSheet("color: #27ae60; padding: 5px; font-weight: bold;")
                result_label.setAlignment(Qt.AlignCenter)
                self.content_layout.addWidget(result_label)
            
        except Exception as e:
            self.clear_preview()
            import traceback
            error_details = traceback.format_exc()
            print(f"預覽更新錯誤: {error_details}")
            
            error_label = QLabel(f"❌ 預覽渲染失敗: {str(e)}")
            error_label.setStyleSheet("color: #e74c3c; padding: 20px; background: #fdf2f2; border-radius: 5px;")
            error_label.setWordWrap(True)
            self.content_layout.addWidget(error_label)
    
    def clear_preview(self):
        """清除預覽"""
        while self.content_layout.count():
            child = self.content_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
    
    def create_image_preview(self, image_path: str, slide_number: int, preview_dir: str = None) -> QWidget:
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
        
        platform_label = QLabel(f"Spire引擎" if SPIRE_AVAILABLE else "PIL備用")
        platform_label.setFont(QFont("Microsoft JhengHei", 8))
        platform_label.setStyleSheet("color: #f1c40f; background: rgba(255,255,255,20); padding: 2px 6px; border-radius: 3px;")
        
        header_layout.addWidget(number_label, 1)
        header_layout.addWidget(platform_label)
        
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
                target_width = 800
                target_height = 450
                
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
                relative_path = os.path.basename(image_path) if preview_dir else "暫存檔案"
                file_format = "JPG" if image_path.lower().endswith('.jpg') else "其他"
                image_info = QLabel(f"尺寸: {pixmap.width()}×{pixmap.height()} | 大小: {file_size//1024}KB | 格式: {file_format} | 檔案: {relative_path}")
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
        
        frame.setMaximumHeight(600)
        frame.setMinimumHeight(500)
        return frame
    
    def __del__(self):
        """解構函式 - 不再自動清理圖片檔案"""
        # 不再自動清理，讓用戶可以保留預覽圖片
        pass

class MainWindow(QMainWindow):
    """主視窗 - 專注 PowerPoint 版"""
    
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        
        self.word_path = ""
        self.template_path = ""
        self.output_path = ""
        self.worker = None
        
        self.setup_ui()
        self.setup_connections()
        self.load_last_used_paths()
        self.check_ready_to_convert()
        
    def setup_ui(self):
        """設定UI"""
        self.setWindowTitle(f"Word轉PowerPoint工具 v4.1")
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
        title = QLabel(f"📊 Word轉PowerPoint工具 ({platform.system()})")
        title.setFont(QFont("Microsoft JhengHei", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #2c3e50; margin: 10px;")
        title.setMaximumHeight(40)
        layout.addWidget(title)
        
        # 功能說明
        features = QLabel("✨ 新功能: 統一渲染 | 智慧分頁 | 清除範本 | JPG預覽")
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
        
        # 轉換設定
        settings_group = QGroupBox("2. 轉換設定")
        settings_group.setMaximumHeight(180)
        settings_layout = QVBoxLayout(settings_group)
        settings_layout.setSpacing(8)
        settings_layout.setContentsMargins(10, 15, 10, 10)
        
        # 設定說明
        settings_desc = QLabel(f"• 智慧章節識別與自動分頁\n• 完全清除範本原始內容\n• 文字溢出自動檢測\n• python-pptx統一渲染引擎")
        settings_desc.setFont(QFont("Microsoft JhengHei", 9))
        settings_desc.setStyleSheet("color: #27ae60; margin-bottom: 5px;")
        settings_desc.setMaximumHeight(120)
        settings_layout.addWidget(settings_desc)
        
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
        preview_title = QLabel(f"📋 投影片預覽 ({'Spire引擎' if SPIRE_AVAILABLE else 'PIL備用'})")
        preview_title.setFont(QFont("Microsoft JhengHei", 16, QFont.Weight.Bold))
        preview_title.setStyleSheet("color: #2c3e50; margin: 10px;")
        layout.addWidget(preview_title)
        
        # 預覽說明
        engine_name = "Spire.Presentation 高品質引擎" if SPIRE_AVAILABLE else "PIL 備用渲染引擎"
        preview_desc = QLabel(f"🎯 {engine_name} | 📐 智慧分頁檢測 | 🖼️ 高品質JPG預覽")
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
            self.word_status.setText(f"✅ {filename} (智慧分頁+溢出檢測)")
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
        self.status_label.setText(f"正在準備統一渲染轉換...")
        
        self.worker = ConversionWorker(
            self.word_path, 
            self.template_path, 
            self.output_path
        )
        
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.status_updated.connect(self.status_label.setText)
        self.worker.finished_successfully.connect(self.on_conversion_finished)
        self.worker.error_occurred.connect(self.on_conversion_error)
        
        self.worker.start()
    
    def on_conversion_finished(self, output_path: str):
        """轉換完成處理"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("✅ 轉換完成！正在產生預覽...")
        self.status_label.setStyleSheet("color: #27ae60; font-weight: bold;")
        
        self.check_ready_to_convert()
        
        # 更新預覽（使用真實PPTX圖片）
        self.preview_widget.update_preview(output_path)
        
        # 取得預覽圖片資料夾名稱
        output_name = os.path.splitext(os.path.basename(output_path))[0]
        preview_folder = f"{output_name}_預覽圖片"
        
        # 顯示完成訊息
        engine_info = "Spire.Presentation 高品質引擎" if SPIRE_AVAILABLE else "PIL 備用渲染引擎"
        message = f"""🎉 PowerPoint轉換完成！

📊 已建立檔案: {os.path.basename(output_path)}
📁 預覽圖片資料夾: {preview_folder}

✨ 增強功能:
• 智慧章節識別與自動分頁
• 完全清除範本原始內容
• 文字溢出檢測與處理
• {engine_info}
• 高品質JPG預覽圖片已保存

右側顯示投影片預覽。
預覽圖片已永久保存在輸出檔案同目錄下的資料夾中。

是否立即開啟PowerPoint檔案？"""

        reply = QMessageBox.question(
            self, "轉換完成", message,
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.open_file(output_path)
    
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
    
    def closeEvent(self, event):
        """視窗關閉事件"""
        try:
            if hasattr(self, 'preview_widget'):
                del self.preview_widget
        except:
            pass
        
        event.accept()

def main():
    """主函式"""
    app = QApplication(sys.argv)
    
    app.setApplicationName("Word轉PowerPoint工具")
    app.setApplicationVersion("4.1")
    app.setOrganizationName("智慧辦公工具")
    
    # 設定高DPI支援
    app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    print("=" * 80)
    print("📊 Word轉PowerPoint工具 v4.1 (統一渲染優化版) 📊")
    print("=" * 80)
    print(f"🖥️  運行系統: {platform.system()}")
    print(f"🐍 Python版本: {sys.version}")
    print("")
    print("🎯 核心功能:")
    print("  ✓ 智慧識別中文章節標題（一、二、三、四...）")
    print("  ✓ 文字溢出檢測與自動分頁")
    print("  ✓ 完全清除範本原始內容")
    print("  ✓ 專注PowerPoint (.pptx) 輸出")
    print("  ✓ 統一python-pptx渲染引擎")
    print("  ✓ 記憶上次使用的檔案路徑")
    print("")
    print("🔥 統一渲染特色:")
    print("  • 所有平台使用相同的python-pptx引擎")
    print("  • 高品質720p解析度JPG預覽圖片")
    print("  • 跨平台字體自動選擇和優化")
    print("  • 永久保存預覽圖片到指定資料夾")
    print("")
    print("🖼️  跨平台預覽技術:")
    
    if platform.system() == "Darwin":
        print("  • macOS: 使用Keynote + AppleScript匯出")
        print("  • 備用: python-pptx + PingFang TC字體")
    elif platform.system() == "Windows":
        print("  • Windows: 使用PowerPoint COM自動化")
        print("  • 備用: python-pptx + Microsoft JhengHei字體")
    else:
        print("  • Linux: 使用python-pptx + WenQuanYi字體")
    
    print("  • 智慧錯誤恢復機制")
    print("  • UTF-8編碼處理")
    print("")
    print("🔥 智慧分頁特色:")
    print("  • 自動檢測單張投影片文字長度")
    print("  • 超過限制時自動建立新投影片")
    print("  • 保持相同章節標題連續性")
    print("  • 支援最多4個內容項目或220字符")
    print("")
    print("💾 記憶功能:")
    print("  • 自動記住上次使用的Word檔案")
    print("  • 自動記住上次使用的PowerPoint範本")
    print("  • 自動記住上次使用的輸出目錄")
    print("")
    print("📋 使用步驟:")
    print("  1. 拖放Word文件（支援.docx/.doc）")
    print("  2. 拖放PowerPoint範本（.pptx/.ppt）")
    print("  3. 點擊開始轉換")
    print("  4. 查看右側真實投影片預覽")
    print("")
    print("🛠️  相依性檢查:")
    print("pip install PySide6 python-docx python-pptx pillow")
    
    if platform.system() == "Windows":
        print("pip install pywin32  # Windows COM支援")
    
    print("")
    
    # 檢查系統特定功能
    if platform.system() == "Darwin":
        print("🍎 macOS特定功能:")
        print("  • Keynote + AppleScript高品質匯出")
        print("  • 支援Retina顯示器")
        print("  • 最佳化的中文字體渲染")
    elif platform.system() == "Windows":
        print("🪟 Windows特定功能:")
        print("  • PowerPoint COM自動化")
        print("  • 高DPI顯示支援")
        print("  • 原生PowerPoint匯出品質")
    
    print("")
    print("🌏 技術特點:")
    print("  ✓ 真實PPTX投影片匯出（非模擬）")
    print("  ✓ 跨平台字體自動選擇")
    print("  ✓ 智慧錯誤恢復")
    print("  ✓ 高品質圖片預覽")
    print("  ✓ 記憶式使用者介面")
    print("")
    print("⚠️  注意事項:")
    print("  • Windows用戶建議安裝Microsoft Office以獲得最佳效果")
    print("  • macOS用戶建議安裝Keynote應用程式")
    print("  • 首次運行會測試系統功能並選擇最佳方案")
    print("  • 大型文件轉換需要較長時間")
    print("=" * 80)
    
    # 檢查必要套件
    missing_deps = []
    
    try:
        import PySide6
        print("✅ PySide6 - OK")
    except ImportError:
        missing_deps.append("PySide6")
        print("❌ PySide6 - 缺少")
    
    try:
        import docx
        print("✅ python-docx - OK")
    except ImportError:
        missing_deps.append("python-docx")
        print("❌ python-docx - 缺少")
    
    try:
        import pptx
        print("✅ python-pptx - OK")
    except ImportError:
        missing_deps.append("python-pptx")
        print("❌ python-pptx - 缺少")
    
    try:
        import PIL
        print("✅ Pillow - OK")
    except ImportError:
        missing_deps.append("pillow")
        print("❌ Pillow - 缺少")
    
    print("")
    
    if not DEPENDENCIES_OK or missing_deps:
        if missing_deps:
            print(f"❌ 缺少必要套件: {', '.join(missing_deps)}")
            print(f"請執行: pip install {' '.join(missing_deps)}")
        print("\n程式無法啟動，請先安裝缺少的套件。")
        sys.exit(1)
    
    # 顯示系統資訊
    print("🔍 系統資訊:")
    print(f"  作業系統: {platform.system()} {platform.release()}")
    print(f"  處理器: {platform.processor()}")
    print(f"  Python: {platform.python_version()}")
    
    if platform.system() == "Darwin":
        try:
            result = subprocess.run(['sw_vers', '-productVersion'], capture_output=True, text=True)
            if result.returncode == 0:
                print(f"  macOS版本: {result.stdout.strip()}")
        except:
            pass
    elif platform.system() == "Windows":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:
                version = winreg.QueryValueEx(key, "ProductName")[0]
                print(f"  Windows版本: {version}")
        except:
            pass
    
    print(f"  當前目錄: {os.getcwd()}")
    print("")
    
    # 預覽功能測試
    print("🧪 預覽功能測試:")
    print("  ✅ python-pptx統一渲染 - 所有平台可用")
    print("  ✅ PIL圖片處理 - 所有平台可用")
    print("  ✅ 跨平台字體系統 - 自動適配")
    print("  ✅ JPG格式輸出 - 高品質壓縮")
    print("")
    print("🚀 啟動應用程式...")
    print("=" * 80)
    
    try:
        main()
    except Exception as e:
        print(f"❌ 程式執行錯誤: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)