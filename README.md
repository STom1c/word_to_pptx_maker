# Word to PowerPoint Converter (通用Word轉PowerPoint工具)

A professional-grade GUI application that converts Microsoft Word documents to PowerPoint presentations with intelligent content parsing and advanced formatting features.

## Features

### Core Functionality
- **Smart Document Analysis**: Intelligent parsing of Word documents with automatic chapter recognition
- **Chapter Recognition**: Advanced detection of Chinese numbering systems (一、二、三、四...)
- **Template Integration**: Seamless integration with PowerPoint templates
- **UTF-8 Encoding Support**: Complete UTF-8 encoding with Chinese character support
- **Real-time Preview**: Live slide preview with 16:9 aspect ratio
- **Path Memory**: Automatic saving and loading of previously used file paths

### Enhanced Features
- **Large Font Optimization**: Automatic font sizing to minimum 32pt for better readability
- **Drag & Drop Interface**: Intuitive file selection with drag-and-drop support
- **Cross-platform Compatibility**: Works on Windows, macOS, and Linux
- **Professional Styling**: Modern GUI with gradient backgrounds and professional layouts
- **Error Resilient**: Comprehensive error handling and recovery mechanisms

### Technical Highlights
- **Multi-encoding Support**: UTF-8, Big5, GB2312, UTF-16 compatibility
- **Font Management**: Intelligent system font detection and Chinese font support
- **Memory Efficient**: Optimized processing for large documents
- **Thread-safe**: Background processing with progress indication

## Installation

### System Requirements
- Python 3.8 or higher
- Windows 10+, macOS 10.14+, or Linux with GUI support

### Quick Install
```bash
# Create virtual environment
python -m venv word_to_pptx_env
source word_to_pptx_env/bin/activate  # macOS/Linux
# word_to_pptx_env\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt
```

### Manual Install
```bash
pip install PySide6>=6.5.0 python-docx>=0.8.11 python-pptx>=0.6.21 pillow>=9.5.0
```

## Usage

### GUI Application
```bash
python universal_word_to_pptx_gui.py
```

### Step-by-Step Process
1. **Select Word Document**: Drag & drop or click to select `.docx` or `.doc` files
2. **Choose PowerPoint Template**: Select your presentation template (`.pptx` or `.ppt`)
3. **Set Output Path**: Choose where to save the converted presentation
4. **Start Conversion**: Click "開始轉換" to begin processing
5. **Preview Results**: View generated slides in the preview panel

## Document Structure Recognition

### Supported Chapter Patterns
- Chinese numerals: `一、二、三、四...`
- Arabic numerals: `1. 2. 3. 4...`
- Prefixed chapters: `第一章、第二章...`
- Special markers: `前言、結論、總結...`

### Content Classification
- **Headers**: Main document titles
- **Chapters**: Section dividers (Level 0)
- **Subtitles**: Subsection headers (Level 1)
- **Content**: Body text and lists (Level 2)

## Configuration

### Automatic Path Memory
The application automatically saves:
- Last used Word document path
- Last used PowerPoint template path
- Last used output directory
- Window geometry settings

Configuration is stored in: `~/.word_to_ppt_config.json`

### Font Handling
Automatic font detection priority:
1. Microsoft JhengHei (Windows Traditional Chinese)
2. Microsoft YaHei (Windows Simplified Chinese)  
3. PingFang SC (macOS Chinese)
4. System fallback fonts

## Technical Architecture

### Core Components
- **WordDocumentAnalyzer**: Intelligent document structure analysis
- **PowerPointTemplateAnalyzer**: Template layout detection
- **ContentToSlideMapper**: Smart content-to-slide mapping
- **SlideImageGenerator**: UTF-8 compatible preview generation
- **ConfigManager**: Persistent settings management

### Data Flow
```
Word Document → Content Analysis → Template Mapping → Slide Generation → Preview Rendering
```

### Error Handling
- File encoding detection and conversion
- Template compatibility validation
- Memory management for large documents
- Graceful degradation for unsupported features

## File Format Support

### Input Formats
- **Word Documents**: `.docx`, `.doc`
- **PowerPoint Templates**: `.pptx`, `.ppt`

### Output Format
- **PowerPoint Presentation**: `.pptx` (Office 2010+ compatible)

## Advanced Features

### UTF-8 Encoding Engine
- Multi-layer encoding detection
- Chinese character normalization
- Font compatibility testing
- Text rendering optimization

### Preview System
- Real-time slide generation
- 16:9 aspect ratio preservation
- High-quality image rendering
- Memory-efficient processing

### Performance Optimization
- Lazy loading for large documents
- Caching for repeated operations
- Background processing threads
- Progress indication

## Troubleshooting

### Common Issues
1. **Font Display Problems**: Ensure Chinese fonts are installed
2. **Encoding Errors**: Check document encoding and save as UTF-8
3. **Template Compatibility**: Verify PowerPoint template structure
4. **Memory Issues**: Close other applications when processing large files

### Error Messages
- "Word文件中沒有找到可轉換的內容": Document contains no parseable content
- "PowerPoint範本檔案不存在": Template file not found or corrupted
- "圖片載入失敗": Preview generation failed, check file permissions

## Development

### Project Structure
```
word_to_pptx_maker/
├── universal_word_to_pptx_gui.py    # Main application
├── requirements.txt                  # Dependencies
├── README.md                        # Documentation
├── .gitignore                       # Git ignore rules
└── .word_to_ppt_config.json        # User configuration (auto-generated)
```

### Building Executables
```bash
# Install PyInstaller
pip install pyinstaller

# Build standalone executable
pyinstaller --onefile --windowed universal_word_to_pptx_gui.py

# Build with icon (optional)
pyinstaller --onefile --windowed --icon=app_icon.ico universal_word_to_pptx_gui.py
```

## Contributing

### Development Setup
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

### Code Style
- Follow PEP 8 guidelines
- Use meaningful variable names
- Add docstrings for all functions
- Include error handling

## License

This project is proprietary software developed for semiconductor testing and document processing applications.

## Version History

- **v2.3**: UTF-8 encoding fixes, enhanced preview system
- **v2.2**: Path memory functionality, improved error handling
- **v2.1**: Advanced chapter recognition, font optimization
- **v2.0**: Complete GUI rewrite with PySide6
- **v1.0**: Initial release with basic conversion features

## Support

For technical support and feature requests, please contact the development team or create an issue in the project repository.

---

**Note**: This tool is optimized for semiconductor engineering documentation and technical presentations. It includes specialized features for processing technical content and maintaining professional formatting standards.