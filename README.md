# HD Image PPT Builder

[中文说明](#中文说明) | [English](#english)

Convert PowerPoint slides to high-resolution images, then rebuild as a layout-locked PPTX with speaker notes preserved. Perfect for conference presentations where formatting must not break.

将 PowerPoint 每页渲染为高清图片，重建为排版零偏移的纯图 PPTX，同时逐字保留演讲者备注。适用于学术会议大屏放映和设备兼容性兜底。

---

<a name="english"></a>
## English

### Why This Tool Exists

Conference presentations often break on the venue computer because of:
- Missing fonts on the projection machine
- Different PowerPoint versions changing line breaks/spacing
- Image scaling differences causing misalignment

This tool renders every slide as a high-resolution image, then rebuilds a pure-image PPTX — **zero layout shift** — while **preserving speaker notes verbatim**.

### Features

- **Layout-lock**: Every slide becomes a single full-bleed image — nothing can shift
- **Speaker notes preserved**: Notes are copied verbatim to the new file
- **Resolution selectable**: 2K (2560×1440), 4K (3840×2160), or 5K (5120×2880)
- **Vector-quality rendering**: Uses PowerPoint COM engine — EMF formulas, charts, and complex layouts render perfectly
- **No source modification**: Reads your file, never writes to it
- **One-click EXE available**: No Python environment needed for end users

### Requirements

- Windows
- Microsoft PowerPoint (not WPS — real Office required for COM rendering)
- Python 3.10+ with: `pip install python-pptx pywin32 Pillow`

### Usage

#### GUI
```
python hd_pptx_gui.py
```
Or download the standalone `HD_PPT_Builder.exe` from [Releases](https://github.com/fyvtbuhnjkml/hd-image-ppt-builder/releases).

#### CLI
```bash
python hd_pptx.py "presentation.pptx"                          # default 5K
python hd_pptx.py "presentation.pptx" "output.pptx" --dpi 3840 # 4K
```

### How It Works

1. Opens the source PPTX via PowerPoint COM (read-only)
2. Exports each slide as PNG at the target resolution
3. Creates a new PPTX with the same canvas dimensions
4. Inserts each PNG as a full-bleed picture (locked layout)
5. Copies speaker notes verbatim from the original
6. Saves the output — source file untouched

### Build EXE

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes hd_pptx_gui.py
```

---

<a name="中文说明"></a>
## 中文说明

### 为什么需要这个工具

在会议现场，你精心排版的 PPT 经常因为以下原因"跑版"：
- 投影电脑缺少你使用的字体
- PowerPoint 版本不同导致自动换行/行距变化
- 图片缩放策略不同导致偏移

本工具把每一页"烧"成一张高清图片，然后重建一个纯图片 PPT——**排版零偏移**，同时**完整保留演讲者备注**，演讲者视图正常使用。

### 功能

- **排版锁定**：每页变为一张整页图片，文字/图片/公式位置绝对不会偏移
- **备注逐字保留**：演讲者视图正常使用
- **分辨率可选**：2K (2560×1440)、4K (3840×2160)、5K (5120×2880)
- **矢量级渲染**：通过 PowerPoint COM 引擎渲染，EMF 公式、复杂图表完美呈现
- **不修改源文件**
- **一键 EXE**：打包后无需 Python 环境即可使用

### 使用方法

#### GUI（图形界面）
```
python hd_pptx_gui.py
```
或从 [Releases](https://github.com/fyvtbuhnjkml/hd-image-ppt-builder/releases) 下载 `HD_PPT_Builder.exe`。

#### 命令行
```bash
python hd_pptx.py "你的文件.pptx"                          # 默认 5K
python hd_pptx.py "你的文件.pptx" "输出.pptx" --dpi 3840   # 4K
```

### 工作原理

1. 通过 PowerPoint COM（只读模式）打开源 PPTX
2. 将每页导出为目标分辨率的 PNG
3. 创建新的 PPTX，画布尺寸与原文件一致
4. 每页插入一张整页图片（排版锁定）
5. 逐字复制演讲者备注
6. 保存输出文件——源文件不做任何修改

### 分辨率建议

| 清晰度 | 像素 | 适用场景 |
|---|---|---|
| 2K | 2560×1440 | 普通投影 |
| 4K | 3840×2160 | 会议大屏（推荐） |
| 5K | 5120×2880 | 4K/5K LED 大屏 |

### 注意事项

- 输出的 PPT 中每一页是一张图片，**不可编辑文字**
- 如需修改内容，请回到原文件修改后重新转换
- 演讲者备注**逐字保留**，演讲者视图正常使用
- 不会修改源文件

### License

MIT
