# 微信长截图分列工具

将微信长截图自动分割成多列，方便打印。

## 功能特点

- ✨ 自动将长截图分割成N列
- 📄 生成单独的列图片和合并的横向版本
- 📊 **导出到 Excel，完美打印布局**
- 📕 **导出到 PDF，按原始尺寸最大化利用页面空间**
- 💻 **可打包成 Windows exe，双击即用**
- 🖨️ 支持设置打印DPI（默认300）
- 🔄 支持列之间重叠，避免内容被切断
- 📁 支持批量处理整个目录的图片

## 安装依赖

```bash
pip install -r requirements.txt
```

或者直接安装：

```bash
pip install Pillow openpyxl reportlab pyinstaller
```

## 使用方法

### 🎯 方案零：Windows 双击运行（最简单，无需安装 Python）

**适合不懂编程的用户**

1. 下载 `长图转PDF工具.exe`（需要先打包，见 [BUILD.md](BUILD.md)）
2. 将 exe 文件放到包含图片的文件夹
3. 双击运行
4. 按提示设置参数（或直接按回车使用默认值）
5. 等待处理完成，PDF 文件会自动生成在 `PDF输出` 文件夹

### 🎯 方案一：直接导出 PDF（推荐）

**一步到位，按原始尺寸生成多列 PDF：**

```bash
python export_to_pdf.py target.jpg
```

生成的 PDF 会自动：
- ✅ 将长图分成3列（横向A4）
- ✅ 最大化利用页面空间
- ✅ 保持原始图片质量
- ✅ 直接打印，无需调整

**自定义列数和方向：**
```bash
# 2列布局
python export_to_pdf.py target.jpg -c 2

# 纵向布局
python export_to_pdf.py target.jpg --orientation portrait

# 调整页边距
python export_to_pdf.py target.jpg --margin 5
```

### 🎯 方案二：Excel 打印（适合需要编辑）

**步骤1：分割长截图**
```bash
python split_long_image.py target.jpg -c 3 --overlap 50
```

**步骤2：导出到 Excel**
```bash
python export_to_excel.py output/
```

**步骤3：在 Excel 中打印**
1. 打开生成的 `打印预览.xlsx`
2. 按 `⌘ + P` (Mac) 或 `Ctrl + P` (Windows) 打印
3. 确认纸张方向为「横向」
4. 点击打印 ✅

### 🎯 方案三：仅分割图片（自己处理）

将单张图片分成2列（默认）：

```bash
python split_long_image.py screenshot.png
```

**基本用法：**

```bash
python split_long_image.py screenshot.png
```

**指定列数（3列）：**

```bash
python split_long_image.py screenshot.png -c 3
```

**批量处理：**

```bash
python split_long_image.py ./screenshots/ -c 2
```

**设置重叠（推荐）：**

```bash
python split_long_image.py screenshot.png -c 2 --overlap 50
```

## 命令行参数说明

### export_to_pdf.py 参数（推荐）

| 参数 | 简写 | 说明 | 默认值 |
|------|------|------|--------|
| `input` | - | 输入图片文件（必需） | - |
| `--output` | `-o` | 输出PDF文件名 | 原文件名_多列打印.pdf |
| `--columns` | `-c` | 每页显示的列数 | 3 |
| `--orientation` | - | 页面方向：landscape/portrait | landscape |
| `--margin` | - | 页边距（mm） | 10 |
| `--overlap` | - | 列之间重叠的像素数 | 0 |

### split_long_image.py 参数

| 参数 | 简写 | 说明 | 默认值 |
|------|------|------|--------|
| `input` | - | 输入图片文件或目录（必需） | - |
| `--columns` | `-c` | 分割的列数 | 2 |
| `--output` | `-o` | 输出目录 | ./output |
| `--overlap` | - | 列之间重叠的像素数 | 0 |
| `--dpi` | `-d` | 输出图片DPI | 300 |

### export_to_excel.py 参数

| 参数 | 简写 | 说明 | 默认值 |
|------|------|------|--------|
| `input_dir` | - | 包含列图片的目录（必需） | - |
| `--output` | `-o` | 输出Excel文件名 | 打印预览.xlsx |
| `--columns` | `-c` | 每页显示的列数 | 3 |
| `--column-width` | - | Excel列宽（字符） | 25 |
| `--row-height` | - | Excel行高（磅） | 150 |
| `--page-break` | - | 每N组图片插入分页符 | 无 |

## 输出说明

### split_long_image.py 输出

程序会生成以下文件：

1. **单独的列图片**：`原文件名_列1.png`, `原文件名_列2.png`, ...
2. **合并的横向图片**：`原文件名_合并_N列.png`（所有列并排显示，方便预览）

所有文件默认保存在 `output` 目录下。

### export_to_excel.py 输出

生成一个 Excel 文件（默认 `打印预览.xlsx`），包含：
- 自动调整的列宽和行高
- 横向页面设置（适合打印）
- 所有列图片按序排列

### export_to_pdf.py 输出

生成一个多列 PDF 文件（默认 `原文件名_多列打印.pdf`），特点：
- 按原始尺寸最大化利用 A4 页面空间
- 自动计算最优布局
- 直接打印，选择「实际大小」即可

## 使用场景

- 📱 微信长截图太长，无法在一页纸上打印
- 📝 需要将长图内容分栏打印，节省纸张
- 🖼️ 需要将长图转换成适合打印的格式
- 📄 需要保持原始图片质量和尺寸

## 示例

### 示例1：直接生成 PDF（最简单）

```bash
python export_to_pdf.py chat.png
```

输出：`chat_多列打印.pdf`（3列横向布局，直接打印）

### 示例2：分割后导出 Excel

假设你有一个微信长截图 `chat.png`，高度为5000像素：

```bash
# 步骤1：分割
python split_long_image.py chat.png -c 2 --overlap 50

# 步骤2：导出到 Excel
python export_to_excel.py output/
```

输出：
- `output/chat_列1.png`（前2525像素）
- `output/chat_列2.png`（从2475像素到5000像素，有50像素重叠）
- `output/chat_合并_2列.png`（两列并排显示）
- `output/打印预览.xlsx`（Excel 打印文件）

## 注意事项

### PDF 打印
- ✅ 打印时选择「实际大小」，不要选「适合页面」
- ✅ 确认纸张大小为 A4
- ✅ 使用高质量打印设置

### 图片分割
- 建议使用 `--overlap` 参数（如50-100像素）来避免文字或图片被切断
- 打印前建议先查看合并版本，确认分割效果
- 默认DPI 300适合高质量打印，如果只是预览可以降低到150节省空间

## 许可证

MIT License
