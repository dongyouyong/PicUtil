# 打包说明

## Windows 打包步骤

### 1. 安装 PyInstaller

```bash
pip install pyinstaller
```

### 2. 打包成 exe

```bash
pyinstaller batch_convert.spec
```

### 3. 生成的文件

打包完成后，在 `dist` 目录下会生成 `长图转PDF工具.exe`

## 使用方法

### Windows 用户

1. 将 `长图转PDF工具.exe` 复制到包含图片的文件夹
2. 双击运行 `长图转PDF工具.exe`
3. 按照提示设置参数（或直接按回车使用默认值）
4. 等待处理完成
5. 生成的 PDF 文件会保存在 `PDF输出` 文件夹中

### 默认参数

- 每页列数: 3
- 列重叠: 50 像素
- 列间隔: 5 mm
- 页面方向: 横向
- 纸张大小: A4

## 注意事项

1. exe 文件需要放在**包含图片的文件夹**中运行
2. 支持的图片格式: JPG, PNG, BMP, GIF, WebP
3. 生成的 PDF 会自动保存在 `PDF输出` 文件夹
4. 打印时请选择「实际大小」，不要缩放

## 高级选项

如果需要修改默认参数，可以编辑 `batch_convert.py` 文件中的默认值，然后重新打包。

## 跨平台打包

- **Windows**: 使用 Windows 系统打包
- **macOS**: 使用 macOS 系统打包，生成 `.app` 或 `.dmg`
- **Linux**: 使用 Linux 系统打包

PyInstaller 生成的可执行文件只能在对应的操作系统上运行。
