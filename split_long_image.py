#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信长截图分列工具
将长截图自动分割成多列，方便打印
"""

import os
import sys
from PIL import Image
import argparse
from pathlib import Path


def split_image_to_columns(input_path, output_dir=None, num_columns=2, overlap=0, dpi=300, column_gap=20):
    """
    将长图分割成多列
    
    参数:
        input_path: 输入图片路径
        output_dir: 输出目录，默认为当前目录下的output文件夹
        num_columns: 分割的列数，默认2列
        overlap: 列之间的重叠像素数，默认0
        dpi: 输出图片的DPI，默认300（适合打印）
        column_gap: 合并图片时列之间的间隔像素，默认20
    """
    try:
        # 打开图片
        img = Image.open(input_path)
        width, height = img.size
        print(f"原图尺寸: {width} x {height} 像素")
        
        # 计算每列的高度
        column_height = height // num_columns
        
        # 如果有重叠，需要调整高度
        if overlap > 0:
            column_height = (height + overlap * (num_columns - 1)) // num_columns
        
        print(f"分割成 {num_columns} 列，每列高度约: {column_height} 像素")
        
        # 创建输出目录
        if output_dir is None:
            output_dir = Path(input_path).parent / "output"
        else:
            output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 获取原文件名（不含扩展名）
        input_filename = Path(input_path).stem
        
        # 分割图片
        columns = []
        for i in range(num_columns):
            # 计算当前列的起始和结束位置
            start_y = i * column_height - (i * overlap if i > 0 else 0)
            end_y = min(start_y + column_height, height)
            
            # 如果是最后一列，确保包含所有剩余内容
            if i == num_columns - 1:
                end_y = height
            
            # 裁剪图片
            column = img.crop((0, start_y, width, end_y))
            columns.append(column)
            
            # 保存单独的列
            output_path = output_dir / f"{input_filename}_列{i+1}.png"
            column.save(output_path, dpi=(dpi, dpi))
            print(f"已保存: {output_path}")
        
        # 创建横向拼接的版本（所有列并排显示，带间隔）
        gap = column_gap  # 列之间的间隔（像素）
        total_width = width * num_columns + gap * (num_columns - 1)
        max_height = max(col.height for col in columns)
        
        combined = Image.new('RGB', (total_width, max_height), (255, 255, 255))
        
        for i, col in enumerate(columns):
            x_position = i * (width + gap)
            combined.paste(col, (x_position, 0))
        
        combined_path = output_dir / f"{input_filename}_合并_{num_columns}列.png"
        combined.save(combined_path, dpi=(dpi, dpi))
        print(f"已保存合并版本: {combined_path}")
        
        print(f"\n✅ 处理完成！共生成 {num_columns + 1} 个文件")
        print(f"输出目录: {output_dir.absolute()}")
        
        return True
        
    except Exception as e:
        print(f"❌ 错误: {str(e)}")
        return False


def process_directory(input_dir, output_dir=None, num_columns=2, overlap=0, dpi=300, column_gap=20):
    """
    批量处理目录中的所有图片
    """
    input_path = Path(input_dir)
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
    
    # 查找所有图片文件
    image_files = [f for f in input_path.iterdir() 
                   if f.is_file() and f.suffix.lower() in image_extensions]
    
    if not image_files:
        print(f"在目录 {input_dir} 中未找到图片文件")
        return
    
    print(f"找到 {len(image_files)} 个图片文件")
    print("=" * 50)
    
    success_count = 0
    for img_file in image_files:
        print(f"\n处理: {img_file.name}")
        if split_image_to_columns(str(img_file), output_dir, num_columns, overlap, dpi, column_gap):
            success_count += 1
        print("=" * 50)
    
    print(f"\n✅ 批量处理完成！成功处理 {success_count}/{len(image_files)} 个文件")


def main():
    parser = argparse.ArgumentParser(
        description='将微信长截图分割成多列，方便打印',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 将单张图片分成2列
  python split_long_image.py screenshot.png
  
  # 将图片分成3列
  python split_long_image.py screenshot.png -c 3
  
  # 批量处理目录中的所有图片
  python split_long_image.py ./screenshots/ -c 2
  
  # 设置列之间的重叠（避免内容被切断）
  python split_long_image.py screenshot.png -c 2 --overlap 50
  
  # 指定输出目录和DPI
  python split_long_image.py screenshot.png -o ./output -d 150
        """
    )
    
    parser.add_argument('input', help='输入图片文件或目录')
    parser.add_argument('-c', '--columns', type=int, default=2, 
                        help='分割的列数（默认: 2）')
    parser.add_argument('-o', '--output', default=None,
                        help='输出目录（默认: ./output）')
    parser.add_argument('--overlap', type=int, default=0,
                        help='列之间重叠的像素数（默认: 0）')
    parser.add_argument('-d', '--dpi', type=int, default=300,
                        help='输出图片DPI，用于打印（默认: 300）')
    parser.add_argument('--column-gap', type=int, default=20,
                        help='合并图片时列之间的间隔像素（默认: 20）')
    
    args = parser.parse_args()
    
    # 检查输入是文件还是目录
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"❌ 错误: 路径不存在: {args.input}")
        sys.exit(1)
    
    if input_path.is_file():
        # 处理单个文件
        split_image_to_columns(
            args.input, 
            args.output, 
            args.columns, 
            args.overlap,
            args.dpi,
            args.column_gap
        )
    elif input_path.is_dir():
        # 批量处理目录
        process_directory(
            args.input,
            args.output,
            args.columns,
            args.overlap,
            args.dpi,
            args.column_gap
        )
    else:
        print(f"❌ 错误: 无效的输入路径: {args.input}")
        sys.exit(1)


if __name__ == "__main__":
    main()
