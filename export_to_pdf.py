#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将长图按原始尺寸转换成多列 A4 PDF，最大化利用页面空间
"""

import os
import sys
from pathlib import Path
import argparse
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm


def calculate_optimal_layout(img_width, img_height, num_columns=3, 
                             page_size=A4, orientation='landscape',
                             margin=10):
    """
    计算最优布局，最大化利用页面空间
    
    返回：每列的宽度和高度（像素）
    """
    # A4 尺寸（单位：mm）
    if orientation == 'landscape':
        page_width, page_height = landscape(page_size)
    else:
        page_width, page_height = page_size
    
    # 转换为点（1mm = 2.83465 points）
    # 可用宽度和高度（减去边距）
    available_width = page_width - (2 * margin * mm)
    available_height = page_height - (2 * margin * mm)
    
    # 每列可用宽度
    column_width_pts = available_width / num_columns
    
    # 计算图片的DPI和缩放比例
    # 假设图片原始DPI为72（默认）
    img_width_pts = img_width * 72 / 96  # 转换为点
    
    # 计算缩放比例，使图片宽度适应列宽
    scale = column_width_pts / img_width_pts
    
    # 计算每列的高度（像素）
    # 考虑缩放后的高度不能超过页面高度
    scaled_height_pts = img_height * scale * 72 / 96
    
    if scaled_height_pts > available_height:
        # 如果高度超过页面，重新计算缩放比例
        scale = available_height / (img_height * 72 / 96)
    
    # 计算每列应该包含的原图高度（像素）
    column_height_px = int(available_height / scale * 96 / 72)
    
    return column_width_pts, available_height, scale, column_height_px


def split_image_to_pdf(input_path, output_pdf=None, num_columns=3, 
                       orientation='landscape', margin=10, overlap=0, column_gap=3):
    """
    将长图分割并转换成多列 A4 PDF
    
    参数:
        input_path: 输入图片路径
        output_pdf: 输出PDF文件名
        num_columns: 每页显示的列数
        orientation: 页面方向 'landscape'(横向) 或 'portrait'(纵向)
        margin: 页边距（单位：mm）
        overlap: 列之间重叠的像素数
        column_gap: 列之间的间隔（单位：mm）
    """
    try:
        # 打开图片
        img = Image.open(input_path)
        img_width, img_height = img.size
        print(f"原图尺寸: {img_width} x {img_height} 像素")
        
        # 计算最优布局
        page_size = landscape(A4) if orientation == 'landscape' else A4
        page_width, page_height = page_size
        
        print(f"页面尺寸: {page_width/mm:.1f} x {page_height/mm:.1f} mm")
        print(f"页面方向: {'横向' if orientation == 'landscape' else '纵向'}")
        print(f"页边距: {margin} mm")
        print(f"列间隔: {column_gap} mm")
        print(f"列数: {num_columns}")
        
        # 可用区域（减去边距和列间隔）
        available_width = page_width - (2 * margin * mm) - ((num_columns - 1) * column_gap * mm)
        available_height = page_height - (2 * margin * mm)
        
        # 每列宽度（点）
        column_width_pts = available_width / num_columns
        
        # 计算每列应包含的图片高度（像素）
        # 图片会按宽度缩放以适应列宽，然后计算能放多高
        scale_for_width = column_width_pts / (img_width * 72 / 96)
        column_height_px = int(available_height / scale_for_width / 72 * 96)
        
        # 计算总共需要多少列（考虑重叠）
        total_segments = 0
        current_pos = 0
        while current_pos < img_height:
            total_segments += 1
            current_pos += column_height_px
            if current_pos < img_height and overlap > 0:
                current_pos -= overlap
        
        # 计算页数
        total_pages = (total_segments + num_columns - 1) // num_columns
        
        print(f"每列高度: {column_height_px} 像素")
        print(f"总共分成: {total_segments} 列")
        print(f"预计页数: {total_pages} 页")
        
        # 输出PDF路径
        if output_pdf is None:
            output_pdf = Path(input_path).parent / f"{Path(input_path).stem}_多列打印.pdf"
        else:
            output_pdf = Path(output_pdf)
        
        # 创建PDF
        c = canvas.Canvas(str(output_pdf), pagesize=page_size)
        
        # 分割图片并添加到PDF
        # 将长图分成多个段，每个段作为一列，每页显示 num_columns 列
        segments = []
        current_y = 0
        
        # 先生成所有列段
        while current_y < img_height:
            start_y = current_y
            end_y = min(current_y + column_height_px, img_height)
            
            # 裁剪列段
            segment = img.crop((0, start_y, img_width, end_y))
            segments.append({
                'image': segment,
                'start_y': start_y,
                'end_y': end_y
            })
            
            # 移动到下一段（考虑重叠）
            current_y = end_y
            if current_y < img_height and overlap > 0:
                current_y -= overlap
        
        print(f"\n已分割成 {len(segments)} 个列段")
        
        # 按页面排列列段
        page_num = 1
        for page_start in range(0, len(segments), num_columns):
            print(f"\n生成第 {page_num} 页...")
            page_segments = segments[page_start:page_start + num_columns]
            
            for col_idx, seg_info in enumerate(page_segments):
                segment = seg_info['image']
                
                # 保存临时图片（使用全局索引避免重名）
                global_idx = page_start + col_idx
                temp_path = output_pdf.parent / f"temp_seg_{global_idx}.png"
                segment.save(temp_path, dpi=(96, 96))
                
                # 计算在PDF中的位置
                x_pos = margin * mm + col_idx * (column_width_pts + column_gap * mm)
                y_pos = page_height - margin * mm  # 从顶部开始
                
                # 计算显示尺寸（保持宽高比，适应列宽）
                display_width = column_width_pts
                display_height = (segment.height / img_width) * display_width
                
                # 确保不超过可用高度
                if display_height > available_height:
                    display_height = available_height
                    display_width = (img_width / segment.height) * display_height
                
                y_pos = y_pos - display_height
                
                # 绘制图片
                c.drawImage(str(temp_path), x_pos, y_pos, 
                           width=display_width, height=display_height,
                           preserveAspectRatio=True)
                
                print(f"  列 {col_idx + 1}: 原图像素 {seg_info['start_y']}-{seg_info['end_y']}")
                
                # 删除临时文件
                temp_path.unlink()
            
            # 如果还有更多段，添加新页
            if page_start + num_columns < len(segments):
                c.showPage()
                page_num += 1
        
        # 保存PDF
        c.save()
        
        print(f"\n✅ 成功！PDF 已保存: {output_pdf.absolute()}")
        print(f"共生成 {page_num} 页")
        print(f"\n📋 打印说明:")
        print(f"  1. 打开 {output_pdf.name}")
        print(f"  2. 使用 Adobe Reader 或系统自带 PDF 阅读器打开")
        print(f"  3. 打印时选择「实际大小」，不要缩放")
        print(f"  4. 确认纸张大小为 A4")
        
        return True
        
    except Exception as e:
        print(f"❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def process_directory(input_dir, output_dir=None, num_columns=3, 
                      orientation='landscape', margin=10, overlap=0, column_gap=3):
    """
    批量处理目录中的所有图片
    """
    input_path = Path(input_dir)
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
    
    # 查找所有图片文件
    image_files = [f for f in input_path.iterdir() 
                   if f.is_file() and f.suffix.lower() in image_extensions]
    
    if not image_files:
        print(f"❌ 在目录 {input_dir} 中未找到图片文件")
        return False
    
    print(f"找到 {len(image_files)} 个图片文件")
    print("=" * 60)
    
    success_count = 0
    for img_file in image_files:
        print(f"\n处理: {img_file.name}")
        
        # 确定输出路径
        if output_dir:
            output_pdf = Path(output_dir) / f"{img_file.stem}_多列打印.pdf"
        else:
            output_pdf = img_file.parent / f"{img_file.stem}_多列打印.pdf"
        
        if split_image_to_pdf(
            str(img_file), 
            str(output_pdf),
            num_columns,
            orientation,
            margin,
            overlap,
            column_gap
        ):
            success_count += 1
        
        print("=" * 60)
    
    print(f"\n✅ 批量处理完成！成功处理 {success_count}/{len(image_files)} 个文件")
    return success_count > 0


def main():
    parser = argparse.ArgumentParser(
        description='将长图按原始尺寸转换成多列 A4 PDF',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 单个文件（3列，横向）
  python export_to_pdf.py target.jpg
  
  # 批量处理文件夹
  python export_to_pdf.py ./images/ -c 3 --overlap 50
  
  # 指定输出文件名（单个文件）
  python export_to_pdf.py target.jpg -o 打印文件.pdf
  
  # 批量处理并指定输出目录
  python export_to_pdf.py ./images/ --output-dir ./pdfs/
  
  # 2列布局
  python export_to_pdf.py target.jpg -c 2
  
  # 纵向布局（适合窄图）
  python export_to_pdf.py target.jpg --orientation portrait
  
  # 自定义页边距和重叠
  python export_to_pdf.py target.jpg --margin 5 --overlap 50
  
  # 批量处理（先用 split_long_image.py 生成列图片）
  python export_to_pdf.py output/target_列1.png -c 1 --orientation portrait
        """
    )
    
    parser.add_argument('input', help='输入图片文件或目录')
    parser.add_argument('-o', '--output', default=None,
                        help='输出PDF文件名（单个文件时使用，默认: 原文件名_多列打印.pdf）')
    parser.add_argument('--output-dir', default=None,
                        help='输出目录（批量处理时使用，默认: 与输入文件相同目录）')
    parser.add_argument('-c', '--columns', type=int, default=3,
                        help='每页显示的列数（默认: 3）')
    parser.add_argument('--orientation', choices=['landscape', 'portrait'], 
                        default='landscape',
                        help='页面方向：landscape(横向) 或 portrait(纵向)，默认: landscape')
    parser.add_argument('--margin', type=float, default=10,
                        help='页边距（单位：mm），默认: 10')
    parser.add_argument('--overlap', type=int, default=0,
                        help='列之间重叠的像素数（默认: 0）')
    parser.add_argument('--column-gap', type=float, default=3,
                        help='列之间的间隔（单位：mm），默认: 3')
    
    args = parser.parse_args()
    
    # 检查输入路径
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"❌ 错误: 路径不存在: {args.input}")
        sys.exit(1)
    
    # 判断是文件还是目录
    if input_path.is_file():
        # 处理单个文件
        if args.output_dir:
            print("⚠️  警告: --output-dir 参数仅在批量处理时有效，已忽略")
        
        success = split_image_to_pdf(
            args.input,
            args.output,
            args.columns,
            args.orientation,
            args.margin,
            args.overlap,
            args.column_gap
        )
        
        if not success:
            sys.exit(1)
    
    elif input_path.is_dir():
        # 批量处理目录
        if args.output:
            print("⚠️  警告: -o/--output 参数仅在处理单个文件时有效，已忽略")
        
        success = process_directory(
            args.input,
            args.output_dir,
            args.columns,
            args.orientation,
            args.margin,
            args.overlap,
            args.column_gap
        )
        
        if not success:
            sys.exit(1)
    
    else:
        print(f"❌ 错误: 无效的输入路径: {args.input}")
        sys.exit(1)


if __name__ == "__main__":
    main()
