#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量处理当前目录的所有图片，转换成PDF
双击运行即可
"""

import os
import sys
from pathlib import Path
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm


def split_image_to_pdf(input_path, output_pdf, num_columns=3, 
                       orientation='landscape', margin=10, overlap=50, column_gap=5):
    """
    将长图分割并转换成多列 A4 PDF
    """
    try:
        # 打开图片
        img = Image.open(input_path)
        img_width, img_height = img.size
        print(f"  原图尺寸: {img_width} x {img_height} 像素")
        
        # 计算最优布局
        page_size = landscape(A4) if orientation == 'landscape' else A4
        page_width, page_height = page_size
        
        # 可用区域（减去边距和列间隔）
        available_width = page_width - (2 * margin * mm) - ((num_columns - 1) * column_gap * mm)
        available_height = page_height - (2 * margin * mm)
        
        # 每列宽度（点）
        column_width_pts = available_width / num_columns
        
        # 计算每列应包含的图片高度（像素）
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
        
        print(f"  分成: {total_segments} 列，{total_pages} 页")
        
        # 创建PDF
        c = canvas.Canvas(str(output_pdf), pagesize=page_size)
        
        # 分割图片并添加到PDF
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
        
        # 按页面排列列段
        page_num = 1
        for page_start in range(0, len(segments), num_columns):
            page_segments = segments[page_start:page_start + num_columns]
            
            for col_idx, seg_info in enumerate(page_segments):
                segment = seg_info['image']
                
                # 保存临时图片（使用全局索引避免重名）
                global_idx = page_start + col_idx
                temp_path = output_pdf.parent / f"temp_seg_{global_idx}.png"
                segment.save(temp_path, dpi=(96, 96))
                
                # 计算在PDF中的位置
                x_pos = margin * mm + col_idx * (column_width_pts + column_gap * mm)
                y_pos = page_height - margin * mm
                
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
                
                # 删除临时文件
                temp_path.unlink()
            
            # 如果还有更多段，添加新页
            if page_start + num_columns < len(segments):
                c.showPage()
                page_num += 1
        
        # 保存PDF
        c.save()
        
        print(f"  ✅ 成功生成: {output_pdf.name}")
        return True
        
    except Exception as e:
        print(f"  ❌ 错误: {str(e)}")
        return False


def main():
    print("=" * 70)
    print("长图转PDF打印工具 - 批量处理模式")
    print("=" * 70)
    print()
    
    # 获取当前目录
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        current_dir = Path(sys.executable).parent
    else:
        # 如果是python脚本
        current_dir = Path.cwd()
    
    print(f"当前目录: {current_dir}")
    print()
    
    # 查找所有图片文件
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
    image_files = [f for f in current_dir.iterdir() 
                   if f.is_file() and f.suffix.lower() in image_extensions]
    
    if not image_files:
        print("❌ 未找到任何图片文件！")
        print()
        print("支持的格式: JPG, PNG, BMP, GIF, WebP")
        input("\n按回车键退出...")
        return
    
    print(f"找到 {len(image_files)} 个图片文件:")
    for img in image_files:
        print(f"  - {img.name}")
    print()
    
    # 询问参数
    print("请设置参数（直接按回车使用默认值）:")
    print()
    
    try:
        columns_input = input("每页列数 (默认: 3): ").strip()
        num_columns = int(columns_input) if columns_input else 3
        
        overlap_input = input("列重叠像素 (默认: 50): ").strip()
        overlap = int(overlap_input) if overlap_input else 50
        
        gap_input = input("列间隔(mm) (默认: 5): ").strip()
        column_gap = float(gap_input) if gap_input else 5
        
    except ValueError:
        print("⚠️  输入无效，使用默认值")
        num_columns = 3
        overlap = 50
        column_gap = 5
    
    print()
    print("=" * 70)
    print("开始处理...")
    print("=" * 70)
    print()
    
    # 创建输出目录
    output_dir = current_dir / "PDF输出"
    output_dir.mkdir(exist_ok=True)
    
    # 批量处理
    success_count = 0
    for idx, img_file in enumerate(image_files, 1):
        print(f"[{idx}/{len(image_files)}] 处理: {img_file.name}")
        
        output_pdf = output_dir / f"{img_file.stem}_打印.pdf"
        
        if split_image_to_pdf(
            img_file,
            output_pdf,
            num_columns=num_columns,
            orientation='landscape',
            margin=10,
            overlap=overlap,
            column_gap=column_gap
        ):
            success_count += 1
        
        print()
    
    print("=" * 70)
    print(f"✅ 处理完成！成功: {success_count}/{len(image_files)}")
    print(f"📁 输出目录: {output_dir.absolute()}")
    print("=" * 70)
    print()
    
    # 询问是否打开输出目录
    open_folder = input("是否打开输出目录? (Y/n): ").strip().lower()
    if open_folder != 'n':
        try:
            if sys.platform == 'win32':
                os.startfile(output_dir)
            elif sys.platform == 'darwin':
                os.system(f'open "{output_dir}"')
            else:
                os.system(f'xdg-open "{output_dir}"')
        except Exception as e:
            print(f"无法打开目录: {e}")
    
    input("\n按回车键退出...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
        input("\n按回车键退出...")
