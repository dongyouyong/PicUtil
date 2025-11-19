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
    将长图分割并转换成多列 A4 PDF - 零信息丢失版本（增强重叠算法）
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
        base_column_height_px = int(available_height / scale_for_width / 72 * 96)
        
        # 计算实际需要的净列高（减去重叠）
        net_column_height = base_column_height_px - overlap
        
        # 如果净高度太小，调整参数
        if net_column_height <= 0:
            print(f"  警告：重叠像素({overlap})过大，自动调整")
            overlap = base_column_height_px // 3
            net_column_height = base_column_height_px - overlap
        
        print(f"  分列参数: 列高{base_column_height_px}px, 重叠{overlap}px")
        
        # 重新设计重叠算法：第一列的最后N像素和第二列的开头N像素重合
        segments = []
        current_y = 0
        segment_index = 0
        
        while current_y < img_height:
            # 计算当前段的结束位置
            segment_end = min(current_y + base_column_height_px, img_height)
            
            # 裁剪图片段
            img_segment = img.crop((0, current_y, img_width, segment_end))
            
            # 添加段信息
            segments.append({
                'image': img_segment,
                'start_y': current_y,
                'end_y': segment_end,
                'height': segment_end - current_y
            })
            
            print(f"    段{segment_index + 1}: Y={current_y}-{segment_end} (高度{segment_end - current_y}px)")
            
            # 如果已经到达图片末尾，结束
            if segment_end >= img_height:
                break
            
            # 计算下一段的开始位置：当前段结束位置向前回退重叠像素
            # 这样确保第一列的最后overlap像素和第二列的开头overlap像素重合
            next_start = segment_end - overlap
            
            # 防止下一段开始位置不合理
            if next_start <= current_y:
                # 如果重叠太大导致下一段开始位置不合理，调整
                next_start = current_y + (base_column_height_px - overlap)
                if next_start >= img_height:
                    break
            
            current_y = next_start
            segment_index += 1
            
            # 防止无限循环
            if segment_index > 50:  # 降低安全限制
                print("  警告：分段数量过多，停止分段")
                break
        
        # 验证重叠性（新的重叠逻辑）
        overlap_check_passed = True
        overlap_details = []
        
        if len(segments) > 1:
            for i in range(1, len(segments)):
                prev_end = segments[i-1]['end_y'] 
                curr_start = segments[i]['start_y']
                actual_overlap = prev_end - curr_start
                overlap_details.append(f"段{i}-{i+1}: 重叠{actual_overlap}px")
                
                if actual_overlap < overlap:
                    print(f"  警告：段{i}与段{i+1}间重叠不足({actual_overlap}px < {overlap}px)")
                    overlap_check_passed = False
        
        # 检查完整覆盖
        coverage_check = segments[-1]['end_y'] >= img_height if segments else False
        if not coverage_check:
            print(f"  调整最后一段以完整覆盖图片")
            if segments:
                # 重新裁剪最后一段
                last_start = segments[-1]['start_y']
                segments[-1] = {
                    'image': img.crop((0, last_start, img_width, img_height)),
                    'start_y': last_start,
                    'end_y': img_height,
                    'height': img_height - last_start
                }
        
        # 计算页数
        total_pages = (len(segments) + num_columns - 1) // num_columns
        
        coverage_range = f"0-{segments[-1]['end_y']}px" if segments else "0-0px"
        overlap_status = "✓" if overlap_check_passed else "⚠"
        print(f"  分成: {len(segments)}段, {total_pages}页, 覆盖:{coverage_range} {overlap_status}")
        
        # 创建PDF
        c = canvas.Canvas(str(output_pdf), pagesize=page_size)
        
        # 按页面排列列段
        for page_start in range(0, len(segments), num_columns):
            page_segments = segments[page_start:page_start + num_columns]
            
            for col_idx, seg_info in enumerate(page_segments):
                segment = seg_info['image']
                
                # 保存临时图片
                global_idx = page_start + col_idx
                temp_path = output_pdf.parent / f"temp_seg_{global_idx}.png"
                segment.save(temp_path, dpi=(96, 96))
                
                # 计算在PDF中的位置
                x_pos = margin * mm + col_idx * (column_width_pts + column_gap * mm)
                
                # 计算显示尺寸（保持宽高比，适应列宽）
                display_width = column_width_pts
                display_height = (segment.height / img_width) * display_width
                
                # 确保不超过可用高度
                if display_height > available_height:
                    display_height = available_height
                    display_width = (img_width / segment.height) * display_height
                
                # Y位置：从页面顶部开始
                y_pos = page_height - margin * mm - display_height
                
                # 绘制图片
                c.drawImage(str(temp_path), x_pos, y_pos, 
                           width=display_width, height=display_height,
                           preserveAspectRatio=True)
                
                # 删除临时文件
                temp_path.unlink(missing_ok=True)
            
            # 如果还有更多段，添加新页
            if page_start + num_columns < len(segments):
                c.showPage()
        
        # 保存PDF
        c.save()
        
        # 最终状态
        final_status = "零信息丢失" if (overlap_check_passed and coverage_check) else "增强覆盖"
        print(f"  ✅ 成功生成: {output_pdf.name} ({final_status})")
        return True
        
    except Exception as e:
        print(f"  ❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
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
