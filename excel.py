import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import json
import os

def extract_cell_style(cell):
    """提取单元格的样式信息"""
    style = {}
    
    # 提取字体信息
    if cell.font:
        style['font'] = {
            'name': cell.font.name,
            'size': cell.font.size,
            'bold': cell.font.bold,
            'italic': cell.font.italic,
            'color': cell.font.color.rgb if cell.font.color else None
        }
    
    # 提取对齐信息
    if cell.alignment:
        style['alignment'] = {
            'horizontal': cell.alignment.horizontal,
            'vertical': cell.alignment.vertical,
            'wrap_text': cell.alignment.wrap_text
        }
    
    # 提取填充信息
    if cell.fill and cell.fill.fill_type != 'none':
        style['fill'] = {
            'fill_type': cell.fill.fill_type,
            'start_color': cell.fill.start_color.rgb if hasattr(cell.fill.start_color, 'rgb') else None,
            'end_color': cell.fill.end_color.rgb if hasattr(cell.fill.end_color, 'rgb') else None
        }
    
    # 提取边框信息
    if any([cell.border.left, cell.border.right, cell.border.top, cell.border.bottom]):
        style['border'] = {}
        for side in ['left', 'right', 'top', 'bottom']:
            border_side = getattr(cell.border, side)
            if border_side and border_side.style:
                style['border'][side] = {
                    'style': border_side.style,
                    'color': border_side.color.rgb if border_side.color else None
                }
    
    # 提取合并单元格信息
    style['merged'] = False
    
    return style

def extract_template_styles(template_file, range_str='A1:F22'):
    """提取模板文件中指定范围的单元格样式"""
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    
    # 解析范围字符串
    start_cell, end_cell = range_str.split(':')
    start_col, start_row = openpyxl.utils.cell.coordinate_from_string(start_cell)
    end_col, end_row = openpyxl.utils.cell.coordinate_from_string(end_cell)
    
    start_col_idx = openpyxl.utils.column_index_from_string(start_col)
    end_col_idx = openpyxl.utils.column_index_from_string(end_col)
    
    # 提取样式
    styles = {}
    merged_cells = ws.merged_cells.ranges
    
    for row in range(start_row, end_row + 1):
        for col_idx in range(start_col_idx, end_col_idx + 1):
            col = openpyxl.utils.get_column_letter(col_idx)
            cell_coord = f"{col}{row}"
            cell = ws[cell_coord]
            
            # 检查是否为合并单元格
            for merged_range in merged_cells:
                if cell_coord in merged_range:
                    if cell_coord == merged_range.coord.split(':')[0]:  # 如果是合并单元格的左上角
                        styles[cell_coord] = extract_cell_style(cell)
                        styles[cell_coord]['merged'] = str(merged_range)
                    break
            else:
                styles[cell_coord] = extract_cell_style(cell)
    
    return styles

def generate_style_application_code(styles):
    """生成应用样式的Python代码"""
    code = []
    code.append("# 应用模板样式")
    code.append("def apply_template_styles(ws):")
    code.append("    # 定义常用样式")
    code.append("    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))")
    code.append("    # 应用单元格样式")
    
    for cell_coord, style in styles.items():
        code.append(f"    # 设置 {cell_coord} 单元格样式")
        
        # 处理合并单元格
        if style.get('merged'):
            code.append(f"    ws.merge_cells('{style['merged']}')")
        
        # 设置字体
        if 'font' in style:
            font = style['font']
            font_params = []
            if font.get('name'):
                font_params.append(f"name='{font['name']}'")
            if font.get('size'):
                font_params.append(f"size={font['size']}")
            if font.get('bold'):
                font_params.append(f"bold={font['bold']}")
            if font.get('italic'):
                font_params.append(f"italic={font['italic']}")
            if font.get('color'):
                font_params.append(f"color='{font['color']}'")
            
            if font_params:
                code.append(f"    ws['{cell_coord}'].font = Font({', '.join(font_params)})")
        
        # 设置对齐
        if 'alignment' in style:
            align = style['alignment']
            align_params = []
            if align.get('horizontal'):
                align_params.append(f"horizontal='{align['horizontal']}'")
            if align.get('vertical'):
                align_params.append(f"vertical='{align['vertical']}'")
            if align.get('wrap_text'):
                align_params.append(f"wrap_text={align['wrap_text']}")
            
            if align_params:
                code.append(f"    ws['{cell_coord}'].alignment = Alignment({', '.join(align_params)})")
        
        # 设置填充
        if 'fill' in style and style['fill'].get('start_color'):
            fill = style['fill']
            code.append(f"    ws['{cell_coord}'].fill = PatternFill(fill_type='{fill['fill_type']}', start_color='{fill['start_color']}', end_color='{fill['end_color'] or fill['start_color']}')")
        
        # 设置边框
        if 'border' in style:
            border = style['border']
            if len(border) == 4:  # 如果四边都有边框
                code.append(f"    ws['{cell_coord}'].border = thin_border")
            else:
                border_params = []
                for side, border_style in border.items():
                    border_params.append(f"{side}=Side(style='{border_style['style']}', color='{border_style['color']}')")
                
                if border_params:
                    code.append(f"    ws['{cell_coord}'].border = Border({', '.join(border_params)})")
    
    return "\n".join(code)

def main():
    template_file = "测试.xlsx"
    output_file = "template_styles.py"
    
    if not os.path.exists(template_file):
        print(f"错误: 模板文件 '{template_file}' 不存在!")
        return
    
    print(f"正在提取 '{template_file}' 的样式...")
    styles = extract_template_styles(template_file)
    
    print(f"生成样式应用代码...")
    code = generate_style_application_code(styles)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(code)
    
    print(f"样式代码已生成到 '{output_file}'")
    print("请将此代码集成到 generate.py 中，并在创建Excel文件时调用 apply_template_styles(ws) 函数")

if __name__ == "__main__":
    main()

