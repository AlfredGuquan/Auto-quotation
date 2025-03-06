# 应用模板样式
def apply_template_styles(ws):
    # 定义常用样式
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    # 应用单元格样式
    # 设置 A1 单元格样式
    ws.merge_cells('A1:F1')
    ws['A1'].font = Font(name='Microsoft YaHei', size=16.0, bold=True, color='FF002060')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A1'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 A2 单元格样式
    ws.merge_cells('A2:B5')
    ws['A2'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws['A2'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 C2 单元格样式
    ws['C2'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C2'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws['C2'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 D2 单元格样式
    ws.merge_cells('D2:F5')
    ws['D2'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D2'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws['D2'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 C3 单元格样式
    ws['C3'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C3'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws['C3'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 C4 单元格样式
    ws['C4'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C4'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws['C4'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 C5 单元格样式
    ws['C5'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C5'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws['C5'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 A6 单元格样式
    ws.merge_cells('A6:F6')
    ws['A6'].font = Font(name='Microsoft YaHei', size=16.0, bold=True, color='FFFFFFFF')
    ws['A6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A6'].fill = PatternFill(fill_type='solid', start_color='FF538DD5', end_color='Values must be of type <class 'str'>')
    ws['A6'].border = Border(left=Side(style='thin', color='Values must be of type <class 'str'>'), right=Side(style='thin', color='Values must be of type <class 'str'>'), bottom=Side(style='thin', color='Values must be of type <class 'str'>'))
    # 设置 A7 单元格样式
    ws['A7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['A7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='Values must be of type <class 'str'>')
    ws['A7'].border = thin_border
    # 设置 B7 单元格样式
    ws['B7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['B7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='Values must be of type <class 'str'>')
    ws['B7'].border = thin_border
    # 设置 C7 单元格样式
    ws['C7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['C7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='Values must be of type <class 'str'>')
    ws['C7'].border = thin_border
    # 设置 D7 单元格样式
    ws['D7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['D7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='Values must be of type <class 'str'>')
    ws['D7'].border = thin_border
    # 设置 E7 单元格样式
    ws['E7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['E7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='Values must be of type <class 'str'>')
    ws['E7'].border = thin_border
    # 设置 F7 单元格样式
    ws['F7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['F7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='Values must be of type <class 'str'>')
    ws['F7'].border = thin_border
    # 设置 A8 单元格样式
    ws['A8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A8'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['A8'].border = thin_border
    # 设置 B8 单元格样式
    ws['B8'].font = Font(name='Microsoft YaHei', size=10.0, color='Values must be of type <class 'str'>')
    ws['B8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B8'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['B8'].border = thin_border
    # 设置 C8 单元格样式
    ws['C8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C8'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['C8'].border = thin_border
    # 设置 D8 单元格样式
    ws['D8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D8'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['D8'].border = thin_border
    # 设置 E8 单元格样式
    ws['E8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E8'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E8'].border = thin_border
    # 设置 F8 单元格样式
    ws['F8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F8'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F8'].border = thin_border
    # 设置 A9 单元格样式
    ws['A9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A9'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['A9'].border = thin_border
    # 设置 B9 单元格样式
    ws['B9'].font = Font(name='Microsoft YaHei', size=10.0, color='Values must be of type <class 'str'>')
    ws['B9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B9'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['B9'].border = thin_border
    # 设置 C9 单元格样式
    ws['C9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C9'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['C9'].border = thin_border
    # 设置 D9 单元格样式
    ws['D9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D9'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['D9'].border = thin_border
    # 设置 E9 单元格样式
    ws['E9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E9'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E9'].border = thin_border
    # 设置 F9 单元格样式
    ws['F9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F9'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F9'].border = thin_border
    # 设置 A10 单元格样式
    ws['A10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A10'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['A10'].border = thin_border
    # 设置 B10 单元格样式
    ws['B10'].font = Font(name='Microsoft YaHei', size=10.0, color='Values must be of type <class 'str'>')
    ws['B10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B10'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['B10'].border = thin_border
    # 设置 C10 单元格样式
    ws['C10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C10'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['C10'].border = thin_border
    # 设置 D10 单元格样式
    ws['D10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D10'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['D10'].border = thin_border
    # 设置 E10 单元格样式
    ws['E10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E10'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E10'].border = thin_border
    # 设置 F10 单元格样式
    ws['F10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F10'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F10'].border = thin_border
    # 设置 A11 单元格样式
    ws['A11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A11'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['A11'].border = thin_border
    # 设置 B11 单元格样式
    ws['B11'].font = Font(name='Microsoft YaHei', size=10.0, color='Values must be of type <class 'str'>')
    ws['B11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B11'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['B11'].border = thin_border
    # 设置 C11 单元格样式
    ws['C11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C11'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['C11'].border = thin_border
    # 设置 D11 单元格样式
    ws['D11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D11'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['D11'].border = thin_border
    # 设置 E11 单元格样式
    ws['E11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E11'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E11'].border = thin_border
    # 设置 F11 单元格样式
    ws['F11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F11'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F11'].border = thin_border
    # 设置 A12 单元格样式
    ws['A12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A12'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['A12'].border = thin_border
    # 设置 B12 单元格样式
    ws['B12'].font = Font(name='Microsoft YaHei', size=10.0, color='Values must be of type <class 'str'>')
    ws['B12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B12'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['B12'].border = thin_border
    # 设置 C12 单元格样式
    ws['C12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C12'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['C12'].border = thin_border
    # 设置 D12 单元格样式
    ws['D12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D12'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['D12'].border = thin_border
    # 设置 E12 单元格样式
    ws['E12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E12'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E12'].border = thin_border
    # 设置 F12 单元格样式
    ws['F12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F12'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F12'].border = thin_border
    # 设置 A13 单元格样式
    ws['A13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A13'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['A13'].border = thin_border
    # 设置 B13 单元格样式
    ws['B13'].font = Font(name='Microsoft YaHei', size=10.0, color='Values must be of type <class 'str'>')
    ws['B13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B13'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['B13'].border = thin_border
    # 设置 C13 单元格样式
    ws['C13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C13'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['C13'].border = thin_border
    # 设置 D13 单元格样式
    ws['D13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D13'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['D13'].border = thin_border
    # 设置 E13 单元格样式
    ws['E13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E13'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E13'].border = thin_border
    # 设置 F13 单元格样式
    ws['F13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F13'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F13'].border = thin_border
    # 设置 A14 单元格样式
    ws.merge_cells('A14:A18')
    ws['A14'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A14'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A14'].fill = PatternFill(fill_type='solid', start_color='Values must be of type <class 'str'>', end_color='Values must be of type <class 'str'>')
    ws['A14'].border = Border(top=Side(style='thin', color='Values must be of type <class 'str'>'))
    # 设置 B14 单元格样式
    ws.merge_cells('B14:D20')
    ws['B14'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B14'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws['B14'].fill = PatternFill(fill_type='solid', start_color='Values must be of type <class 'str'>', end_color='Values must be of type <class 'str'>')
    ws['B14'].border = Border(top=Side(style='thin', color='Values must be of type <class 'str'>'))
    # 设置 E14 单元格样式
    ws['E14'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E14'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E14'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E14'].border = thin_border
    # 设置 F14 单元格样式
    ws['F14'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F14'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F14'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F14'].border = thin_border
    # 设置 E15 单元格样式
    ws['E15'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E15'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E15'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E15'].border = thin_border
    # 设置 F15 单元格样式
    ws['F15'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F15'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F15'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F15'].border = thin_border
    # 设置 E16 单元格样式
    ws['E16'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E16'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E16'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E16'].border = thin_border
    # 设置 F16 单元格样式
    ws['F16'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F16'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F16'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F16'].border = thin_border
    # 设置 E17 单元格样式
    ws['E17'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E17'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E17'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['E17'].border = thin_border
    # 设置 F17 单元格样式
    ws['F17'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F17'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F17'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    ws['F17'].border = thin_border
    # 设置 E18 单元格样式
    ws['E18'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['E18'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 F18 单元格样式
    ws['F18'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['F18'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 A19 单元格样式
    ws['A19'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['A19'].alignment = Alignment(vertical='center')
    ws['A19'].fill = PatternFill(fill_type='solid', start_color='Values must be of type <class 'str'>', end_color='Values must be of type <class 'str'>')
    # 设置 E19 单元格样式
    ws['E19'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['E19'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 F19 单元格样式
    ws['F19'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['F19'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 A20 单元格样式
    ws['A20'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['A20'].alignment = Alignment(vertical='center')
    ws['A20'].fill = PatternFill(fill_type='solid', start_color='Values must be of type <class 'str'>', end_color='Values must be of type <class 'str'>')
    # 设置 E20 单元格样式
    ws['E20'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['E20'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 F20 单元格样式
    ws['F20'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['F20'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 A21 单元格样式
    ws['A21'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['A21'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 B21 单元格样式
    ws['B21'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['B21'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 C21 单元格样式
    ws['C21'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['C21'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 D21 单元格样式
    ws['D21'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['D21'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 E21 单元格样式
    ws['E21'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['E21'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 F21 单元格样式
    ws['F21'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['F21'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 A22 单元格样式
    ws['A22'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['A22'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 B22 单元格样式
    ws['B22'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['B22'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 C22 单元格样式
    ws['C22'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['C22'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 D22 单元格样式
    ws['D22'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['D22'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 E22 单元格样式
    ws['E22'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['E22'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')
    # 设置 F22 单元格样式
    ws['F22'].font = Font(name='宋体', size=11.0, color='Values must be of type <class 'str'>')
    ws['F22'].fill = PatternFill(fill_type='None', start_color='00000000', end_color='00000000')