
def apply_template_styles(ws):
    # 定义常用样式
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # 应用合并单元格
    ws.merge_cells('A1:F1')
    ws.merge_cells('A2:B5')
    ws.merge_cells('D2:F5')
    ws.merge_cells('A6:F6')
    ws.merge_cells('A14:A18')
    ws.merge_cells('B14:D20')

    # 应用单元格样式
    # 设置 A1 单元格样式
    ws['A1'].value = 'ZeroErr Control Co., Ltd.'
    ws['A1'].font = Font(name='Microsoft YaHei', size=16.0, bold=True, color='FF002060')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # 设置 B1 单元格样式
    ws['B1'].font = Font(name='宋体', size=11.0)

    # 设置 C1 单元格样式
    ws['C1'].font = Font(name='宋体', size=11.0)

    # 设置 D1 单元格样式
    ws['D1'].font = Font(name='宋体', size=11.0)

    # 设置 E1 单元格样式
    ws['E1'].font = Font(name='宋体', size=11.0)

    # 设置 F1 单元格样式
    ws['F1'].font = Font(name='宋体', size=11.0)

    # 设置 A2 单元格样式
    ws['A2'].value = 'To:  Vinothkumar Viswanathan\nCompany: CSIRO\nStreet Address:Data61, 1 Technology Court,Pullenvale, \nCity, ST  ZIP Code: Qld 4069, Australia\nPhone:+617 3327 4077'
    ws['A2'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # 设置 B2 单元格样式
    ws['B2'].font = Font(name='宋体', size=11.0)

    # 设置 C2 单元格样式
    ws['C2'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C2'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 设置 D2 单元格样式
    ws['D2'].value = 'Date：5.3.2025\nQuotation #: 2025030503\nStreet Address：Fuyuan 1st, Fuhai City, ZIP Code：Bao\'an，Shen Zhen, 518103\nPhone：+86 18922807806\n'
    ws['D2'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D2'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # 设置 E2 单元格样式
    ws['E2'].font = Font(name='宋体', size=11.0)

    # 设置 F2 单元格样式
    ws['F2'].font = Font(name='宋体', size=11.0)

    # 设置 A3 单元格样式
    ws['A3'].font = Font(name='宋体', size=11.0)

    # 设置 B3 单元格样式
    ws['B3'].font = Font(name='宋体', size=11.0)

    # 设置 C3 单元格样式
    ws['C3'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C3'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 设置 D3 单元格样式
    ws['D3'].font = Font(name='宋体', size=11.0)

    # 设置 E3 单元格样式
    ws['E3'].font = Font(name='宋体', size=11.0)

    # 设置 F3 单元格样式
    ws['F3'].font = Font(name='宋体', size=11.0)

    # 设置 A4 单元格样式
    ws['A4'].font = Font(name='宋体', size=11.0)

    # 设置 B4 单元格样式
    ws['B4'].font = Font(name='宋体', size=11.0)

    # 设置 C4 单元格样式
    ws['C4'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C4'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 设置 D4 单元格样式
    ws['D4'].font = Font(name='宋体', size=11.0)

    # 设置 E4 单元格样式
    ws['E4'].font = Font(name='宋体', size=11.0)

    # 设置 F4 单元格样式
    ws['F4'].font = Font(name='宋体', size=11.0)

    # 设置 A5 单元格样式
    ws['A5'].font = Font(name='宋体', size=11.0)

    # 设置 B5 单元格样式
    ws['B5'].font = Font(name='宋体', size=11.0)

    # 设置 C5 单元格样式
    ws['C5'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C5'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 设置 D5 单元格样式
    ws['D5'].font = Font(name='宋体', size=11.0)

    # 设置 E5 单元格样式
    ws['E5'].font = Font(name='宋体', size=11.0)

    # 设置 F5 单元格样式
    ws['F5'].font = Font(name='宋体', size=11.0)

    # 设置 A6 单元格样式
    ws['A6'].value = 'Quotation List '
    ws['A6'].font = Font(name='Microsoft YaHei', size=16.0, bold=True, color='FFFFFFFF')
    ws['A6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A6'].fill = PatternFill(fill_type='solid', start_color='FF538DD5', end_color='FF538DD5')
    ws['A6'].border = Border(left=Side(style='thin'), right=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B6 单元格样式
    ws['B6'].font = Font(name='宋体', size=11.0)
    ws['B6'].border = Border(bottom=Side(style='thin'))

    # 设置 C6 单元格样式
    ws['C6'].font = Font(name='宋体', size=11.0)
    ws['C6'].border = Border(bottom=Side(style='thin'))

    # 设置 D6 单元格样式
    ws['D6'].font = Font(name='宋体', size=11.0)
    ws['D6'].border = Border(bottom=Side(style='thin'))

    # 设置 E6 单元格样式
    ws['E6'].font = Font(name='宋体', size=11.0)
    ws['E6'].border = Border(bottom=Side(style='thin'))

    # 设置 F6 单元格样式
    ws['F6'].font = Font(name='宋体', size=11.0)
    ws['F6'].border = Border(right=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A7 单元格样式
    ws['A7'].value = 'IMAGES'
    ws['A7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['A7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='FFDCE6F1')
    ws['A7'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B7 单元格样式
    ws['B7'].value = 'MODELS'
    ws['B7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['B7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='FFDCE6F1')
    ws['B7'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 C7 单元格样式
    ws['C7'].value = 'QUANTITY\n(PC)'
    ws['C7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['C7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='FFDCE6F1')
    ws['C7'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 D7 单元格样式
    ws['D7'].value = 'WEIGHT\n(KG/PC)'
    ws['D7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['D7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='FFDCE6F1')
    ws['D7'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 E7 单元格样式
    ws['E7'].value = 'UNIT PRICE\n(USD/PC)'
    ws['E7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['E7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='FFDCE6F1')
    ws['E7'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F7 单元格样式
    ws['F7'].value = 'AMOUNT\n(USD)'
    ws['F7'].font = Font(name='Microsoft YaHei', size=11.0, bold=True, color='FF002060')
    ws['F7'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F7'].fill = PatternFill(fill_type='solid', start_color='FFDCE6F1', end_color='FFDCE6F1')
    ws['F7'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A8 单元格样式
    ws['A8'].value = '=_xlfn.DISPIMG("ID_7B1E0C7B5C53469CB0A42E2E1ED2308A",1)'
    ws['A8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A8'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B8 单元格样式
    ws['B8'].value = 'eRob110H100I-BHM-18ET[V6]'
    ws['B8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B8'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 C8 单元格样式
    ws['C8'].value = 6
    ws['C8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C8'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 D8 单元格样式
    ws['D8'].value = 2.68
    ws['D8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D8'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 E8 单元格样式
    ws['E8'].value = '=1156+46+23+23+92'
    ws['E8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E8'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F8 单元格样式
    ws['F8'].value = '=PRODUCT(C8,E8)'
    ws['F8'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F8'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F8'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A9 单元格样式
    ws['A9'].value = '=_xlfn.DISPIMG("ID_C004F17782744D88A68B2D406AD95193",1)'
    ws['A9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A9'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B9 单元格样式
    ws['B9'].value = 'eRob80H50I-BHM-18ET[V6]'
    ws['B9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B9'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 C9 单元格样式
    ws['C9'].value = 8
    ws['C9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C9'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 D9 单元格样式
    ws['D9'].value = 1.03
    ws['D9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D9'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 E9 单元格样式
    ws['E9'].value = '=974+46+23+23+78'
    ws['E9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E9'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F9 单元格样式
    ws['F9'].value = '=PRODUCT(C9,E9)'
    ws['F9'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F9'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F9'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A10 单元格样式
    ws['A10'].value = '=_xlfn.DISPIMG("ID_BECB8E922C1F4EE9B4C0257A1A74009E",1)'
    ws['A10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A10'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B10 单元格样式
    ws['B10'].value = 'eRob90H100I-BHM-18ET[V6]'
    ws['B10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B10'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 C10 单元格样式
    ws['C10'].value = 6
    ws['C10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C10'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 D10 单元格样式
    ws['D10'].value = 1.62
    ws['D10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D10'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 E10 单元格样式
    ws['E10'].value = '=1104+180+46+23+23'
    ws['E10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E10'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F10 单元格样式
    ws['F10'].value = '=PRODUCT(C10,E10)'
    ws['F10'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F10'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F10'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A11 单元格样式
    ws['A11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A11'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B11 单元格样式
    ws['B11'].value = 'eLine - RJ45 ECAT -30'
    ws['B11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B11'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 C11 单元格样式
    ws['C11'].value = '=C12'
    ws['C11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C11'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 D11 单元格样式
    ws['D11'].value = 0.03
    ws['D11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D11'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 E11 单元格样式
    ws['E11'].value = 11
    ws['E11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E11'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F11 单元格样式
    ws['F11'].value = '=PRODUCT(C11,E11)'
    ws['F11'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F11'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F11'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A12 单元格样式
    ws['A12'].value = '=_xlfn.DISPIMG("ID_A2E589D82096455190DF7C332474A5FA",1)'
    ws['A12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A12'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B12 单元格样式
    ws['B12'].value = 'eRob Universal Accessories Kit'
    ws['B12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B12'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 C12 单元格样式
    ws['C12'].value = '=SUM(C8:C10)'
    ws['C12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C12'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 D12 单元格样式
    ws['D12'].value = 0.2
    ws['D12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D12'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 E12 单元格样式
    ws['E12'].value = 35
    ws['E12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E12'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F12 单元格样式
    ws['F12'].value = '=PRODUCT(C12,E12)'
    ws['F12'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F12'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F12'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A13 单元格样式
    ws['A13'].value = '=_xlfn.DISPIMG("ID_98D5881DD6FB43E9A01277EE4846B5C3",1)'
    ws['A13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A13'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 B13 单元格样式
    ws['B13'].value = 'eRob to PC Connector'
    ws['B13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['B13'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 C13 单元格样式
    ws['C13'].value = 5
    ws['C13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['C13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['C13'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 D13 单元格样式
    ws['D13'].value = 0.2
    ws['D13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['D13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['D13'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 E13 单元格样式
    ws['E13'].value = 66
    ws['E13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['E13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E13'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F13 单元格样式
    ws['F13'].value = '=PRODUCT(C13,E13)'
    ws['F13'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F13'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F13'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A14 单元格样式
    ws['A14'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['A14'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['A14'].fill = PatternFill(fill_type='solid', start_color=None, end_color=None)
    ws['A14'].border = Border(top=Side(style='thin'))

    # 设置 B14 单元格样式
    ws['B14'].value = 'Remarks:\n1. Price term: DAP Australia\n2. Payment term: T/T. 100% advance payment.\n3. Leading time: 12 working days after the payment. \n4.The price needs to be updated if the exchange rate fluctuate more than 10%.             '
    ws['B14'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['B14'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws['B14'].fill = PatternFill(fill_type='solid', start_color=None, end_color=None)
    ws['B14'].border = Border(top=Side(style='thin'))

    # 设置 C14 单元格样式
    ws['C14'].font = Font(name='宋体', size=11.0)
    ws['C14'].border = Border(top=Side(style='thin'))

    # 设置 D14 单元格样式
    ws['D14'].font = Font(name='宋体', size=11.0)
    ws['D14'].border = Border(top=Side(style='thin'))

    # 设置 E14 单元格样式
    ws['E14'].value = 'SUBTOTAL'
    ws['E14'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E14'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E14'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F14 单元格样式
    ws['F14'].value = '=SUM(F1:F13)'
    ws['F14'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F14'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F14'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A15 单元格样式
    ws['A15'].font = Font(name='宋体', size=11.0)

    # 设置 B15 单元格样式
    ws['B15'].font = Font(name='宋体', size=11.0)

    # 设置 C15 单元格样式
    ws['C15'].font = Font(name='宋体', size=11.0)

    # 设置 D15 单元格样式
    ws['D15'].font = Font(name='宋体', size=11.0)

    # 设置 E15 单元格样式
    ws['E15'].value = 'FREIGHT'
    ws['E15'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E15'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E15'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F15 单元格样式
    ws['F15'].value = '=ROUND(4830*1.1/6.5,0)'
    ws['F15'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F15'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F15'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A16 单元格样式
    ws['A16'].font = Font(name='宋体', size=11.0)

    # 设置 B16 单元格样式
    ws['B16'].font = Font(name='宋体', size=11.0)

    # 设置 C16 单元格样式
    ws['C16'].font = Font(name='宋体', size=11.0)

    # 设置 D16 单元格样式
    ws['D16'].font = Font(name='宋体', size=11.0)

    # 设置 E16 单元格样式
    ws['E16'].value = 'OTHER'
    ws['E16'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E16'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E16'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F16 单元格样式
    ws['F16'].value = 0
    ws['F16'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F16'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F16'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A17 单元格样式
    ws['A17'].font = Font(name='宋体', size=11.0)

    # 设置 B17 单元格样式
    ws['B17'].font = Font(name='宋体', size=11.0)

    # 设置 C17 单元格样式
    ws['C17'].font = Font(name='宋体', size=11.0)

    # 设置 D17 单元格样式
    ws['D17'].font = Font(name='宋体', size=11.0)

    # 设置 E17 单元格样式
    ws['E17'].value = 'TOTAL'
    ws['E17'].font = Font(name='Microsoft YaHei', size=10.0, bold=True, italic=True, color='FF002060')
    ws['E17'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['E17'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 F17 单元格样式
    ws['F17'].value = '=SUM(F14:F15)'
    ws['F17'].font = Font(name='Microsoft YaHei', size=10.0)
    ws['F17'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws['F17'].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 设置 A18 单元格样式
    ws['A18'].font = Font(name='宋体', size=11.0)

    # 设置 B18 单元格样式
    ws['B18'].font = Font(name='宋体', size=11.0)

    # 设置 C18 单元格样式
    ws['C18'].font = Font(name='宋体', size=11.0)

    # 设置 D18 单元格样式
    ws['D18'].font = Font(name='宋体', size=11.0)

    # 设置 E18 单元格样式
    ws['E18'].font = Font(name='宋体', size=11.0)

    # 设置 F18 单元格样式
    ws['F18'].font = Font(name='宋体', size=11.0)

    # 设置 A19 单元格样式
    ws['A19'].font = Font(name='宋体', size=11.0)
    ws['A19'].alignment = Alignment(vertical='center')
    ws['A19'].fill = PatternFill(fill_type='solid', start_color=None, end_color=None)

    # 设置 B19 单元格样式
    ws['B19'].font = Font(name='宋体', size=11.0)

    # 设置 C19 单元格样式
    ws['C19'].font = Font(name='宋体', size=11.0)

    # 设置 D19 单元格样式
    ws['D19'].font = Font(name='宋体', size=11.0)

    # 设置 E19 单元格样式
    ws['E19'].font = Font(name='宋体', size=11.0)

    # 设置 F19 单元格样式
    ws['F19'].font = Font(name='宋体', size=11.0)

    # 设置 A20 单元格样式
    ws['A20'].font = Font(name='宋体', size=11.0)
    ws['A20'].alignment = Alignment(vertical='center')
    ws['A20'].fill = PatternFill(fill_type='solid', start_color=None, end_color=None)

    # 设置 B20 单元格样式
    ws['B20'].font = Font(name='宋体', size=11.0)

    # 设置 C20 单元格样式
    ws['C20'].font = Font(name='宋体', size=11.0)

    # 设置 D20 单元格样式
    ws['D20'].font = Font(name='宋体', size=11.0)

    # 设置 E20 单元格样式
    ws['E20'].font = Font(name='宋体', size=11.0)

    # 设置 F20 单元格样式
    ws['F20'].font = Font(name='宋体', size=11.0)

    # 设置 A21 单元格样式
    ws['A21'].font = Font(name='宋体', size=11.0)

    # 设置 B21 单元格样式
    ws['B21'].font = Font(name='宋体', size=11.0)

    # 设置 C21 单元格样式
    ws['C21'].font = Font(name='宋体', size=11.0)

    # 设置 D21 单元格样式
    ws['D21'].font = Font(name='宋体', size=11.0)

    # 设置 E21 单元格样式
    ws['E21'].font = Font(name='宋体', size=11.0)

    # 设置 F21 单元格样式
    ws['F21'].font = Font(name='宋体', size=11.0)

    # 设置 A22 单元格样式
    ws['A22'].font = Font(name='宋体', size=11.0)

    # 设置 B22 单元格样式
    ws['B22'].font = Font(name='宋体', size=11.0)

    # 设置 C22 单元格样式
    ws['C22'].font = Font(name='宋体', size=11.0)

    # 设置 D22 单元格样式
    ws['D22'].font = Font(name='宋体', size=11.0)

    # 设置 E22 单元格样式
    ws['E22'].font = Font(name='宋体', size=11.0)

    # 设置 F22 单元格样式
    ws['F22'].font = Font(name='宋体', size=11.0)

