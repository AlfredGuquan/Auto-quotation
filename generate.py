import re
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image
from datetime import datetime
import os
import math
# 在文件顶部导入
from template_styles import apply_template_styles


# 在这里添加从template_styles.py复制的apply_template_styles函数
# 应用模板样式
def apply_template_styles(ws):
    # 定义常用样式
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    # 应用单元格样式
    # 设置 A1 单元格样式
    # ... 这里是从template_styles.py复制的所有样式代码 ...

class ERobPriceCalculator:
    def __init__(self):
        # 基础价格表 - 根据型号和数量范围
        self.price_table = {
            'eRob70F': {
                'retail': 1032.00,
                '5-10': 938.00,
                '10-99': 893.00,
                '100-499': 851.00,
                '500-999': 810.00,
                '1000-1999': 772.00,
                '2000+': 670.00
            },
            'eRob70H': {
                'retail': 1011.00,
                '5-10': 919.00,
                '10-99': 876.00,
                '100-499': 834.00,
                '500-999': 794.00,
                '1000-1999': 756.00,
                '2000+': 657.00
            },
            'eRob80H': {
                'retail': 1072.00,
                '5-10': 974.00,
                '10-99': 928.00,
                '100-499': 884.00,
                '500-999': 842.00,
                '1000-1999': 802.00,
                '2000+': 696.00
            },
            'eRob90H': {
                'retail': 1215.00,
                '5-10': 1104.00,
                '10-99': 1052.00,
                '100-499': 1002.00,
                '500-999': 954.00,
                '1000-1999': 908.00,
                '2000+': 789.00
            },
            'eRob110H': {
                'retail': 1272.00,
                '5-10': 1156.00,
                '10-99': 1101.00,
                '100-499': 1049.00,
                '500-999': 999.00,
                '1000-1999': 951.00,
                '2000+': 826.00
            },
            'eRob142H': {
                'retail': 1338.00,
                '5-10': 1216.00,
                '10-99': 1158.00,
                '100-499': 1103.00,
                '500-999': 1050.00,
                '1000-1999': 1000.00,
                '2000+': 868.00
            },
            'eRob170H': {
                'retail': 2008.00,
                '5-10': 1826.00,
                '10-99': 1739.00,
                '100-499': 1656.00,
                '500-999': 1577.00,
                '1000-1999': 1502.00,
                '2000+': 1304.00
            }
        }
        
        # 选项价格表
        self.option_prices = {
            'without_brake': -15,
            'multiturn': 23,
            'ethercat': 23,
            'high_precision': 46,
            't_type': {
                70: 77,
                80: 77,
                90: 77,
                110: 138,
                142: 200,
                170: 200
            },
            'version': {
                'V6': {
                    80: 78,
                    90: 180,
                    110: 92
                }
            },
            'leaderdrive_gear': {
                70: 83,
                80: 111,
                90: 129,
                110: 138,
                142: 258,
                170: 277
            }
        }
        
        # 默认减速比映射
        self.default_ratio = {
            70: 100,
            80: 120,
            90: 120,
            110: 120,
            142: 120,
            170: 120
        }
        
        # 产品重量（kg）
        self.weights = {
            70: 0.85,
            80: 1.03,
            90: 1.62,
            110: 2.68,
            142: 4.50,
            170: 7.20,
            'eRob Universal Accessories Kit': 0.2,
            'eLine - RJ45 ECAT -30': 0.03,
            'eRob to PC Connector': 0.2
        }
        
        # 配件价格
        self.accessories = {
            'eRob Universal Accessories Kit': 35,
            'eLine - RJ45 ECAT -30': 11,
            'eRob to PC Connector': 66
        }
        
        # 产品图片路径
        self.images = {
            # eRob系列产品图片
            'eRob70F': 'images/70F.png',
            'eRob70H': 'images/70I.png',
            'eRob80H': 'images/80I.png',
            'eRob90H': 'images/90I.png',
            'eRob110H': 'images/110I.png',
            'eRob142H': 'images/142I.png',
            'eRob170H': 'images/170I.png',
            # 配件图片
            'eRob Universal Accessories Kit': 'images/Kit.png',
            'eLine - RJ45 ECAT -30': 'images/RJ45.png',
            'eRob to PC Connector': 'images/PC.png'
        }
    
    def normalize_model_code(self, model_code):
        """标准化型号编码，补全缺失部分"""
        # 处理特殊版本标记
        version = ""
        if '[' in model_code:
            base_code, version = model_code.split('[')
            version = version.strip(']')
            model_code = base_code.strip()
        
        # 如果不是以eRob开头，添加eRob前缀
        if not model_code.startswith('eRob'):
            # 检查是否以数字开头（如80H120I-BHM-18CN）
            if model_code[0].isdigit():
                model_code = 'eRob' + model_code
        
        # 解析型号主体部分
        parts = model_code.split('-')
        base_info = parts[0]
        
        # 提取直径和齿轮类型
        if 'eRob' in base_info:
            diameter_start = base_info.find('eRob') + 4
            diameter_end = diameter_start
            while diameter_end < len(base_info) and base_info[diameter_end].isdigit():
                diameter_end += 1
            
            if diameter_end > diameter_start:
                diameter = int(base_info[diameter_start:diameter_end])
                gear_type = base_info[diameter_end] if diameter_end < len(base_info) else 'H'  # 默认H型
            else:
                # 无法提取直径，使用默认值
                diameter = 80
                gear_type = 'H'
        else:
            # 无法识别eRob，使用默认值
            diameter = 80
            gear_type = 'H'
        
        # 检查是否包含减速比
        ratio_match = re.search(r'(\d+)[IT]', base_info)
        if not ratio_match:
            # 添加默认减速比
            default_ratio = self.default_ratio.get(diameter, 120)
            if 'I' in base_info or 'T' in base_info:
                # 已有形状标识，在其前添加减速比
                form_type_pos = base_info.find('I') if 'I' in base_info else base_info.find('T')
                base_info = base_info[:form_type_pos] + str(default_ratio) + base_info[form_type_pos:]
            else:
                # 无形状标识，添加减速比和默认I型
                base_info = base_info + str(default_ratio) + 'I'
        
        # 确保有形状标识（I或T）
        if 'I' not in base_info and 'T' not in base_info:
            base_info = base_info + 'I'  # 默认I型
        
        # 重建完整型号
        normalized_code = base_info
        if len(parts) > 1:
            normalized_code += '-' + parts[1]
        else:
            normalized_code += '-BHM'  # 默认配置
        
        if len(parts) > 2:
            normalized_code += '-' + parts[2]
        else:
            normalized_code += '-18CN'  # 默认接口
        
        # 添加版本标记（如果有）
        if version:
            normalized_code += f"[{version}]"
        
        return normalized_code
    
    def parse_model_code(self, model_code):
        """解析型号编码，提取各部分信息"""
        # 标准化型号
        normalized_model = self.normalize_model_code(model_code)
        
        # 处理特殊版本标记
        version = ""
        if '[' in normalized_model:
            base_model, version_part = normalized_model.split('[')
            version = version_part.strip(']')
            normalized_model = base_model.strip()
        
        # 解析型号主体部分
        parts = normalized_model.split('-')
        base_info = parts[0]
        
        # 提取直径和齿轮类型
        diameter_start = base_info.find('eRob') + 4
        diameter_end = diameter_start
        while diameter_end < len(base_info) and base_info[diameter_end].isdigit():
            diameter_end += 1
        
        diameter = int(base_info[diameter_start:diameter_end])
        gear_type = base_info[diameter_end]
        
        # 提取减速比
        ratio_match = re.search(r'(\d+)[IT]', base_info)
        ratio = int(ratio_match.group(1)) if ratio_match else self.default_ratio.get(diameter, 120)
        
        # 提取形状类型
        form_type_match = re.search(r'\d+([IT])', base_info)
        form_type = form_type_match.group(1) if form_type_match else 'I'
        
        # 提取配置和接口
        config = parts[1] if len(parts) > 1 else 'BHM'
        interface = parts[2] if len(parts) > 2 else '18CN'
        
        # 确定基本型号
        base_model = f"eRob{diameter}{gear_type}"
        
        # 检查形状类型是否为T型
        is_t_type = form_type == 'T'
        
        # 检查接口是否为EtherCAT
        is_ethercat = 'ET' in interface
        
        # 检查多圈和高精度
        is_multiturn = 'M' in config
        is_high_precision = 'H' in config and 'P' in config
        
        # 检查是否有刹车
        has_brake = 'B' in config
        
        # 返回解析结果
        result = {
            'model_code': normalized_model,
            'full_model': normalized_model + (f"[{version}]" if version else ""),
            'base_model': base_model,
            'diameter': diameter,
            'gear_type': gear_type,
            'ratio': ratio,
            'form_type': form_type,
            'config': config,
            'interface': interface,
            'is_t_type': is_t_type,
            'is_ethercat': is_ethercat,
            'is_multiturn': is_multiturn,
            'is_high_precision': is_high_precision,
            'has_brake': has_brake,
            'version': version
        }
        
        return result
    
    def get_price_range_index(self, quantity):
        """根据数量确定价格范围"""
        if quantity < 5:
            return 'retail'
        elif quantity < 10:
            return '5-10'
        elif quantity < 100:
            return '10-99'
        elif quantity < 500:
            return '100-499'
        elif quantity < 1000:
            return '500-999'
        elif quantity < 2000:
            return '1000-1999'
        else:
            return '2000+'
    
    def calculate_price(self, model_code, quantity=1):
        """计算单个产品的价格"""
        parsed_model = self.parse_model_code(model_code)
        
        # 获取基本型号价格
        base_model = parsed_model['base_model']
        price_range = self.get_price_range_index(quantity)
        
        if base_model in self.price_table:
            base_price = self.price_table[base_model][price_range]
        else:
            # 未找到型号，使用默认价格
            base_price = 1000.00
        
        # 计算附加选项价格
        additional_price = 0.0
        
        # T型选项价格
        if parsed_model['is_t_type']:
            t_price = self.option_prices['t_type'].get(parsed_model['diameter'], 0)
            additional_price += t_price
        
        # 特殊版本价格（如V6）
        if parsed_model['version'] == 'V6':
            # 只有特定型号支持V6版本
            v6_supported_diameters = self.option_prices['version']['V6'].keys()
            if parsed_model['diameter'] in v6_supported_diameters:
                v6_price = self.option_prices['version']['V6'][parsed_model['diameter']]
                additional_price += v6_price
        
        # EtherCAT选项
        if parsed_model['is_ethercat']:
            additional_price += self.option_prices['ethercat']
        
        # 多圈选项
        if parsed_model['is_multiturn']:
            additional_price += self.option_prices['multiturn']
        
        # 高精度选项
        if parsed_model['is_high_precision']:
            additional_price += self.option_prices['high_precision']
        
        # 无刹车时减价
        if not parsed_model['has_brake']:
            additional_price += self.option_prices['without_brake']
        
        # 计算最终价格
        final_price = base_price + additional_price
        
        # 获取产品重量
        weight = self.weights.get(parsed_model['diameter'], 1.0)
        
        return {
            'model': parsed_model['full_model'],
            'quantity': quantity,
            'base_price': base_price,
            'additional_price': additional_price,
            'unit_price': final_price,
            'total_price': final_price * quantity,
            'weight': weight,
            'parsed_info': parsed_model
        }
    
    def calculate_batch(self, model_codes, quantities, customer_info):
        """计算多个产品的价格"""
        if not model_codes:
            return {"items": [], "subtotal": 0, "freight": 0, "grand_total": 0, "customer_info": customer_info}
        
        results = []
        
        # 计算每个产品的价格
        for i, model_code in enumerate(model_codes):
            qty = quantities[i] if i < len(quantities) else 1
            item_result = self.calculate_price(model_code, qty)
            results.append(item_result)
        
        # 统计eRob总数
        erob_total_qty = sum(item['quantity'] for item in results)
        
        # 添加配件 - eRob Universal Accessories Kit
        kit_qty = erob_total_qty
        if kit_qty > 0:
            kit_price = self.accessories['eRob Universal Accessories Kit']
            results.append({
                'model': 'eRob Universal Accessories Kit',
                'quantity': kit_qty,
                'base_price': kit_price,
                'additional_price': 0,
                'unit_price': kit_price,
                'total_price': kit_price * kit_qty,
                'weight': self.weights['eRob Universal Accessories Kit'],
                'parsed_info': None
            })
        
        # 检查是否有EtherCAT产品，添加RJ45连接器
        has_ethercat = any(item['parsed_info'] and item['parsed_info']['is_ethercat'] for item in results if item['parsed_info'])
        if has_ethercat:
            rj45_qty = erob_total_qty
            rj45_price = self.accessories['eLine - RJ45 ECAT -30']
            results.append({
                'model': 'eLine - RJ45 ECAT -30',
                'quantity': rj45_qty,
                'base_price': rj45_price,
                'additional_price': 0,
                'unit_price': rj45_price,
                'total_price': rj45_price * rj45_qty,
                'weight': self.weights['eLine - RJ45 ECAT -30'],
                'parsed_info': None
            })
        
        # 添加PC连接器（默认数量为1）
        pc_connector_price = self.accessories['eRob to PC Connector']
        pc_connector_qty = 5 if erob_total_qty > 20 else 1  # 如果eRob数量大于20，提供5个PC连接器
        results.append({
            'model': 'eRob to PC Connector',
            'quantity': pc_connector_qty,
            'base_price': pc_connector_price,
            'additional_price': 0,
            'unit_price': pc_connector_price,
            'total_price': pc_connector_price * pc_connector_qty,
            'weight': self.weights['eRob to PC Connector'],
            'parsed_info': None
        })
        
        # 计算总价和合计
        subtotal = sum(item['total_price'] for item in results)
        
        # 计算运费
        total_weight = sum(item['weight'] * item['quantity'] for item in results)
        freight = round(4830 * 1.1 / 6.5, 0)  # 使用固定公式计算运费
        
        # 计算总计
        grand_total = subtotal + freight
        
        return {
            "items": results,
            "subtotal": subtotal,
            "freight": freight,
            "grand_total": grand_total,
            "customer_info": customer_info
        }
    
    def export_to_quotation(self, results):
        """导出报价单到Excel文件"""
        # 创建工作簿
        wb = Workbook()
        ws = wb.active
        
        # 首先应用标准模板样式
        apply_template_styles(ws)
        
        # 获取当前日期
        today = datetime.now()
        date_str = today.strftime('%Y.%m.%d')
        
        # 生成报价单号
        quotation_number = today.strftime('%Y%m%d') + '01'  # 以年月日+序号01组成
        
        # 填充客户信息
        customer_info = results['customer_info']
        
        # 设置标题和客户信息（保留模板样式的基础上修改内容）
        # 标题已经在apply_template_styles中设置了
        
        # 客户信息 - To部分
        to_text = f"To:  {customer_info.get('name', '')}\n"
        to_text += f"Company: {customer_info.get('company', '')}\n"
        to_text += f"Street Address: {customer_info.get('address', '')}\n"
        to_text += f"City, ST  ZIP Code: {customer_info.get('city', '')}, {customer_info.get('zip', '')}\n"
        to_text += f"Phone: {customer_info.get('phone', '')}"
        ws['A2'].value = to_text
        
        # 日期和报价号部分
        date_text = f"Date：{date_str}\n"
        date_text += f"Quotation #: {quotation_number}\n"
        date_text += f"Street Address：Fuyuan 1st, Fuhai City, ZIP Code：Bao'an，Shen Zhen, 518103\n"
        date_text += f"Phone：+86 18922807806"
        ws['D2'].value = date_text
        
        # 填充表头 (A6:F6) - 已在apply_template_styles中设置
        
        # 开始填充产品列表
        row = 8  # 从第8行开始插入产品
        
        for i, item in enumerate(results['items']):
            # 插入产品图片
            model_code = item['model']
            if model_code in self.images and os.path.exists(self.images[model_code]):
                try:
                    img = Image(self.images[model_code])
                    # 调整图片大小
                    img.width = 60
                    img.height = 60
                    # 设置图片在单元格中的位置
                    ws.add_image(img, f'A{row}')
                except Exception as e:
                    print(f"插入图片时出错: {e}")
            
            # 填充产品信息，保持单元格样式
            ws.cell(row=row, column=2).value = item['model']  # B列：型号
            ws.cell(row=row, column=3).value = item['quantity']  # C列：数量
            ws.cell(row=row, column=4).value = item['weight']  # D列：重量
            ws.cell(row=row, column=5).value = item['unit_price']  # E列：单价
            
            # F列：总价，使用货币格式
            cell = ws.cell(row=row, column=6)
            cell.value = f"$ {item['total_price']:,.2f}"
            cell.alignment = Alignment(horizontal='right')
            
            row += 1
        
        # 设置备注信息
        remarks_row = row + 1
        ws.cell(row=remarks_row, column=2).value = f"1. Price term: DAP {customer_info.get('country', 'Australia')}"
        ws.cell(row=remarks_row+1, column=2).value = "2. Payment term: T/T, 100% advance payment."
        ws.cell(row=remarks_row+2, column=2).value = "3. Leading time: 12 working days after received payment."
        ws.cell(row=remarks_row+3, column=2).value = "4.The price needs to be updated if the exchange rate fluctuate more than 10%."
        
        # 设置合计金额
        ws.cell(row=remarks_row, column=5).value = "SUBTOTAL"
        ws.cell(row=remarks_row, column=6).value = f"$ {results['subtotal']:,.2f}"
        ws.cell(row=remarks_row, column=6).alignment = Alignment(horizontal='right')
        
        ws.cell(row=remarks_row+1, column=5).value = "FREIGHT"
        ws.cell(row=remarks_row+1, column=6).value = f"$ {results['freight']:,.2f}"
        ws.cell(row=remarks_row+1, column=6).alignment = Alignment(horizontal='right')
        
        ws.cell(row=remarks_row+2, column=5).value = "OTHER"
        ws.cell(row=remarks_row+2, column=6).value = "$ -"
        ws.cell(row=remarks_row+2, column=6).alignment = Alignment(horizontal='right')
        
        ws.cell(row=remarks_row+3, column=5).value = "TOTAL"
        ws.cell(row=remarks_row+3, column=6).value = f"$ {results['grand_total']:,.2f}"
        ws.cell(row=remarks_row+3, column=6).alignment = Alignment(horizontal='right')
        
        # 保存文件
        customer_name = customer_info.get('name', 'Customer').replace(' ', '_')
        filename = f"ZeroErr_Quotation_{customer_name}_{today.strftime('%Y%m%d')}.xlsx"
        wb.save(filename)
        return filename

def main():
    calculator = ERobPriceCalculator()
    
    print("eRob产品报价单生成系统")
    print("请输入客户信息:")
    customer_name = input("客户姓名: ")
    customer_company = input("公司名称: ")
    customer_address = input("街道地址: ")
    customer_city = input("城市: ")
    customer_zip = input("邮编: ")
    customer_country = input("国家: ")
    customer_phone = input("电话: ")
    
    customer_info = {
        'name': customer_name,
        'company': customer_company,
        'address': customer_address,
        'city': customer_city,
        'zip': customer_zip,
        'country': customer_country,
        'phone': customer_phone
    }
    
    print("\n请输入产品型号和数量，多个产品请用逗号分隔")
    print("例如: eRob90H160I-BHM-18ET[V6],80H120I-BHM-18CN")
    print("数量例如: 6,8")
    
    model_input = input("产品型号: ")
    quantity_input = input("对应数量: ")
    
    model_codes = [code.strip() for code in model_input.split(',')]
    quantities = [int(qty.strip()) for qty in quantity_input.split(',')]
    
    # 确保数量与型号数量匹配
    while len(quantities) < len(model_codes):
        quantities.append(1)  # 默认数量为1
    
    results = calculator.calculate_batch(model_codes, quantities, customer_info)
    
    # 打印结果
    print("\n计算结果:")
    print("-" * 80)
    print(f"{'序号':<5}{'产品型号':<30}{'数量':<8}{'单价':<12}{'总价':<12}")
    print("-" * 80)
    
    for i, item in enumerate(results['items'], 1):
        print(f"{i:<5}{item['model']:<30}{item['quantity']:<8}{item['unit_price']:<12.2f}{item['total_price']:<12.2f}")
    
    print("-" * 80)
    print(f"{'小计':<43}{results['subtotal']:<12.2f}")
    print(f"{'运费':<43}{results['freight']:<12.2f}")
    print(f"{'总计':<43}{results['grand_total']:<12.2f}")
    
    # 导出报价单
    excel_file = calculator.export_to_quotation(results)
    print(f"\n报价单已生成: {excel_file}")

if __name__ == "__main__":
    main()