#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import xlrd

# 模拟code_sheet_mapping数据，根据截图中的工作表名称创建
# 注意：这里只是一个示例，实际应该从Excel文件中读取工作表名称
code_sheet_mapping = [
    ('1001', 'Sheet1'),  # 假设Sheet1对应科目编码1001
    ('1002', 'Sheet2'),  # 假设Sheet2对应科目编码1002
    ('2171', '主营业务收入'),  # 主营业务收入对应科目编码2171
    ('其他科目1', '主营业务税金及附加'),
    ('其他科目2', '其他应付款'),
    ('其他科目3', '应交税金'),
    ('其他科目4', '应付职工'),
    # 根据实际情况添加更多的映射
]

def get_account_sorting_with_details():
    """
    从余额表1获取科目编码、名称和位置的详细信息
    """
    try:
        # 文件路径
        file_path = r'D:\old\e\双录日常问题解决方案\个人\汉和2024序时账.xls'
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return []
        
        # 加载工作簿
        book = xlrd.open_workbook(file_path)
        
        # 查找余额表1工作表
        balance_sheet = None
        for name in book.sheet_names():
            if '余额表1' in name:
                balance_sheet = book.sheet_by_name(name)
                break
        
        if not balance_sheet:
            print("未找到余额表1工作表")
            return []
        
        # 从列3（索引2）提取科目编码，从列6（索引5）提取科目名称
        account_data = []
        for row_idx in range(1, balance_sheet.nrows):
            code_value = balance_sheet.cell_value(row_idx, 2)  # 索引从0开始，所以列3对应的索引是2
            name_value = balance_sheet.cell_value(row_idx, 5)  # 索引从0开始，所以列6对应的索引是5
            
            if code_value is not None and code_value != '' and name_value is not None and name_value != '':
                account_data.append({
                    'code': str(code_value).strip(),
                    'name': str(name_value).strip(),
                    'index': row_idx - 1  # 记录在余额表中的位置索引
                })
        
        return account_data
    except Exception as e:
        print(f"获取余额表信息时出错: {str(e)}")
        return []

def final_verification():
    """
    最终验证排序逻辑是否符合用户的期望
    """
    # 获取余额表1中的详细信息
    account_data = get_account_sorting_with_details()
    
    # 创建科目编码到详细信息的映射
    code_to_details = {item['code']: item for item in account_data}
    
    # 打印余额表1中的所有科目信息
    print("余额表1中的所有科目信息:")
    for idx, item in enumerate(account_data, 1):
        print(f"{idx}. 索引: {item['index']}, 编码: {item['code']}, 名称: {item['name']}")
    
    # 创建科目编码到工作表名称的映射字典
    code_to_sheet = {code: sheet_name for code, sheet_name in code_sheet_mapping}
    
    # 创建科目名称到编码的映射
    name_to_code = {item['name']: item['code'] for item in account_data}
    
    # 创建工作表名称到余额表索引的映射，并添加额外的信息以处理相同索引的情况
    sheet_to_balance_info = {}
    
    # 遍历所有工作表名称，尝试找到对应的余额表信息
    for code, sheet_name in code_sheet_mapping:
        # 检查科目编码是否在余额表1中
        if code in code_to_details:
            # 获取科目在余额表1中的信息
            details = code_to_details[code]
            sheet_to_balance_info[sheet_name] = {
                'index': details['index'],
                'code': details['code'],
                'name': details['name'],
                'type': 'code_match'  # 标记为通过编码匹配
            }
        # 如果科目编码不在余额表1中，尝试通过名称匹配
        else:
            found = False
            for balance_code, balance_name in name_to_code.items():
                if balance_name in sheet_name or sheet_name in balance_name:
                    if balance_code in code_to_details:
                        details = code_to_details[balance_code]
                        sheet_to_balance_info[sheet_name] = {
                            'index': details['index'],
                            'code': details['code'],
                            'name': details['name'],
                            'type': 'name_match'  # 标记为通过名称匹配
                        }
                        found = True
                        break
            if not found:
                sheet_to_balance_info[sheet_name] = {
                    'index': float('inf'),
                    'code': code,
                    'name': '未知',
                    'type': 'no_match'  # 标记为未匹配
                }
        
    # 打印工作表名称与余额表信息的匹配结果
    print("\n工作表名称与余额表信息的匹配结果:")
    for sheet_name, info in sheet_to_balance_info.items():
        match_type = "通过编码匹配" if info['type'] == 'code_match' else "通过名称匹配" if info['type'] == 'name_match' else "未匹配"
        print(f"{sheet_name}: 余额表索引 = {info['index']}, 科目编码 = {info['code']}, 科目名称 = {info['name']}, 匹配方式 = {match_type}")
    
    # 按照余额表1的索引排序工作表，如果索引相同，则按照工作表名称排序
    sorted_sheets_by_balance = sorted(code_sheet_mapping, key=lambda x: (sheet_to_balance_info[x[1]]['index'], x[1]))
    
    # 将排序后的工作表名称添加到结果列表中
    sorted_sheets = [sheet_name for _, sheet_name in sorted_sheets_by_balance]
    
    # 打印最终排序结果
    print("\n最终排序后的工作表顺序:")
    for i, sheet_name in enumerate(sorted_sheets, 1):
        print(f"{i}. {sheet_name}")
        
    # 特别检查截图中出现的工作表顺序
    screenshot_order = ['Sheet1', 'Sheet2', '主营业务收入', '主营业务税金及附加', '其他应付款', '应交税金', '应付职工']
    print("\n截图中的工作表顺序:")
    for i, sheet_name in enumerate(screenshot_order, 1):
        print(f"{i}. {sheet_name}")
    
    # 比较排序结果与截图顺序
    print("\n比较排序结果与截图顺序:")
    for i, sheet_name in enumerate(sorted_sheets):
        if i < len(screenshot_order) and sheet_name == screenshot_order[i]:
            print(f"位置{i+1}: {sheet_name} - 顺序一致")
        else:
            if i < len(screenshot_order):
                print(f"位置{i+1}: {sheet_name} vs {screenshot_order[i]} - 顺序不一致")
            else:
                print(f"位置{i+1}: {sheet_name} - 排序结果中多余的工作表")
    
    # 检查关键工作表的位置
    key_sheets = ['主营业务收入', '主营业务税金及附加', '其他应付款', '应交税金', '应付职工']
    print("\n关键工作表在排序结果中的位置:")
    for sheet_name in key_sheets:
        if sheet_name in sorted_sheets:
            pos = sorted_sheets.index(sheet_name) + 1
            info = sheet_to_balance_info[sheet_name]
            print(f"{sheet_name}: 排序位置 = {pos}, 余额表索引 = {info['index']}, 余额表位置 = {info['index']+1}")
        else:
            print(f"{sheet_name}: 未找到")
    
    # 生成最终的排序建议
    print("\n最终排序建议:")
    for i, sheet_name in enumerate(sorted_sheets, 1):
        print(f"{i}. {sheet_name}")
    
    return sorted_sheets

if __name__ == "__main__":
    print("开始最终验证排序逻辑...")
    final_verification()
    print("验证完成")