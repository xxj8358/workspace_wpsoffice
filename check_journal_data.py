import pandas as pd
import xlrd
import os

# 设置文件路径
file_path = r'D:\old\e\双录日常问题解决方案\个人\汉和2024序时账.xls'

# 检查文件是否存在
if not os.path.exists(file_path):
    print(f"文件不存在: {file_path}")
else:
    try:
        # 加载工作簿（使用xlrd读取.xls文件）
        book = xlrd.open_workbook(file_path)
        
        # 检查是否有序时账工作表
        sheet_names = book.sheet_names()
        journal_sheet = None
        for name in sheet_names:
            if '序时账' in name:
                journal_sheet = book.sheet_by_name(name)
                break
        
        if journal_sheet:
            # 获取表头
            headers = [journal_sheet.cell_value(0, col_idx) for col_idx in range(journal_sheet.ncols)]
            print('序时账列标题:')
            for i, header in enumerate(headers):
                print(f"列{i+1}: {header}")
            
            # 查找科目编码列
            code_column_index = -1
            code_column_name = None
            for col_idx in range(journal_sheet.ncols):
                header = headers[col_idx]
                if header and (any(keyword in str(header) for keyword in ['科目编码', '科目代码'])):
                    code_column_name = header
                    code_column_index = col_idx
                    break
            
            print(f'识别的科目编码列: "{code_column_name}" (列{code_column_index+1})')
            
            # 如果找到科目编码列，读取数据
            if code_column_index != -1:
                # 读取前100行的科目编码数据
                data = []
                for row_idx in range(1, min(101, journal_sheet.nrows)):  # 跳过表头行
                    cell_value = journal_sheet.cell_value(row_idx, code_column_index)
                    if cell_value is not None and cell_value != '':
                        data.append(str(cell_value).strip())
                
                # 获取唯一的科目编码
                unique_codes = list(set(data))
                print(f'前100行中的唯一科目编码 ({len(unique_codes)}个):')
                for code in unique_codes:
                    print(f'- {code}')
                
                # 尝试按数值排序（如果可能）
                try:
                    sorted_numeric = sorted([code for code in unique_codes if code.replace('.', '', 1).isdigit()], key=float)
                    sorted_non_numeric = sorted([code for code in unique_codes if not code.replace('.', '', 1).isdigit()])
                    sorted_codes = sorted_numeric + sorted_non_numeric
                    print('按数值从小到大排序后的科目编码:')
                    for code in sorted_codes:
                        print(f'- {code}')
                except Exception as e:
                    print(f'排序时出错: {e}')
        else:
            print("文件中未找到'序时账'工作表")
    except Exception as e:
        print(f"读取文件时出错: {e}")

print("\n程序执行完毕")