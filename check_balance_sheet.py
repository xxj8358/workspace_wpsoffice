import os
import xlrd

# 设置文件路径
file_path = r'D:\old\e\双录日常问题解决方案\个人\汉和2024序时账.xls'

def get_account_sorting():
    """
    从余额表1获取科目名称的排序顺序
    返回：按余额表1中原始顺序排列的科目名称列表
    """
    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        return []
    
    try:
        # 加载工作簿
        book = xlrd.open_workbook(file_path)
        
        # 检查是否有余额表1工作表
        sheet_names = book.sheet_names()
        balance_sheet = None
        for name in sheet_names:
            if '余额表1' in name:
                balance_sheet = book.sheet_by_name(name)
                break
        
        if balance_sheet:
            # 直接指定科目名称列索引（F列，索引从0开始，所以列6对应的索引是5）
            name_column_index = 5  # 索引从0开始，所以列6对应的索引是5
            
            # 收集科目名称（从第二行开始读取数据）
            account_names = []
            for row_idx in range(1, balance_sheet.nrows):
                name_value = balance_sheet.cell_value(row_idx, name_column_index)
                if name_value is not None and name_value != '':
                    name_str = str(name_value).strip()
                    account_names.append(name_str)
            
            return account_names
        else:
            print("文件中未找到'余额表1'工作表")
            return []
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return []

# 主程序执行
if __name__ == "__main__":
    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
    else:
        try:
            # 加载工作簿
            book = xlrd.open_workbook(file_path)
            
            # 检查是否有余额表1工作表
            sheet_names = book.sheet_names()
            balance_sheet = None
            for name in sheet_names:
                if '余额表1' in name:
                    balance_sheet = book.sheet_by_name(name)
                    break
            
            if balance_sheet:
                print(f'找到余额表1工作表，共有{balance_sheet.nrows}行，{balance_sheet.ncols}列')
                
                # 获取表头
                headers = [balance_sheet.cell_value(0, col_idx) for col_idx in range(balance_sheet.ncols)]
                print('余额表1列标题:')
                for i, header in enumerate(headers):
                    print(f"列{i+1}: {header}")
                
                # 查找科目编码列和科目名称列
                code_column_index = -1
                name_column_index = -1
                
                # 直接指定列索引，因为从输出看科目编码是列3，科目名称是列6
                code_column_index = 2  # 索引从0开始，所以列3对应的索引是2
                name_column_index = 5  # 索引从0开始，所以列6对应的索引是5
                
                print(f'识别的科目编码列: 列{code_column_index+1}')
                print(f'识别的科目名称列: 列{name_column_index+1}')
                
                # 如果同时找到科目编码列和科目名称列
                if code_column_index != -1 and name_column_index != -1:
                    print('\n科目编码和名称映射关系:')
                    code_name_mapping = {}
                    
                    # 从第二行开始读取数据
                    for row_idx in range(1, min(101, balance_sheet.nrows)):
                        code_value = balance_sheet.cell_value(row_idx, code_column_index)
                        name_value = balance_sheet.cell_value(row_idx, name_column_index)
                        
                        if code_value is not None and code_value != '' and name_value is not None and name_value != '':
                            code_str = str(code_value).strip()
                            name_str = str(name_value).strip()
                            code_name_mapping[code_str] = name_str
                            # 只显示前20个
                            if row_idx <= 21:
                                print(f'{code_str}: {name_str}')
                    
                    print(f'\n共找到{len(code_name_mapping)}个科目编码和名称的映射关系')
                    
                    # 尝试按科目编码数值排序
                    try:
                        print('\n按科目编码数值从小到大排序后的映射关系:')
                        sorted_codes = []
                        numeric_codes = []
                        non_numeric_codes = []
                        
                        for code in code_name_mapping.keys():
                            if code.replace('.', '', 1).isdigit():
                                numeric_codes.append((float(code), code))
                            else:
                                non_numeric_codes.append(code)
                        
                        # 对数字编码按数值排序
                        numeric_codes.sort(key=lambda x: x[0])
                        sorted_codes.extend([code for _, code in numeric_codes])
                        # 对非数字编码按字符串排序
                        sorted_codes.extend(sorted(non_numeric_codes))
                        
                        # 显示排序后的前20个
                        for i, code in enumerate(sorted_codes[:20]):
                            print(f'{code}: {code_name_mapping[code]}')
                    except Exception as e:
                        print(f'排序时出错: {e}')
                else:
                    print("未找到科目编码列或科目名称列")
            else:
                print("文件中未找到'余额表1'工作表")
        except Exception as e:
            print(f"读取文件时出错: {e}")

    print("\n程序执行完毕")