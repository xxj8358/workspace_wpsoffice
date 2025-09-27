#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证按照科目名称排序的功能是否正确
"""

import os
import sys
import traceback


def verify_name_sorting():
    """验证按照科目名称排序的功能"""
    print("开始验证按照余额表1的F列科目名称排序功能...")
    
    try:
        # 导入check_balance_sheet模块，获取科目名称顺序
        from check_balance_sheet import get_account_sorting
        sorted_names = get_account_sorting()
        
        print(f"\n从余额表1获取的科目名称顺序（前20个）:")
        for i, name in enumerate(sorted_names[:20], 1):
            print(f"{i}. {name}")
        
        # 模拟几个工作表名称用于测试排序
        test_sheet_names = [
            "主营业务收入", "主营业务税金及附加", "其他应付款", 
            "应交税金", "应付职工", "测试科目1", "测试科目2"
        ]
        
        print(f"\n待排序的测试工作表名称:")
        print(test_sheet_names)
        
        # 按照余额表中的科目名称顺序排序
        sheet_to_position = {}
        for sheet_name in test_sheet_names:
            # 查找工作表名称在余额表1中的位置
            found = False
            for i, balance_name in enumerate(sorted_names):
                if balance_name in sheet_name or sheet_name in balance_name:
                    sheet_to_position[sheet_name] = i
                    found = True
                    break
            if not found:
                sheet_to_position[sheet_name] = float('inf')
        
        # 按照余额表中的位置排序
        sorted_sheets = sorted(test_sheet_names, key=lambda x: sheet_to_position.get(x, float('inf')))
        
        print(f"\n按照余额表1科目名称排序后的工作表顺序:")
        for i, sheet_name in enumerate(sorted_sheets, 1):
            matched_name = None
            pos = sheet_to_position[sheet_name]
            if pos < float('inf'):
                matched_name = sorted_names[pos]
            print(f"{i}. {sheet_name} (匹配位置: {pos}, 匹配名称: {matched_name if matched_name else '未匹配'})")
        
        # 验证test_copy_table.py中的排序逻辑
        print(f"\n验证test_copy_table.py中的排序逻辑...")
        
        # 尝试导入test_copy_table模块，检查是否存在错误
        try:
            import test_copy_table
            print("成功导入test_copy_table模块，无语法错误")
        except ImportError as e:
            print(f"导入test_copy_table模块失败: {str(e)}")
        except Exception as e:
            print(f"导入test_copy_table模块时发生其他错误: {str(e)}")
            traceback.print_exc()
        
        print("\n验证完成。请运行主程序test_copy_table.py来实际测试排序功能。")
        
    except Exception as e:
        print(f"验证过程中发生错误: {str(e)}")
        traceback.print_exc()


if __name__ == "__main__":
    verify_name_sorting()