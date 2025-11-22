import pandas as pd
import os
import shutil
import tkinter as tk
from tkinter import filedialog
import sys

def select_excel_file():
    """打开文件选择对话框，让用户选择Excel文件"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 设置文件类型过滤
    file_types = [('Excel files', '*.xlsx;*.xls')]
    
    # 打开文件选择对话框
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=file_types
    )
    
    return file_path

def select_voucher_folder():
    """打开文件夹选择对话框，让用户选择凭证文件夹"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 打开文件夹选择对话框
    folder_path = filedialog.askdirectory(
        title="选择凭证文件夹"
    )
    
    return folder_path

def copy_voucher_files(file_path, base_dir, vouchers_dir):
    """根据Excel文件的H列凭证号信息，复制对应文件到创建的文件夹中"""
    try:
        # 检查凭证文件夹是否存在
        if not vouchers_dir:
            print("警告：未选择凭证文件夹")
            return 0, 0
            
        if not os.path.exists(vouchers_dir):
            print(f"警告：凭证文件夹不存在: {vouchers_dir}")
            print("请确认凭证文件夹路径是否正确")
            return 0, 0
        
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 获取C列和H列的数据
        # C列用于确定目标文件夹，H列用于确定凭证号
        c_column_index = ord('C') - 65
        h_column_index = ord('H') - 65
        
        # 检查列是否存在
        if c_column_index >= len(df.columns) or h_column_index >= len(df.columns):
            print("警告：Excel文件中找不到C列或H列")
            return 0, 0
        
        # 获取所有非空的行
        valid_rows = df.dropna(subset=[df.columns[c_column_index], df.columns[h_column_index]])
        
        copied_files_count = 0
        not_found_files_count = 0
        
        # 遍历每一行数据
        for _, row in valid_rows.iterrows():
            # 获取文件夹名称（C列数据）
            folder_name = str(row.iloc[c_column_index]).strip()
            # 替换不能在文件名中使用的字符
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                folder_name = folder_name.replace(char, '_')
            
            # 获取凭证号（H列数据）
            voucher_number = str(row.iloc[h_column_index]).strip()
            
            # 构建目标文件夹路径
            target_folder = os.path.join(base_dir, folder_name)
            
            # 检查目标文件夹是否存在
            if not os.path.exists(target_folder):
                print(f"警告：文件夹 '{folder_name}' 不存在，跳过复制文件")
                continue
            
            # 在凭证目录中查找匹配的文件
            found = False
            match_attempts = []  # 记录匹配尝试信息
            try:
                # 首先尝试直接匹配
                match_attempts.append(f"尝试直接匹配: '{voucher_number}'")
                for filename in os.listdir(vouchers_dir):
                    # 检查文件名是否包含凭证号
                    if voucher_number in filename:
                        source_file = os.path.join(vouchers_dir, filename)
                        target_file = os.path.join(target_folder, filename)
                        
                        # 复制文件
                        try:
                            shutil.copy2(source_file, target_file)
                            print(f"已复制: {filename} -> {folder_name}/{filename}")
                            copied_files_count += 1
                            found = True
                            break
                        except Exception as e:
                            print(f"复制文件时出错 '{filename}': {e}")
                
                # 如果直接匹配失败，尝试处理年份缺失的情况
                if not found and '-' in voucher_number:
                    parts = voucher_number.split('-')
                    # 检查是否为日期格式 (YYYY-MM-DD 或 YYYY-M-D 或 MM-DD 等)
                    if len(parts) >= 2:
                        # 尝试处理年份缺失的情况
                        # 例如: "2023-1-1" -> "13-1-1"
                        if len(parts[0]) == 4 and parts[0].isdigit():
                            # 提取年份后两位
                            year_suffix = parts[0][2:]
                            # 构建可能的短格式日期
                            short_voucher_number = f"{year_suffix}-{'-'.join(parts[1:])}"
                            match_attempts.append(f"尝试年份缺失匹配: '{short_voucher_number}' (原始: '{voucher_number}')")
                            
                            # 再次搜索凭证文件夹
                            for filename in os.listdir(vouchers_dir):
                                if short_voucher_number in filename:
                                    source_file = os.path.join(vouchers_dir, filename)
                                    target_file = os.path.join(target_folder, filename)
                                    
                                    # 复制文件
                                    try:
                                        shutil.copy2(source_file, target_file)
                                        print(f"已复制(年份缺失匹配): {filename} -> {folder_name}/{filename}")
                                        print(f"  原始凭证号: {voucher_number}")
                                        print(f"  匹配的短格式: {short_voucher_number}")
                                        copied_files_count += 1
                                        found = True
                                        break
                                    except Exception as e:
                                        print(f"复制文件时出错 '{filename}': {e}")
                
                # 尝试更宽松的匹配 - 只匹配月日部分
                if not found and '-' in voucher_number:
                    parts = voucher_number.split('-')
                    if len(parts) >= 2:
                        # 提取月日部分
                        month_day_part = '-'.join(parts[1:])
                        match_attempts.append(f"尝试月日部分匹配: '{month_day_part}' (从原始: '{voucher_number}')")
                        
                        # 再次搜索凭证文件夹
                        for filename in os.listdir(vouchers_dir):
                            if month_day_part in filename:
                                source_file = os.path.join(vouchers_dir, filename)
                                target_file = os.path.join(target_folder, filename)
                                
                                # 复制文件
                                try:
                                    shutil.copy2(source_file, target_file)
                                    print(f"已复制(月日部分匹配): {filename} -> {folder_name}/{filename}")
                                    print(f"  原始凭证号: {voucher_number}")
                                    print(f"  匹配的月日部分: {month_day_part}")
                                    copied_files_count += 1
                                    found = True
                                    break
                                except Exception as e:
                                    print(f"复制文件时出错 '{filename}': {e}")
            except Exception as e:
                error_msg = f"读取凭证文件夹时出错: {e}"
                print(error_msg)
                match_attempts.append(f"错误: {error_msg}")
            
            if not found:
                print(f"\n未找到凭证文件: {voucher_number}")
                print("匹配尝试详情:")
                for attempt in match_attempts:
                    print(f"  - {attempt}")
                # 列出凭证文件夹中的文件，帮助用户排查
                print("\n凭证文件夹中可用的文件:")
                try:
                    files_in_dir = os.listdir(vouchers_dir)
                    if files_in_dir:
                        # 只显示最多5个文件作为示例
                        for i, filename in enumerate(files_in_dir[:5]):
                            print(f"  - {filename}")
                        if len(files_in_dir) > 5:
                            print(f"  ... 以及其他 {len(files_in_dir) - 5} 个文件")
                    else:
                        print("  凭证文件夹为空")
                except Exception as e:
                    print(f"  无法读取凭证文件夹内容: {e}")
                
                not_found_files_count += 1
                print()  # 添加空行，使输出更清晰
        
        return copied_files_count, not_found_files_count
        
    except Exception as e:
        print(f"复制凭证文件时出错: {e}")
        import traceback
        traceback.print_exc()
        return 0, 0

def create_folders_from_column(file_path, column='C', vouchers_dir=None):
    """根据Excel文件指定列的数据创建文件夹"""
    try:
        # 获取Excel文件所在目录
        base_dir = os.path.dirname(file_path)
        
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 获取指定列的数据
        column_data = df.iloc[:, ord(column.upper()) - 65].dropna().unique()  # 将列字母转换为索引
        
        # 创建文件夹
        created_folders = []
        for value in column_data:
            # 确保文件夹名称有效（去除特殊字符）
            folder_name = str(value).strip()
            # 替换不能在文件名中使用的字符
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                folder_name = folder_name.replace(char, '_')
            
            # 创建文件夹路径
            folder_path = os.path.join(base_dir, folder_name)
            
            # 如果文件夹不存在，则创建
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                created_folders.append(folder_name)
                print(f"创建文件夹: {folder_name}")
            else:
                print(f"文件夹已存在: {folder_name}")
        
        # 复制凭证文件
        print("\n开始复制凭证文件...")
        copied_count, not_found_count = copy_voucher_files(file_path, base_dir, vouchers_dir)
        print(f"\n复制完成：成功复制 {copied_count} 个文件，未找到 {not_found_count} 个文件")
        
        return created_folders
    
    except Exception as e:
        print(f"处理文件时出错: {e}")
        import traceback
        traceback.print_exc()
        return []

def main():
    print("Excel文件分类创建文件夹工具")
    print("=" * 50)
    print("功能：根据Excel文件C列创建文件夹，并根据H列凭证号复制对应文件")
    print("=" * 50)
    
    # 检查命令行参数
    file_path = None
    if len(sys.argv) > 1:
        # 如果提供了命令行参数，则使用第一个参数作为文件路径
        file_path = sys.argv[1]
        if not os.path.exists(file_path):
            print(f"错误：文件 '{file_path}' 不存在")
            file_path = None
    
    # 如果没有提供有效的命令行参数，则使用文件选择对话框
    if not file_path:
        file_path = select_excel_file()
    
    if not file_path:
        print("未选择文件，程序退出。")
        return
    
    print(f"已选择文件: {file_path}")
    
    # 选择凭证文件夹
    print("\n请选择存放凭证文件的文件夹：")
    vouchers_dir = select_voucher_folder()
    
    # 根据C列创建文件夹
    column = 'C'  # 默认使用C列
    print(f"\n正在根据{column}列数据创建文件夹...")
    
    created_folders = create_folders_from_column(file_path, column, vouchers_dir)
    
    print("\n" + "=" * 50)
    if created_folders:
        print(f"成功创建 {len(created_folders)} 个文件夹")
        for folder in created_folders[:10]:  # 只显示前10个文件夹，避免输出过多
            print(f"- {folder}")
        if len(created_folders) > 10:
            print(f"... 等共 {len(created_folders)} 个文件夹")
    else:
        print("未创建新文件夹，可能是所有文件夹已存在或处理出错")
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()