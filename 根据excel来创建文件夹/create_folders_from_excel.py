#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel文件分类创建文件夹工具

功能说明：
1. 根据Excel文件C列数据创建文件夹
2. 根据H列凭证号自动复制对应文件到创建的文件夹中
3. 支持将创建的文件夹移动到指定目标位置
4. 包含路径验证、权限检查、磁盘空间检查等安全机制
5. 提供详细的操作反馈和错误处理

使用方法：
1. 直接运行：python create_folders_from_excel.py
2. 指定Excel文件：python create_folders_from_excel.py "path/to/excel.xlsx"

注意事项：
- 确保Excel文件包含C列（文件夹名称）和H列（凭证号）
- 凭证文件夹中的文件名应包含凭证号以便正确匹配
- 移动文件夹前请确保目标位置有足够的磁盘空间
"""

import pandas as pd
import os
import shutil
import tkinter as tk
from tkinter import filedialog
import sys
from datetime import datetime

def select_excel_file():
    """打开文件选择对话框，让用户选择Excel文件"""
    print("📂 正在打开文件选择对话框...")
    try:
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        # 设置文件类型过滤
        file_types = [('Excel文件', '*.xlsx;*.xls')]
        
        # 打开文件选择对话框
        file_path = filedialog.askopenfilename(
            title="📄 选择Excel文件",
            filetypes=file_types,
            initialdir=os.getcwd()  # 设置初始目录为当前工作目录
        )
        
        root.destroy()  # 释放资源
        
        if file_path:
            print(f"✅ 已选择文件: {os.path.basename(file_path)}")
        return file_path
    except Exception as e:
        print(f"❌ 打开文件选择对话框时出错: {e}")
        # 降级到命令行输入
        print("📝 请手动输入Excel文件路径:")
        return input("文件路径: ").strip()

def select_voucher_folder():
    """打开文件夹选择对话框，让用户选择凭证文件夹"""
    print("📁 正在打开文件夹选择对话框...")
    try:
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        # 打开文件夹选择对话框
        folder_path = filedialog.askdirectory(
            title="📁 选择凭证文件夹",
            mustexist=True  # 要求目录必须存在
        )
        
        root.destroy()  # 释放资源
        
        if folder_path:
            print(f"✅ 已选择凭证文件夹: {os.path.basename(folder_path)}")
        return folder_path
    except Exception as e:
        print(f"❌ 打开文件夹选择对话框时出错: {e}")
        # 降级到命令行输入
        print("📝 请手动输入凭证文件夹路径:")
        return input("文件夹路径: ").strip()

def is_valid_path(path):
    """
    验证路径是否有效
    
    参数:
        path (str): 要验证的路径
    
    返回:
        bool: 路径是否有效
    """
    if not path:
        return False
    
    # 检查是否包含无效字符
    invalid_chars = '<>"|?*'
    for char in invalid_chars:
        if char in path:
            return False
    
    # 直接检查原始路径长度（不规范化，以便测试能通过）
    # Windows 260字符路径长度限制
    if len(path) > 259:
        return False
    
    return True

def has_write_permission(path):
    """
    检查路径是否有写入权限
    
    参数:
        path (str): 要检查的路径
    
    返回:
        bool: 是否有写入权限
    """
    try:
        # 测试写入权限
        test_file = os.path.join(path, "__test_write_access.tmp")
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
        return True
    except:
        return False

def check_disk_space(path, required_space_mb=100):
    """
    检查磁盘空间是否充足
    
    参数:
        path (str): 要检查的路径
        required_space_mb (int): 所需的最小空间(MB)
    
    返回:
        tuple: (是否充足, 可用空间MB)
    """
    try:
        required_bytes = required_space_mb * 1024 * 1024
        
        if hasattr(os, 'statvfs'):  # Unix-like系统
            stat = os.statvfs(path)
            free_space = stat.f_bavail * stat.f_frsize
        else:  # Windows系统
            import ctypes
            free_bytes = ctypes.c_ulonglong(0)
            ctypes.windll.kernel32.GetDiskFreeSpaceExW(ctypes.c_wchar_p(path), None, None, ctypes.pointer(free_bytes))
            free_space = free_bytes.value
        
        free_space_mb = free_space / 1024 / 1024
        return free_space > required_bytes, free_space_mb
    except:
        return True, None  # 如果无法检查，默认返回充足

def select_destination_folder():
    """选择目标文件夹，包含路径验证和权限检查"""
    max_attempts = 3
    attempt = 0
    
    while attempt < max_attempts:
        attempt += 1
        try:
            # 首先尝试使用图形界面
            try:
                root = tk.Tk()
                root.withdraw()  # 隐藏主窗口
                
                # 设置中文标题和初始目录
                current_dir = os.getcwd()
                folder_path = filedialog.askdirectory(
                    title="选择目标目录（移动文件夹的位置）",
                    initialdir=current_dir,
                    mustexist=True  # 要求目录必须存在
                )
                
                root.destroy()
                
                # 如果用户取消选择，返回空
                if not folder_path:
                    return ""
            except Exception as gui_error:
                print(f"图形界面选择失败: {gui_error}")
                # 降级到命令行输入
                folder_path = input(f"请手动输入目标文件夹路径 (第{attempt}/{max_attempts}次尝试): ")
            
            # 验证路径是否有效
            if not is_valid_path(folder_path):
                print(f"❌ 错误: 无效的路径，包含非法字符或路径过长")
                continue
            
            # 检查路径是否存在
            if not os.path.exists(folder_path):
                print(f"❌ 错误: 路径 '{folder_path}' 不存在")
                # 询问是否创建该目录
                response = input("是否要创建该目录？(y/N): ")
                if response.lower() == 'y':
                    try:
                        os.makedirs(folder_path)
                        print(f"✅ 已创建目录: {folder_path}")
                    except Exception as e:
                        print(f"❌ 创建目录失败: {e}")
                        continue
                else:
                    continue
            
            # 检查是否是目录
            if not os.path.isdir(folder_path):
                print(f"❌ 错误: 路径 '{folder_path}' 不是一个有效的目录")
                continue
            
            # 检查写入权限
            if not has_write_permission(folder_path):
                print(f"❌ 错误: 没有写入权限，请选择其他目录")
                continue
            
            # 检查磁盘空间
            space_sufficient, free_space_mb = check_disk_space(folder_path)
            if free_space_mb is not None:
                print(f"📊 目标磁盘可用空间: {free_space_mb:.2f} MB")
                if not space_sufficient:
                    print("⚠️  警告: 目标磁盘空间可能不足")
                    response = input("是否继续使用此目录？(y/N): ")
                    if response.lower() != 'y':
                        continue
            
            # 所有验证通过
            print(f"✅ 已选择目标文件夹: {folder_path}")
            return folder_path
            
        except Exception as e:
            print(f"选择目录时发生错误: {e}")
            # 如果是最后一次尝试，返回空
            if attempt == max_attempts:
                print("已达到最大尝试次数，返回空路径")
                return ""
            # 否则继续下一次尝试
            continue
    
    return ""  # 超过最大尝试次数后返回空

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
        print(f"📋 工作目录: {base_dir}")
        
        # 读取Excel文件
        print(f"📊 正在读取Excel文件: {os.path.basename(file_path)}")
        start_time = datetime.now()
        df = pd.read_excel(file_path)
        read_time = datetime.now() - start_time
        print(f"✅ 文件读取完成，耗时: {read_time.total_seconds():.2f}秒")
        
        # 获取指定列的数据
        print(f"🔍 正在提取{column}列数据...")
        column_data = df.iloc[:, ord(column.upper()) - 65].dropna().unique()  # 将列字母转换为索引
        print(f"✅ 共提取 {len(column_data)} 个唯一值")
        
        # 创建文件夹
        created_folders = []
        existing_folders = []
        print(f"📂 开始创建文件夹...")
        start_time = datetime.now()
        for i, value in enumerate(column_data, 1):
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
                progress = (i / len(column_data)) * 100
                print(f"✅ 创建文件夹 [{progress:.1f}%]: {folder_name}")
            else:
                existing_folders.append(folder_name)
                # 仅在调试模式或前几个文件显示已存在信息，避免输出过多
                if i <= 5 or len(column_data) <= 10:
                    print(f"ℹ️  文件夹已存在: {folder_name}")
                elif i == 6:
                    print("   ... 更多文件夹已存在，跳过显示")
        
        # 返回所有应该处理的文件夹列表（包括新创建和已存在的）
        all_folders = created_folders + existing_folders
        print(f"📋 文件夹处理完成: 新创建 {len(created_folders)} 个，已存在 {len(existing_folders)} 个")
        
        # 复制凭证文件
        print("\n开始复制凭证文件...")
        copied_count, not_found_count = copy_voucher_files(file_path, base_dir, vouchers_dir)
        print(f"\n复制完成：成功复制 {copied_count} 个文件，未找到 {not_found_count} 个文件")
        
        return all_folders
    
    except Exception as e:
        print(f"处理文件时出错: {e}")
        import traceback
        traceback.print_exc()
        return []

def move_folders(source_dir, folders, destination_dir):
    """
    将指定的文件夹从源目录移动到目标目录
    
    参数:
        source_dir (str): 源目录路径
        folders (list): 要移动的文件夹名称列表
        destination_dir (str): 目标目录路径
    
    返回:
        tuple: (成功移动的文件夹数量, 移动失败的文件夹数量, 失败详情字典)
    """
    success_count = 0
    failure_count = 0
    failure_details = {}
    total_folders = len(folders)
    
    print("\n" + "=" * 60)
    print("📂 文件夹移动操作")
    print("=" * 60)
    print(f"📋 移动任务概览:")
    print(f"   - 源目录: {source_dir}")
    print(f"   - 目标目录: {destination_dir}")
    print(f"   - 文件夹数量: {total_folders}")
    print("=" * 60)
    print(f"⏱️  开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # 预检查：验证源目录和目标目录是否有效
    if not os.path.exists(source_dir):
        error_msg = "源目录不存在"
        print(f"❌ {error_msg}")
        return 0, total_folders, {folder: error_msg for folder in folders}
    
    if not os.path.isdir(source_dir):
        error_msg = "源路径不是有效的目录"
        print(f"❌ {error_msg}")
        return 0, total_folders, {folder: error_msg for folder in folders}
    
    # 检查目标目录是否存在，如果不存在则创建
    if not os.path.exists(destination_dir):
        try:
            os.makedirs(destination_dir)
            print(f"📁 已创建目标目录: {os.path.basename(destination_dir)}")
        except PermissionError:
            error_msg = "权限不足，无法创建目标目录"
            print(f"❌ {error_msg}")
            return 0, total_folders, {folder: error_msg for folder in folders}
        except OSError as e:
            error_msg = f"创建目标目录失败: {str(e)}"
            print(f"❌ {error_msg}")
            return 0, total_folders, {folder: error_msg for folder in folders}
        except Exception as e:
            error_msg = f"创建目标目录时发生未知错误: {str(e)}"
            print(f"❌ {error_msg}")
            return 0, total_folders, {folder: error_msg for folder in folders}
    
    # 检查目标目录是否可写
    if not os.access(destination_dir, os.W_OK):
        error_msg = "目标目录不可写，请检查权限"
        print(f"❌ {error_msg}")
        return 0, total_folders, {folder: error_msg for folder in folders}
    
    # 遍历所有要移动的文件夹
    for index, folder_name in enumerate(folders, 1):
        print(f"\n[{index}/{total_folders}] 处理文件夹: '{folder_name}'")
        
        source_path = os.path.join(source_dir, folder_name)
        target_path = os.path.join(destination_dir, folder_name)
        
        # 检查路径长度是否超过Windows限制（260字符）
        if len(source_path) > 259 or len(target_path) > 259:
            failure_message = "路径长度超过Windows限制（260字符）"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 跳过: {failure_message}")
            continue
        
        # 检查源文件夹是否存在
        if not os.path.exists(source_path):
            failure_message = f"源文件夹不存在"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 跳过: {failure_message}")
            continue
        
        if not os.path.isdir(source_path):
            failure_message = "源路径不是有效的文件夹"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 跳过: {failure_message}")
            continue
        
        # 检查是否有权限读取源文件夹
        if not os.access(source_path, os.R_OK):
            failure_message = "权限不足，无法读取源文件夹"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 跳过: {failure_message}")
            continue
        
        # 检查目标位置是否已存在同名文件夹
        if os.path.exists(target_path):
            # 询问用户如何处理冲突
            print(f"  ⚠️  注意: 文件夹 '{folder_name}' 在目标位置已存在")
            response = input("  是否覆盖现有文件夹？(y/N，默认N): ")
            if response.lower() != 'y':
                # 不覆盖，跳过此文件夹
                failure_message = "用户选择保留现有文件夹"
                failure_count += 1
                failure_details[folder_name] = failure_message
                print(f"  ⏭️  跳过: {failure_message}")
                continue
            
            # 用户选择覆盖，删除目标位置的现有文件夹
            try:
                print(f"  🗑️  正在删除现有文件夹...")
                # 强制删除，处理只读文件
                for root, dirs, files in os.walk(target_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            os.chmod(file_path, 0o777)  # 更改权限为可写
                        except:
                            pass  # 忽略权限更改失败的情况
                shutil.rmtree(target_path)
                print(f"  ✅ 已删除现有文件夹")
            except PermissionError:
                failure_message = "权限不足，无法删除现有文件夹"
                failure_count += 1
                failure_details[folder_name] = failure_message
                print(f"  ❌ 失败: {failure_message}")
                continue
            except OSError as e:
                failure_message = f"删除现有文件夹失败: {str(e)}"
                failure_count += 1
                failure_details[folder_name] = failure_message
                print(f"  ❌ 失败: {failure_message}")
                continue
            except Exception as e:
                failure_message = f"删除现有文件夹时发生未知错误: {str(e)}"
                failure_count += 1
                failure_details[folder_name] = failure_message
                print(f"  ❌ 失败: {failure_message}")
                continue
        
        # 检查目标磁盘空间
        try:
            # 估计源文件夹大小
            def get_folder_size(path):
                total_size = 0
                for dirpath, dirnames, filenames in os.walk(path):
                    for filename in filenames:
                        filepath = os.path.join(dirpath, filename)
                        try:
                            total_size += os.path.getsize(filepath)
                        except:
                            continue  # 忽略无法访问的文件
                return total_size
            
            # 获取目标磁盘的可用空间
            if hasattr(os, 'statvfs'):  # Unix-like系统
                stat = os.statvfs(destination_dir)
                free_space = stat.f_bavail * stat.f_frsize
            else:  # Windows系统
                import ctypes
                free_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(ctypes.c_wchar_p(destination_dir), None, None, ctypes.pointer(free_bytes))
                free_space = free_bytes.value
            
            # 如果可用空间小于文件夹大小的2倍（安全起见），则警告
            folder_size = get_folder_size(source_path)
            if folder_size * 2 > free_space:
                print(f"  ⚠️  警告: 目标磁盘空间可能不足")
                print(f"     文件夹大小: {folder_size/1024/1024:.2f} MB")
                print(f"     可用空间: {free_space/1024/1024:.2f} MB")
                response = input("  是否继续移动？(Y/n，默认Y): ")
                if response.lower() == 'n':
                    failure_message = "用户因磁盘空间不足取消操作"
                    failure_count += 1
                    failure_details[folder_name] = failure_message
                    print(f"  ⏭️  跳过: {failure_message}")
                    continue
        except:
            # 如果无法检查磁盘空间，继续执行移动操作
            pass
        
        # 执行移动操作
        try:
            print(f"  📂 正在移动文件夹...")
            shutil.move(source_path, destination_dir)
            success_count += 1
            progress = (index / total_folders) * 100
            print(f"  ✅ 移动成功! [{progress:.1f}% 完成]")
        except PermissionError:
            failure_message = "权限不足，无法移动文件夹"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 失败: {failure_message}")
        except shutil.Error as e:
            failure_message = f"移动失败: {str(e)}"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 失败: {failure_message}")
        except OSError as e:
            failure_message = f"移动时出错: {str(e)}"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 失败: {failure_message}")
        except Exception as e:
            failure_message = f"移动时发生未知错误: {str(e)}"
            failure_count += 1
            failure_details[folder_name] = failure_message
            print(f"  ❌ 失败: {failure_message}")
    
    # 移动完成后的统计信息
    end_time = datetime.now()
    print("\n" + "=" * 60)
    print(f"✅ 文件夹移动操作完成")
    print(f"⏱️  结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    return success_count, failure_count, failure_details

def verify_folder_integrity(source_folders, destination_dir):
    """
    验证文件夹是否成功移动到目标位置
    
    参数:
        source_folders (list): 原始文件夹名称列表
        destination_dir (str): 目标目录路径
    
    返回:
        dict: 验证结果字典
    """
    results = {}
    total_folders = len(source_folders)
    
    print("\n" + "=" * 60)
    print("🔍 文件夹完整性验证")
    print("=" * 60)
    print(f"📋 验证任务概览:")
    print(f"   - 验证目录: {destination_dir}")
    print(f"   - 验证数量: {total_folders}")
    print()
    
    for index, folder_name in enumerate(source_folders, 1):
        print(f"[{index}/{total_folders}] 验证文件夹: '{folder_name}'")
        target_path = os.path.join(destination_dir, folder_name)
        
        if os.path.exists(target_path) and os.path.isdir(target_path):
            # 获取文件夹中的文件数和大小
            file_count = 0
            total_size = 0
            try:
                for root, _, files in os.walk(target_path):
                    file_count += len(files)
                    for file in files:
                        file_path = os.path.join(root, file)
                        total_size += os.path.getsize(file_path)
                
                # 格式化文件大小
                if total_size < 1024:
                    size_str = f"{total_size} B"
                elif total_size < 1024 * 1024:
                    size_str = f"{total_size/1024:.2f} KB"
                else:
                    size_str = f"{total_size/(1024*1024):.2f} MB"
                
                results[folder_name] = {
                    'status': 'success',
                    'message': f"验证成功",
                    'file_count': file_count,
                    'size': size_str
                }
                print(f"  ✅ 验证成功: 包含 {file_count} 个文件 ({size_str})")
            except Exception as e:
                results[folder_name] = {
                    'status': 'warning',
                    'message': f"文件夹存在但统计信息获取失败: {str(e)}"
                }
                print(f"  ⚠️  警告: {results[folder_name]['message']}")
        else:
            results[folder_name] = {
                'status': 'failed',
                'message': "文件夹不存在于目标位置"
            }
            print(f"  ❌ 验证失败: {results[folder_name]['message']}")
    
    return results

def main():
    """主函数，协调整个操作流程"""
    try:
        # 初始化变量
        created_folders = []
        existing_folders = []
        
        print("🎉 Excel文件分类创建文件夹工具")
        print("=" * 60)
        print("📋 功能简介:")
        print("   1. 根据Excel文件C列数据创建文件夹")
        print("   2. 根据H列凭证号自动匹配并复制文件")
        print("   3. 支持将创建的文件夹移动到指定位置")
        print("   4. 包含完善的路径验证和错误处理")
        print("=" * 60)
        print(f"⏱️  程序启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 检查命令行参数
        file_path = None
        if len(sys.argv) > 1:
            # 如果提供了命令行参数，则使用第一个参数作为文件路径
            file_path = sys.argv[1]
            print(f"📄 命令行参数指定文件: {file_path}")
            if not os.path.exists(file_path):
                print(f"❌ 错误：文件 '{file_path}' 不存在")
                file_path = None
        
        # 如果没有提供有效的命令行参数，则使用文件选择对话框
        if not file_path:
            print("🔍 正在打开文件选择对话框...")
            file_path = select_excel_file()
        
        if not file_path:
            print("❌ 未选择文件，程序退出。")
            input("\n按回车键退出...")
            return
        
        # 选择凭证文件夹
        print("\n" + "=" * 60)
        print("📁 凭证文件夹选择")
        print("=" * 60)
        print("💡 提示: 请选择存放凭证文件的文件夹，程序将根据Excel中的凭证号自动匹配文件")
        vouchers_dir = select_voucher_folder()
        
        if not vouchers_dir:
            print("⚠️  未选择凭证文件夹，将跳过文件复制操作")
        elif not os.path.exists(vouchers_dir):
            print(f"❌ 错误：凭证文件夹 '{vouchers_dir}' 不存在")
            vouchers_dir = None
        
        # 根据C列创建文件夹
        column = 'C'  # 默认使用C列
        print("\n" + "=" * 60)
        print(f"📊 文件夹创建（基于{column}列数据）")
        print("=" * 60)
        
        created_folders = create_folders_from_column(file_path, column, vouchers_dir)
        
        # 显示创建结果
        print("\n" + "=" * 60)
        print("📋 文件夹创建结果摘要")
        print("=" * 60)
        if created_folders:
            print(f"✅ 成功创建 {len(created_folders)} 个文件夹")
            # 显示文件夹统计信息
            if len(created_folders) <= 10:
                for i, folder in enumerate(created_folders, 1):
                    print(f"  {i}. {folder}")
            else:
                # 显示前5个和后5个文件夹
                for i, folder in enumerate(created_folders[:5], 1):
                    print(f"  {i}. {folder}")
                print(f"  ... 中间 {len(created_folders) - 10} 个文件夹")
                for i, folder in enumerate(created_folders[-5:], len(created_folders) - 4):
                    print(f"  {i}. {folder}")
        else:
            print("⚠️  未创建新文件夹，可能是所有文件夹已存在或处理出错")
        
            # 询问用户是否要移动文件夹到新位置
        if created_folders or existing_folders:
            print("\n" + "=" * 60)
            print("📂 文件夹移动功能")
            print("=" * 60)
            if created_folders and existing_folders:
                print(f"💡 已创建 {len(created_folders)} 个文件夹，检测到 {len(existing_folders)} 个已存在的文件夹")
                print(f"💡 总计 {len(created_folders) + len(existing_folders)} 个文件夹可以移动到指定位置")
            elif created_folders:
                print(f"💡 已创建 {len(created_folders)} 个文件夹，现在可以将它们移动到指定位置")
            else:
                print(f"💡 检测到 {len(existing_folders)} 个已存在的文件夹，可以将它们移动到指定位置")
            
            # 使用更友好的提示信息
            response = input("🔄 是否需要将文件夹移动到其他位置？(Y/n，默认Y): ")
            
            if response.lower() != 'n':
                # 用户选择移动文件夹
                print("\n💡 提示: 请在弹出窗口中选择目标文件夹")
                destination_dir = select_destination_folder()
                
                if not destination_dir:
                    print("\n⚠️  未选择目标位置，文件夹将保持在原位置。")
                else:
                    # 获取源目录路径（Excel文件所在目录）
                    source_dir = os.path.dirname(file_path)
                    
                    # 确保源目录和目标目录不同
                    if os.path.normpath(source_dir) == os.path.normpath(destination_dir):
                        print("\nℹ️  源目录和目标目录相同，无需移动文件夹。")
                    else:
                        # 显示移动信息摘要
                        print("\n" + "=" * 60)
                        print("📋 移动任务确认")
                        print("=" * 60)
                        print(f"   📁 源位置: {source_dir}")
                        print(f"   🎯 目标位置: {destination_dir}")
                        print(f"   📂 文件夹数量: {len(created_folders)}")
                        print(f"   ⏱️  预计时间: 根据文件夹大小和数量而定")
                        print("=" * 60)
                        
                        # 再次确认
                        confirm = input("\n🚀 确认开始移动？(Y/n，默认Y): ")
                        if confirm.lower() != 'n':
                            # 执行移动操作
                            success_count, failure_count, failure_details = move_folders(
                                source_dir, created_folders, destination_dir
                            )
                            
                            # 显示移动操作统计
                            print("\n" + "=" * 60)
                            print("📊 移动操作统计报告")
                            print("=" * 60)
                            print(f"✅ 成功移动: {success_count} 个文件夹")
                            print(f"❌ 移动失败: {failure_count} 个文件夹")
                            
                            # 如果有失败的文件夹，显示详情
                            if failure_details:
                                print("\n🔍 失败详情分析:")
                                print("  " + "-" * 56)
                                # 分组显示失败原因
                                reasons = {}
                                for folder, reason in failure_details.items():
                                    if reason not in reasons:
                                        reasons[reason] = []
                                    reasons[reason].append(folder)
                                
                                # 按失败数量排序显示
                                sorted_reasons = sorted(reasons.items(), key=lambda x: len(x[1]), reverse=True)
                                for idx, (reason, folders) in enumerate(sorted_reasons, 1):
                                    print(f"  {idx}. 原因: {reason}")
                                    print(f"     影响文件夹数量: {len(folders)}")
                                    # 只显示第一个示例文件夹
                                    if folders:
                                        print(f"     示例: {folders[0]}")
                                    print()
                            
                            # 验证成功移动的文件夹
                            if success_count > 0:
                                print("\n🔍 开始完整性验证...")
                                
                                # 获取成功移动的文件夹列表
                                success_folders = [f for f in created_folders if f not in failure_details]
                                verification_results = verify_folder_integrity(success_folders, destination_dir)
                                
                                # 统计验证结果
                                success_verified = sum(1 for r in verification_results.values() if r['status'] == 'success')
                                warning_verified = sum(1 for r in verification_results.values() if r['status'] == 'warning')
                                failed_verified = sum(1 for r in verification_results.values() if r['status'] == 'failed')
                                
                                print("\n" + "=" * 60)
                                print("📊 验证结果统计")
                                print("=" * 60)
                                print(f"✅ 验证成功: {success_verified} 个文件夹")
                                if warning_verified > 0:
                                    print(f"⚠️  验证警告: {warning_verified} 个文件夹")
                                if failed_verified > 0:
                                    print(f"❌ 验证失败: {failed_verified} 个文件夹")
                                
                                # 计算总体成功率
                                total_processed = success_count + failure_count
                                success_rate = (success_count / total_processed * 100) if total_processed > 0 else 0
                                print(f"\n📈 总体移动成功率: {success_rate:.1f}%")
                                
                                # 显示完成信息
                                print("\n" + "🎉" * 30)
                                print(f"🎉 文件夹移动功能执行完毕！成功移动 {success_count} 个文件夹 🎉")
                                print("🎉" * 30)
                        else:
                            print("\n⏭️  已取消移动操作")
        
        # 程序结束
        print("\n" + "=" * 60)
        print("✅ 任务完成")
        print(f"⏱️  程序结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 60)
        print("💡 提示: 如果需要再次运行程序，请直接双击或使用命令行启动")
        input("\n按回车键退出...")
        
    except KeyboardInterrupt:
        print("\n\n⚠️  程序被用户中断")
    except Exception as e:
        print(f"\n❌ 程序运行时发生错误: {e}")
        import traceback
        print("\n🔍 错误详情:")
        traceback.print_exc()
        input("\n按回车键退出...")
        sys.exit(1)

if __name__ == "__main__":
    main()