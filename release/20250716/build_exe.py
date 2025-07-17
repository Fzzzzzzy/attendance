import os
import sys
import subprocess
import shutil
import importlib
from subprocess import run, PIPE

def check_pyinstaller():
    """检查 PyInstaller 是否可用"""
    try:
        result = run(['python', '-m', 'PyInstaller', '--version'], stdout=PIPE, stderr=PIPE, text=True)
        return result.returncode == 0
    except FileNotFoundError:
        return False

def check_dependencies():
    """检查必要的依赖是否已安装"""
    required_packages = [
        'pandas',
        'numpy',
        'openpyxl'
    ]
    
    print("检查依赖项...")
    missing_packages = []
    
    # 检查普通包
    for package in required_packages:
        try:
            importlib.import_module(package.lower())
        except ImportError:
            missing_packages.append(package.lower())
    
    # 特别检查 PyInstaller
    if not check_pyinstaller():
        missing_packages.append('pyinstaller')
    
    if missing_packages:
        print("\n缺少以下依赖包：")
        for package in missing_packages:
            print(f"- {package}")
        print("\n请使用以下命令安装依赖：")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    print("✓ 所有依赖项已安装")
    return True

def clean_previous_build():
    """清理之前的构建文件"""
    print("\n清理之前的构建文件...")
    dirs_to_clean = ['build', 'dist']
    files_to_clean = ['attendance_system.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"✓ 已删除 {dir_name} 目录")
    
    for file_name in files_to_clean:
        if os.path.exists(file_name):
            os.remove(file_name)
            print(f"✓ 已删除 {file_name}")

def build_exe():
    """构建可执行文件"""
    # 检查依赖
    if not check_dependencies():
        sys.exit(1)
    
    # 清理之前的构建
    clean_previous_build()
    
    print("\n开始打包...")
    
    # PyInstaller 命令
    cmd = [
        'python', '-m', 'PyInstaller',
        '--noconfirm',    # 覆盖输出目录
        '--clean',        # 清理临时文件
        '--onefile',      # 打包成单个文件
        '--name', 'attendance_system',  # 输出文件名
        '--icon', 'NONE', # 不使用图标
        '--log-level', 'WARN',  # 只显示警告和错误
        'execute.py'      # 主程序
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print("\n✓ 打包成功！")
        print("\n使用说明：")
        print("1. 将 dist 目录中的 attendance_system.exe 复制到新目录")
        print("2. 确保以下数据文件与exe在同一目录：")
        print("   - 临时卡.xlsx")
        print("   - 休假单.xlsx")
        print("   - 出差单.xlsx")
        print("   - 原始数据.xlsx")
        print("   - 员工花名册.xlsx")
        print("   - 外出单.xlsx")
        print("   - 日历.xlsx")
        
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 打包失败：{e}")
        sys.exit(1)

if __name__ == '__main__':
    build_exe() 