import os
import subprocess
import sys

def check_requirements():
    # 检查是否安装了必要的工具
    try:
        subprocess.run(['sips', '--version'], capture_output=True)
        subprocess.run(['iconutil', '--help'], capture_output=True)
    except FileNotFoundError:
        print('错误：缺少必要的工具 sips 或 iconutil')
        sys.exit(1)

def create_icon():
    print('正在生成应用图标...')
    try:
        # 运行图标生成脚本
        subprocess.run(['python3', 'create_icns.py'], check=True)
        if not os.path.exists('icon.icns'):
            raise Exception('图标文件生成失败')
    except Exception as e:
        print(f'生成图标时出错：{str(e)}')
        sys.exit(1)

def build_app():
    print('正在构建应用...')
    try:
        # 清理之前的构建
        subprocess.run(['rm', '-rf', 'build', 'dist'], check=True)
        
        # 使用py2app构建应用
        subprocess.run(['python3', 'setup.py', 'py2app'], check=True)
        
        # 检查应用是否成功构建
        app_path = 'dist/word2pdf_app.app'
        if not os.path.exists(app_path):
            raise Exception('应用构建失败')
            
        # 设置应用权限
        subprocess.run(['chmod', '-R', '755', app_path], check=True)
        
        print(f'\n应用已成功构建！\n应用路径：{os.path.abspath(app_path)}')
        print('\n请注意：首次运行时，您需要在系统设置中授予以下权限：')
        print('1. 辅助功能')
        print('2. 文件与文件夹')
        print('3. 完全磁盘访问权限')
        
    except Exception as e:
        print(f'构建应用时出错：{str(e)}')
        sys.exit(1)

def main():
    print('=== Word2PDF 应用构建工具 ===')
    check_requirements()
    create_icon()
    build_app()

if __name__ == '__main__':
    main()