import os
import subprocess
import time

def convert_word_to_pdf(input_folder, output_folder):
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)

    # 启动Word
    subprocess.run(['open', '-a', 'Microsoft Word'])

    # 给Word一些时间启动
    time.sleep(5)

    # 遍历输入文件夹中的所有Word文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx") or filename.endswith(".doc"):
            # 构建完整文件路径
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf")

            # AppleScript脚本
            applescript = f'''
            tell application "Microsoft Word"
                open "{input_path}"
                set activeDoc to active document
                save as activeDoc file name "{output_path}" file format format PDF
                close activeDoc saving no
            end tell
            '''

            # 执行AppleScript
            try:
                subprocess.run(['osascript', '-e', applescript], check=True)
                print(f"已转换: {filename} -> {os.path.basename(output_path)}")
            except subprocess.CalledProcessError as e:
                print(f"转换失败: {filename}")
                print(f"错误信息: {e}")

    # 关闭Word
    subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to quit'])

if __name__ == "__main__":
    input_folder = "/Users/jonny/Downloads/建设工程法规及相关知识教材精讲班-2025版"
    output_folder = "/Users/jonny/Downloads/法规精讲PDF"
    
    convert_word_to_pdf(input_folder, output_folder)