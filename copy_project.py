import os
import shutil

def copy_project():
    source_dir = os.path.expanduser('~/Library/Mobile Documents/com~apple~CloudDocs/Software/word2pdf')
    target_dir = '/Users/jonny/Downloads/word2pdf'
    
    # 确保目标目录存在
    os.makedirs(target_dir, exist_ok=True)
    
    # 复制所有文件
    for item in os.listdir(source_dir):
        source_item = os.path.join(source_dir, item)
        target_item = os.path.join(target_dir, item)
        
        if os.path.isdir(source_item):
            if os.path.exists(target_item):
                shutil.rmtree(target_item)
            shutil.copytree(source_item, target_item)
        else:
            shutil.copy2(source_item, target_item)
    
    print('项目文件已成功复制到新位置！')

if __name__ == '__main__':
    copy_project()