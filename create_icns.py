import os
import subprocess

def create_iconset():
    # 创建临时iconset目录
    iconset_name = 'icon.iconset'
    if not os.path.exists(iconset_name):
        os.makedirs(iconset_name)

    # 使用Inkscape将SVG转换为PNG
    sizes = [(16,16), (32,32), (64,64), (128,128), (256,256), (512,512), (1024,1024)]
    
    for size in sizes:
        output_name = f'icon_{size[0]}x{size[0]}.png'
        output_path = os.path.join(iconset_name, output_name)
        
        # 转换SVG到PNG
        subprocess.run([
            'sips',
            '-s', 'format', 'png',
            '-z', str(size[0]), str(size[1]),
            'icon.svg',
            '--out', output_path
        ])

        # 为Retina显示创建2x版本
        if size[0] <= 512:
            retina_name = f'icon_{size[0]}x{size[0]}@2x.png'
            retina_path = os.path.join(iconset_name, retina_name)
            subprocess.run([
                'sips',
                '-s', 'format', 'png',
                '-z', str(size[0]*2), str(size[1]*2),
                'icon.svg',
                '--out', retina_path
            ])

    # 使用iconutil将iconset转换为icns
    subprocess.run(['iconutil', '-c', 'icns', iconset_name])

    # 清理临时文件
    subprocess.run(['rm', '-rf', iconset_name])

if __name__ == '__main__':
    create_iconset()