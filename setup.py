from setuptools import setup

APP = ['word2pdf_app.py']
DATA_FILES = [
    ('', ['grant_access.scpt', 'icon.svg'])
]
OPTIONS = {
    'argv_emulation': False,
    'packages': ['PyQt6'],
    'includes': ['PyQt6.QtCore', 'PyQt6.QtWidgets', 'PyQt6.QtGui'],
    'excludes': ['tkinter'],
    'iconfile': 'icon.icns',
    'plist': {
        'CFBundleName': 'Word2PDF',
        'CFBundleDisplayName': 'Word2PDF',
        'CFBundleIdentifier': 'com.word2pdf.converter',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': '© 2024',
        'NSHighResolutionCapable': True,
        'LSEnvironment': {
            'LANG': 'zh_CN.UTF-8',
            'PATH': '/usr/bin:/bin:/usr/sbin:/sbin'
        },
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)