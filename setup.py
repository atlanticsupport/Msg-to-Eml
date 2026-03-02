from setuptools import setup

APP = ['main.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['extract_msg', 'email'],
    'plist': {
        'LSUIElement': True, # Faz com que a app corra no background sem ícone na Dock
        'CFBundleDocumentTypes': [
            {
                'CFBundleTypeExtensions': ['msg'],
                'CFBundleTypeName': 'Microsoft Outlook Message',
                'CFBundleTypeRole': 'Viewer',
                'LSHandlerRank': 'Owner',
            }
        ],
    }
}

setup(
    app=APP,
    name='MSGtoOutlook',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
