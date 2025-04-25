from setuptools import setup

APP = ['pdfwatcherapp1.py']
OPTIONS = {
    'argv_emulation': True,
    'packages': [
        'fitz',
        'PySimpleGUI',
        'watchdog',
        'openpyxl',
        'jaraco.text',
        'jaraco.context',
        'jaraco.functools',
        'autocommand',
    ],
    'includes': [
        'jaraco',
        'jaraco.text',
        'jaraco.context',
        'jaraco.functools',
        'autocommand',
    ],
    # 'iconfile': 'Resources/pdfwatcher.icns',
}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
