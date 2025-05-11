# hook-pptx.py
# PyInstallerのhooks
from PyInstaller.utils.hooks import collect_data_files

# python-pptxに必要なデータファイルを収集
datas = collect_data_files('pptx')
