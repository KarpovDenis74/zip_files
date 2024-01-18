from openpyxl import load_workbook, Workbook
# from PIL import ImageGrab
import zipfile
from pathlib import Path
import os
from datetime import datetime, date
import shutil


zip_file = '1.zip'

dir = Path(datetime.strftime(datetime.now(), '%Y%m%d%H%M%S'))
# with zipfile.ZipFile(zip_file, 'r') as zf:
#     test_zf = zf.testzip()
#     list_file = zf.namelist()
#     print(f'{test_zf=}')
#     print(f'{list_file=}')
#     zf.extractall(dir)



list_files = list()
list_dir = list()
with zipfile.ZipFile(zip_file, 'r') as zf:
    for name in zf.namelist(): 
        zf.extract(name, path=dir)
        name = str(dir) + '/' + name
        if Path(name).is_dir():
            os.rename(name,  name.encode('cp437').decode('cp866'))
            list_dir.append(name)
        else:
            os.rename(name,name.encode('cp437').decode('cp866'))
            list_files.append(name)
            # os.remove(name)

    print(f'{list_files=}')
    print(f'{list_dir=}')
    shutil.rmtree(list_dir[0])
# list_dir = [dir]
# while list_dir:
#     for l_d in list_dir:
#         for file in l_d.iterdir():
#             if file.is_dir():
#                 list_dir.append(file)
#             file_name = pathlib.Path(str(file).encode('cp437').decode('cp866'))
#             print(f'{file=}')
#             print(f'{file_name.name=}')
#             file.rename(file_name)
#         list_dir.remove(file)


# workbook = load_workbook(filename='excel_images.xlsx')
# sheet_name = workbook['1']
# print(f'{sheet_name=}')
# cell = sheet_name['C1']
# print(f'{cell._hyperlink=}')

# # image = ImageGrab.grabclipboard()
# # image.save('1.jpg')