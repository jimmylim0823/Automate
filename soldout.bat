@echo off & python -x "%~f0" %* & pause & goto :eof
import sys
import subprocess
import pkg_resources

required = {'openpyxl', 'xlwings'} 
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed


if missing:
    # implement pip as a subprocess:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install',*missing])

# 단종여부를 모두 X로 만들어주는 파일입니다 xlwings 모듈 사용
# upload 폴더 아래에 브랜드 폴더와 같은 레벨에 해당 파일을 위치하게 하세요

import os
import xlwings

folder_path = os.getcwd()
folder_path = folder_path + '\\'
folder_list = os.listdir(folder_path)

for folder in folder_list:
    if (folder[-3:] == '.py') or (folder[-4:] == '.bat'):
        continue
    else:
        file_path = folder_path + folder
        file_list = os.listdir(file_path)

        for file in file_list:
            if file[-5:] != '.xlsx':
                continue
            else:
                file_name = file_path + '\\' + file
                wb = xlwings.Book(file_name)
                ws = wb.sheets['cos']
                ws.range('C13').value = "X"
                wb.save(file_name)
                xlwings.apps.active.kill()
                print("{} 파일을 변경하였습니다".format(file_name))

print("\n파일 변경을 완료하였습니다.")
