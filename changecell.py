# 특정셀에 같은 식을 입력하는 파일입니다 xlwings 모듈 사용
# upload 폴더와 같은 레벨 (브랜드 폴더의상위)에 해당 파일을 위치하게 하세요

import os
import xlwings

print("자동반복입력기 v1.0.1")
sht_idx = input("1. liv 2. cos: ")
sht = ['liv', 'cos'][int(sht_idx) - 1]

inp = input("입력하려는 텍스트 또는 식(등호포함)을 입력하세요: ")
loc = input("입력하는 위치를 입력하세요 예: C5 : ")
brand_count = 0
file_count = 0

upper_path = os.getcwd()
upper_path = upper_path + '\\'
upper_list = os.listdir(upper_path)

for upper in upper_list: #상위폴더 탐색
    if (upper[-3:] == '.py') or (upper[-4:] == '.bat'):
        continue
    else:
        folder_path = upper_path + upper + '\\'
        folder_list = os.listdir(folder_path)

        for folder in folder_list:
            if (folder[-3:] == '.py') or (folder[-4:] == '.bat'):
                continue
            else:
                brand_count = brand_count + 1
                file_path = folder_path + folder
                file_list = os.listdir(file_path)

                for file in file_list:
                    if file[-5:] != '.xlsx':
                        continue
                    else:
                        file_count = file_count + 1
                        file_name = file_path + '\\' + file
                        wb = xlwings.Book(file_name)
                        ws = wb.sheets[sht]
                        ws.range(loc).value = inp
                        wb.save(file_name)
                        xlwings.apps.active.kill()
                        print("{} 파일을 변경하였습니다".format(file_name))

print("\n브랜드 {}개 파일 {}개변경을 완료하였습니다.".format(brand_count, file_count))
