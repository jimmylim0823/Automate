import os
import time
import pandas as pd
import webbrowser
import pyautogui

#if locatOnWindow attribute is not found in pyscreeze copy following in __init__ under below locateAllOnScreen

#def locateOnWindow(image, title, **kwargs):
#    """
#    TODO
#    """
#    if _PYGETWINDOW_UNAVAILABLE:
#        raise PyScreezeException('locateOnWindow() failed because PyGetWindow is not installed or is unsupported on this platform.')
#
#    matchingWindows = pygetwindow.getWindowsWithTitle(title)
#    if len(matchingWindows) == 0:
#        raise PyScreezeException('Could not find a window with %s in the title' % (title))
#    elif len(matchingWindows) > 1:
#        raise PyScreezeException('Found multiple windows with %s in the title: %s' % (title, [str(win) for win in matchingWindows]))
#
#    win = matchingWindows[0]
#    win.activate()
#    return locateOnScreen(image, region=(win.left, win.top, win.width, win.height), **kwargs)


print('가격정보 수집 도구 v1.0.10')
print('해당 실행 파일의 위치가 불러오려는 엑셀파일과 동일한 경로에 있는지 확인하세요.')
file_default = input('불러오려는 파일명이 prices.xlsx 가 맞습니까? [y/n]: ').lower()
if file_default == 'y':
    file_name = 'prices.xlsx'
else:
    file_name = input('파일명을 입력해주세요 (예: prices.xlsx): ')
file_path = os.getcwd()
file_path = file_path + '\\' + file_name
df = pd.read_excel(file_path)
df = df.fillna('')

from_bottom = input('가장 최근에 링크가 입력된 제품까지는 넘어가겠습니까? [y/n]: ').lower()
last = 0  # 다른 조건이 없으면 0 이 default
if from_bottom == 'y':  # 이미 입력한 제품 넘어갈 수 있게 위치 확인
    last_list = list(df['iframe'])  # 칼럼명 확인하기
    last_list.reverse()
    for i in range(len(last_list)):
        if last_list[i] != '':
            last = len(last_list) - i
            break

print('\n(o)pen: 창을 엽니다\n(n)ext: 다음 제품으로 넘어갑니다\n(p)revious: 이전 제품으로 돌아갑니다\n(t)erminate: 프로그램 종료\n')

idx = last
while True:
    if idx + 1 >= len(list(df['productId'])):
        end = input('End of the List! (p)revious (t)erminate: ').lower()
        if end == 'p':
            idx = idx - 1
            continue
        else:
            break
    print("id: {} | brand: {} | product: {}".format(df['productId'][idx], df['brandName'][idx], df['productName'][idx]))
    choice = input('(o)pen (n)ext (p)revious (t)erminate: ').lower()
    if choice == 'o':
        prd_id = str(int(df['productId'][idx]))
        brand = str(df['brandName'][idx])
        if '(' in str(df['brandName'][idx]):  # 영어/한국어 번역이 있는 경우 번역 제거 및 숫자 브랜드 고려
            brand = brand[:brand.index('(')]
        product = str(df['productName'][idx])
        if brand in product:  # 제품명에 브랜드 명이 있는 경우 브랜드명 제거
            product = product[len(brand) + 1:]
        if '리뉴얼' in product:  # 리뉴얼 제품의 경우 리뉴얼 부분 제거
            product = product[product.rfind('('):]
        search = brand + ' ' + product
        search = search.replace(' ', '%20')
        webbrowser.open('https://partners.coupang.com/#affiliate/ws/link/0/{}'.format(search))
        webbrowser.open('https://admin.momguide.co.kr/products/details/{}'.format(prd_id))
        time.sleep(0.1)
        pyautogui.keyDown("ctrl")
        pyautogui.press("n")
        pyautogui.keyUp("ctrl")
        time.sleep(0.1)
        webbrowser.open('https://search.shopping.naver.com/search/all?query={}'.format(search))
        time.sleep(0.1)
        pyautogui.keyDown("win")
        pyautogui.press("right")
        pyautogui.press("space")
        pyautogui.keyUp("win")
        idx = idx + 1
        continue

    elif choice == 'n':
        idx = idx + 1
        continue

    elif choice == 'p':
        idx = idx - 1
        continue

    elif choice == 't':  # terminate
        break

    else:  # wrong input
        print('Error: Wrong Input! Retry')
        continue

print('프로그램을 종료합니다')
# end of program
