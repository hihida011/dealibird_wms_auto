import selenium.webdriver.support.ui as ui

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from datetime import datetime, timedelta


import pyautogui
import time
import sys
import json
import re

import gspread
from oauth2client.service_account import ServiceAccountCredentials


#########################
# 사입 요청 파일 다운로드
# https://docs.google.com/spreadsheets/d/19X1duCg7N2npHQHGu_pPcDaji_9pDWdI/edit#gid=1830395100
# 해당 엑셀 파일을 c:\test\ 에 저장

deal_test_saip_excel_upload = 'C:\\test\\덕규_자동화_사입요청.xlsx' # 사입 요청 파일 정보


#########################
# 딜리버드 테스트 기본 설정
deal_admin_login_id = 'hihida@deali.net'
deal_admin_login_password = '!incasys0'
deal_admin_url = 'https://dealibird.qa.sinsang.market/ssm_admins/sign_in'
deal_seller_login_id = 'chy_soqa09'
deal_seller_login_password = 'tt'
deal_seller_url = 'https://vat.qa.sinsang.market/'

# WMS 테스트 기본 설정
wms_login_id = 'hihida' 									# WMS 로그인 ID
wms_login_passWord = 'elffltutm1!'  						# WMS 로그인 비번
wms_url = 'https://matrix-web.qa.sinsang.market/signin'

# 구글 시트 연동
json_file_name = 'C:\\auto_json\\fulfillment-371610-dee41b117bdb.json'	# 구글 시트 jSON

scope = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive',
]

gc = gspread.service_account(filename=json_file_name)
google_url = 'https://docs.google.com/spreadsheets/d/1fMf-pNUosGPMJ6evQJLTRRhAGpqZwEIhmC48wxLXHMQ/edit#gid=651282265' # 테스트 시나리오 엑셀 주소

google_doc = gc.open_by_url(google_url)
google_sheet = 'TC' # 구글 시트


google_sheet = google_doc.worksheet(google_sheet)
google_email = 'client_email: fulfillment-test@fulfillment-371610.iam.gserviceaccount.com'


#########################
# 크롭 탭 3개 실행
chrome_options = Options()
chrome_options.add_argument('--start-maximized')

driver = webdriver.Chrome(chrome_options=chrome_options)
driver.execute_script('window.open("about:blank", "_blank");')
driver.execute_script('window.open("about:blank", "_blank");')

tabs = driver.window_handles

driver.switch_to.window(tabs[0])
driver.get(deal_seller_url) #신상마켓 소매 -> 딜리버드 진입

driver.maximize_window()

driver.switch_to.window(tabs[1])
driver.get(deal_admin_url)

#driver.set_window_size(6000, 1024)
driver.switch_to.window(tabs[2])
driver.get(wms_url)

action = ActionChains(driver)


########################################hihida##############################




#########################
##어드민 -> WMS 이동
driver.switch_to.window(tabs[2])
#time.sleep(200) # 사입 마감 처리 시간 후 로그인 시도 hihida
time.sleep(2) # 테스트 임시

# 테스트를 위한 임시 저장 hihida
deal_wms_purchase_number = "19654"

# wms 각 항목 -> 조회 -> 리스트의 스크롤이 생기면서 데이터 로드의 어려움(현재 보여지는 화면의 데이터만 가져옴)
# pyautogui을 사용 -> wms 로그인 화면에서 글꼴 축소 -> 항목 조회 -> 리스트의 모든 데이터가 보이게 됨
#pyautogui_count = 0
#while pyautogui_count < 8:
#    pyautogui.hotkey('ctrl', '-') # 화면 축소
#    pyautogui_count = pyautogui_count + 1



#WMS 로그인 진행
driver.find_element(By.ID, 'login').send_keys(wms_login_id)
driver.find_element(By.ID, 'password').send_keys(wms_login_passWord)
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/form/div[2]/button').click()

time.sleep(5)



##### 입고 관리 - 입고 검수진행 #####




##### 입고 관리 - 입고 진행현황 #####
driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[1]/div/a[6]').click() # 입고 관리 -> 입고 진행현황 이동
time.sleep(2)
print("##########")
print("입고 관리 - 입고 진행현황 이동")

################################################
## 열(컬럼) 메뉴 - 특정 컬럼만 선택
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/div[2]/div/div/div[3]/div/div[2]/div[3]/div[1]/div[1]/button').click() # 화면 중간 오른쪽 -> 열(컬럼) 메뉴 버튼 클릭
time.sleep(5)


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click() #전체 선택 버튼 클릭해 일단 전체 컬럼 선택
time.sleep(5)


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)


# 사입 성공 수량 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("사입성공수량") # 컬럼 검색 필드 - 사입 성공 수량 입력
time.sleep(5)


wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 사입성공수량 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 장끼수량 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("장끼수량") # 컬럼 검색 필드 - 장끼수량
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 장끼수량 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 낱개수량 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("낱개수량") # 컬럼 검색 필드 - 낱개수량
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 낱개수량 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 센터입고수량 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("센터입고수량") # 컬럼 검색 필드 - 장끼수량
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 센터입고수량 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 입고상태 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("입고상태") # 컬럼 검색 필드 - 장끼수량
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 입고상태 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 주문번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("입고상태") # 컬럼 검색 필드 - 주문번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합




################################################



cell_data = google_sheet.acell('H74').value 

driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매 매장명']").send_keys(cell_data) # "도매 매장명"에 도매 명 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매 매장명']").send_keys(Keys.ENTER) # 검색 적용

driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='주문번호 필터 입력']").send_keys(deal_wms_purchase_number) # 테이블(리스트) -> 주문번호
time.sleep(2)

# ROWS 데이터 가져오기
wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[3]/div/div[3]/div[1]/div/span[2]')
wms_test_result_check = wms_test_result.text
#wms_test_result_check = wms_test_result_check.split(" ") # 공백만 제거 하고 배열에 입력
#wms_test_result_check = wms_test_result_check[2]
wms_search_list_row = wms_test_result_check # 하단에서 리스트의 ROW수 계산을 위한 데이터 입력
print(wms_search_list_row)


# 입고 요청 데이터 확인(테이블 - 리스트)

wms_search_list_row_index_ini = wms_search_list_row
wms_search_list_row = 0

wms_search_list_row_index_ini = int(wms_search_list_row_index_ini) # 반복문을 실행하기 위한 int -> 형변환
wms_search_list_row_index_ini = wms_search_list_row_index_ini - 1  # 리스트의 배열이 0부터 시작하기 떄문에 -1

wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

while wms_search_list_row_index_ini >= wms_search_list_row : # 리스트 row 수 만큼 실행
    wms_search_list_row = str(wms_search_list_row) # row 0 부터 wms_search_list.find_element에 입력하기 위해 str -> 형변환
    wms_search_list_row_index_str = "div[row-index=\"" + wms_search_list_row + "\"]" # wms_search_list.find_element 에 row 0부터 조회를 위한 값 합치기
    wms_search_list_row_id_result = wms_search_list_row_result.find_element(By.CSS_SELECTOR, wms_search_list_row_index_str) # 리스트 row 마다 ID가져오기
    
    wms_search_list_row_result_check_1 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 합계(사입성공수량) : 27
    wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
    print("wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)
    
    ######################## 화면을 최대한 줄여도 데이터의 스크롤이 생겨서 데이터를 불러 오지 못함...
    wms_search_list_row_result_check_2 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="28]') # 합계(장끼수량) : 28
    wms_search_list_row_result_check_2 = wms_search_list_row_result_check_2.text
    print("wms_search_list_row_result_check_2   ", wms_search_list_row_result_check_2)
    
    
    wms_search_list_row_result_check_3 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="29"]') # 합계(낱개수량) : 29
    wms_search_list_row_result_check_3 = wms_search_list_row_result_check_3.text
    print("wms_search_list_row_result_check_3   ", wms_search_list_row_result_check_3)
    
    
    wms_search_list_row_result_check_4 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="30"]') # 합계(센터입고수량) : 30
    wms_search_list_row_result_check_4 = wms_search_list_row_result_check_4.text
    print("wms_search_list_row_result_check_4   ", wms_search_list_row_result_check_4)
    
    
    wms_search_list_row_result_check_5 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="8"]') # 입고상태  : 8
    wms_search_list_row_result_check_5 = wms_search_list_row_result_check_5.text
    print("wms_search_list_row_result_check_5   ", wms_search_list_row_result_check_5)


    print("##########")
    print("입고 요청 데이터 확인(테이블 - 리스트) 종료")
   

    wms_search_list_row = int(wms_search_list_row) # while을 실행하기 위해 int -> 형변환
    
    if wms_search_list_row == 0 :
        cell_data = google_sheet.acell('H79').value # 도매 상품명_자동화001 -> 합계(사입성공수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I79', 'Pass')
        else:
            google_sheet.update_acell('I79', 'Failed')
   
        
        cell_data = google_sheet.acell('H80').value # 도매 상품명_자동화001 -> 합계(장끼수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I80', 'Pass')
        else:
            google_sheet.update_acell('I80', 'Failed')



        cell_data = google_sheet.acell('H81').value # 도매 상품명_자동화001 -> 합계(낱개수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I81', 'Pass')
        else:
            google_sheet.update_acell('I81', 'Failed')
 
 
        cell_data = google_sheet.acell('H82').value # 도매 상품명_자동화001 -> 합계(센터입고수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I82', 'Pass')
        else:
            google_sheet.update_acell('I82', 'Failed')


        cell_data = google_sheet.acell('H83').value # 도매 상품명_자동화001 -> 입고상태 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_5:
            google_sheet.update_acell('I83', 'Pass')
        else:
            google_sheet.update_acell('I83', 'Failed')


 
   
    if wms_search_list_row == 1 :
        cell_data = google_sheet.acell('H85').value # 도매 상품명_자동화002 -> 합계(사입성공수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I85', 'Pass')
        else:
            google_sheet.update_acell('I85', 'Failed')
   
        
        cell_data = google_sheet.acell('H86').value # 도매 상품명_자동화002 -> 합계(장끼수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I86', 'Pass')
        else:
            google_sheet.update_acell('II861', 'Failed')



        cell_data = google_sheet.acell('H87').value # 도매 상품명_자동화002 -> 합계(낱개수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I87', 'Pass')
        else:
            google_sheet.update_acell('I87', 'Failed')
 
 
        cell_data = google_sheet.acell('H88').value # 도매 상품명_자동화002 -> 합계(센터입고수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I88', 'Pass')
        else:
            google_sheet.update_acell('I88', 'Failed')


        cell_data = google_sheet.acell('H89').value # 도매 상품명_자동화002 -> 입고상태 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_5:
            google_sheet.update_acell('I89', 'Pass')
        else:
            google_sheet.update_acell('I89', 'Failed')

   
    if wms_search_list_row == 2 :
        cell_data = google_sheet.acell('H91').value # 도매 상품명_자동화003 -> 합계(사입성공수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I91', 'Pass')
        else:
            google_sheet.update_acell('I91', 'Failed')
   
        
        cell_data = google_sheet.acell('H92').value # 도매 상품명_자동화003 -> 합계(장끼수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I92', 'Pass')
        else:
            google_sheet.update_acell('I92', 'Failed')



        cell_data = google_sheet.acell('H93').value # 도매 상품명_자동화003 -> 합계(낱개수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I93', 'Pass')
        else:
            google_sheet.update_acell('I93', 'Failed')
 
 
        cell_data = google_sheet.acell('H94').value # 도매 상품명_자동화003 -> 합계(센터입고수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I94', 'Pass')
        else:
            google_sheet.update_acell('I94', 'Failed')


        cell_data = google_sheet.acell('H95').value # 도매 상품명_자동화003 -> 입고상태 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_5:
            google_sheet.update_acell('I95', 'Pass')
        else:
            google_sheet.update_acell('I95', 'Failed')

   
    wms_search_list_row = wms_search_list_row + 1  # row 증가
   












#########################
#########################
#########################
#########################
#########################
#########################
#########################
#########################
#########################




while(True):
    	pass
 