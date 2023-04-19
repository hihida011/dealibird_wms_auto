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




#########################
##어드민 -> WMS 이동
driver.switch_to.window(tabs[2])


# 테스트를 위한 임시 저장 hihida
deal_wms_purchase_number = "19694"



#time.sleep(80) # 사입 마감 처리 시간 후 로그인 시도 hihida
time.sleep(2) # 테스트 임시


#WMS 로그인 진행
driver.find_element(By.ID, 'login').send_keys(wms_login_id)
driver.find_element(By.ID, 'password').send_keys(wms_login_passWord)
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/form/div[2]/button').click()
time.sleep(2)
print("##########")
print("wms 로그인 완료")



##### 입고 관리 - 입고 검수진행 #####
#########################################################################################################################
##### 입고 관리 - 입고 검수진행 #####
time.sleep(2)

########################################################################################################################
##### 입고 관리 - 입고 확정 가능 #####
driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[1]/div/a[4]').click() # 입고 관리 -> 입고 확정 가능 이동
time.sleep(2)
print("##########")
print("입고 관리 - 입고 확정 가능 이동")

################################################
## 열(컬럼) 메뉴 - 특정 컬럼만 선택
wms_test_result = driver.find_elements(By.CLASS_NAME, 'ag-side-button-button') # 버튼의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result[0].click() # 중간 오른쪽의 열 컬럼 버튼 클릭
time.sleep(3) # 0[열 컬럼] / 1[필터]


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)


# 체크박스 컬럼 선택
wms_test_result = driver.find_element(By.CSS_SELECTOR, "div>[aria-posinset='1']") # 체크박스 컬럼의 유일값을 찾음, aria-posinset='1'
time.sleep(5)

wms_test_result = wms_test_result.get_attribute("aria-describedby") # 체크박스 유일값으로 동적으로 변화되는 aria-describedby ID값 취득
time.sleep(5)

wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() 
time.sleep(5)



# 딜리버드 상품코드 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("딜리버드 상품코드") # 컬럼 검색 필드 - 딜리버드 상품코드 입력
time.sleep(5)


wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 딜리버드 상품코드 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 입고확정가능수량(정상) 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("입고확정가능수량(정상)") # 컬럼 검색 필드 - 입고확정가능수량(정상)
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 입고확정가능수량(정상) 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 주문번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("주문번호") # 컬럼 검색 필드 - 주문번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


################################################
################################################

wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[4]/div/div[3]/div[1]/div/span[2]')

wms_test_result_check = wms_test_result.text
#wms_test_result_check = wms_test_result_check.split(" ") # 공백만 제거 하고 배열에 입력
#wms_test_result_check = wms_test_result_check[2]
wms_search_list_row_all = wms_test_result_check # 하단에서 리스트의 ROW수 계산을 위한 데이터 입력

wms_search_list_row_all = " ~ " + wms_search_list_row_all
print("wms_search_list_row_all", wms_search_list_row_all)

cell_data = google_sheet.acell('H100').value 

driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(cell_data) # "도매 매장명"에 도매 명 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
time.sleep(2)

# hihida 230102 WMS 오류로 임시 주석
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='주문번호 필터 입력']").send_keys(deal_wms_purchase_number) # 테이블(리스트) -> 주문번호

time.sleep(2)

# ROWS 데이터 가져오기
wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[4]/div/div[3]/div[1]/div/span[2]')

wms_test_result_check = wms_test_result.text
#wms_test_result_check = wms_test_result_check.split(" ") # 공백만 제거 하고 배열에 입력
#wms_test_result_check = wms_test_result_check[2]
#wms_search_list_row = wms_test_result_check # 하단에서 리스트의 ROW수 계산을 위한 데이터 입력
wms_search_list_row = wms_test_result_check.replace(wms_search_list_row_all,'') # 텍스트에서 수량만 취득하기 위해 나머지 텍스트 삭제

print("wms_search_list_rowASD",wms_search_list_row)

time.sleep(5)

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
    
    wms_search_list_row_result_check_1 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="3"]') # 딜리버드 상품코드 : 3
    wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
    print("wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)
    
    wms_search_list_row_result_check_2 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="22"]') # 입고확정가능수량(정상) : 22
    wms_search_list_row_result_check_2 = wms_search_list_row_result_check_2.text
    print("wms_search_list_row_result_check_2   ", wms_search_list_row_result_check_2)
    

    print("##########")
    print("입고 요청 데이터 확인(테이블 - 리스트) 종료")
   

    wms_search_list_row = int(wms_search_list_row) # while을 실행하기 위해 int -> 형변환
    
    if wms_search_list_row == 0 :
        cell_data = google_sheet.acell('H103').value # 도매 상품명_자동화001 -> 딜리버드 상품코드 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I103', 'Pass')
        else:
            google_sheet.update_acell('I103', 'Failed')
   
        
        cell_data = google_sheet.acell('H104').value # 도매 상품명_자동화001 -> 입고확정가능수량(정상) 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I104', 'Pass')
        else:
            google_sheet.update_acell('I104', 'Failed')

 
   
    if wms_search_list_row == 1 :
        cell_data = google_sheet.acell('H106').value # 도매 상품명_자동화002 -> 딜리버드 상품코드 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I106', 'Pass')
        else:
            google_sheet.update_acell('I106', 'Failed')
   
        
        cell_data = google_sheet.acell('H107').value # 도매 상품명_자동화002 -> 입고확정가능수량(정상) 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I107', 'Pass')
        else:
            google_sheet.update_acell('II107', 'Failed')


   
    if wms_search_list_row == 2 :
        cell_data = google_sheet.acell('H109').value # 도매 상품명_자동화003 -> 딜리버드 상품코드 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I109', 'Pass')
        else:
            google_sheet.update_acell('I109', 'Failed')
   
        
        cell_data = google_sheet.acell('H110').value # 도매 상품명_자동화003 -> 입고확정가능수량(정상) 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I110', 'Pass')
        else:
            google_sheet.update_acell('I110', 'Failed')


  
    wms_search_list_row = wms_search_list_row + 1  # row 증가
while(True):
    	pass