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
########## 어드민 -> 딜리버드 셀러 이동 ##########
driver.switch_to.window(tabs[0])


#어드민 로그인 진행
#driver.find_element(By.ID, 'ssm_admin_email').send_keys(deal_admin_login_id)
#driver.find_element(By.ID, 'ssm_admin_password').send_keys(deal_admin_login_password)
#driver.find_element(By.NAME, 'commit').click()

#어드민에서 테스트 셀러 계정 검색 -> 딜리버드 파트너 센터 이동
#driver.find_element(By.ID, 'inline_search').send_keys(deal_seller_id)
#driver.find_element(By.ID, 'inline_search').send_keys(Keys.ENTER)
#element = driver.find_element(By.ID,'sel_date_1month')# 라디오 버튼 클릭
#driver.execute_script("arguments[0].click();", element)

#driver.implicitly_wait(5)
#셀러 ID 클릭
#time.sleep(5)
#driver.find_element(By.XPATH, '//*[@id="purchasesList"]/tbody/tr[1]/td[3]/a').click()


##신상마켓 소매 -> 딜리버드 이동
#신상마켓 로그인
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/header/div/div[2]/div[3]').click() # 로그인 버튼(페이지 상단)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/div[1]/input').send_keys(deal_seller_login_id) # 모달 / ID 입력
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/div[2]/input').send_keys(deal_seller_login_password) # 모달 / 비번 입력
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/div[3]/div[2]/div/button').click() # 모달 / 로그인 버튼
print("##########")
print("신상마켓 소매 로그인")



#딜리버드 바로가기 클릭
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/div[1]/div/ul/li[1]/div/span').click()
print("##########")
print("딜리버드 이동")



#### 딜리버드 -> 사입 요청 #####
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[2]/a').click()
time.sleep(5)
print("##########")
print("사입 요청 시작")


# 기존에 등록했던 사입 요청 리스트가 있다면 입력 초기화 하고 새로 입력
deal_test_result = driver.find_element(By.XPATH,'//*[@id="purchase_totalCount"]') # 페이지 중간 왼쪽 -> 사입 요청 : X 값
deal_test_result_check = deal_test_result.text # 사입 요청 값에서 테스트 값을 저장

if deal_test_result_check != "-": # 사입 요청 건수가 0건일 경우 실행 되지 않음
    deal_table_count = int(deal_test_result_check) # 텍스트 값을 int형으로 형변환
    
    if deal_table_count > 0: # 사입 요청 건 수가 1건 이상 있을 경우
        driver.find_element(By.XPATH, '//*[@id="purchasesList_wrapper"]/div[1]/div/div/button[1]').click() # 페이지 중간의 [입력초기화] 버튼
        time.sleep(3)
    
        driver.find_element(By.XPATH, '/html/body/div[5]/div/div[3]/button[3]').click() # 얼럿 -> 모두 초기화 ~ [예] 버튼
        time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="purchasesList_wrapper"]/div[1]/div/div/button[6]').click() # 페이지 중간의 [엑셀 업로드] 버튼
time.sleep(2)

#파일 업로드
up_load_file = driver.find_element(By.XPATH, '//*[@id="excel_file"]') # 모달 / 엑셀 업로드 양식 선태 -> 엑셀 파일 선택 Browse 버튼
up_load_file.send_keys(deal_test_saip_excel_upload)
driver.find_element(By.NAME, 'commit').click() # [저장] 버튼

print("##########")
print("파일 업로드 완료")


time.sleep(3)


#사입 요청 수량 체크
deal_test_result = driver.find_element(By.XPATH,'//*[@id="purchase_totalCount"]') # 페이지 중간 왼쪽 -> 사입 요청 : X 값
deal_test_result_check = deal_test_result.text

print(deal_test_result_check)
cell_data = google_sheet.acell('H8').value # 사입 요청 수량(리스트 수량) / 사입 요청 수량을 확인한다

if cell_data == deal_test_result_check:
    google_sheet.update_acell('I8', 'Pass') 
else:
    google_sheet.update_acell('I8', 'Failed')


#요청 가능 수량 체크
deal_test_result = driver.find_element(By.XPATH,'//*[@id="purchase_able_to_count"]') # 페이지 중간 왼쪽 -> 요청 가능 : X 값
deal_test_result_check = deal_test_result.text

print(deal_test_result_check)
cell_data = google_sheet.acell('H9').value # 요청 가능 수량(리스트 수량) / 요청 가능 수량을 확인한다.

if cell_data == deal_test_result_check:
    google_sheet.update_acell('I9', 'Pass') 
else:
    google_sheet.update_acell('I9', 'Failed')


#요청 불가능 수량 체크
deal_test_result = driver.find_element(By.XPATH,'//*[@id="purchase_unable_to_count"]') # 페이지 중간 왼쪽 -> 요청 불가능 : X 값
deal_test_result_check = deal_test_result.text

print(deal_test_result_check)
cell_data = google_sheet.acell('H10').value # 요청 불가능 수량(리스트 수량) / 요청 불가능 수량을 확인한다.

if cell_data == deal_test_result_check:
    google_sheet.update_acell('I10', 'Pass') 
else:
    google_sheet.update_acell('I10', 'Failed')
    
    


# 사입 요청 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="purchasesList"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
deal_saib_do_store_name = "" # 도매 매장명


for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        
        if deal_list_count == 11:
            cell_data = google_sheet.acell('H11').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I11', 'Pass')
            else:
                google_sheet.update_acell('I11', 'Failed')           
            
        if deal_list_count == 12:
            cell_data = google_sheet.acell('H12').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I12', 'Pass')
            else:
                google_sheet.update_acell('I12', 'Failed')           
        
        if deal_list_count == 13: # 테스트 데이터를 위해 엑셀에 업로드, 도매 매장명
            deal_saib_do_store_name = td.get_attribute("innerText")
            google_sheet.update_acell('H31', deal_saib_do_store_name)        
            google_sheet.update_acell('H35', deal_saib_do_store_name)
            google_sheet.update_acell('H39', deal_saib_do_store_name)
            google_sheet.update_acell('H43', deal_saib_do_store_name)
            google_sheet.update_acell('H74', deal_saib_do_store_name)
            google_sheet.update_acell('H100', deal_saib_do_store_name)
            google_sheet.update_acell('H115', deal_saib_do_store_name)
            google_sheet.update_acell('H126', deal_saib_do_store_name)        
                
        if deal_list_count == 37:
            cell_data = google_sheet.acell('H13').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I13', 'Pass')
            else:
                google_sheet.update_acell('I13', 'Failed')           
                
        if deal_list_count == 38:
            cell_data = google_sheet.acell('H14').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I14', 'Pass')
            else:
                google_sheet.update_acell('I14', 'Failed')           
                
        if deal_list_count == 63:
            cell_data = google_sheet.acell('H15').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I15', 'Pass')
            else:
                google_sheet.update_acell('I15', 'Failed')           
                
        if deal_list_count == 64:
            cell_data = google_sheet.acell('H16').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I16', 'Pass')
            else:
                google_sheet.update_acell('I16', 'Failed')           
   
        deal_list_count += 1


driver.find_element(By.XPATH, '//*[@id="purchasesList_wrapper"]/div[1]/div/div/button[9]').click() # 페이지 중간 오른쪽 [사입 요청하기] 버튼
time.sleep(2)


# 사입 요청 버튼
driver.find_element(By.XPATH, '/html/body/div[5]/div/div[3]/button[3]').click() # 얼럿 / X건의 상품을~~ 진행하시겠습니까? -> [네] 버튼
time.sleep(5)
print("##########")
print("사입 요청 완료")


# 결제 정보 모달
driver.find_element(By.ID, 'method_SINSANGPOINT').click() # 모달 / 결제 수단 선택 -> 신상캐시 버튼
time.sleep(2)

driver.find_element(By.XPATH,'//*[@id="confirmCollapse"]/div[2]/div/label').click() # 모달 -> 이용 약관 -> 전체 동의합니다. 버튼
driver.find_element(By.XPATH, '//*[@id="payment_button"]').click() # 모달 -> [결제하기] 버튼
time.sleep(10)

print("##########")
print("결제 완료")


# 중요 hihida 딜리버드 주문번호 저장
deal_wms_purchase_number = driver.find_element(By.XPATH, '//*[@id="page-wrapper"]/div[2]/div[2]/div/div/div[1]/div[1]/h4/span')
deal_wms_purchase_number = deal_wms_purchase_number.text

# 딜리버드 주문번호 -> 구글 시트에 업데이트
google_sheet.update_acell('H46', deal_wms_purchase_number)
google_sheet.update_acell('H54', deal_wms_purchase_number)
google_sheet.update_acell('H78', deal_wms_purchase_number)
google_sheet.update_acell('H102', deal_wms_purchase_number)
google_sheet.update_acell('H118', deal_wms_purchase_number)


# 중요 hihida 딜리버드 상품코드 저장
deal_product_id_1 = driver.find_element(By.XPATH, '//*[@id="purchasesDetail"]/tbody/tr[1]/td[2]')
deal_product_id_2 = driver.find_element(By.XPATH, '//*[@id="purchasesDetail"]/tbody/tr[2]/td[2]')
deal_product_id_3 = driver.find_element(By.XPATH, '//*[@id="purchasesDetail"]/tbody/tr[3]/td[2]')

deal_product_id_1 = deal_product_id_1.text
deal_product_id_2 = deal_product_id_2.text
deal_product_id_3 = deal_product_id_3.text

# 딜리버드 상품코드  -> 구글 시트에 업데이트
google_sheet.update_acell('H30', deal_product_id_1)
google_sheet.update_acell('H103', deal_product_id_1)
google_sheet.update_acell('H119', deal_product_id_1)
google_sheet.update_acell('H127', deal_product_id_1)

google_sheet.update_acell('H34', deal_product_id_2)
google_sheet.update_acell('H106', deal_product_id_2)
google_sheet.update_acell('H121', deal_product_id_2)
google_sheet.update_acell('H130', deal_product_id_2)

google_sheet.update_acell('H38', deal_product_id_3)
google_sheet.update_acell('H109', deal_product_id_3)
google_sheet.update_acell('H123', deal_product_id_3)
google_sheet.update_acell('H133', deal_product_id_3)





########## 어드민 사입마감 처리 ##########
driver.switch_to.window(tabs[1])


#어드민 로그인 진행
driver.find_element(By.ID, 'ssm_admin_email').send_keys(deal_admin_login_id)
driver.find_element(By.ID, 'ssm_admin_password').send_keys(deal_admin_login_password)
driver.find_element(By.NAME, 'commit').click()
time.sleep(2)
print("##########")
print("어드민 로그인 완료")



driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[8]/a').click() # 사입 마감 시간 설정
time.sleep(2)

driver.find_element(By.XPATH,'//*[@id="scheduleList_wrapper"]/div[1]/div[2]/div/button').click() # [기본 사입 일시 설정] 버튼
time.sleep(2)

test_now = datetime.now() # 테스트 시점 날짜 시간 구하기
test_date = test_now.date() # 날짜만
test_time = test_now + timedelta(minutes=1) # 현재 시간에서 1분 더하기

test_date = str(test_date) # 필드 입력을 위해 string로 변환
test_time = str(test_time) # 필드 입력을 위해 string로 변환

test_time = test_time[11:16] # 3분 추가된 시간에서 시간과 분만 저장 # 2022-12-22 17:03:48.128988

driver.find_element(By.ID, 'group_name').send_keys("QA 테스트") # 모달 -> 반복 설정 이름
time.sleep(2)

# 모달 뒤에 있는 페이지에도 동일한 ID, NAME이 있어 클릭이나 키 입력이 되지 않음
# "div>input[placeholder='검색에 필요한 텍스트']으로 검색 시 정상 동작
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='시작일 선택']").send_keys(test_date) # 모달 -> 반복 설정 기간 -> 시작 날짜 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='종료일 선택']").send_keys(test_date) # 모달 -> 반복 설정 기간 -> 종료 날짜 입력
time.sleep(2)


driver.find_element(By.XPATH,'//*[@id="defaultScheduleModal"]/div/div/div[1]/div[1]/div/div/div/div[3]/div/div/div/div/input').send_keys(test_time) # 모달 -> 반복 설정 시간 -> 3분 추가된 시간 입력
driver.find_element(By.XPATH,'//*[@id="defaultScheduleModal"]/div/div/div[1]/div[1]/div/div/div/div[3]/div/div/div/div/input').send_keys(Keys.ENTER) # 모달 -> 엔터로 시간 입력

driver.find_element(By.XPATH, '//*[@id="defaultScheduleModal"]/div/div/div[1]/div[1]/div/div/div/div[4]/div/button').click() # 모달 -> 설정 추가 하기 버튼
time.sleep(2)

driver.find_element(By.XPATH, '/html/body/div[7]/div/div[3]/button[1]').click() # 모달 -> 설정 추가 하기 버튼 -> [확인] 버튼

driver.find_element(By.XPATH, '//*[@id="defaultScheduleModal"]/div/div/div[1]/button').click() # 모달 -> 모달 종료 버튼 [X]

print("##########")
print("사입 마감 처리 완료")



#########################
##어드민 -> WMS 이동
driver.switch_to.window(tabs[2])


# 테스트를 위한 임시 저장 hihida
# deal_wms_purchase_number = "19590"

# wms 각 항목 -> 조회 -> 리스트의 스크롤이 생기면서 데이터 로드의 어려움(현재 보여지는 화면의 데이터만 가져옴)
# pyautogui을 사용 -> wms 로그인 화면에서 글꼴 축소 -> 항목 조회 -> 리스트의 모든 데이터가 보이게 됨
pyautogui_count = 0
while pyautogui_count < 8:
    pyautogui.hotkey('ctrl', '-') # 화면 축소
    pyautogui_count = pyautogui_count + 1

time.sleep(80) # 사입 마감 처리 시간 후 로그인 시도 hihida
#time.sleep(2) # 테스트 임시


#WMS 로그인 진행
driver.find_element(By.ID, 'login').send_keys(wms_login_id)
driver.find_element(By.ID, 'password').send_keys(wms_login_passWord)
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/form/div[2]/button').click()
time.sleep(2)
print("##########")
print("wms 로그인 완료")



##### 입고 관리 - 입고 요청 #####
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='셀러ID 필터 입력']").send_keys(deal_seller_login_id) # 테이블(리스트) -> 셀러ID 입력
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='주문번호 필터 입력']").send_keys(deal_wms_purchase_number) # 테이블(리스트) -> 주문번호
time.sleep(2)
print("##########")
print("입고 관리 - 입고 요청 이동")


# 총 주문 수 갯수 가져오기
wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[3]/div/div/div/h2[1]')
wms_test_result_check = wms_test_result.text
wms_test_result_check = wms_test_result_check.split(" ") # 공백만 제거 하고 배열에 입력
wms_test_result_check = wms_test_result_check[3]

cell_data = google_sheet.acell('H27').value # 총 주문수를 확인한다.

if cell_data == wms_test_result_check:
    google_sheet.update_acell('I27', 'Pass')
else:
    google_sheet.update_acell('I27', 'Failed')

print("##########")
print("총 주문 수 갯수 가져오기")


#총 상품수(sku) 갯수 가져오기
wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[3]/div/div/div/h2[2]')
wms_test_result_check = wms_test_result.text
wms_test_result_check = wms_test_result_check.split(" ") # 공백만 제거 하고 배열에 입력
wms_test_result_check = wms_test_result_check[2]
wms_search_list_row = wms_test_result_check # 하단에서 리스트의 ROW수 계산을 위한 데이터 입력

cell_data = google_sheet.acell('H28').value # 총 상품 수를 확인한다

if cell_data == wms_test_result_check:
    google_sheet.update_acell('I28', 'Pass')
else:
    google_sheet.update_acell('I28', 'Failed')

print("##########")
print("총 상품수(sku) 갯수 가져오기")



#총 상품수량 갯수 가져오기
wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[3]/div/div/div/h2[3]')
wms_test_result_check = wms_test_result.text
wms_test_result_check = wms_test_result_check.split(" ") # 공백만 제거 하고 배열에 입력
wms_test_result_check = wms_test_result_check[2]

cell_data = google_sheet.acell('H29').value # 총 상품 수량을 확인한다.

if cell_data == wms_test_result_check:
    google_sheet.update_acell('I29', 'Pass')
else:
    google_sheet.update_acell('I29', 'Failed')

print("##########")
print("총 상품수량 갯수 가져오기")


# 입고 요청 데이터 확인(테이블 - 리스트)
wms_search_list = driver.find_element(By.CSS_SELECTOR,'div[class="ag-pinned-left-cols-container"]') # 총 SKU 수 확인, 각 행마다 데이터 검증을 위한 row 확인

wms_search_list_row_index_ini = wms_search_list_row
wms_search_list_row = 0

wms_search_list_row_index_ini = int(wms_search_list_row_index_ini) # 반복문을 실행하기 위한 int -> 형변환
wms_search_list_row_index_ini = wms_search_list_row_index_ini - 1  # 리스트의 배열이 0부터 시작하기 떄문에 -1

wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

while wms_search_list_row_index_ini >= wms_search_list_row : # 리스트 row 수 만큼 실행
    wms_search_list_row = str(wms_search_list_row) # row 0 부터 wms_search_list.find_element에 입력하기 위해 str -> 형변환
    wms_search_list_row_index_str = "div[row-index=\"" + wms_search_list_row + "\"]" # wms_search_list.find_element 에 row 0부터 조회를 위한 값 합치기
    wms_search_list_row_id_result = wms_search_list_row_result.find_element(By.CSS_SELECTOR, wms_search_list_row_index_str) # 리스트 row 마다 ID가져오기
    
    wms_search_list_row_result_check_1 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="13"]') # 딜리버드 상품코드 : 13
    wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
    
    wms_search_list_row_result_check_2 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 도매처정보(WMS) 컬럼 >  도매매장명 : 27
    wms_search_list_row_result_check_2 = wms_search_list_row_result_check_2.text
    
    wms_search_list_row_result_check_3 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="22"]') # 도매단가 :  22
    wms_search_list_row_result_check_3 = wms_search_list_row_result_check_3.text
    
    wms_search_list_row_result_check_4 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="23"]') # 합계(주문수량) : 23
    wms_search_list_row_result_check_4 = wms_search_list_row_result_check_4.text
    
    wms_search_list_row = int(wms_search_list_row) # while을 실행하기 위해 int -> 형변환
    
    if wms_search_list_row == 0 :
        cell_data = google_sheet.acell('H30').value # 도매 상품명_자동화001 -> 딜리버드 상품코드 컬럼  확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I30', 'Pass')
        else:
            google_sheet.update_acell('I30', 'Failed')
   
        
        cell_data = google_sheet.acell('H31').value # 도매 상품명_자동화001 -> 도매처정보(WMS) 컬럼 >  도매매장명 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I31', 'Pass')
        else:
            google_sheet.update_acell('I31', 'Failed')



        cell_data = google_sheet.acell('H32').value # 도매 상품명_자동화001 -> 도매단가 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I32', 'Pass')
        else:
            google_sheet.update_acell('I32', 'Failed')
 
 
        cell_data = google_sheet.acell('H33').value # 도매 상품명_자동화001 -> 합계(주문수량) 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I33', 'Pass')
        else:
            google_sheet.update_acell('I33', 'Failed')

 
    
    if wms_search_list_row == 1 :
        cell_data = google_sheet.acell('H34').value # 도매 상품명_자동화002 -> 딜리버드 상품코드 컬럼  확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I34', 'Pass')
        else:
            google_sheet.update_acell('I34', 'Failed')

        
        cell_data = google_sheet.acell('H35').value # 도매 상품명_자동화002 -> 도매처정보(WMS) 컬럼 >  도매매장명 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I35', 'Pass')
        else:
            google_sheet.update_acell('I35', 'Failed')
    
        
        cell_data = google_sheet.acell('H36').value # 도매 상품명_자동화002 -> 도매단가 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I36', 'Pass')
        else:
            google_sheet.update_acell('I36', 'Failed')
     
        
        cell_data = google_sheet.acell('H37').value # 도매 상품명_자동화002 -> 합계(주문수량) 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I37', 'Pass')
        else:
            google_sheet.update_acell('I37', 'Failed')
    
    
    if wms_search_list_row == 2 :
        cell_data = google_sheet.acell('H38').value # 도매 상품명_자동화003 -> 딜리버드 상품코드 컬럼  확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I38', 'Pass')
        else:
            google_sheet.update_acell('I38', 'Failed')
        
        
        cell_data = google_sheet.acell('H39').value # 도매 상품명_자동화003 -> 도매처정보(WMS) 컬럼 >  도매매장명 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I39', 'Pass')
        else:
            google_sheet.update_acell('I39', 'Failed')
        
        
        cell_data = google_sheet.acell('H40').value # 도매 상품명_자동화003 -> 도매단가 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I40', 'Pass')
        else:
            google_sheet.update_acell('I40', 'Failed')
        
        
        cell_data = google_sheet.acell('H41').value # 도매 상품명_자동화003 -> 합계(주문수량) 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I41', 'Pass')
        else:
            google_sheet.update_acell('I41', 'Failed')
        
    
    wms_search_list_row = wms_search_list_row + 1  # row 증가

print("##########")
print("입고 요청 데이터 확인(테이블 - 리스트)")




##### 입고 관리 - 입고 검수진행 #####
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[1]/div/a[3]').click() # 입고 관리 -> 입고 검수진행 이동
time.sleep(2)
print("##########")
print("입고 관리 - 입고 검수진행 이동")



cell_data = google_sheet.acell('H43').value 

driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매처를 입력해 주세요']").send_keys(cell_data) # 도매처를 입력해주세요에 도매 명 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매처를 입력해 주세요']").send_keys(Keys.ENTER) # 입력 필드 하단에 도매 드룹다운 메뉴 출력
time.sleep(2)

element = driver.find_element(By.XPATH, '/html/body/div[3]/div[3]/ul/li') # 해당 도매를 클릭
element.send_keys(Keys.ENTER)

print("##########")
print("도매를 클릭")


#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='셀러ID 필터 입력']").send_keys(deal_seller_login_id) # 테이블(리스트) -> 셀러ID 입력
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='주문번호 필터 입력']").send_keys(deal_wms_purchase_number) # 테이블(리스트) -> 주문번호
time.sleep(2)


# 신규 주문 수량 확인
wms_test_result = driver.find_element(By.CLASS_NAME, 'MuiButtonBase-root.MuiTab-root.MuiTab-textColorPrimary.Mui-selected.css-1fs0d0o') # 버튼의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("ID") # 동적으로 변화되는 ID값 취득

wms_search_list_row_index_str = "//*[@id=\"" + wms_test_result + "\"]/span[1]" # 버튼[(신규 주문 (X건)]의 xPATH값 조합

wms_test_result = driver.find_element(By.XPATH, wms_search_list_row_index_str) # 버튼[(신규 주문 (X건)] -> 신규 주문 (X건) 값 취득
wms_test_result_check = wms_test_result.text

wms_test_result_check = wms_test_result_check.replace('신규 주문 (','')
wms_test_result_check = wms_test_result_check.replace('건)','') # 텍스트에서 수량만 취득하기 위해 나머지 텍스트 삭제


cell_data = google_sheet.acell('H47').value # 신규 주문 수량 확인한다.

if cell_data == wms_test_result_check:
    google_sheet.update_acell('I47', 'Pass')
else:
    google_sheet.update_acell('I47', 'Failed')

print("##########")
print("신규 주문 수량 확인")



#총 도매단가, 총 장끼 수량 가져오기
time.sleep(20) # 총 도매 단가를 가져오려면 일정 시간 대기를 해야 로드 할 수 있음

wms_test_result_check = driver.find_element(By.CLASS_NAME, 'MuiTypography-root.MuiTypography-h2.css-i8nswv').text

cell_data = google_sheet.acell('H48').value # 총 도매단가 확인한다.

if cell_data == wms_test_result_check:
    google_sheet.update_acell('I48', 'Pass')
else:
    google_sheet.update_acell('I48', 'Failed')

time.sleep(20)
#wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[4]/h2[4]') # 총 장끼 수량 / 위치 강제
#wms_test_result_check = wms_test_result.text
wms_test_result_check = driver.find_element(By.CLASS_NAME, 'MuiTypography-root.MuiTypography-h2.css-1tq6ygv').text

cell_data = google_sheet.acell('H49').value # 총 장끼수량 확인한다.
if cell_data == wms_test_result_check:
    google_sheet.update_acell('I49', 'Pass')
else:
    google_sheet.update_acell('I49', 'Failed')

print("##########")
print("총 도매단가, 총 장끼 수량 가져오기")


#driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[9]/div[2]/button[2]').click() # 하단 -> [입고 검수 진행 처리 (소봉 바코드 출력) ] 버튼
print("##########")
print("소봉 바코드 출력) ] 버튼 시작")
# xpath 오류로 삭제 driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[8]/div[2]/button[2]').click() # 하단 -> [입고 검수 진행 처리 (소봉 바코드 출력) ] 버튼
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeSmall.MuiButton-containedSizeSmall.MuiButtonBase-root.jss119.css-1ugafsf') # 버튼의 상위 그룹의 클래스 네임으로 정보 취득

wms_test_result[1].click() # 하단 -> [입고 검수 진행 처리 (소봉 바코드 출력) ] 버튼
time.sleep(20) # 0[선택 상품 바코드 출력] / 1[입고 검수 진행 처리 (소봉 바코드 출력)] / 2[장끼 수량 초기화]
######
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 사입성공수량 체크박스 xPATH값 조합
######





alert = driver.switch_to.alert 
alert.accept() # 얼럿 확인
# alert.dismiss()# 취소
print("##########")
print("소봉 바코드 출력) ] 버튼 종료")
time.sleep(5)

alert = driver.switch_to.alert 
alert.accept() # 얼럿 확인
# alert.dismiss()# 취소
print("##########")
print("확인버튼) ] 버튼 종료")
time.sleep(5)



#바코드 출력 후 총 도매단가, 총 장끼 수량 가져오기
#wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[5]/h2[2]') # 총 도매단가 / 위치 강제
#wms_test_result_check = wms_test_result.text
time.sleep(20) # 총 도매 단가를 가져오려면 일정 시간 대기를 해야 로드 할 수 있음

wms_test_result_check = driver.find_element(By.CLASS_NAME, 'MuiTypography-root.MuiTypography-h2.css-i8nswv').text
print("wms_test_result_check!", wms_test_result_check,("!"))

time.sleep(5)
cell_data = google_sheet.acell('H52').value # 총 도매단가 확인한다.
print("cell_data!", cell_data,("!"))

wms_test_result_check = wms_test_result_check.replace(' ','')
print("wms_test_result_check.replace!", wms_test_result_check,("!"))
cell_data = cell_data.replace(' ','')
print("cell_data.replace!", cell_data,("!"))


if cell_data == wms_test_result_check: # hihida 확인해야 함 45,400  @@@@  / 45,400 @@@@  -> 각 값의 뒷 공백 길이가 다름, 그래서 failed
    google_sheet.update_acell('I52', 'Pass')
else:
    google_sheet.update_acell('I52', 'Failed')

time.sleep(20)
#wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[5]/h2[4]') # 총 장끼 수량 / 위치 강제
#wms_test_result_check = wms_test_result.text
wms_test_result_check = driver.find_element(By.CLASS_NAME, 'MuiTypography-root.MuiTypography-h2.css-1tq6ygv').text

cell_data = google_sheet.acell('H53').value # 총 장끼수량 확인한다.
if cell_data == wms_test_result_check:
    google_sheet.update_acell('I53', 'Pass')
else:
    google_sheet.update_acell('I53', 'Failed')


print("##########")
print("바코드 출력 후 총 도매단가, 총 장끼 수량 가져오기 종료")





# 입고 요청 데이터 확인(테이블 - 리스트)
#wms_search_list = driver.find_element(By.CSS_SELECTOR,'div[class="ag-pinned-left-cols-container"]') # 총 SKU 수 확인, 각 행마다 데이터 검증을 위한 row 확인
print("##########")
print("입고 요청 데이터 확인(테이블 - 리스트) 시작")

wms_search_list_row_index_ini = 3 # XPATH 값이 지속적으로 동적으로 변해, 테스트 row를 강제로 적용 ex)//*[@id="mui-p-2903-P-purchase"]~~
wms_search_list_row = int(0)

wms_search_list_row_index_ini = int(wms_search_list_row_index_ini) # 반복문을 실행하기 위한 int -> 형변환
wms_search_list_row_index_ini = wms_search_list_row_index_ini - 1  # 리스트의 배열이 0부터 시작하기 떄문에 -1

wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

while wms_search_list_row_index_ini >= wms_search_list_row : # 리스트 row 수 만큼 실행
    wms_search_list_row = str(wms_search_list_row) # row 0 부터 wms_search_list.find_element에 입력하기 위해 str -> 형변환
    wms_search_list_row_index_str = "div[row-index=\"" + wms_search_list_row + "\"]" # wms_search_list.find_element 에 row 0부터 조회를 위한 값 합치기
    wms_search_list_row_id_result = wms_search_list_row_result.find_element(By.CSS_SELECTOR, wms_search_list_row_index_str) # 리스트 row 마다 ID가져오기
    
    wms_search_list_row_result_check_1 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="13"]') # 합계(주문수량) : 23
    wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
    print("wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)
    
    
    wms_search_list_row_result_check_2 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="14"]') # 합계(장끼수량) : 14 -> 하위 버튼값 확인
    wms_search_list_row_result_check_2 = wms_search_list_row_result_check_2.text
    print("wms_search_list_row_result_check_2   ", wms_search_list_row_result_check_2)
    
    
    wms_search_list_row_result_check_3 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 합계(낱개수량) : 15
    wms_search_list_row_result_check_3 = wms_search_list_row_result_check_3.text
    print("wms_search_list_row_result_check_3   ", wms_search_list_row_result_check_3)
    
    
    wms_search_list_row_result_check_4 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 합계(센터입고수량) : 16 -> 하위 버튼값 확인
    wms_search_list_row_result_check_4 = wms_search_list_row_result_check_4.text
    print("wms_search_list_row_result_check_4   ", wms_search_list_row_result_check_4)
    
    
    wms_search_list_row_result_check_5 = wms_search_list_row_id_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="19"]') # 입고상태  : 19
    wms_search_list_row_result_check_5 = wms_search_list_row_result_check_5.text
    print("wms_search_list_row_result_check_5   ", wms_search_list_row_result_check_5)


    print("##########")
    print("입고 요청 데이터 확인(테이블 - 리스트) 종료")
   

    wms_search_list_row = int(wms_search_list_row) # while을 실행하기 위해 int -> 형변환
    
    if wms_search_list_row == 0 :
        cell_data = google_sheet.acell('H55').value # 도매 상품명_자동화001 -> 합계(사입성공수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I55', 'Pass')
        else:
            google_sheet.update_acell('I55', 'Failed')
   
        
        cell_data = google_sheet.acell('H56').value # 도매 상품명_자동화001 -> 합계(장끼수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I56', 'Pass')
        else:
            google_sheet.update_acell('I56', 'Failed')



        cell_data = google_sheet.acell('H57').value # 도매 상품명_자동화001 -> 합계(낱개수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I57', 'Pass')
        else:
            google_sheet.update_acell('I57', 'Failed')
 
 
        cell_data = google_sheet.acell('H58').value # 도매 상품명_자동화001 -> 합계(센터입고수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I58', 'Pass')
        else:
            google_sheet.update_acell('I58', 'Failed')


        cell_data = google_sheet.acell('H59').value # 도매 상품명_자동화001 -> 입고상태 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_5:
            google_sheet.update_acell('I59', 'Pass')
        else:
            google_sheet.update_acell('I59', 'Failed')


 
   
    if wms_search_list_row == 1 :
        cell_data = google_sheet.acell('H60').value # 도매 상품명_자동화002 -> 합계(사입성공수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I60', 'Pass')
        else:
            google_sheet.update_acell('I60', 'Failed')
   
        
        cell_data = google_sheet.acell('H61').value # 도매 상품명_자동화002 -> 합계(장끼수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I61', 'Pass')
        else:
            google_sheet.update_acell('I61', 'Failed')



        cell_data = google_sheet.acell('H62').value # 도매 상품명_자동화002 -> 합계(낱개수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I62', 'Pass')
        else:
            google_sheet.update_acell('I62', 'Failed')
 
 
        cell_data = google_sheet.acell('H63').value # 도매 상품명_자동화002 -> 합계(센터입고수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I63', 'Pass')
        else:
            google_sheet.update_acell('I63', 'Failed')


        cell_data = google_sheet.acell('H64').value # 도매 상품명_자동화002 -> 입고상태 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_5:
            google_sheet.update_acell('I64', 'Pass')
        else:
            google_sheet.update_acell('I64', 'Failed')

   
    if wms_search_list_row == 2 :
        cell_data = google_sheet.acell('H65').value # 도매 상품명_자동화003 -> 합계(사입성공수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_1:
            google_sheet.update_acell('I65', 'Pass')
        else:
            google_sheet.update_acell('I65', 'Failed')
   
        
        cell_data = google_sheet.acell('H66').value # 도매 상품명_자동화003 -> 합계(장끼수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_2:
            google_sheet.update_acell('I66', 'Pass')
        else:
            google_sheet.update_acell('I66', 'Failed')



        cell_data = google_sheet.acell('H67').value # 도매 상품명_자동화003 -> 합계(낱개수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_3:
            google_sheet.update_acell('I67', 'Pass')
        else:
            google_sheet.update_acell('I67', 'Failed')
 
 
        cell_data = google_sheet.acell('H68').value # 도매 상품명_자동화003 -> 합계(센터입고수량)컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_4:
            google_sheet.update_acell('I68', 'Pass')
        else:
            google_sheet.update_acell('I68', 'Failed')


        cell_data = google_sheet.acell('H69').value # 도매 상품명_자동화003 -> 입고상태 컬럼 확인한다.
        if cell_data == wms_search_list_row_result_check_5:
            google_sheet.update_acell('I69', 'Pass')
        else:
            google_sheet.update_acell('I69', 'Failed')

   
    wms_search_list_row = wms_search_list_row + 1  # row 증가
   


# 전체 바코드 출력하기
time.sleep(10)
print("##########")
print("전체 바코드 출력) ] 버튼 시작") # [전체 바코드 출력]버튼 클릭한다.
driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/button').click() # 중간 -> [전체 바코드 출력] 버튼 클릭
time.sleep(10)


driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/div[2]/div/div/button').click() # 모달 -> 전체 상품 바코드 출력 -> 진행 클릭
time.sleep(10)

# 젙체 바코드 출력 후 시리얼 저장
alert = driver.switch_to.alert 
alert_barcode_all_print_text = alert.text


alert.accept() # 얼럿 확인
print("##########")
print("전체 바코드 출력) ] 버튼 종료")
time.sleep(5)




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
 