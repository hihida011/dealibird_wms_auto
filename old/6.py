import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.ui import Select

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
import openpyxl

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import requests
from bs4 import BeautifulSoup

print("2.정상출고 시작")
# 정산 입고 - 사입앱 실행파일에서 인자 전달해 줌

'''
deal_product_id_1 = '1000026699'
deal_product_id_2 = '1000026700'
deal_product_id_3 = '1000026701'

deal_seller_name = 'QA자동화_소매2'

deal_sell_product_name_1 = '판매 상품명_자동화001'
deal_sell_product_name_2 = '판매 상품명_자동화002'
deal_sell_product_name_3 = '판매 상품명_자동화003'

deal_out_excel_upload_name = 'C:\\test\\자동화_출고요청_QA사입앱자동화_20230201_094905.xlsx'            
'''
#########################################################################################################################
#########################################################################################################################
# 딜리버드 테스트 기본 설정
deal_admin_login_id = 'hihida@deali.net'
deal_admin_login_password = '!incasys0'
deal_admin_url = 'https://dealibird.qa.sinsang.market/ssm_admins/sign_in'
deal_seller_login_id = 'auto_soqa2'
deal_seller_login_password = '!dealisys21'
deal_seller_url = 'https://vat.qa.sinsang.market/'

# WMS 테스트 기본 설정
wms_login_id = 'auto_hihida' 									# WMS 로그인 ID
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
google_sheet = '2.정상출고' # 구글 시트


google_sheet = google_doc.worksheet(google_sheet)
google_email = 'client_email: fulfillment-test@fulfillment-371610.iam.gserviceaccount.com'

#########################################################################################################################
## 출고 시작 전 구글 시트에 정보 업데이트
# login ID / WMS 로그인 정보

#########################################################################################################################
## 데이터저장을 위한 변수 선언
# 딜리버드 배송요청 번호
deal_ship_now_normal = '210287'  # 자동화상품001(바로-일반)
deal_ship_now_today = '210288'   # 자동화상품001(바로-일반)
deal_ship_das_normal = '210289'  # 자동화상품002(다스-일반)
deal_ship_das_today = '210290'   # 자동화상품002(다스-일반)
deal_ship_each_normal = '210291' # 자동화상품003(개별-일반)
deal_ship_each_today = '210292'  # 자동화상품003(개별-일반)
deal_ship_number_temp = '' # 배송 요청 번호 임시 저장

# 배송 방법
deal_ship_type_temp = ''   # 배송 방법 임시 저장

# 딜리버드 송장번호
deal_invoice_number_now_normal = '570408852286'  # 자동화상품001(바로-일반)
deal_invoice_number_now_today = '076600040009'   # 자동화상품001(바로-일반)
deal_invoice_number_das_normal  = '570408852290'  # 자동화상품002(다스-일반)
deal_invoice_number_das_today = '076600040010'   # 자동화상품002(다스-일반)
deal_invoice_number_each_normal = '570408852301' # 자동화상품003(개별-일반)
deal_invoice_number_each_today = '076600040011'  # 자동화상품003(개별-일반)


# WMS - 출고 관리 - 출고 회자 주문별 조회
wms_sku_barcode_1 = '' # 딜리버드 상품에 대한 WMS의 SKU 정보
wms_sku_barcode_2 = '' # 딜리버드 상품에 대한 WMS의 SKU 정보
wms_sku_barcode_3 = '' # 딜리버드 상품에 대한 WMS의 SKU 정보

wms_ship_out_type_1 = '' # 배송타입 + 출고방식 = 바로-일반
wms_ship_out_type_2 = '' # 배송타입 + 출고방식 = 바로-일반
wms_ship_out_type_3 = '' # 배송타입 + 출고방식 = 다스-일반
wms_ship_out_type_4 = '' # 배송타입 + 출고방식 = 다스-일반
wms_ship_out_type_5 = '' # 배송타입 + 출고방식 = 개별-일반
wms_ship_out_type_6 = '' # 배송타입 + 출고방식 = 개별-일반

# WMS -출고 회자 번호
# WMS -출고 회자 번호
wms_out_round_number_now_normal = '19'  # 자동화상품001(바로-일반)
wms_out_round_number_now_today = '20'   # 자동화상품001(바로-일반)
wms_out_round_number_das_normal  = '21'  # 자동화상품002(다스-일반)
wms_out_round_number_das_today = '22'   # 자동화상품002(다스-일반)
wms_out_round_number_each_normal = '23' # 자동화상품003(개별-일반)
wms_out_round_number_each_today = '24'  # 자동화상품003(개별-일반)


# 피킹 바코드
wms_picking_barcode_now_normal = '1589'  # 자동화상품001(바로-일반)
wms_picking_barcode_now_today = '1590'   # 자동화상품001(바로-일반)
wms_picking_barcode_das_normal  = '1591'  # 자동화상품002(다스-일반)
wms_picking_barcode_das_today = '1592'   # 자동화상품002(다스-일반)
wms_picking_barcode_each_normal = '1593' # 자동화상품003(개별-일반)
wms_picking_barcode_each_today = '1594'  # 자동화상품003(개별-일반)

#########################################################################################################################
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

# hihida 221230
#########################################################################################################################
##어드민 -> WMS 이동
#########################################################################################################################
driver.switch_to.window(tabs[2])


# 테스트를 위한 임시 저장 hihida
#deal_wms_purchase_number = "19674"

# wms 각 항목 -> 조회 -> 리스트의 스크롤이 생기면서 데이터 로드의 어려움(현재 보여지는 화면의 데이터만 가져옴)
# pyautogui을 사용 -> wms 로그인 화면에서 글꼴 축소 -> 항목 조회 -> 리스트의 모든 데이터가 보이게 됨
#pyautogui_count = 0
#while pyautogui_count < 8:
#    pyautogui.hotkey('ctrl', '-') # 화면 축소
#    pyautogui_count = pyautogui_count + 1

# time.sleep(130) # 사입 마감 처리 시간 후 로그인 시도 hihida
time.sleep(5) # 테스트 임시


#WMS 로그인 진행
driver.find_element(By.ID, 'login').send_keys(wms_login_id)
driver.find_element(By.ID, 'password').send_keys(wms_login_passWord)
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/form/div[2]/button').click()
time.sleep(2)
print("##########")
print("wms 로그인 완료")



#########################################################################################################################
##### 출고 관리 - 출고 처리 - DAS 출고 처리 진행 #####
time.sleep(2)

deal_link = driver.find_element(By.LINK_TEXT, ("출고 관리"))
deal_link.click()
time.sleep(2)

deal_link = driver.find_element(By.LINK_TEXT, ("출고 처리"))
deal_link.click()
time.sleep(2)

deal_link = driver.find_element(By.LINK_TEXT, ("바로 출고 처리"))
deal_link.click()
time.sleep(2)

deal_link = driver.find_element(By.LINK_TEXT, ("개별 출고 처리"))
deal_link.click()
time.sleep(2)

deal_link = driver.find_element(By.LINK_TEXT, ("DAS 출고 처리"))
deal_link.click()
time.sleep(2)

# 출고 관리 -> 출고 처리 이동(230207)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

google_sheet.update_acell('I236', 'OK')



print("##########")
print("출고 관리 - 출고 처리 - DAS 출고 처리 이동")


print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(다스-일반) 시작")
cell_data = google_sheet.acell('H237').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I219', 'OK') 
time.sleep(2)



# tds = driver.find_elements_by_css_selector('td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall.css-9epf82')
# 출고 회차 확인

cell_data = google_sheet.acell('H238').value 
wms_str_loop_result = cell_data + "회차"

wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    wms_str_loop_text = wms_str_loop.text
    if wms_str_loop_result == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
wms_str_loop_result = wms_str_loop_result.replace('회차','')
if cell_data == wms_str_loop_result: # 출고 회자 확인
    google_sheet.update_acell('I238', 'Pass')            
else:
    google_sheet.update_acell('I238', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H240').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I240', 'OK') 
time.sleep(2)



# 셀러명(대표자) 가져오기
cell_data = google_sheet.acell('H241').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 셀러명(대표자) 확인
    google_sheet.update_acell('I241', 'Pass')            
else:
    google_sheet.update_acell('I241', 'Failed')

time.sleep(2)



# SKU 데이터 가져오기
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="2"]') # SKU : 2
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 처리 - DAS 출고 처리 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H242').value 
if cell_data == wms_search_list_row_result_check_1: # SKU 확인
    google_sheet.update_acell('I242', 'Pass')            
else:
    google_sheet.update_acell('I242', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H243').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I243', 'OK') 
time.sleep(2)









print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(다스-당일) 시작")
cell_data = google_sheet.acell('H245').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I245', 'OK') 
time.sleep(2)



# tds = driver.find_elements_by_css_selector('td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall.css-9epf82')

# 출고 회차 확인
cell_data = google_sheet.acell('H246').value 
wms_str_loop_result = cell_data + "회차"

wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    wms_str_loop_text = wms_str_loop.text
    if wms_str_loop_result == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
wms_str_loop_result = wms_str_loop_result.replace('회차','')
if cell_data == wms_str_loop_result: # 출고 회자 확인
    google_sheet.update_acell('I246', 'Pass')            
else:
    google_sheet.update_acell('I246', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H248').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I248', 'OK') 
time.sleep(2)



# 셀러명(대표자) 가져오기
cell_data = google_sheet.acell('H249').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 셀러명(대표자) 확인
    google_sheet.update_acell('I249', 'Pass')            
else:
    google_sheet.update_acell('I249', 'Failed')

time.sleep(2)



# SKU 데이터 가져오기
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="2"]') # SKU : 2
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 처리 - DAS 출고 처리 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H250').value 
if cell_data == wms_search_list_row_result_check_1: # SKU 확인
    google_sheet.update_acell('I250', 'Pass')            
else:
    google_sheet.update_acell('I250', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H251').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I251', 'OK') 
time.sleep(2)








#########################################################################################################################
##### 출고 관리 - 출고 처리 - 바로 출고 처리 진행 #####
time.sleep(2)

# 출고 관리 -> 출고 처리 이동(230207)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "바로 출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

google_sheet.update_acell('I236', 'OK')



print("##########")
print("출고 관리 - 출고 처리 - 바로 출고 처리 이동")


print("출고 관리 - 출고 처리 - 바로 출고 처리 자동화상품001(바로-일반) 시작")
cell_data = google_sheet.acell('H237').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I219', 'OK') 
time.sleep(2)



# tds = driver.find_elements_by_css_selector('td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall.css-9epf82')
# 출고 회차 확인

cell_data = google_sheet.acell('H238').value 
wms_str_loop_result = cell_data + "회차"

wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    wms_str_loop_text = wms_str_loop.text
    if wms_str_loop_result == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
wms_str_loop_result = wms_str_loop_result.replace('회차','')
if cell_data == wms_str_loop_result: # 출고 회자 확인
    google_sheet.update_acell('I238', 'Pass')            
else:
    google_sheet.update_acell('I238', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H240').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I240', 'OK') 
time.sleep(2)



# 셀러명(대표자) 가져오기
cell_data = google_sheet.acell('H241').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 셀러명(대표자) 확인
    google_sheet.update_acell('I241', 'Pass')            
else:
    google_sheet.update_acell('I241', 'Failed')

time.sleep(2)



# SKU 데이터 가져오기
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="2"]') # SKU : 2
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 처리 - DAS 출고 처리 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H242').value 
if cell_data == wms_search_list_row_result_check_1: # SKU 확인
    google_sheet.update_acell('I242', 'Pass')            
else:
    google_sheet.update_acell('I242', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H243').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I243', 'OK') 
time.sleep(2)

















































# 출고 회자 정보 취득

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'p.MuiTypography-root.MuiTypography-body1')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    wms_str_loop_text = wms_str_loop.text
    if "회차" in wms_str_loop_text:
        if "출고회차:" not in wms_str_loop_text:
            print("회차 텍스트 확인 시도\n")
            wms_str_loop_result = wms_str_loop_text
            print("회차 텍스트 확인 완료\n")
            break
wms_str_loop_result = wms_str_loop_result.replace('회차','')
print(wms_str_loop_result)
time.sleep(2)


cell_data = google_sheet.acell('H220').value 
if cell_data == wms_str_loop_result: # 출고 회자 확인
    google_sheet.update_acell('I220', 'Pass')            
else:
    google_sheet.update_acell('I220', 'Failed')


# 주문수 정보 취득
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'p.MuiTypography-root.MuiTypography-body1')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    wms_str_loop_text = wms_str_loop.text
    if "개" in wms_str_loop_text:
        print("주문수 텍스트 확인 시도\n")
        wms_str_loop_result = wms_str_loop_text
        print("주문수 텍스트 확인 완료\n")
        break

time.sleep(2)


cell_data = google_sheet.acell('H221').value 
if cell_data == wms_str_loop_result: # 출고 회자 확인
    google_sheet.update_acell('I221', 'Pass')            
else:
    google_sheet.update_acell('I221', 'Failed')


try:
    # 총 주문수량 확인
    # 총 수량 앞 수량을 가져오는 부분이 클래스 값과 css 값(css-bqeta0)을 가져 오고 있어
    # 혹시 css 값이 변경 될 경우 오류가 날 수 있어 try문으로 작업함
    wms_test_result_text = ""
    
    # 총 주문 수량의 앞 수량 가져오기 ex) 0 / 5 의 앞 숫자 부분(0)
    wms_test_result = driver.find_elements(By.CLASS_NAME,'MuiTypography-root.MuiTypography-body1.css-bqeta0')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        wms_str_loop_text = wms_str_loop.text
    
    wms_test_result_text = wms_str_loop_text
    
    # 총 주문 수량의 뒤 수량 가져오기 ex) 0 / 5 의 앞 숫자 부분(/5)
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'p.MuiTypography-root.MuiTypography-body1')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        wms_str_loop_text = wms_str_loop.text
        if "/" in wms_str_loop_text:
            print("주문수 텍스트 확인 시도\n")
            wms_str_loop_result = wms_str_loop_text
            print("주문수 텍스트 확인 완료\n")
            break
    
    wms_test_result_text = wms_test_result_text + wms_str_loop_result
    wms_test_result_text = wms_test_result_text.replace(" ", "")
    time.sleep(2)
    
    cell_data = google_sheet.acell('H222').value 
    if cell_data == wms_test_result_text: # 총 주문수량 확인
        google_sheet.update_acell('I222', 'Pass')            
    else:
        google_sheet.update_acell('I222', 'Failed')
 
except:
    print("출고 처리 - 총 주문 수량 체크 오류 -> except")
    google_sheet.update_acell('I222', 'N/A')
    google_sheet.update_acell('I233', 'N/A')
    pass




# DAS 출고 - 1번째 테이블의 정보 확인자동화상품002(다스-일반))
# 하단의 4개 컬럼의 클래스 다름. 값이 각자 다름 확인해야 함
# 화면에 2개의 테이블이 존재 클래스 값이 중복으로 있을 경우 하단과 같이 작성
wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-range-left.ag-cell-value[col-id="das_cell_no"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H223').value 
    if cell_data == element.text: # DAS 셀번호
        google_sheet.update_acell('I223', 'Pass')            
    else:
        google_sheet.update_acell('I223', 'Failed')    


wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="external_order_id"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H225').value 
    if cell_data == element.text: # 배송요청
        google_sheet.update_acell('I225', 'Pass')            
    else:
        google_sheet.update_acell('I225', 'Failed')    


wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="invoice_delivery_corp"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H226').value 
    if cell_data == element.text: # 택배사
        google_sheet.update_acell('I226', 'Pass')            
    else:
        google_sheet.update_acell('I226', 'Failed')   
        

wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="invoice_delivery_number"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H227').value 
    if cell_data == element.text: # 송장번호
        google_sheet.update_acell('I227', 'Pass')            
    else:
        google_sheet.update_acell('I227', 'Failed')    


wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-range-right.ag-cell-value[col-id="out_stock_count"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H228').value 
    if cell_data == element.text: # 주문수량
        google_sheet.update_acell('I228', 'Pass')            
    else:
        google_sheet.update_acell('I228', 'Failed')    


# SKU 코드 입력하기
try:
    cell_data = google_sheet.acell('H229').value 
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    google_sheet.update_acell('I229', 'OK') 
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    google_sheet.update_acell('I231', 'OK') 
    time.sleep(2)

except:
    print("#SKU 코드 입력하기 오류 발생")
    pass



try:
    # 총 주문수량 확인
    # 총 수량 앞 수량을 가져오는 부분이 클래스 값과 css 값(css-cm47wr)을 가져 오고 있어
    # 혹시 css 값이 변경 될 경우 오류가 날 수 있어 try문으로 작업함
    wms_test_result_text = ""
    
    # 총 주문 수량의 앞 수량 가져오기 ex) 0 / 5 의 앞 숫자 부분(0)
    wms_test_result = driver.find_elements(By.CLASS_NAME,'MuiTypography-root.MuiTypography-body1.css-cm47wr')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        wms_str_loop_text = wms_str_loop.text
    
    wms_test_result_text = wms_str_loop_text
    
    # 총 주문 수량의 뒤 수량 가져오기 ex) 0 / 5 의 앞 숫자 부분(/5)
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'p.MuiTypography-root.MuiTypography-body1')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        wms_str_loop_text = wms_str_loop.text
        if "/" in wms_str_loop_text:
            print("주문수 텍스트 확인 시도\n")
            wms_str_loop_result = wms_str_loop_text
            print("주문수 텍스트 확인 완료\n")
            break
    
    wms_test_result_text = wms_test_result_text + wms_str_loop_result
    wms_test_result_text = wms_test_result_text.replace(" ", "")
    time.sleep(2)
    
    cell_data = google_sheet.acell('H233').value 
    if cell_data == wms_test_result_text: # 총 주문수량 확인
        google_sheet.update_acell('I233', 'Pass')            
    else:
        google_sheet.update_acell('I233', 'Failed')
 
except:
    print("출고 처리 - 총 주문 수량 체크 오류 -> except")    
    pass


cell_data = google_sheet.acell('H234').value 

driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I234', 'OK') 
google_sheet.update_acell('I235', 'OK') 
time.sleep(2)

print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(다스-일반) 완료")


print("##########")
print("출고 관리 - 출고 처리 - DAS 출고 처리 완료")






































while(True):
    	pass
