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

deal_product_id_1 = '1000026945'   # 딜리버드 상품번호(자동화상품 001)
deal_product_id_2 = '1000026946'
deal_product_id_3 = '1000026947'

deal_seller_name = 'QA자동화_소매2'

deal_sell_product_name_1 = '판매 상품명_자동화001'
deal_sell_product_name_2 = '판매 상품명_자동화002'
deal_sell_product_name_3 = '판매 상품명_자동화003'

# deal_out_excel_upload_name = 'C:\\test\\자동화_출고요청_QA사입앱자동화_20230126_165646.xlsx'            

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

#
#########################################################################################################################
## 데이터저장을 위한 변수 선언
# 딜리버드 배송요청 번호
deal_ship_now_normal = '211407'  # 자동화상품001(바로-일반)
deal_ship_now_today = '211408'   # 자동화상품001(바로-당일)
deal_ship_das_normal = '211409'  # 자동화상품002(다스-일반)
deal_ship_das_today = '211410'   # 자동화상품002(다스-당일)
deal_ship_each_normal = '211411' # 자동화상품003(개별-일반)
deal_ship_each_today = '211412'  # 자동화상품003(개별-당일)
deal_ship_number_temp = '' # 배송 요청 번호 임시 저장


# 딜리버드 송장번호
deal_invoice_number_now_normal = '570408962643'  # 자동화상품001(바로-일반)
deal_invoice_number_now_today = '076800040002'   # 자동화상품001(바로-당일)
deal_invoice_number_das_normal  = '570408962654'  # 자동화상품002(다스-일반)
deal_invoice_number_das_today = '076600040010'   # 자동화상품002(다스-당일)
deal_invoice_number_each_normal = '570408962665' # 자동화상품003(개별-일반)
deal_invoice_number_each_today = '076800040003'  # 자동화상품003(개별-당일)


# 배송 방법
deal_ship_type_temp = ''   # 배송 방법 임시 저장


# WMS - 출고 관리 - 출고 현황 조회
wms_sku_barcode_1 = '' # 딜리버드 상품에 대한 WMS의 SKU 정보
wms_sku_barcode_2 = '' # 딜리버드 상품에 대한 WMS의 SKU 정보
wms_sku_barcode_3 = '' # 딜리버드 상품에 대한 WMS의 SKU 정보

wms_ship_out_type_1 = '' # 배송타입 + 출고방식 = 바로-일반
wms_ship_out_type_2 = '' # 배송타입 + 출고방식 = 바로-당일
wms_ship_out_type_3 = '' # 배송타입 + 출고방식 = 다스-일반
wms_ship_out_type_4 = '' # 배송타입 + 출고방식 = 다스-당일
wms_ship_out_type_5 = '' # 배송타입 + 출고방식 = 개별-일반
wms_ship_out_type_6 = '' # 배송타입 + 출고방식 = 개별-당일


# WMS -출고 회자 번호
wms_out_round_number_now_normal = ''  # 자동화상품001(바로-일반)
wms_out_round_number_now_today = ''   # 자동화상품001(바로-당일)
wms_out_round_number_das_normal  = ''  # 자동화상품002(다스-일반)
wms_out_round_number_das_today = ''   # 자동화상품002(다스-당일)
wms_out_round_number_each_normal = '' # 자동화상품003(개별-일반)
wms_out_round_number_each_today = ''  # 자동화상품003(개별-당일)


# 피킹 바코드
wms_picking_barcode_now_normal = ''  # 자동화상품001(바로-일반)
wms_picking_barcode_now_today = ''   # 자동화상품001(바로-당일)
wms_picking_barcode_das_normal  = ''  # 자동화상품002(다스-일반)
wms_picking_barcode_das_today = ''   # 자동화상품002(다스-당일)
wms_picking_barcode_each_normal = '' # 자동화상품003(개별-일반)
wms_picking_barcode_each_today = ''  # 자동화상품003(개별-당일)




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
##### 출고 관리 - 출고 대상 리스트 조회 진행 #####
time.sleep(10)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 관리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(10)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 대상 리스트 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(10)


wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 대기":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

print("출고 관리 - 출고 대상 리스트 조회 이동")
google_sheet.update_acell('I113', 'OK') 


#if wms_search_list_row_all.rfind('0') == -1: # ~가 포함되어 있지 않다면 참, ! 포함되어 있다면 else로
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-status-name-value.ag-status-panel.ag-status-panel-total-and-filtered-row-count') # row 카운트 가져오기
wms_search_list_row_all = wms_test_result.text # 검색 전 전체 row 수 획득
wms_search_list_row_all = wms_search_list_row_all.replace('ROWS','')
wms_search_list_row_all = wms_search_list_row_all.replace(':','')
wms_search_list_row_all = wms_search_list_row_all.replace(' ','')
wms_search_list_row_all = wms_search_list_row_all.replace('\n','')
#else : 
#    pass

################################################
# 특정 페이지 검색 시 1 ~ 10으로 표현, 해당 부분을 회피하기 위한 코드
if wms_search_list_row_all.rfind('~') == -1: # ~가 포함되어 있지 않다면 참, ! 포함되어 있다면 else로
    wms_search_list_row_all = "~" + wms_search_list_row_all # ~ 과 전체 row의 값을 합치기
    time.sleep(5)
else :
    time.sleep(5)

################################################
## 열(컬럼) 메뉴 - 특정 컬럼만 선택
wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break

time.sleep(3)


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)

driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)


# 체크박스 컬럼 선택
wms_test_result = driver.find_element(By.CSS_SELECTOR, "div>[aria-posinset='1']") # 체크박스 컬럼의 유일값을 찾음, aria-posinset='2'
time.sleep(5)

wms_test_result = wms_test_result.get_attribute("aria-describedby") # 체크박스 유일값으로 동적으로 변화되는 aria-describedby ID값 취득
time.sleep(5)

wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() 
time.sleep(5)


# 배송요청번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("배송요청번호") # 컬럼 검색 필드 - 배송요청번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 배송요청번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 송장번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("송장번호") # 컬럼 검색 필드 - 송장번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 송장번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


cell_data = google_sheet.acell('H114').value 

driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(cell_data) #  셀러명(대표자) 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I114', 'OK') 
time.sleep(2)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H115').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I115', 'Pass')            
else:
    google_sheet.update_acell('I115', 'Failed')

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
google_sheet.update_acell('I116', 'OK') 
print("자동화상품001(바로-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_today) 
print("송장번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)

# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H117').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I117', 'Pass')            
else:
    google_sheet.update_acell('I117', 'Failed')

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
google_sheet.update_acell('I118', 'OK') 
print("자동화상품001(바로-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_normal) 
print("송장번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)

# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H119').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I119', 'Pass')            
else:
    google_sheet.update_acell('I119', 'Failed')

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
google_sheet.update_acell('I120', 'OK')
print("자동화상품002(다스-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_today) 
print("송장번호 검색 -> 자동화상품002(다스-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H121').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I121', 'Pass')            
else:
    google_sheet.update_acell('I121', 'Failed')

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
google_sheet.update_acell('I122', 'OK')
print("자동화상품002(다스-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)





# [선택출고지시] 버튼 클릭
try:
    driver.find_element(By.XPATH, '추후 입력 예정').click() # [선택출고지시] 버튼 클릭
except:
    print("try try try try try try")
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeSmall.MuiButton-containedSizeSmall.MuiButtonBase-root')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        if wms_str_loop.text == "선택출고지시":
            #print("클릭 시도\n")
            wms_str_loop.click()
            google_sheet.update_acell('I123', 'OK') 
            break
            #print("클릭 완료\n")

time.sleep(2)


alert = driver.switch_to.alert 
alert.accept() # 얼럿 확인
google_sheet.update_acell('I124', 'OK') 
time.sleep(3)
# alert.dismiss()# 취소


# [회차생성] 버튼 클릭
try:
    driver.find_element(By.XPATH, '추후 입력 예정').click() # [선택출고지시] 버튼 클릭
except:
    print("try try try try try try")
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeMedium.MuiButton-containedSizeMedium.MuiButtonBase-root')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        if wms_str_loop.text == "회차 생성":
            #print("클릭 시도\n")
            wms_str_loop.click()
            google_sheet.update_acell('I125', 'OK') 
            break
            #print("클릭 완료\n")

time.sleep(2)


try:
    alert = driver.switch_to.alert
    alert.accept() # 얼럿 확인
    time.sleep(3)
except:
    print("try try try try try try")
    pass



# 테이블(리스트) -> 송장번호 검색 자동화상품003(개별-일반)-> 
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H127').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I127', 'Pass')            
else:
    google_sheet.update_acell('I127', 'Failed')

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
google_sheet.update_acell('I128', 'OK')
print("자동화상품003(개별-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 송장번호 검색 자동화상품003(개별-당일)-> 
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_today) 
print("송장번호 검색 -> 자동화상품003(개별-당일)")
time.sleep(2)

# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H129').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I129', 'Pass')            
else:
    google_sheet.update_acell('I129', 'Failed')

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
google_sheet.update_acell('I130', 'OK')
print("자동화상품003(개별-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# [개별출고지시] 버튼 클릭
try:
    driver.find_element(By.XPATH, '추후 입력 예정').click() # [개별출고지시] 버튼 클릭
except:
    print("try try try try try try")
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeSmall.MuiButton-containedSizeSmall.MuiButtonBase-root')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        if wms_str_loop.text == "개별출고지시":
            #print("클릭 시도\n")
            wms_str_loop.click()
            google_sheet.update_acell('I131', 'OK') 
            break
            #print("클릭 완료\n")

time.sleep(2)


alert = driver.switch_to.alert 
alert.accept() # 얼럿 확인
google_sheet.update_acell('I132', 'OK') 
time.sleep(3)
# alert.dismiss()# 취소



#########################################################################################################################
##### 출고 관리 - 출고 현황 조회 진행 #####
time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 관리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)
'''
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 대상 리스트 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break
'''
time.sleep(2)

# 출고 관리 -> 출고 현황 조회 이동(230201)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 현황 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

# wms_test_result = driver.find_element(By.CSS_SELECTOR,'input.PrivateSwitchBase-input[value="7"]')
# wms_test_result.click()

time.sleep(2)
try:
    driver.find_element(By.XPATH, '추후 입력 예정').click() # [검색] 버튼 클릭
except:
    print("try try try try try try")
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeSmall.MuiButton-containedSizeSmall.MuiButtonBase-root')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        if wms_str_loop.text == "검색":
            #print("클릭 시도\n")
            wms_str_loop.click()
            #print("클릭 완료\n")

time.sleep(2)

print("##########")
print("출고 관리 - 출고 현황 조회 이동")
google_sheet.update_acell('I74', 'OK') 


################################################
## 열(컬럼) 메뉴 - 특정 컬럼만 선택
wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break

time.sleep(3)




driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)

driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)


# 배송요청번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("배송요청번호") # 컬럼 검색 필드 - 배송요청번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 배송요청번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# SKU 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("SKU") # 컬럼 검색 필드 - SKU
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # SKU 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 배송타입 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("배송타입") # 컬럼 검색 필드 - 배송타입
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 배송타입 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 진행상태 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("진행상태") # 컬럼 검색 필드 - 진행상태
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 진행상태 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 출고방식 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고방식") # 컬럼 검색 필드 - 출고방식
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 출고방식 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 송장번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("송장번호") # 컬럼 검색 필드 - 송장번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 송장번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 배송요청번호 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H77').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I77', 'Pass')            
else:
    google_sheet.update_acell('I77', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="14"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# SKU 업데이트
wms_sku_barcode_1 = wms_search_list_row_result_check_1
google_sheet.update_acell('I78', 'OK') # SKU 기준 데이터 : 첫 조회
google_sheet.update_acell('H78', wms_sku_barcode_1)
google_sheet.update_acell('H84', wms_sku_barcode_1)
google_sheet.update_acell('H256', wms_sku_barcode_1)
google_sheet.update_acell('H262', wms_sku_barcode_1)
google_sheet.update_acell('H382', wms_sku_barcode_1)


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송타입 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H79').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I79', 'Pass')            
else:
    google_sheet.update_acell('I79', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="20"]') # 진행상태 : 20
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H80').value 
if cell_data == wms_search_list_row_result_check_1: # 진행상태 확인
    google_sheet.update_acell('I80', 'Pass')            
else:
    google_sheet.update_acell('I80', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 출고방식 : 27
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H81').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I81', 'Pass')            
else:
    google_sheet.update_acell('I81', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="30"]') # 송장번호 : 30
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H82').value 
if cell_data == wms_search_list_row_result_check_1: # 송장번호 확인
    google_sheet.update_acell('I82', 'Pass')            
else:
    google_sheet.update_acell('I82', 'Failed')


print("자동화상품001(바로-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_today) 
print("송장번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 배송요청번호 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H83').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I83', 'Pass')            
else:
    google_sheet.update_acell('I83', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="14"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H84').value 
if cell_data == wms_search_list_row_result_check_1: # SKU 확인
    google_sheet.update_acell('I84', 'Pass')            
else:
    google_sheet.update_acell('I84', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송타입 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H85').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I85', 'Pass')            
else:
    google_sheet.update_acell('I85', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="20"]') # 진행상태 : 20
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H86').value 
if cell_data == wms_search_list_row_result_check_1: # 진행상태 확인
    google_sheet.update_acell('I86', 'Pass')            
else:
    google_sheet.update_acell('I86', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 출고방식 : 27
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H87').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I87', 'Pass')            
else:
    google_sheet.update_acell('I87', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="30"]') # 송장번호 : 30
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H88').value 
if cell_data == wms_search_list_row_result_check_1: # 송장번호 확인
    google_sheet.update_acell('I88', 'Pass')            
else:
    google_sheet.update_acell('I88', 'Failed')


print("자동화상품001(바로-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



###################################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_normal) 
print("송장번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 배송요청번호 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H89').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I89', 'Pass')            
else:
    google_sheet.update_acell('I89', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="14"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# SKU 업데이트
wms_sku_barcode_2 = wms_search_list_row_result_check_1
google_sheet.update_acell('I91', 'OK') # SKU 기준 데이터 : 첫 조회
google_sheet.update_acell('H91', wms_sku_barcode_2)
google_sheet.update_acell('H96', wms_sku_barcode_2)
google_sheet.update_acell('H211', wms_sku_barcode_2)
google_sheet.update_acell('H213', wms_sku_barcode_2)
google_sheet.update_acell('H229', wms_sku_barcode_2)
google_sheet.update_acell('H231', wms_sku_barcode_2)
google_sheet.update_acell('H242', wms_sku_barcode_2)
google_sheet.update_acell('H250', wms_sku_barcode_2)
google_sheet.update_acell('H386', wms_sku_barcode_2)



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송타입 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H91').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I91', 'Pass')            
else:
    google_sheet.update_acell('I91', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="20"]') # 진행상태 : 20
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H92').value 
if cell_data == wms_search_list_row_result_check_1: # 진행상태 확인
    google_sheet.update_acell('I92', 'Pass')            
else:
    google_sheet.update_acell('I92', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 출고방식 : 27
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H93').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I93', 'Pass')            
else:
    google_sheet.update_acell('I93', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="30"]') # 송장번호 : 30
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H94').value 
if cell_data == wms_search_list_row_result_check_1: # 송장번호 확인
    google_sheet.update_acell('I94', 'Pass')            
else:
    google_sheet.update_acell('I94', 'Failed')


print("자동화상품002(다스-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



#########################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_today) 
print("송장번호 검색 -> 자동화상품002(다스-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 배송요청번호 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H95').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I95', 'Pass')            
else:
    google_sheet.update_acell('I95', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="14"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H96').value 
if cell_data == wms_search_list_row_result_check_1: # SKU 확인
    google_sheet.update_acell('I96', 'Pass')            
else:
    google_sheet.update_acell('I96', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송타입 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H97').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I97', 'Pass')            
else:
    google_sheet.update_acell('I97', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="20"]') # 진행상태 : 20
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H98').value 
if cell_data == wms_search_list_row_result_check_1: # 진행상태 확인
    google_sheet.update_acell('I98', 'Pass')            
else:
    google_sheet.update_acell('I98', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 출고방식 : 27
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H99').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I99', 'Pass')            
else:
    google_sheet.update_acell('I99', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="30"]') # 송장번호 : 30
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H100').value 
if cell_data == wms_search_list_row_result_check_1: # 송장번호 확인
    google_sheet.update_acell('I100', 'Pass')            
else:
    google_sheet.update_acell('I100', 'Failed')


print("자동화상품002(다스-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



####################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품003(개별-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_normal) 
print("송장번호 검색 -> 자동화상품003(개별-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 배송요청번호 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H101').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I101', 'Pass')            
else:
    google_sheet.update_acell('I101', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="14"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)


# SKU 업데이트
wms_sku_barcode_3 = wms_search_list_row_result_check_1
google_sheet.update_acell('I102', 'OK') # SKU 기준 데이터 : 첫 조회
google_sheet.update_acell('H102', wms_sku_barcode_3)
google_sheet.update_acell('H108', wms_sku_barcode_3)
google_sheet.update_acell('H269', wms_sku_barcode_3)
google_sheet.update_acell('H277', wms_sku_barcode_3)
google_sheet.update_acell('H390', wms_sku_barcode_3)



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송타입 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H103').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I103', 'Pass')            
else:
    google_sheet.update_acell('I103', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="20"]') # 진행상태 : 20
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H104').value 
if cell_data == wms_search_list_row_result_check_1: # 진행상태 확인
    google_sheet.update_acell('I104', 'Pass')            
else:
    google_sheet.update_acell('I104', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 출고방식 : 27
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H105').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I105', 'Pass')            
else:
    google_sheet.update_acell('I105', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="30"]') # 송장번호 : 30
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H106').value 
if cell_data == wms_search_list_row_result_check_1: # 송장번호 확인
    google_sheet.update_acell('I106', 'Pass')            
else:
    google_sheet.update_acell('I106', 'Failed')


print("자동화상품003(개별-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



##############################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품003(개별-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_today) 
print("송장번호 검색 -> 자동화상품003(개별-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 배송요청번호 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H107').value 
if cell_data == wms_search_list_row_result_check_1: # 배송요청번호 확인
    google_sheet.update_acell('I107', 'Pass')            
else:
    google_sheet.update_acell('I107', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="14"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H108').value 
if cell_data == wms_search_list_row_result_check_1: # SKU 확인
    google_sheet.update_acell('I108', 'Pass')            
else:
    google_sheet.update_acell('I108', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송타입 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H109').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I109', 'Pass')            
else:
    google_sheet.update_acell('I109', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="20"]') # 진행상태 : 20
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H110').value 
if cell_data == wms_search_list_row_result_check_1: # 진행상태 확인
    google_sheet.update_acell('I110', 'Pass')            
else:
    google_sheet.update_acell('I110', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="27"]') # 출고방식 : 27
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H111').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I111', 'Pass')            
else:
    google_sheet.update_acell('I111', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="30"]') # 송장번호 : 30
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H112').value 
if cell_data == wms_search_list_row_result_check_1: # 송장번호 확인
    google_sheet.update_acell('I112', 'Pass')            
else:
    google_sheet.update_acell('I112', 'Failed')


print("자동화상품003(개별-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



print("##########")
print("출고 관리 - 출고 현황 조회 완료")



#########################################################################################################################
##### 출고 관리 - 출고 회자 주문별 조회 진행 #####
time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 관리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 회자 주문별 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)



# wms_test_result = driver.find_element(By.CSS_SELECTOR,'input.PrivateSwitchBase-input[value="7"]')
# wms_test_result.click()

time.sleep(2)
try:
    driver.find_element(By.XPATH, '추후 입력 예정').click() # [검색] 버튼 클릭
except:
    print("try try try try try try")
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeSmall.MuiButton-containedSizeSmall.MuiButtonBase-root')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        if wms_str_loop.text == "검색":
            #print("클릭 시도\n")
            wms_str_loop.click()
            #print("클릭 완료\n")

time.sleep(2)

print("##########")
print("출고 관리 - 출고 회자 주문별 조회 이동")
# google_sheet.update_acell('I74', 'OK') 


################################################
## 열(컬럼) 메뉴 - 특정 컬럼만 선택
wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break

time.sleep(3)




driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)

# 지시자ID  3 : wms_login_id
# 출고회차번호  8 : 변수 선언
# 송장번호 16 : 

# 지시자ID 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("지시자ID") # 컬럼 검색 필드 - 지시자ID
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 지시자ID 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 출고회차번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고회차번호") # 컬럼 검색 필드 - 출고회차번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 출고회차번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 송장번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("송장번호") # 컬럼 검색 필드 - 송장번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 송장번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색")
time.sleep(2)




# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="8"]') # 출고회차번호 : 8
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회자 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_now_normal = wms_search_list_row_result_check_1
google_sheet.update_acell('H137', wms_out_round_number_now_normal)
google_sheet.update_acell('H255', wms_out_round_number_now_normal)
google_sheet.update_acell('H286', wms_out_round_number_now_normal)


print("자동화상품001(바로-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_today) 
print("송장번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="8"]') # 출고회차번호 : 8
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회자 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_now_today = wms_search_list_row_result_check_1
google_sheet.update_acell('H148', wms_out_round_number_now_today)
google_sheet.update_acell('H261', wms_out_round_number_now_today)
google_sheet.update_acell('H294', wms_out_round_number_now_today)
print("자동화상품001(바로-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



###################################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_normal) 
print("송장번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="8"]') # 출고회차번호 : 8
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회자 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_das_normal = wms_search_list_row_result_check_1
google_sheet.update_acell('H159', wms_out_round_number_das_normal)
google_sheet.update_acell('H203', wms_out_round_number_das_normal)
google_sheet.update_acell('H238', wms_out_round_number_das_normal)
google_sheet.update_acell('H302', wms_out_round_number_das_normal)

print("자동화상품002(다스-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



#########################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_today) 
print("송장번호 검색 -> 자동화상품002(다스-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="8"]') # 출고회차번호 : 8
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회자 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_das_today = wms_search_list_row_result_check_1
google_sheet.update_acell('H170', wms_out_round_number_das_today)
google_sheet.update_acell('H220', wms_out_round_number_das_today)
google_sheet.update_acell('H246', wms_out_round_number_das_today)
google_sheet.update_acell('H310', wms_out_round_number_das_today)

print("자동화상품002(다스-당일) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



####################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품003(개별-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_normal) 
print("송장번호 검색 -> 자동화상품003(개별-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="8"]') # 출고회차번호 : 8
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회자 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_each_normal = wms_search_list_row_result_check_1
google_sheet.update_acell('H181', wms_out_round_number_each_normal)
google_sheet.update_acell('H318', wms_out_round_number_each_normal)

print("자동화상품003(개별-일반) 데이터 확인 완료")


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



##############################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품003(개별-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_today) 
print("송장번호 검색 -> 자동화상품003(개별-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="8"]') # 출고회차번호 : 8
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회자 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_each_today = wms_search_list_row_result_check_1
google_sheet.update_acell('H192', wms_out_round_number_each_today)
google_sheet.update_acell('H326', wms_out_round_number_each_today)


# 출고테이블 - 테스트 제외
google_sheet.update_acell('I138', "N/A")
google_sheet.update_acell('I287', "N/A")
google_sheet.update_acell('I149', "N/A")
google_sheet.update_acell('I295', "N/A")
google_sheet.update_acell('I160', "N/A")
google_sheet.update_acell('I303', "N/A")
google_sheet.update_acell('I171', "N/A")
google_sheet.update_acell('I311', "N/A")
google_sheet.update_acell('I182', "N/A")
google_sheet.update_acell('I319', "N/A")
google_sheet.update_acell('I193', "N/A")
google_sheet.update_acell('I327', "N/A")



print("자동화상품003(개별-당일) 데이터 확인 완료")

print("##########")
print("출고 관리 - 출고 회자 주문별 조회 완료")




#########################################################################################################################
##### 출고 관리 - 출고 회차별 조회 진행 #####
time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 관리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

# 출고 관리 -> 출고 회차별 조회 이동(230201)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 회차별 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

# wms_test_result = driver.find_element(By.CSS_SELECTOR,'input.PrivateSwitchBase-input[value="7"]')
# wms_test_result.click()

time.sleep(2)
try:
    driver.find_element(By.XPATH, '추후 입력 예정').click() # [검색] 버튼 클릭
except:
    print("try try try try try try")
    wms_test_result = driver.find_elements(By.CSS_SELECTOR,'button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeSmall.MuiButton-containedSizeSmall.MuiButtonBase-root')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        if wms_str_loop.text == "검색":
            #print("클릭 시도\n")
            wms_str_loop.click()
            #print("클릭 완료\n")

time.sleep(2)

print("##########")
print("출고 관리 - 출고 회차별 조회 이동")
google_sheet.update_acell('I133', 'OK') 


################################################
## 열(컬럼) 메뉴 - 특정 컬럼만 선택
wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break

time.sleep(3)




driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
time.sleep(5)


# 지시자ID 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("지시자ID") # 컬럼 검색 필드 - 지시자ID
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 지시자ID 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 배송타입 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("배송타입") # 컬럼 검색 필드 - 배송타입
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 배송타입 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 출고방식 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고방식") # 컬럼 검색 필드 - 출고방식
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 출고방식 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 출고회차번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고회차번호") # 컬럼 검색 필드 - 출고회차번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 출고회차번호 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 출고진행 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고진행") # 컬럼 검색 필드 - 출고진행
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 출고진행 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


# 출고완료 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고완료") # 컬럼 검색 필드 - 출고완료
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 출고완료 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 피킹리스트출력 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("피킹리스트출력") # 컬럼 검색 필드 - 피킹리스트출력
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 피킹리스트출력 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색")
google_sheet.update_acell('I134', 'OK') 
time.sleep(2)



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_now_normal) 
print("출고회차번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H135').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I135', 'Pass')            
else:
    google_sheet.update_acell('I135', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H136').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I136', 'Pass')            
else:
    google_sheet.update_acell('I136', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 출고회차번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H137').value 
if cell_data == wms_search_list_row_result_check_1: # 출고회차번호 확인
    google_sheet.update_acell('I137', 'Pass')            
else:
    google_sheet.update_acell('I137', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H139').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I139', 'Pass')            
else:
    google_sheet.update_acell('I139', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H140').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I140', 'Pass')            
else:
    google_sheet.update_acell('I140', 'Failed')


print("자동화상품001(바로-일반) 데이터 확인 완료")


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        google_sheet.update_acell('I141', 'OK') 

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        google_sheet.update_acell('I142', 'OK') 
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_now_normal = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H143', wms_picking_barcode_now_normal) 
google_sheet.update_acell('H254', wms_picking_barcode_now_normal) 
google_sheet.update_acell('I143', 'OK') 

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H145').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I145', 'Pass')            
else:
    google_sheet.update_acell('I145', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)





# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_now_today) 
print("출고회차번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H146').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I146', 'Pass')            
else:
    google_sheet.update_acell('I146', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H147').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I147', 'Pass')            
else:
    google_sheet.update_acell('I147', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 출고회차번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H148').value 
if cell_data == wms_search_list_row_result_check_1: # 출고회차번호 확인
    google_sheet.update_acell('I148', 'Pass')            
else:
    google_sheet.update_acell('I148', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H150').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I150', 'Pass')            
else:
    google_sheet.update_acell('I150', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H151').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I151', 'Pass')            
else:
    google_sheet.update_acell('I151', 'Failed')


print("자동화상품001(바로-당일) 데이터 확인 완료")


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        google_sheet.update_acell('I152', 'OK') 

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        google_sheet.update_acell('I153', 'OK') 
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_now_today = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H154', wms_picking_barcode_now_today) 
google_sheet.update_acell('H260', wms_picking_barcode_now_today) 
google_sheet.update_acell('I154', 'OK') 

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H156').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I156', 'Pass')            
else:
    google_sheet.update_acell('I156', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_das_normal) 
print("출고회차번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H157').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I157', 'Pass')            
else:
    google_sheet.update_acell('I157', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H158').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I158', 'Pass')            
else:
    google_sheet.update_acell('I158', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 출고회차번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H159').value 
if cell_data == wms_search_list_row_result_check_1: # 출고회차번호 확인
    google_sheet.update_acell('I159', 'Pass')            
else:
    google_sheet.update_acell('I159', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H161').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I161', 'Pass')            
else:
    google_sheet.update_acell('I161', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H162').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I162', 'Pass')            
else:
    google_sheet.update_acell('I162', 'Failed')


print("자동화상품002(다스-일반) 데이터 확인 완료")


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        google_sheet.update_acell('I163', 'OK') 

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        google_sheet.update_acell('I164', 'OK') 
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_das_normal = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H165', wms_picking_barcode_das_normal) 
google_sheet.update_acell('H202', wms_picking_barcode_das_normal) 
google_sheet.update_acell('H216', wms_picking_barcode_das_normal) 
google_sheet.update_acell('H237', wms_picking_barcode_das_normal) 
google_sheet.update_acell('I165', 'OK') 

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H167').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I167', 'Pass')            
else:
    google_sheet.update_acell('I167', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)





# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품002(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_das_today) 
print("출고회차번호 검색 -> 자동화상품002(다스-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H168').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I168', 'Pass')            
else:
    google_sheet.update_acell('I168', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H169').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I169', 'Pass')            
else:
    google_sheet.update_acell('I169', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 출고회차번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H170').value 
if cell_data == wms_search_list_row_result_check_1: # 출고회차번호 확인
    google_sheet.update_acell('I170', 'Pass')            
else:
    google_sheet.update_acell('I170', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H172').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I172', 'Pass')            
else:
    google_sheet.update_acell('I172', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H173').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I173', 'Pass')            
else:
    google_sheet.update_acell('I173', 'Failed')


print("자동화상품002(다스-당일) 데이터 확인 완료")


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        google_sheet.update_acell('I174', 'OK') 

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        google_sheet.update_acell('I175', 'OK') 
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_das_today = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H176', wms_picking_barcode_das_today) 
google_sheet.update_acell('H219', wms_picking_barcode_das_today) 
google_sheet.update_acell('H234', wms_picking_barcode_das_today) 
google_sheet.update_acell('H245', wms_picking_barcode_das_today) 
google_sheet.update_acell('I176', 'OK') 

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H178').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I178', 'Pass')            
else:
    google_sheet.update_acell('I178', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품003(개별-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_each_normal) 
print("출고회차번호 검색 -> 자동화상품003(개별-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H179').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I179', 'Pass')            
else:
    google_sheet.update_acell('I179', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H180').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I180', 'Pass')            
else:
    google_sheet.update_acell('I180', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 출고회차번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H181').value 
if cell_data == wms_search_list_row_result_check_1: # 출고회차번호 확인
    google_sheet.update_acell('I181', 'Pass')            
else:
    google_sheet.update_acell('I181', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H183').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I183', 'Pass')            
else:
    google_sheet.update_acell('I183', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H173').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I184', 'Pass')            
else:
    google_sheet.update_acell('I184', 'Failed')


print("자동화상품003(개별-일반) 데이터 확인 완료")


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        google_sheet.update_acell('I185', 'OK') 

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        google_sheet.update_acell('I186', 'OK') 
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_each_normal = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H187', wms_picking_barcode_each_normal) 
google_sheet.update_acell('I187', 'OK') 

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H189').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I189', 'Pass')            
else:
    google_sheet.update_acell('I189', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품003(개별-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_each_today) 
print("출고회차번호 검색 -> 자동화상품003(개별-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H190').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I190', 'Pass')            
else:
    google_sheet.update_acell('I190', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H191').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I191', 'Pass')            
else:
    google_sheet.update_acell('I191', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 출고회차번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H192').value 
if cell_data == wms_search_list_row_result_check_1: # 출고회차번호 확인
    google_sheet.update_acell('I192', 'Pass')            
else:
    google_sheet.update_acell('I192', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H194').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I194', 'Pass')            
else:
    google_sheet.update_acell('I194', 'Failed')



wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H195').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I195', 'Pass')            
else:
    google_sheet.update_acell('I195', 'Failed')


print("자동화상품003(개별-당일) 데이터 확인 완료")


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        google_sheet.update_acell('I196', 'OK') 

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        google_sheet.update_acell('I197', 'OK') 
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_each_today = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H198', wms_picking_barcode_each_today) 
google_sheet.update_acell('I198', 'OK') 

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H200').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I200', 'Pass')            
else:
    google_sheet.update_acell('I200', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)


print("##########")
print("출고 관리 - 출고 회차별 조회 완료")




#########################################################################################################################
##### 출고 관리 - DAS 피킹 진행 #####
time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 관리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

# 출고 관리 -> DAS 피킹 이동(230201)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "DAS 피킹":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

# wms_test_result = driver.find_element(By.CSS_SELECTOR,'input.PrivateSwitchBase-input[value="7"]')
# wms_test_result.click()


print("##########")
print("출고 관리 - DAS 피킹 이동")
google_sheet.update_acell('I201', 'OK') 


cell_data = google_sheet.acell('H202').value 

print("출고 관리 - DAS 피킹 자동화상품002(다스-일반) 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I202', 'OK') 
time.sleep(2)


# 출고 회자 정보 취득
wms_str_loop_result = ""
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


cell_data = google_sheet.acell('H203').value 
if cell_data == wms_str_loop_result: # 출고 회자 확인
    google_sheet.update_acell('I203', 'Pass')            
else:
    google_sheet.update_acell('I203', 'Failed')


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


cell_data = google_sheet.acell('H204').value 
if cell_data == wms_str_loop_result: # 출고 회자 확인
    google_sheet.update_acell('I204', 'Pass')            
else:
    google_sheet.update_acell('I204', 'Failed')


try:
    # 총 주문수량 확인
    # 총 수량 앞 수량을 가져오는 부분이 클래스 값과 css 값(css-bqeta0)을 가져 오고 있어
    # 혹시 css 값이 변경 될 경우 오류가 날 수 있어 try문으로 작업함
    wms_test_result_text = ""
    
    # 총 주문 수량의 앞 수량 가져오기 ex) 0 / 3 의 앞 숫자 부분(0)
    wms_test_result = driver.find_elements(By.CLASS_NAME,'MuiTypography-root.MuiTypography-body1.css-bqeta0')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        wms_str_loop_text = wms_str_loop.text
    
    wms_test_result_text = wms_str_loop_text
    
    # 총 주문 수량의 뒤 수량 가져오기 ex) 0 / 3 의 앞 숫자 부분(/3)
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
    
    cell_data = google_sheet.acell('H205').value 
    if cell_data == wms_test_result_text: # 총 주문수량 확인
        google_sheet.update_acell('I205', 'Pass')            
    else:
        google_sheet.update_acell('I205', 'Failed')
 
except:
    print("DAS 피킹 - 총 주문 수량 체크 오류 -> except")
    google_sheet.update_acell('I205', 'N/A')
    google_sheet.update_acell('I215', 'N/A')
    pass




# DAS 출고 - 1번째 테이블의 정보 확인자동화상품002(다스-일반))
# 하단의 4개 컬럼의 클래스 다름. 값이 각자 다름 확인해야 함
# 화면에 2개의 테이블이 존재 클래스 값이 중복으로 있을 경우 하단과 같이 작성
wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-range-left.ag-cell-value[col-id="das_cell_no"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H206').value 
    if cell_data == element.text: # DAS 셀번호
        google_sheet.update_acell('I206', 'Pass')            
    else:
        google_sheet.update_acell('I206', 'Failed')    


wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="external_order_id"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H208').value 
    if cell_data == element.text: # 배송요청
        google_sheet.update_acell('I208', 'Pass')            
    else:
        google_sheet.update_acell('I208', 'Failed')    


wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="invoice_delivery_number"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H209').value 
    if cell_data == element.text: # 송장번호
        google_sheet.update_acell('I209', 'Pass')            
    else:
        google_sheet.update_acell('I209', 'Failed')    


wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-range-right.ag-cell-value[col-id="out_stock_count"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H210').value 
    if cell_data == element.text: # 주문수량
        google_sheet.update_acell('I210', 'Pass')            
    else:
        google_sheet.update_acell('I210', 'Failed')    


# SKU 코드 입력하기
try:
    cell_data = google_sheet.acell('H211').value 
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    google_sheet.update_acell('I211', 'OK') 
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    google_sheet.update_acell('I213', 'OK') 
    time.sleep(2)

except:
    print("#SKU 코드 입력하기 오류 발생")
    pass



try:
    # 총 주문수량 확인
    # 총 수량 앞 수량을 가져오는 부분이 클래스 값과 css 값(css-cm47wr)을 가져 오고 있어
    # 혹시 css 값이 변경 될 경우 오류가 날 수 있어 try문으로 작업함
    wms_test_result_text = ""
    
    # 총 주문 수량의 앞 수량 가져오기 ex) 0 / 3 의 앞 숫자 부분(0)
    wms_test_result = driver.find_elements(By.CLASS_NAME,'MuiTypography-root.MuiTypography-body1.css-cm47wr')
    for wms_str_loop in wms_test_result:
        #print(wms_str_loop.text,"\n")
        wms_str_loop_text = wms_str_loop.text
    
    wms_test_result_text = wms_str_loop_text
    
    # 총 주문 수량의 뒤 수량 가져오기 ex) 0 / 3 의 앞 숫자 부분(/3)
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
    
    cell_data = google_sheet.acell('H215').value 
    if cell_data == wms_test_result_text: # 총 주문수량 확인
        google_sheet.update_acell('I215', 'Pass')            
    else:
        google_sheet.update_acell('I215', 'Failed')
 
except:
    print("DAS 피킹 - 총 주문 수량 체크 오류 -> except")
    pass


cell_data = google_sheet.acell('H216').value 

driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I216', 'OK') 
google_sheet.update_acell('I217', 'OK') 
time.sleep(2)

print("출고 관리 - DAS 피킹 자동화상품002(다스-일반) 완료")



cell_data = google_sheet.acell('H219').value 

print("출고 관리 - DAS 피킹 자동화상품002(다스-당일) 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I219', 'OK') 
time.sleep(2)


# 출고 회자 정보 취득
wms_str_loop_result = ""
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
    print("DAS 피킹 - 총 주문 수량 체크 오류 -> except")
    google_sheet.update_acell('I222', 'N/A')
    google_sheet.update_acell('I233', 'N/A')
    pass




# DAS 출고 - 1번째 테이블의 정보 확인자동화상품002(다스-당일))
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
    print("DAS 피킹 - 총 주문 수량 체크 오류 -> except")    
    pass


cell_data = google_sheet.acell('H234').value 

driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I234', 'OK') 
google_sheet.update_acell('I235', 'OK') 
time.sleep(2)

print("출고 관리 - DAS 피킹 자동화상품002(다스-당일) 완료")


print("##########")
print("출고 관리 - DAS 피킹 완료")





#########################################################################################################################
##### 출고 관리 - 출고 처리 - DAS 출고 처리 진행 #####
time.sleep(2)

# 출고 관리 -> DAS 출고 처리 이동(230207)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(5)


wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "DAS 출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

google_sheet.update_acell('I236', 'OK')



print("##########")
print("출고 관리 - 출고 처리 - DAS 출고 처리 이동")


print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(DAS-일반) 시작")
cell_data = google_sheet.acell('H237').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I237', 'OK') 
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





print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(DAS-당일) 시작")
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
#########################################################################################################################
#########################################################################################################################







#########################################################################################################################
##### 출고 관리 - 출고 처리 - 바로 출고 처리 진행 #####
time.sleep(2)

# 출고 관리 -> DAS 출고 처리 이동(230207)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "바로 출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

google_sheet.update_acell('I253', 'OK')



print("##########")
print("출고 관리 - 출고 처리 - 바로 출고 처리 이동")


print("출고 관리 - 출고 처리 - 바로 출고 처리 자동화상품002(바로-일반) 시작")
cell_data = google_sheet.acell('H254').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I254', 'OK') 
time.sleep(2)






# tds = driver.find_elements_by_css_selector('td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall.css-9epf82')
# 출고 회차 확인

cell_data = google_sheet.acell('H255').value 
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
    google_sheet.update_acell('I255', 'Pass')            
else:
    google_sheet.update_acell('I255', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H256').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(cell_data) #  SKU 상품 바코드 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I256', 'OK') 
time.sleep(2)



# 상품 바코드 입력란 > 상품 바코드(SKU) 입력 -> 송장 출력 정보 가져오기
cell_data = google_sheet.acell('H257').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 송장 출력 확인
    google_sheet.update_acell('I257', 'Pass')            
else:
    google_sheet.update_acell('I257', 'Failed')

time.sleep(2)


# 상품 바코드 입력란 > 상품 바코드(SKU) 입력 -> 판매상품명 출력 정보 가져오기
cell_data = google_sheet.acell('H258').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 판매상품명 출력 확인
    google_sheet.update_acell('I258', 'Pass')            
else:
    google_sheet.update_acell('I258', 'Failed')

time.sleep(2)



cell_data = google_sheet.acell('H259').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I259', 'OK') 
time.sleep(2)






print("출고 관리 - 출고 처리 - 바로 출고 처리 자동화상품002(바로-당일) 시작")
cell_data = google_sheet.acell('H260').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I260', 'OK') 
time.sleep(2)



# tds = driver.find_elements_by_css_selector('td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall.css-9epf82')
# 출고 회차 확인

cell_data = google_sheet.acell('H261').value 
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
    google_sheet.update_acell('I261', 'Pass')            
else:
    google_sheet.update_acell('I261', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H262').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(cell_data) #  SKU 상품 바코드 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I262', 'OK') 
time.sleep(2)



# 상품 바코드 입력란 > 상품 바코드(SKU) 입력 -> 송장 출력 정보 가져오기
cell_data = google_sheet.acell('H263').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 송장 출력 확인
    google_sheet.update_acell('I263', 'Pass')            
else:
    google_sheet.update_acell('I263', 'Failed')

time.sleep(2)


# 상품 바코드 입력란 > 상품 바코드(SKU) 입력 -> 판매상품명 출력 정보 가져오기
cell_data = google_sheet.acell('H264').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 판매상품명 출력 확인
    google_sheet.update_acell('I264', 'Pass')            
else:
    google_sheet.update_acell('I264', 'Failed')

time.sleep(2)



cell_data = google_sheet.acell('H265').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
google_sheet.update_acell('I265', 'OK') 
time.sleep(2)























print("출고 일부 테스트 완료")



while(True):
    	pass
