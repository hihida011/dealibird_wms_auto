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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


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
# google_sheet = '2.정상출고' # 구글 시트
google_sheet = '1.정상입고-사입앱' # 구글 시트


google_sheet = google_doc.worksheet(google_sheet)
google_email = 'client_email: fulfillment-test@fulfillment-371610.iam.gserviceaccount.com'

#########################################################################################################################
## 출고 시작 전 구글 시트에 정보 업데이트
# login ID / WMS 로그인 정보

#########################################################################################################################
## 데이터저장을 위한 변수 선언
# 딜리버드 배송요청 번호
deal_ship_now_normal = '216210'  # 자동화상품001(바로-일반)
deal_ship_now_today = '216211'   # 자동화상품001(바로-당일)
deal_ship_das_normal = '216212'  # 자동화상품002(다스-일반)
deal_ship_das_today = '216213'   # 자동화상품002(다스-당일)
deal_ship_each_normal = '216214' # 자동화상품003(개별-일반)
deal_ship_each_today = '216215'  # 자동화상품003(개별-당일)
deal_ship_number_temp = '' # 배송 요청 번호 임시 저장

# 배송 방법
deal_ship_type_temp = ''   # 배송 방법 임시 저장

# 딜리버드 송장번호
deal_invoice_number_now_normal = '570408998925'  # 자동화상품001(바로-일반)
deal_invoice_number_now_today = '077300040033'   # 자동화상품001(바로-당일)
deal_invoice_number_das_normal  = '570408998936'  # 자동화상품002(다스-일반)
deal_invoice_number_das_today = '077300040034'   # 자동화상품002(다스-당일)
deal_invoice_number_each_normal = '570408998940' # 자동화상품003(개별-일반)
deal_invoice_number_each_today = '077300040035'  # 자동화상품003(개별-당일)


# WMS - 출고 관리 - 출고 회자 주문별 조회
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
wms_out_round_number_now_normal = '19'  # 자동화상품001(바로-일반)
wms_out_round_number_now_today = '20'   # 자동화상품001(바로-당일)
wms_out_round_number_das_normal  = '21'  # 자동화상품002(다스-일반)
wms_out_round_number_das_today = '22'   # 자동화상품002(다스-당일)
wms_out_round_number_each_normal = '23' # 자동화상품003(개별-일반)
wms_out_round_number_each_today = '24'  # 자동화상품003(개별-당일)


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
time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 관리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

# time.sleep(30)
print("google_sheet API 제한으로 30초 딜레이")

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 대상 리스트 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)


try:
    # 모든 ag-icon ag-icon-cancel 버튼을 찾음
    cancel_buttons = driver.find_elements(By.CSS_SELECTOR, "span.ag-icon.ag-icon-cancel")

    # 모든 ag-icon ag-icon-cancel 버튼을 클릭
    for cancel_button in cancel_buttons:
        cancel_button.click()
        break

    cancel_buttons = driver.find_elements(By.CSS_SELECTOR, "span.ag-icon.ag-icon-cancel")

    # 모든 ag-icon ag-icon-cancel 버튼을 클릭
    for cancel_button in cancel_buttons:
        cancel_button.click()
        break


    cancel_buttons = driver.find_elements(By.CSS_SELECTOR, "span.ag-icon.ag-icon-cancel")

    # 모든 ag-icon ag-icon-cancel 버튼을 클릭
    for cancel_button in cancel_buttons:
        cancel_button.click()
        break# 모든 ag-icon ag-icon-cancel 버튼을 클릭


    cancel_buttons = driver.find_elements(By.CSS_SELECTOR, "span.ag-icon.ag-icon-cancel")

    for cancel_button in cancel_buttons:
        cancel_button.click()
        break


    cancel_buttons = driver.find_elements(By.CSS_SELECTOR, "span.ag-icon.ag-icon-cancel")

    for cancel_button in cancel_buttons:
        cancel_button.click()
        break

except:
    pass

'''
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.ag-column-drop-cell-text.ag-column-drop-horizontal-cell-text')
wms_test_result_2 = driver.find_elements(By.CSS_SELECTOR,'span.ag-icon.ag-icon-cancel')

for wms_str_loop in wms_test_result:
    for wms_str_loop_2 in wms_test_result_2: 
        wms_str_loop_2.click()
        time.sleep(3)
'''    

print("출고 관리 - 출고 대상 리스트 조회 이동")
# google_sheet.update_acell('I113', 'OK')
print("113 PASS")


# google_sheet.update_acell('I132', 'OK')
print("132 PASS")
time.sleep(10)
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

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 대상 리스트 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break



















while(True):
    	pass
