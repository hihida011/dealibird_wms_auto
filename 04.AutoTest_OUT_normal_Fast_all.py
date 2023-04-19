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
from selenium.common.exceptions import TimeoutException


import tkinter as tk
from tkinter import filedialog

from openpyxl.styles import Font

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

#########################################################################################################################
## 출고에 전달할 정보 : 원본 파일에서 데이터 가져오기
in_to_out_excel_path = r"c:\test\info_in_to_out.xlsx" # Excel 파일 경로 설정
try:
    info_in_to_out_workbook = openpyxl.load_workbook(in_to_out_excel_path)
    info_in_to_out_sheet = info_in_to_out_workbook.active
    
    deal_product_id_1 = info_in_to_out_sheet['B1'].value  # 딜리버드 상품번호(자동화상품 001)
    deal_product_id_2 = info_in_to_out_sheet['B2'].value  # 딜리버드 상품번호(자동화상품 002)
    deal_product_id_3 = info_in_to_out_sheet['B3'].value  # 딜리버드 상품번호(자동화상품 003)

    deal_seller_name =  info_in_to_out_sheet['B4'].value   # 딜리버드 셀러명(대포자)

    deal_sell_product_name_1 = info_in_to_out_sheet['B5'].value # 딜리버드 판매 상품명(자동화상품 001)
    deal_sell_product_name_2 = info_in_to_out_sheet['B6'].value # 딜리버드 판매 상품명(자동화상품 002)
    deal_sell_product_name_3 = info_in_to_out_sheet['B7'].value # 딜리버드 판매 상품명(자동화상품 003)
except FileNotFoundError:
    # 파일이 없을 경우 오류 발생
    print("c:\test\info_in_to_out.xlsx 경로에 파일 정보를 확인하세요")
    exit


## 계정 정보 입력
info_file_path = 'C:\\test\\info.xlsx'
try:
    info_workbook = openpyxl.load_workbook(info_file_path)
    info_sheet = info_workbook.active
    
    deal_admin_login_id = info_sheet['A1'].value
    deal_admin_login_password = info_sheet['A2'].value

    deal_seller_login_id = info_sheet['A3'].value
    deal_seller_login_password = info_sheet['A4'].value

    wms_login_id = info_sheet['A5'].value
    wms_login_passWord = info_sheet['A6'].value

except FileNotFoundError:
    # 파일이 없을 경우 오류 발생
    print("c:\test\info.xlsx 경로에 파일 정보를 확인하세요")
    exit


# 불러오기 창 생성
root = tk.Tk()
root.withdraw()
deal_out_excel_upload_name = filedialog.askopenfilename()


#########################################################################################################################
# 딜리버드 테스트 기본 설정
deal_admin_url = 'https://dealibird.qa.sinsang.market/ssm_admins/sign_in'
deal_seller_url = 'https://vat.qa.sinsang.market/'

# WMS 테스트 기본 설정
wms_url = 'https://matrix-web.qa.sinsang.market/signin'

#########################################################################################################################
# 구글 시트 연동
font_color = Font(color='FFFFFF')
# 구글 json 경로
json_file_name_cell = info_sheet['A7']

if json_file_name_cell.value is None:
    json_file_name = filedialog.askopenfilename()    
    json_file_name_cell.value = json_file_name
    info_sheet['A7'].font = font_color
    info_workbook.save(info_file_path)
    # print("1번째", deal_admin_login_id)
else:
    json_file_name = json_file_name_cell.value
    #print("2번째", deal_admin_login_id)


scope = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive',
]

gc = gspread.service_account(filename=json_file_name)


# 구글 테스트 시나리오 엑셀 주소
google_url_cell = info_sheet['A8']

if google_url_cell.value is None:
    google_url_input = (input("구글 테스트 시나리오 엑셀 주소: "))
    google_url_cell.value = google_url_input
    info_sheet['A8'].font = font_color
    info_workbook.save(info_file_path)
    google_url = google_url_input
    #print("1번째", deal_seller_login_password)

else:
    google_url = google_url_cell.value
    #print("2번째", deal_seller_login_password)


# 구글 API 사용 이메일
google_email_cell = info_sheet['A9']

if google_email_cell.value is None:
    google_email_input = (input("구글 API 사용 이메일: "))
    google_email_cell.value = google_email_input
    info_sheet['A9'].font = font_color
    info_workbook.save(info_file_path)
    google_email = google_email_input
    #print("1번째", deal_seller_login_password)
else:
    google_email = google_email_cell.value

google_doc = gc.open_by_url(google_url)


# 구글 시트
google_sheet_cell = info_sheet['A11']

if google_sheet_cell.value is None:
    google_sheet_input = (input("구글 시트: "))
    google_sheet_cell.value = google_sheet_input
    info_sheet['A11'].font = font_color
    info_workbook.save(info_file_path)
    google_sheet = google_sheet_input
    #print("1번째", deal_seller_login_password)
else:
    google_sheet = google_sheet_cell.value

# google_sheet = '1.정상입고-사입앱' # 구글 시트


google_sheet = google_doc.worksheet(google_sheet)


#########################################################################################################################
## 데이터저장을 위한 변수 선언
# 딜리버드 배송요청 번호
deal_ship_now_normal = ''  # 자동화상품001(바로-일반)
deal_ship_now_today = ''   # 자동화상품001(바로-당일)
deal_ship_das_normal = ''  # 자동화상품002(다스-일반)
deal_ship_das_today = ''   # 자동화상품002(다스-당일)
deal_ship_each_normal = '' # 자동화상품003(개별-일반)
deal_ship_each_today = ''  # 자동화상품003(개별-당일)
deal_ship_number_temp = '' # 배송 요청 번호 임시 저장


# 딜리버드 송장번호
deal_invoice_number_now_normal = ''  # 자동화상품001(바로-일반)
deal_invoice_number_now_today = ''   # 자동화상품001(바로-당일)
deal_invoice_number_das_normal  = ''  # 자동화상품002(다스-일반)
deal_invoice_number_das_today = ''   # 자동화상품002(다스-당일)
deal_invoice_number_each_normal = '' # 자동화상품003(개별-일반)
deal_invoice_number_each_today = ''  # 자동화상품003(개별-당일)


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

driver.switch_to.window(tabs[2])
driver.get(wms_url)

time.sleep(2)

# wms 각 항목 -> 조회 -> 리스트의 스크롤이 생기면서 데이터 로드의 어려움(현재 보여지는 화면의 데이터만 가져옴)
# pyautogui을 사용 -> wms 로그인 화면에서 글꼴 축소 -> 항목 조회 -> 리스트의 모든 데이터가 보이게 됨

#pyautogui_count = 0
#while pyautogui_count < 8:
#    pyautogui.hotkey('ctrl', '-') # 화면 축소
#    pyautogui_count = pyautogui_count + 1

#time.sleep(2)
#action = ActionChains(driver)


# hihida 221230
#########################################################################################################################
########### 어드민 -> 딜리버드 셀러 이동 ##########
#########################################################################################################################
driver.switch_to.window(tabs[0])

##신상마켓 소매 -> 딜리버드 이동
#신상마켓 로그인
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/header/div/div[2]/div[3]').click() # 로그인 버튼(페이지 상단)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/div[1]/input').send_keys(deal_seller_login_id) # 모달 / ID 입력
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/div[2]/input').send_keys(deal_seller_login_password) # 모달 / 비번 입력
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/div[3]/div[2]/div/button').click() # 모달 / 로그인 버튼
print("##########")
print("신상마켓 소매 로그인")

time.sleep(3)
# 팝업 광고 있을 경우 닫기 클릭
try:
    print("팝업 광고 시작")
    deal_test_result = driver.find_element(By.CLASS_NAME, "close-button")
    #deal_test_result = driver.find_element(By.CSS_SELECTOR, 'div[class="button.close-button.full-btn"]')
    #deal_test_result = driver.find_element(By.CSS_SELECTOR, ".popup-footer__button-area .close-button")
    print("팝업 광고 시작2")
    deal_test_result.click()
    print("팝업 광고 시작3")
except:
    print("팝업 광고 오류")
    pass

#딜리버드 바로가기 클릭
time.sleep(5)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/div[1]/div/ul/li[1]/div/span').click()
print("##########")
print("딜리버드 이동")
# print(" 5PASS")

#########################################################################################################################
#### 딜리버드 -> 배송요청 #####
time.sleep(5)
try:
    driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[5]/a').click()
except:
    deal_link = driver.find_element(By.LINK_TEXT, ("배송요청"))
    deal_link.click()
    
print("6 PASS")
time.sleep(5)
print("##########")
print("배송 요청 시작")



# 기존에 등록되어 있던 부분이 있을 경우 [요청 전체 삭제] 버튼 클릭 해야함
deal_test_result = driver.find_element(By.XPATH,'//*[@id="ready_total_count"]') # 페이지 상단 오른쪽 -> 배송 요청 가능(총 x건) : X 값
deal_test_result_check = deal_test_result.text # 배송 요청 값에서 테스트 값을 저장

if deal_test_result_check != "": # 배송 요청 건수가 0건일 경우 실행 되지 않음
    print("기존에 요청했던 배송 요청 가능 건수가 있어 삭제 합니다.")
    driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]/div[1]/div[2]/div/button[3]').click() # 페이지 중간의 [입력초기화] 버튼
    time.sleep(3)

    driver.find_element(By.XPATH, '/html/body/div[5]/div/div[3]/button[3]').click() # 얼럿 -> 모두 초기화 ~ [예] 버튼
    time.sleep(3)
    
    driver.find_element(By.XPATH, '/html/body/div[5]/div/div[3]/button[1]').click() # 얼럿 -> X건 삭제 완료 : [OK] 버튼
    time.sleep(2)


# 엑셀 업로드 버튼 클릭
driver.find_element(By.CLASS_NAME, 'btn.btn-outline-secondary.order_excel_upload').click()
print("업로드 클릭 시작\n")
print("7 PASS")


print("파일 업로드 선택 완료\n")
print("9 PASS")


#파일 업로드
up_load_file = driver.find_element(By.XPATH, '//*[@id="orders"]') # 모달 / 엑셀 업로드 양식 선태 -> 엑셀 파일 선택 Browse 버튼
up_load_file.send_keys(deal_out_excel_upload_name)
time.sleep(3)
driver.find_element(By.ID, 'excel_order_import_btn').click() # [업로드] 버튼

print("10 PASS")
print("업로드 클릭 완료\n")


#########################################################################################################################
#### 딜리버드 -> 배송요청 #####
time.sleep(5)

try:
    driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[5]/a').click()
except:
    deal_link = driver.find_element(By.LINK_TEXT, ("배송요청"))
    deal_link.click()
    

time.sleep(5)
# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로


# HIHIDA 230308
# 230216 eof를 잡기 위해 개발팀에서 스레트를 1개에서 3개로 변경
# 당일 출고 요청시 : 1초에 2번 요청하면 중복 송장 발생으로 오류 발생
# 전체 출고 요청에서 각 출고요청 건수마다 [배송 요청하기] 버튼 클릭
# 6개의 행에서 순차적으로 버튼을 실행하고자 했으나
# 0번째 배송 요청하기 버튼을 클릭하면 행에서 사라지게 되는 현상으로
# 셀레니움에서 0번째 버튼이 없어져 오류 발생
# 그래서 버튼 객체를 불러오고 한번씩 6번 실행으로 작성
deal_tbody_int = int(0)
deal_tbody_row = int(7)


try:
    while deal_tbody_row >= deal_tbody_int : # 리스트 row 수 만큼 실행
        print("while : ", deal_tbody_int)
        deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
        order_request_buttons = deal_tbody.find_elements(By.CLASS_NAME, 'order_request')
        # 배송 요청하기 버튼 클릭
        for button in order_request_buttons:
            print("for : ", deal_tbody_int)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'order_request')))
            button.click()
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm')))
            driver.find_element(By.CLASS_NAME, 'swal2-confirm').send_keys(Keys.ENTER)
            time.sleep(5) 
            break
        deal_tbody_int = deal_tbody_int +1   
        print("deal_tbody_int = deal_tbody_int +1 : ", deal_tbody_int)
        time.sleep(5)  

except:
    pass


print("40 PASS")

#########################################################################################################################
#### 딜리버드 -> 배송현황 #####
time.sleep(5)

try:
    driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[6]/a').click()
except:
    deal_link = driver.find_element(By.LINK_TEXT, ("배송현황"))
    deal_link.click()
    
time.sleep(3)

driver.find_element(By.XPATH,'//*[@id="page-wrapper"]/div[2]/div[1]/div[1]/ul[1]/li[3]/a').click()
print("41 PASS")
print("##########")
print("배송현황 시작")
time.sleep(2)


# 드롭다운 메뉴에서 [배송 요청 번호]선택 
deal_test_result = Select(driver.find_element(By.XPATH,'//*[@id="search_field"]'))
deal_test_result.select_by_visible_text("배송 요청 번호")
time.sleep(2)


## 자동화상품001(바로-일반) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품001(바로-일반)")
cell_data = deal_ship_now_normal # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
print("42 PASS")
time.sleep(2)
print("배송 요청 번호 검색 완료 : # 자동화상품001(바로-일반)")


print("배송 현황 리스트(테이블) 체크 시작 : 자동화상품001(바로-일반)")
# 배송 현황 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 자동화상품001(바로-일반) - 배송상태, 송장번호 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        # print("배송 현황 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 7: # 송장번호
            deal_invoice_number_now_normal = td.get_attribute("innerText")
            print("44 PASS")
            break
        
        deal_list_count += 1
print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품001(바로-일반)")

deal_name_search_sendkeys(driver, "searchWord")
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




## 자동화상품001(바로-당일) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품001(바로-당일)")
cell_data = deal_ship_now_today # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
print("45 PASS")
time.sleep(2)
print("배송 요청 번호 검색 완료 : # 자동화상품001(바로-당일)")


print("배송 현황 리스트(테이블) 체크 시작 : 자동화상품001(바로-당일)")
# 배송 현황 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 자동화상품001(바로-당일) - 배송상태, 송장번호 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        # print("배송 현황 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        if deal_list_count == 7: # 송장번호
            deal_invoice_number_now_today = td.get_attribute("innerText")
            print("47 PASS")
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품001(바로-당일)")

deal_name_search_sendkeys(driver, "searchWord")
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)



## 자동화상품002(다스-일반) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품002(다스-일반)")
cell_data = deal_ship_das_normal # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
print("48 PASS")
time.sleep(2)
print("배송 요청 번호 검색 완료 : # 자동화상품002(다스-일반)")


print("배송 현황 리스트(테이블) 체크 시작 : 자동화상품002(다스-일반)")
# 배송 현황 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 자동화상품002(다스-일반) - 배송상태, 송장번호 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        # print("배송 현황 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        if deal_list_count == 7: # 송장번호
            deal_invoice_number_das_normal = td.get_attribute("innerText")
            print("50 PASS")
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품002(다스-일반)")

deal_name_search_sendkeys(driver, "searchWord")
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




## 자동화상품002(다스-당일) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품002(다스-당일)")
cell_data = deal_ship_das_today # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
print("51 PASS")
time.sleep(2)
print("배송 요청 번호 검색 완료 : # 자동화상품002(다스-당일)")


print("배송 현황 리스트(테이블) 체크 시작 : 자동화상품002(다스-당일)")
# 배송 현황 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 자동화상품002(다스-당일) - 배송상태, 송장번호 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        if deal_list_count == 7: # 송장번호
            deal_invoice_number_das_today = td.get_attribute("innerText")
            print("53 PASS")
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품002(다스-당일)")

deal_name_search_sendkeys(driver, "searchWord")
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)





## 자동화상품003(개별-일반) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품003(개별-일반)")
cell_data = deal_ship_each_normal # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
print("54 PASS")
time.sleep(2)
print("배송 요청 번호 검색 완료 : # 자동화상품003(개별-일반)")


print("배송 현황 리스트(테이블) 체크 시작 : 자동화상품003(개별-일반)")
# 배송 현황 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 자동화상품003(개별-일반) - 배송상태, 송장번호 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        if deal_list_count == 7: # 송장번호
            deal_invoice_number_each_normal = td.get_attribute("innerText")
            print("56 PASS")
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품003(개별-일반)")

deal_name_search_sendkeys(driver, "searchWord")
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




## 자동화상품003(개별-당일) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품003(개별-당일)")
cell_data = deal_ship_each_today # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
print("57 PASS")
time.sleep(2)
print("배송 요청 번호 검색 완료 : # 자동화상품003(개별-당일)")


print("배송 현황 리스트(테이블) 체크 시작 : 자동화상품003(개별-당일)")
# 배송 현황 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 자동화상품003(개별-당일) - 배송상태, 송장번호 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        # print("배송 현황 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        if deal_list_count == 7: # 송장번호
            deal_invoice_number_each_today = td.get_attribute("innerText")
            print("59 PASS")
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품003(개별-당일)")

deal_name_search_sendkeys(driver, "searchWord")
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




# hihida 230206
# hihida 221230
#########################################################################################################################
##어드민 -> WMS 이동
#########################################################################################################################
driver.switch_to.window(tabs[2])


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



wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 대상 리스트 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)



print("출고 관리 - 출고 대상 리스트 조회 이동")
print("113 PASS")


# 230213리스트 상단에 선택되어 있는 칼럼들 삭제
# for문으로 돌려서 순차적으로 종료 하고 싶었지만, StaleElementReferenceException 오류 벌생
# 간단하게 첫번째를 여러번 돌리는 방법으로 긴급하게 수정
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

time.sleep(10)


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


# 체크박스 컬럼 선택
wms_test_result = driver.find_element(By.CSS_SELECTOR, "div>[aria-posinset='1']") # 체크박스 컬럼의 유일값을 찾음, aria-posinset='2'
time.sleep(5)

wms_test_result_chk = wms_test_result.get_attribute("aria-describedby") # 체크박스 유일값으로 동적으로 변화되는 aria-describedby ID값 취득
time.sleep(5)

wms_test_result.find_element(By.ID, wms_test_result_chk).click()


# 송장번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("송장번호") # 컬럼 검색 필드 - 송장번호
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)


wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break

cell_data = deal_seller_name 

driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(cell_data) #  셀러명(대표자) 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
print("114 PASS")
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(5)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result_chk = wms_test_result.find_element(By.CLASS_NAME, 'ag-input-field-input.ag-checkbox-input')
time.sleep(5)
wms_test_result_chk.click()
print("116 PASS")
print("자동화상품001(바로-일반) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_today) 
print("송장번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)

# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-당일))
#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result_chk = wms_test_result.find_element(By.CLASS_NAME, 'ag-input-field-input.ag-checkbox-input')
time.sleep(5)
wms_test_result_chk.click()
print("118 PASS")
print("자동화상품001(바로-당일) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_normal) 
print("송장번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)

# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-일반))
#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result_chk = wms_test_result.find_element(By.CLASS_NAME, 'ag-input-field-input.ag-checkbox-input')
time.sleep(5)
wms_test_result_chk.click()

#wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
#wms_test_result_chk = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
#time.sleep(5)
#wms_test_result.find_element(By.ID, wms_test_result_chk).click()

#wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
#wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
#wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
#driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
print("120 PASS")
print("자동화상품002(다스-일반) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_today) 
print("송장번호 검색 -> 자동화상품002(다스-당일)")
time.sleep(2)

#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result_chk = wms_test_result.find_element(By.CLASS_NAME, 'ag-input-field-input.ag-checkbox-input')
time.sleep(5)
wms_test_result_chk.click()

print("122 PASS")
print("자동화상품002(다스-당일) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
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
            print("123 PASS")
            break
            #print("클릭 완료\n")

time.sleep(10)


alert = driver.switch_to.alert 
alert.accept() # 얼럿 확인
print("124 PASS")
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
            print("125 PASS")
            break
            #print("클릭 완료\n")

time.sleep(15)


try:
    alert = driver.switch_to.alert
    alert.accept() # 얼럿 확인
    time.sleep(3)
except:
    print("try try try try try try")
    pass


try:
    # 일반, DAS 의 선택 출고 지시 이 후
    # 개별 출고 항목 2개가 남았는데, 체크박스가 체크 되어 있을 경우 체크 해제
    # 전체 체크가 되어 있을 경우 ;전체 행 선택 (체크됨)'의 값을 가져 올 수 있음
    # 가져 오지 못한다면 except 이동
    wms_test_result = driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 행 선택 (체크됨)']")    
    print("일반, DAS 의 선택 출고 지시 이 후 전체 체크박스 선택 되어 있는 부분 체크 해제 시작")
    #리스트의 전체 체크 박스 선택
    wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
    wms_test_result_chk = wms_test_result.find_element(By.CLASS_NAME, 'ag-input-field-input.ag-checkbox-input')
    time.sleep(5)
    wms_test_result_chk.click()
    
    print("일반, DAS 의 선택 출고 지시 이 후 전체 체크박스 선택 되어 있는 부분 체크 해제 완료")
except:
    pass

time.sleep(15)

# 테이블(리스트) -> 송장번호 검색 자동화상품003(개별-일반)-> 
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result_chk = wms_test_result.find_element(By.CLASS_NAME, 'ag-input-field-input.ag-checkbox-input')
time.sleep(5)
wms_test_result_chk.click()
print("128 PASS")
print("자동화상품003(개별-일반) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 송장번호 검색 자동화상품003(개별-당일)-> 
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_today) 
print("송장번호 검색 -> 자동화상품003(개별-당일)")
time.sleep(2)

# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-당일))
#리스트의 전체 체크 박스 선택
wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-header-cell.ag-header-cell-sortable.ag-focus-managed') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result_chk = wms_test_result.find_element(By.CLASS_NAME, 'ag-input-field-input.ag-checkbox-input')
time.sleep(5)
wms_test_result_chk.click()

print("130 PASS")
print("자동화상품003(개별-당일) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
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
            print("131 PASS")
            break
            #print("클릭 완료\n")

time.sleep(10)


# 1번째 alert 처리
try:
    print("1번째 alert 처리 시작")
    alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
    alert_text = alert.text
    alert.accept()
    print("1번째 alert 처리 완료")
except TimeoutException:
    pass

time.sleep(5)
# 2번째 alert 처리
try:
    print("2번째 alert 처리 시작")
    alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
    alert_text = alert.text
    alert.accept()
    print("2번째 alert 처리 완료")
except TimeoutException:
    pass

    

print("132 PASS")




#########################################################################################################################
##### 출고 관리 - 출고 현황 조회 진행 #####
time.sleep(5)

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

time.sleep(3)



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


# 230213리스트 상단에 선택되어 있는 칼럼들 삭제
# for문으로 돌려서 순차적으로 종료 하고 싶었지만, StaleElementReferenceException 오류 벌생
# 간단하게 첫번째를 여러번 돌리는 방법으로 긴급하게 수정
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


# SKU 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("SKU") # 컬럼 검색 필드 - SKU
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='필터 컬럼 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


# 송장번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("송장번호") # 컬럼 검색 필드 - 송장번호
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)

wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="skuId"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# SKU 업데이트
wms_sku_barcode_1 = wms_search_list_row_result_check_1
print("78 PASS")# SKU 기준 데이터 : 첫 조회
time.sleep(5)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_today) 
print("송장번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)


print("자동화상품001(바로-당일) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



###################################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_normal) 
print("송장번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="skuId"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# SKU 업데이트
wms_sku_barcode_2 = wms_search_list_row_result_check_1
print("91 PASS")# SKU 기준 데이터 : 첫 조회
print("자동화상품002(다스-당일) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



####################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품003(개별-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_normal) 
print("송장번호 검색 -> 자동화상품003(개별-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="skuId"]') # SKU : 14
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 현황 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)


# SKU 업데이트
wms_sku_barcode_3 = wms_search_list_row_result_check_1
print("102 PASS")# SKU 기준 데이터 : 첫 조회

print("##########")
print("출고 관리 - 출고 현황 조회 완료")



#########################################################################################################################
##### 출고 관리 - 출고 회차 주문별 조회 진행 #####
time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 관리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(3)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 회차 주문별 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

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
print("출고 관리 - 출고 회차 주문별 조회 이동")

# 230213리스트 상단에 선택되어 있는 칼럼들 삭제
# for문으로 돌려서 순차적으로 종료 하고 싶었지만, StaleElementReferenceException 오류 벌생
# 간단하게 첫번째를 여러번 돌리는 방법으로 긴급하게 수정
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

# 지시자ID  3 : wms_login_id
# 출고회차번호  8 : 변수 선언
# 송장번호 16 : 

# 지시자ID 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("지시자ID") # 컬럼 검색 필드 - 지시자ID
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)

wms_css_selector_input_sendkeys(driver, "div>input[aria-label='필터 컬럼 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
#time.sleep(5)



# 출고회차번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고회차번호") # 컬럼 검색 필드 - 출고회차번호
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='필터 컬럼 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


# 송장번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("송장번호") # 컬럼 검색 필드 - 송장번호
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)

wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break


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


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="wave_seq"]') # 출고회차번호
#wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="waveSeq"]') # 출고회차번호
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_now_normal = wms_search_list_row_result_check_1

print("자동화상품001(바로-일반) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_today) 
print("송장번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="wave_seq"]') # 출고회차번호
#wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="waveSeq"]') # 출고회차번호
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_now_today = wms_search_list_row_result_check_1

print("자동화상품001(바로-당일) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



###################################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_normal) 
print("송장번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="wave_seq"]') # 출고회차번호
#wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="waveSeq"]') # 출고회차번호
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_das_normal = wms_search_list_row_result_check_1


print("자동화상품002(다스-일반) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



#########################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품002(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_das_today) 
print("송장번호 검색 -> 자동화상품002(다스-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품002(다스-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="wave_seq"]') # 출고회차번호
#wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="waveSeq"]') # 출고회차번호
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_das_today = wms_search_list_row_result_check_1


print("자동화상품002(다스-당일) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



####################################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품003(개별-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_normal) 
print("송장번호 검색 -> 자동화상품003(개별-일반)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="wave_seq"]') # 출고회차번호
#wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="waveSeq"]') # 출고회차번호
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_each_normal = wms_search_list_row_result_check_1


print("자동화상품003(개별-일반) 데이터 확인 완료")


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='송장번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



##############################
# 테이블(리스트) -> 송장번호 검색 -> 자동화상품003(개별-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_today) 
print("송장번호 검색 -> 자동화상품003(개별-당일)")
time.sleep(2)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')


wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="wave_seq"]') # 출고회차번호
#wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="waveSeq"]') # 출고회차번호
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차 주문별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

# 출고회차번호 업데이트
wms_out_round_number_each_today = wms_search_list_row_result_check_1

print("자동화상품003(개별-당일) 데이터 확인 완료")
print("##########")
print("출고 관리 - 출고 회차 주문별 조회 완료")





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

time.sleep(3)

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
print("133 PASS")


# 230213리스트 상단에 선택되어 있는 칼럼들 삭제
# for문으로 돌려서 순차적으로 종료 하고 싶었지만, StaleElementReferenceException 오류 벌생
# 간단하게 첫번째를 여러번 돌리는 방법으로 긴급하게 수정
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


# 지시자ID 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("지시자ID") # 컬럼 검색 필드 - 지시자ID
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='필터 컬럼 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)



# 출고회차번호 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("출고회차번호") # 컬럼 검색 필드 - 출고회차번호
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)


wms_css_selector_input_sendkeys(driver, "div>input[aria-label='필터 컬럼 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


# 피킹리스트출력 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("피킹리스트출력") # 컬럼 검색 필드 - 피킹리스트출력
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.ENTER)
time.sleep(2)
#time.sleep(5)

wms_side_columns_button = driver.find_elements(By.CLASS_NAME, 'ag-side-button-label')
#print("wms_side_columns_button", wms_side_columns_button)
for wms_side_button_ck in wms_side_columns_button:
    print("wms_side_button_ck", wms_side_button_ck)
    print("wms_side_button_ck.text",wms_side_button_ck.text)
    if wms_side_button_ck.text == "열(컬럼)":
        print("OK")
        wms_side_button_ck.click()
        break


# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색(바로-일반)")
print("134 PASS")
time.sleep(2)



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_now_normal) 
print("출고회차번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        print("141 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        print("142 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_now_normal = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트

print("143 PASS")

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()

time.sleep(2)

# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색(바로-일반)출고진행")
time.sleep(2)



wms_css_selector_input_sendkeys(driver, "div>input[aria-label='출고회차번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_now_today) 
print("출고회차번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)

# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        print("152 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        print("153 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_now_today = wms_test_result

print("154 PASS")

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()
time.sleep(2)

wms_css_selector_input_sendkeys(driver, "div>input[aria-label='출고회차번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)





# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품002(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_das_normal) 
print("출고회차번호 검색 -> 자동화상품002(다스-일반)")
time.sleep(2)


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        print("163 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        print("164 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_das_normal = wms_test_result

print("165 PASS")

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()
time.sleep(2)


# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색(다스-일반)출고진행")
time.sleep(2)

wms_css_selector_input_sendkeys(driver, "div>input[aria-label='출고회차번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품002(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_das_today) 
print("출고회차번호 검색 -> 자동화상품002(다스-당일)")
time.sleep(2)


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        print("174 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        print("175 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_das_today = wms_test_result
print("176 PASS")

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()
time.sleep(2)

# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색(다스-당일)출고진행")
time.sleep(2)

wms_css_selector_input_sendkeys(driver, "div>input[aria-label='출고회차번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품003(개별-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_each_normal) 
print("출고회차번호 검색 -> 자동화상품003(개별-일반)")
time.sleep(2)


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        print("185 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        print("186 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_each_normal = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트

print("187 PASS")

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()

time.sleep(2)
# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색(개별-일반)출고진행")
time.sleep(2)

wms_css_selector_input_sendkeys(driver, "div>input[aria-label='출고회차번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)





# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품003(개별-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_each_today) 
print("출고회차번호 검색 -> 자동화상품003(개별-당일)")
time.sleep(2)


# 피킹리스트출력 프린트 버튼 클릭
wms_test_result = driver.find_elements(By.CLASS_NAME, 'MuiButtonBase-root')
for button in wms_test_result:
    if "PrintIcon" in button.get_attribute("innerHTML"):
        button.click()
        print("196 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        print("197 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_each_today = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트

print("198 PASS")

# esc 클릭해서 피킹리스트 종료
actions = ActionChains(driver)
actions.send_keys(Keys.ESCAPE)
actions.perform()

time.sleep(2)
# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색(개별-당일)출고진행")
time.sleep(2)

wms_css_selector_input_sendkeys(driver, "div>input[aria-label='출고회차번호 필터 입력']")
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
#time.sleep(2)
#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
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


time.sleep(3)

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

print("##########")
print("출고 관리 - DAS 피킹 이동")
print("201 PASS")


cell_data = wms_picking_barcode_das_normal

print("출고 관리 - DAS 피킹 자동화상품002(다스-일반) 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("202 PASS")
time.sleep(2)

# DAS 피킹 - 1번째 테이블의 정보 확인자동화상품002(다스-일반))

# SKU 코드 입력하기
try:
    cell_data = wms_sku_barcode_2
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    print("211 PASS")
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)
        
    print("213 PASS")
    time.sleep(2)

except:
    print("#SKU 코드 입력하기 오류 발생")
    pass

cell_data = wms_picking_barcode_das_normal

driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("216 PASS")
time.sleep(2)

print("출고 관리 - DAS 피킹 자동화상품002(다스-일반) 완료")



cell_data = wms_picking_barcode_das_today 

print("출고 관리 - DAS 피킹 자동화상품002(다스-당일) 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("219 PASS")
time.sleep(2)


# DAS 피킹 - 1번째 테이블의 정보 확인자동화상품002(다스-당일))
# SKU 코드 입력하기
try:
    cell_data = wms_sku_barcode_2 
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    print("229 PASS")
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)
    
    print("231 PASS")

except:
    print("#SKU 코드 입력하기 오류 발생")
    pass




cell_data = wms_picking_barcode_das_today

driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("234 PASS")
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

time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "DAS 출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)
print("236 PASS")



print("##########")
print("출고 관리 - 출고 처리 - DAS 출고 처리 이동")


print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(DAS-일반) 시작")
cell_data = wms_picking_barcode_das_normal
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("219 PASS")
time.sleep(2)

cell_data = deal_invoice_number_das_normal
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("240 PASS")
time.sleep(2)


cell_data = deal_invoice_number_das_normal
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("243 PASS")
time.sleep(2)



#########################################################################################################################




print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(DAS-당일) 시작")
cell_data = wms_picking_barcode_das_today 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("245 PASS")
time.sleep(2)


cell_data = deal_invoice_number_das_today 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("248 PASS")
time.sleep(2)


cell_data = deal_invoice_number_das_today 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("251 PASS")
time.sleep(2)



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
print("253 PASS")



print("##########")
print("출고 관리 - 출고 처리 - 바로 출고 처리 이동")


print("출고 관리 - 출고 처리 - 바로 출고 처리 자동화상품002(바로-일반) 시작")
cell_data = wms_picking_barcode_now_normal
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("254 PASS")
time.sleep(2)


cell_data = wms_sku_barcode_1
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(cell_data) #  SKU 상품 바코드 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(Keys.ENTER) # 검색 적용
print("256 PASS")
time.sleep(2)


cell_data = deal_invoice_number_now_normal
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("259 PASS")
time.sleep(2)






print("출고 관리 - 출고 처리 - 바로 출고 처리 자동화상품002(바로-당일) 시작")
cell_data = wms_picking_barcode_now_today 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
print("260 PASS")
time.sleep(2)


cell_data = wms_sku_barcode_1
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(cell_data) #  SKU 상품 바코드 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='sku_id']").send_keys(Keys.ENTER) # 검색 적용
print("262 PASS")
time.sleep(2)



cell_data = deal_invoice_number_now_today 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("265 PASS")
time.sleep(2)




#########################################################################################################################
##### 출고 관리 - 출고 처리 - 개별 출고 처리 진행 #####
time.sleep(2)

# 출고 관리 -> DAS 출고 처리 이동(230207)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "개별 출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)
print("266 PASS")



print("##########")
print("출고 관리 - 출고 처리 - 개별 출고 처리 이동")


print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-일반) 시작")

cell_data = deal_invoice_number_each_normal
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("267 PASS")
time.sleep(2)


cell_data = wms_sku_barcode_3
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
print("269 PASS")
time.sleep(2)




# SKU 코드 입력하기
try:
    cell_data = wms_sku_barcode_3 
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    print("272 PASS")
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    
    
    print("272 PASS")
    time.sleep(2)

except:
    print("#SKU 코드 입력하기 오류 발생 또는 횟수 모자람")
    time.sleep(2)
    pass


cell_data = deal_invoice_number_each_normal 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("274 PASS")
print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-일반) 종료")
time.sleep(2)





print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-당일) 시작")

cell_data = deal_invoice_number_each_today
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("275 PASS")
time.sleep(2)


cell_data = wms_sku_barcode_3 
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
print("277 PASS")
time.sleep(2)


# SKU 코드 입력하기
try:
    cell_data = wms_sku_barcode_3
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    print("280 PASS")
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)
    
    print("280 PASS")

except:
    print("#SKU 코드 입력하기 오류 발생 또는 횟수 모자람")
    time.sleep(2)
    pass

cell_data = deal_invoice_number_each_today
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
print("282 PASS")
time.sleep(2)


print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-당일) 종료")


# 배송 완료 진행 여부 확인
Delivery_completed = input("어드민에서 배송 완료를 진행 할 경우 y를 누르세요")
if Delivery_completed != 'y':
    exit()


#########################################################################################################################
########## 딜리버드 -> 어드민 이동 ##########
#########################################################################################################################
driver.switch_to.window(tabs[1])
time.sleep(2)

#어드민 로그인 진행
driver.find_element(By.ID, 'ssm_admin_email').send_keys(deal_admin_login_id)
driver.find_element(By.ID, 'ssm_admin_password').send_keys(deal_admin_login_password)
driver.find_element(By.NAME, 'commit').click()
time.sleep(2)
print("##########")
print("어드민 로그인 완료")

link = driver.find_element(By.LINK_TEXT, ("강제 조정 기능"))
link.click()
time.sleep(2)


link = driver.find_element(By.LINK_TEXT, ("배송 상태값 변경"))
link.click()
time.sleep(2)

print("344 PASS")


print("강제 조정 - 배송 상태값 변경 시작")
#########################################################################################################################
## 강제 조정 - 배송 상태값 변경 - 자동화상품001(바로-일반)
print("강제 조정 - 배송 상태값 변경 - 자동화상품001(바로-일반) -> 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(deal_ship_now_normal) # 배송 요청 번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
print("345 PASS")

# 배송 상태값 변경 - 배송요청 번호 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
time.sleep(3)

# 자동화상품001(바로-일반) - 배송 상태 확인
# 배송 완료 처리 버튼 클릭
try:
    deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
    request_completed_button = deal_tbody.find_element(By.CLASS_NAME, 'request_completed')

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'request_completed')))
    request_completed_button.click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm')))
    driver.find_element(By.CLASS_NAME, 'swal2-confirm').send_keys(Keys.ENTER)
    print("347 PASS")
    time.sleep(5) 

except:
    pass

print("강제 조정 - 배송 상태값 변경 - 자동화상품001(바로-일반) -> 종료")

wms_css_selector_input_sendkeys(driver, "div>input[placeholder='검색어를 입력해주세요.']")
#driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


#########################################################################################################################
## 강제 조정 - 배송 상태값 변경 - 자동화상품001(바로-당일)
print("강제 조정 - 배송 상태값 변경 - 자동화상품001(바로-당일) -> 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(deal_ship_now_today) # 배송 요청 번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
print("351 PASS")

# 배송 상태값 변경 - 배송요청 번호 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
time.sleep(3)

# 자동화상품001(바로-당일) - 배송 상태 확인
# 배송 완료 처리 버튼 클릭
try:
    deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
    request_completed_button = deal_tbody.find_element(By.CLASS_NAME, 'request_completed')

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'request_completed')))
    request_completed_button.click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm')))
    driver.find_element(By.CLASS_NAME, 'swal2-confirm').send_keys(Keys.ENTER)    
    print("353 PASS")
    time.sleep(5) 

except:
    pass

print("강제 조정 - 배송 상태값 변경 - 자동화상품001(바로-당일) -> 종료")


wms_css_selector_input_sendkeys(driver, "div>input[placeholder='검색어를 입력해주세요.']")
#driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


#########################################################################################################################
## 강제 조정 - 배송 상태값 변경 - 자동화상품002(다스-일반)
print("강제 조정 - 배송 상태값 변경 - 자동화상품002(다스-일반) -> 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(deal_ship_das_normal) # 배송 요청 번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
print("357 PASS")

# 배송 상태값 변경 - 배송요청 번호 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
time.sleep(3)

# 자동화상품002(다스-일반) - 배송 상태 확인

# 배송 완료 처리 버튼 클릭
try:
    deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
    request_completed_button = deal_tbody.find_element(By.CLASS_NAME, 'request_completed')

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'request_completed')))
    request_completed_button.click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm')))
    driver.find_element(By.CLASS_NAME, 'swal2-confirm').send_keys(Keys.ENTER)
    print("359 PASS")
    time.sleep(5) 

except:
    pass

print("강제 조정 - 배송 상태값 변경 - 자동화상품002(다스-일반) -> 종료")


wms_css_selector_input_sendkeys(driver, "div>input[placeholder='검색어를 입력해주세요.']")
#driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


#########################################################################################################################
## 강제 조정 - 배송 상태값 변경 - 자동화상품002(다스-당일)
print("강제 조정 - 배송 상태값 변경 - 자동화상품002(다스-당일) -> 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(deal_ship_das_today) # 배송 요청 번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
print("363 PASS")

# 배송 상태값 변경 - 배송요청 번호 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
time.sleep(3)

# 자동화상품002(다스-당일) - 배송 상태 확인
# 배송 완료 처리 버튼 클릭
try:
    deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
    request_completed_button = deal_tbody.find_element(By.CLASS_NAME, 'request_completed')

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'request_completed')))
    request_completed_button.click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm')))
    driver.find_element(By.CLASS_NAME, 'swal2-confirm').send_keys(Keys.ENTER)
    print("365 PASS")
    time.sleep(5) 

except:
    pass

print("강제 조정 - 배송 상태값 변경 - 자동화상품002(다스-당일) -> 종료")


wms_css_selector_input_sendkeys(driver, "div>input[placeholder='검색어를 입력해주세요.']")
#driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


#########################################################################################################################
## 강제 조정 - 배송 상태값 변경 - 자동화상품003(개별-일반)
print("강제 조정 - 배송 상태값 변경 - 자동화상품003(개별-일반) -> 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(deal_ship_each_normal) # 배송 요청 번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
print("369 PASS")

# 배송 상태값 변경 - 배송요청 번호 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
time.sleep(3)

# 자동화상품003(개별-일반) - 배송 상태 확인
# 배송 완료 처리 버튼 클릭
try:
    deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
    request_completed_button = deal_tbody.find_element(By.CLASS_NAME, 'request_completed')

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'request_completed')))
    request_completed_button.click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm')))
    driver.find_element(By.CLASS_NAME, 'swal2-confirm').send_keys(Keys.ENTER)
    print("371 PASS")
    time.sleep(5) 

except:
    pass

print("강제 조정 - 배송 상태값 변경 - 자동화상품003(개별-일반) -> 종료")


wms_css_selector_input_sendkeys(driver, "div>input[placeholder='검색어를 입력해주세요.']")
#driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)


#########################################################################################################################
## 강제 조정 - 배송 상태값 변경 - 자동화상품003(개별-당일)
print("강제 조정 - 배송 상태값 변경 - 자동화상품003(개별-당일) -> 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(deal_ship_each_today) # 배송 요청 번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='검색어를 입력해주세요.']").send_keys(Keys.ENTER) # 검색 적용
print("375 PASS")

# 배송 상태값 변경 - 배송요청 번호 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
time.sleep(3)

# 자동화상품003(개별-당일) - 배송 상태 확인
# 배송 완료 처리 버튼 클릭
try:
    deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
    request_completed_button = deal_tbody.find_element(By.CLASS_NAME, 'request_completed')

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'request_completed')))
    request_completed_button.click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm')))
    driver.find_element(By.CLASS_NAME, 'swal2-confirm').send_keys(Keys.ENTER)
    print("377 PASS")
    time.sleep(5) 

except:
    pass
print("강제 조정 - 배송 상태값 변경 - 자동화상품003(개별-당일) -> 종료")


print("##########")
print("강제 조정 - 배송 상태값 변경 종료")


#########################################################################################################################



print("출고 테스트 완료")

exit()


#while(True):
#    	pass
