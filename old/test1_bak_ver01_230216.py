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
time.sleep(30)
print("google_sheet API 제한으로 30초 딜레이")
# 정산 입고 - 사입앱 실행파일에서 인자 전달해 줌
file_info = json.loads(sys.argv[1])
print(file_info)

deal_out_excel_upload_name =  file_info["path"]     # 배송 요청 엑셀 파일 경로
deal_product_id_1 = file_info["deal_product_id_1"]  # 딜리버드 상품번호(자동화상품 001)
deal_product_id_2 = file_info["deal_product_id_2"]  # 딜리버드 상품번호(자동화상품 002)
deal_product_id_3 = file_info["deal_product_id_3"]  # 딜리버드 상품번호(자동화상품 003)

deal_seller_name = file_info["deal_seller_name"]    # 딜리버드 셀러명(대포자)

deal_sell_product_name_1 = file_info["deal_sell_product_name_1"] # 딜리버드 판매 상품명(자동화상품 001)
deal_sell_product_name_2 = file_info["deal_sell_product_name_2"] # 딜리버드 판매 상품명(자동화상품 002)
deal_sell_product_name_3 = file_info["deal_sell_product_name_3"] # 딜리버드 판매 상품명(자동화상품 003)

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

#########################################################################################################################
## 출고 시작 전 구글 시트에 정보 업데이트
# login ID / WMS 로그인 정보
google_sheet.update_acell('H134', wms_login_id)
google_sheet.update_acell('H283', wms_login_id)
google_sheet.update_acell('H291', wms_login_id)
google_sheet.update_acell('H299', wms_login_id)
google_sheet.update_acell('H307', wms_login_id)
google_sheet.update_acell('H315', wms_login_id)
google_sheet.update_acell('H323', wms_login_id)
print("login ID / WMS 로그인 정보 : 구글 시트 업데이트 완료")


# 셀러명 대표자
google_sheet.update_acell('H76', deal_seller_name)
google_sheet.update_acell('H114', deal_seller_name)
google_sheet.update_acell('H241', deal_seller_name)
google_sheet.update_acell('H249', deal_seller_name)
google_sheet.update_acell('H268', deal_seller_name)
google_sheet.update_acell('H276', deal_seller_name)
google_sheet.update_acell('H344', deal_seller_name)
google_sheet.update_acell('H350', deal_seller_name)
google_sheet.update_acell('H356', deal_seller_name)
google_sheet.update_acell('H362', deal_seller_name)
google_sheet.update_acell('H368', deal_seller_name)
google_sheet.update_acell('H374', deal_seller_name)
print("셀러명 대표자 : 구글 시트 업데이트 완료", deal_seller_name)


# 어드민 계정
google_sheet.update_acell('H349', deal_admin_login_id)
google_sheet.update_acell('H355', deal_admin_login_id)
google_sheet.update_acell('H361', deal_admin_login_id)
google_sheet.update_acell('H367', deal_admin_login_id)
google_sheet.update_acell('H373', deal_admin_login_id)
google_sheet.update_acell('H379', deal_admin_login_id)
print("어드민 계정 : 구글 시트 업데이트 완료")


# 딜리버드 판매 상품명
google_sheet.update_acell('H12', deal_sell_product_name_1)
google_sheet.update_acell('H258', deal_sell_product_name_1)
google_sheet.update_acell('H264', deal_sell_product_name_1)

google_sheet.update_acell('H17', deal_sell_product_name_2)
google_sheet.update_acell('H22', deal_sell_product_name_3)
print("딜리버드 판매 상품명 : 구글 시트 업데이트 완료", deal_sell_product_name_1)


# 딜리버드 상품번호(SKU)
google_sheet.update_acell('H13', deal_product_id_1)
google_sheet.update_acell('H61', deal_product_id_1)
google_sheet.update_acell('H144', deal_product_id_1)
google_sheet.update_acell('H155', deal_product_id_1)
google_sheet.update_acell('H332', deal_product_id_1)
google_sheet.update_acell('H381', deal_product_id_1)
print("딜리버드 상품번호(SKU) 1번 : 구글 시트 업데이트 완료", deal_product_id_1)

google_sheet.update_acell('H18', deal_product_id_2)
google_sheet.update_acell('H65', deal_product_id_2)
google_sheet.update_acell('H166', deal_product_id_2)
google_sheet.update_acell('H177', deal_product_id_2)
google_sheet.update_acell('H336', deal_product_id_2)
google_sheet.update_acell('H385', deal_product_id_2)
print("딜리버드 상품번호(SKU) 2번 : 구글 시트 업데이트 완료", deal_product_id_2)

google_sheet.update_acell('H23', deal_product_id_3)
google_sheet.update_acell('H69', deal_product_id_3)
google_sheet.update_acell('H188', deal_product_id_3)
google_sheet.update_acell('H199', deal_product_id_3)
google_sheet.update_acell('H340', deal_product_id_3)
google_sheet.update_acell('H389', deal_product_id_3)
print("딜리버드 상품번호(SKU) 3번 : 구글 시트 업데이트 완료", deal_product_id_3)

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

#driver.set_window_size(6000, 1024)
driver.switch_to.window(tabs[2])
driver.get(wms_url)

action = ActionChains(driver)

# hihida 221230
#########################################################################################################################
########### 어드민 -> 딜리버드 셀러 이동 ##########
#########################################################################################################################
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
# google_sheet.update_acell('I5', 'OK')
# print(" 5PASS")

#########################################################################################################################
#### 딜리버드 -> 배송요청 #####
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[5]/a').click()
'''
# 전체 텍스트를 한번에 가지고 와서 for과 if문을 실행 할 수 없음 a/nb/n~~~~
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'ul.navbar-nav.mr-auto')
for wms_str_loop in wms_test_result:
    print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "배송요청":
        print("배송요청 클릭 시도\n")
        wms_str_loop.click()
        print("배송요청 클릭 완료\n")
        break
'''
# google_sheet.update_acell('I6', 'OK')
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
# google_sheet.update_acell('I7', 'OK')
print("7 PASS")
# google_sheet.update_acell('I8', 'OK')


print("파일 업로드 선택 완료\n")
# google_sheet.update_acell('I9', 'OK')
print("9 PASS")


#파일 업로드
up_load_file = driver.find_element(By.XPATH, '//*[@id="orders"]') # 모달 / 엑셀 업로드 양식 선태 -> 엑셀 파일 선택 Browse 버튼
up_load_file.send_keys(deal_out_excel_upload_name)
time.sleep(3)
driver.find_element(By.ID, 'excel_order_import_btn').click() # [업로드] 버튼

# google_sheet.update_acell('I10', 'OK')
print("10 PASS")
print("업로드 클릭 완료\n")


#########################################################################################################################
#### 딜리버드 -> 상품 및 재고 #####
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[8]/a').click()

# google_sheet.update_acell('I11', 'OK')
print("11 PASS")
time.sleep(5)
print("##########")
print("상품 및 재고 시작")


# 도매 상품명_자동화001 상품 및 재고 확인
cell_data = google_sheet.acell('H13').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
# google_sheet.update_acell('I13', 'OK')
print("13 PASS")
time.sleep(2)

# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="productList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 도매 상품명_자동화001 상품 및 재고 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        #print("재고 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 3: # 판매 상품명
            cell_data = google_sheet.acell('H12').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I12', 'Pass')
            else:
                google_sheet.update_acell('I12', 'Failed')        
        
        if deal_list_count == 13: # 총재고
            cell_data = google_sheet.acell('H14').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                # google_sheet.update_acell('I14', 'OK')
                print("14 PASS")            
                google_sheet.update_acell('H62', deal_test_result_check)  # 구글 시트 기대 결과 업데이트
                google_sheet.update_acell('H333', deal_test_result_check) # 구글 시트 기대 결과 업데이트
            else:
                google_sheet.update_acell('I14', 'Failed')           
            
        if deal_list_count == 14: # 정상재고
            cell_data = google_sheet.acell('H15').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                # google_sheet.update_acell('I15', 'OK')
                print("15 PASS")
                google_sheet.update_acell('H63', deal_test_result_check)  # 구글 시트 기대 결과 업데이트
                google_sheet.update_acell('H334', deal_test_result_check) # 구글 시트 기대 결과 업데이트
            else:
                google_sheet.update_acell('I15', 'Failed')
            
        if deal_list_count == 17: # 배송 요청 가능
            cell_data = google_sheet.acell('H16').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I16', 'Pass')
            else:
                google_sheet.update_acell('I16', 'Failed')
        
        deal_list_count += 1

driver.find_element(By.ID,'search_text').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 도매 상품명_자동화002 상품 및 재고 확인
cell_data = google_sheet.acell('H18').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
# google_sheet.update_acell('I18', 'OK')
print("18 PASS")
time.sleep(2)

# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="productList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 도매 상품명_자동화002 상품 및 재고 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        #print("재고 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 3: # 판매 상품명
            cell_data = google_sheet.acell('H17').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I17', 'Pass')
            else:
                google_sheet.update_acell('I17', 'Failed')        
        
        if deal_list_count == 13: # 총재고
            cell_data = google_sheet.acell('H19').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                # google_sheet.update_acell('I19', 'OK')
                print("19 PASS")
                google_sheet.update_acell('H67', deal_test_result_check)  # 구글 시트 기대 결과 업데이트
                google_sheet.update_acell('H337', deal_test_result_check) # 구글 시트 기대 결과 업데이트
            else:
                google_sheet.update_acell('I19', 'Failed')           
            
        if deal_list_count == 14: # 정상재고
            cell_data = google_sheet.acell('H20').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                # google_sheet.update_acell('I20', 'OK')
                print("20 PASS")
                google_sheet.update_acell('H66', deal_test_result_check)  # 구글 시트 기대 결과 업데이트
                google_sheet.update_acell('H338', deal_test_result_check) # 구글 시트 기대 결과 업데이트
            else:
                google_sheet.update_acell('I20', 'Failed')
            
        if deal_list_count == 17: # 배송 요청 가능
            cell_data = google_sheet.acell('H21').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I21', 'Pass')
            else:
                google_sheet.update_acell('I21', 'Failed')
        
        deal_list_count += 1

driver.find_element(By.ID,'search_text').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 도매 상품명_자동화003 상품 및 재고 확인
cell_data = google_sheet.acell('H23').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
# google_sheet.update_acell('I23', 'OK')
print("23 PASS")
time.sleep(2)

# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="productList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 도매 상품명_자동화003 상품 및 재고 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        #print("재고 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 3: # 판매 상품명
            cell_data = google_sheet.acell('H22').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I22', 'Pass')
            else:
                google_sheet.update_acell('I22', 'Failed')        
        
        if deal_list_count == 13: # 총재고
            cell_data = google_sheet.acell('H24').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                # google_sheet.update_acell('I24', 'OK')
                print("24 PASS")
                google_sheet.update_acell('H70', deal_test_result_check)  # 구글 시트 기대 결과 업데이트
                google_sheet.update_acell('H341', deal_test_result_check) # 구글 시트 기대 결과 업데이트
            else:
                google_sheet.update_acell('I24', 'Failed')           
            
        if deal_list_count == 14: # 정상재고
            cell_data = google_sheet.acell('H25').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                # google_sheet.update_acell('I25', 'OK')
                print("25 PASS")
                google_sheet.update_acell('H71', deal_test_result_check)  # 구글 시트 기대 결과 업데이트
                google_sheet.update_acell('H342', deal_test_result_check) # 구글 시트 기대 결과 업데이트
            else:
                google_sheet.update_acell('I25', 'Failed')
            
        if deal_list_count == 17: # 배송 요청 가능
            cell_data = google_sheet.acell('H26').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I26', 'Pass')
            else:
                google_sheet.update_acell('I26', 'Failed')
        
        deal_list_count += 1



#########################################################################################################################
#### 딜리버드 -> 배송요청 #####
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[5]/a').click()

time.sleep(5)
# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')

deal_list_tr_count = 0
deal_list_td_count = 0
time.sleep(3)



# 배송 가능 요청 리스트 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    deal_list_td_count = 0
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        #print(deal_list_tr_count, "번 ROW : ", deal_list_td_count,"번 - ", deal_test_result_check)
                
        if deal_list_td_count == 2: # 배송 요청 번호
            deal_ship_number_temp = td.get_attribute("innerText")  
            
        if deal_list_td_count == 3: # 배송 방법
            deal_ship_type_temp = td.get_attribute("innerText")  
                
        if deal_list_td_count == 8: # 고객사 주문번호 
            deal_test_result_check = td.get_attribute("innerText")
            deal_test_result_check = deal_test_result_check.replace("\n",'')
            deal_test_result_check = deal_test_result_check.replace("( )",'')
            
            if "바로출고-일반배송" == deal_test_result_check:
                google_sheet.update_acell('H27', deal_ship_number_temp) # 딜리버드 배송요청 번호 -> 구글 시트에 업데이트
                google_sheet.update_acell('H28', deal_ship_type_temp)   # 딜리버드 배송 방법 -> 구글 시트에 업데이트
                # google_sheet.update_acell('I27', 'OK')
                print("26 PASS")
                # google_sheet.update_acell('I28', 'OK')                
                deal_ship_now_normal = deal_ship_number_temp
                print(deal_list_tr_count, "번 ROW : ", deal_test_result_check," -> ", "배송 요청 번호 - ", deal_ship_now_normal, " , 배송 방법 - ", deal_ship_type_temp)
                break
                
            elif "바로출고-당일배송" == deal_test_result_check:
                google_sheet.update_acell('H29', deal_ship_number_temp)
                google_sheet.update_acell('H30', deal_ship_type_temp)
                # google_sheet.update_acell('I29', 'OK')
                print("29 PASS")
                # google_sheet.update_acell('I30', 'OK')                
                deal_ship_now_today = deal_ship_number_temp
                print(deal_list_tr_count, "번 ROW : ", deal_test_result_check," -> ", "배송 요청 번호 - ", deal_ship_now_today, " , 배송 방법 - ", deal_ship_type_temp)
                break
                
            elif "다스출고-일반배송" == deal_test_result_check:
                google_sheet.update_acell('H31', deal_ship_number_temp)
                google_sheet.update_acell('H32', deal_ship_type_temp)
                # google_sheet.update_acell('I31', 'OK')
                print("31 PASS")
                # google_sheet.update_acell('I32', 'OK')                
                deal_ship_das_normal = deal_ship_number_temp
                print(deal_list_tr_count, "번 ROW : ", deal_test_result_check," -> ", "배송 요청 번호 - ", deal_ship_das_normal, " , 배송 방법 - ", deal_ship_type_temp)
                break
            
            elif "다스출고-당일배송" == deal_test_result_check:
                google_sheet.update_acell('H33', deal_ship_number_temp)
                google_sheet.update_acell('H34', deal_ship_type_temp)
                # google_sheet.update_acell('I33', 'OK')
                print("33 PASS")
                # google_sheet.update_acell('I34', 'OK')                
                deal_ship_das_today = deal_ship_number_temp
                print(deal_list_tr_count, "번 ROW : ", deal_test_result_check," -> ", "배송 요청 번호 - ", deal_ship_das_today, " , 배송 방법 - ", deal_ship_type_temp)
                break
            
            elif "개별출고-일반배송" == deal_test_result_check:
                google_sheet.update_acell('H35', deal_ship_number_temp)
                google_sheet.update_acell('H36', deal_ship_type_temp)
                # google_sheet.update_acell('I35', 'OK')
                print("35 PASS")
                # google_sheet.update_acell('I36', 'OK')
                deal_ship_each_normal = deal_ship_number_temp
                print(deal_list_tr_count, "번 ROW : ", deal_test_result_check," -> ", "배송 요청 번호 - ", deal_ship_each_normal, " , 배송 방법 - ", deal_ship_type_temp)
                break
                
            elif "개별출고-당일배송" == deal_test_result_check:
                google_sheet.update_acell('H37', deal_ship_number_temp)
                google_sheet.update_acell('H38', deal_ship_type_temp)
                # google_sheet.update_acell('I37', 'OK')
                print("37 PASS")
                # google_sheet.update_acell('I38', 'OK')
                deal_ship_each_today = deal_ship_number_temp
                print(deal_list_tr_count, "번 ROW : ", deal_test_result_check," -> ", "배송 요청 번호 - ", deal_ship_each_today, " , 배송 방법 - ", deal_ship_type_temp)
                break
        
        deal_list_td_count += 1
    deal_list_tr_count += 1


## 딜리버드 배송요청 번호 -> 구글 시트에 업데이트
# 자동화상품001(바로-일반)
print("딜리버드 배송요청 번호 -> 구글 시트에 업데이트 시작")
google_sheet.update_acell('H42', deal_ship_now_normal)
google_sheet.update_acell('H77', deal_ship_now_normal)
google_sheet.update_acell('H115', deal_ship_now_normal)
google_sheet.update_acell('H345', deal_ship_now_normal)

# 자동화상품001(바로-당일)
google_sheet.update_acell('H45', deal_ship_now_today)
google_sheet.update_acell('H83', deal_ship_now_today)
google_sheet.update_acell('H117', deal_ship_now_today)
google_sheet.update_acell('H351', deal_ship_now_today)

# 자동화상품002(다스-일반)
google_sheet.update_acell('H48', deal_ship_das_normal)
google_sheet.update_acell('H89', deal_ship_das_normal)
google_sheet.update_acell('H119', deal_ship_das_normal)
google_sheet.update_acell('H208', deal_ship_das_normal)
google_sheet.update_acell('H357', deal_ship_das_normal)

# 자동화상품002(다스-당일)
google_sheet.update_acell('H51', deal_ship_das_today)
google_sheet.update_acell('H95', deal_ship_das_today)
google_sheet.update_acell('H121', deal_ship_das_today)
google_sheet.update_acell('H225', deal_ship_das_today)
google_sheet.update_acell('H363', deal_ship_das_today)

# 자동화상품003(개별-일반)
google_sheet.update_acell('H54', deal_ship_each_normal)
google_sheet.update_acell('H101', deal_ship_each_normal)
google_sheet.update_acell('H127', deal_ship_each_normal)
google_sheet.update_acell('H369', deal_ship_each_normal)

# 자동화상품003(개별-당일)
google_sheet.update_acell('H57', deal_ship_each_today)
google_sheet.update_acell('H107', deal_ship_each_today)
google_sheet.update_acell('H129', deal_ship_each_today)
google_sheet.update_acell('H375', deal_ship_each_today)
print("딜리버드 배송요청 번호 -> 구글 시트에 업데이트 종료")


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



''' 

# [전체 배송 요청] 버튼 클릭
driver.find_element(By.XPATH, '//*[@id="orderList_wrapper"]/div[1]/div[2]/div/button[4]').click()
# google_sheet.update_acell('I39', 'OK')
print("39 PASS")
print("[전체 배송 요청] 버튼 클릭 완료")
time.sleep(3)

# alert [예] 클릭
driver.find_element(By.XPATH, '/html/body/div[5]/div/div[3]/button[3]').click()
print("alert [예] 클릭 완료")
time.sleep(3)

# alert [OK] 클릭
driver.find_element(By.XPATH, '/html/body/div[5]/div/div[3]/button[1]').click()
'''
# google_sheet.update_acell('I40', 'OK')
print("40 PASS")
print("alert [OK] 클릭 완료")
time.sleep(30)
print("google_sheet API 제한으로 30초 딜레이")

#########################################################################################################################
#### 딜리버드 -> 배송현황 #####
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[6]/a').click()
time.sleep(3)

driver.find_element(By.XPATH,'//*[@id="page-wrapper"]/div[2]/div[1]/div[1]/ul[1]/li[3]/a').click()
# google_sheet.update_acell('I41', 'OK')
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
cell_data = google_sheet.acell('H42').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
# google_sheet.update_acell('I42', 'OK')
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
        
        if deal_list_count == 6: # 배송 상태
            cell_data = google_sheet.acell('H43').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I43', 'Pass')
            else:
                google_sheet.update_acell('I43', 'Failed')        
        
        if deal_list_count == 7: # 송장번호
            deal_test_result_check = td.get_attribute("innerText")
            # google_sheet.update_acell('I44', 'OK')
            print("44 PASS")
            deal_invoice_number_now_normal = deal_test_result_check
            google_sheet.update_acell('H44', deal_invoice_number_now_normal)  # 딜리버드 송장 번호 -> 구글 시트에 업데이트
            google_sheet.update_acell('H82', deal_invoice_number_now_normal)
            google_sheet.update_acell('H257', deal_invoice_number_now_normal)
            google_sheet.update_acell('H259', deal_invoice_number_now_normal)
            break
        
        deal_list_count += 1
print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품001(바로-일반)")

driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




## 자동화상품001(바로-당일) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품001(바로-당일)")
cell_data = google_sheet.acell('H45').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
# google_sheet.update_acell('I45', 'OK')
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
        
        if deal_list_count == 6: # 배송 상태
            cell_data = google_sheet.acell('H46').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I46', 'Pass')
            else:
                google_sheet.update_acell('I46', 'Failed')        
        
        if deal_list_count == 7: # 송장번호
            deal_test_result_check = td.get_attribute("innerText")
            # google_sheet.update_acell('I47', 'OK')
            print("47 PASS")
            deal_invoice_number_now_today = deal_test_result_check
            google_sheet.update_acell('H47', deal_invoice_number_now_today)  # 딜리버드 송장 번호 -> 구글 시트에 업데이트
            google_sheet.update_acell('H88', deal_invoice_number_now_today)
            google_sheet.update_acell('H263', deal_invoice_number_now_today)
            google_sheet.update_acell('H265', deal_invoice_number_now_today)
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품001(바로-당일)")

driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)



## 자동화상품002(다스-일반) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품002(다스-일반)")
cell_data = google_sheet.acell('H48').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
# google_sheet.update_acell('I48', 'OK')
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
        
        if deal_list_count == 6: # 배송 상태
            cell_data = google_sheet.acell('H49').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I49', 'Pass')
            else:
                google_sheet.update_acell('I49', 'Failed')        
        
        if deal_list_count == 7: # 송장번호
            deal_test_result_check = td.get_attribute("innerText")
            # google_sheet.update_acell('I50', 'OK')
            print("50 PASS")
            deal_invoice_number_das_normal  = deal_test_result_check
            google_sheet.update_acell('H50', deal_invoice_number_das_normal)  # 딜리버드 송장 번호 -> 구글 시트에 업데이트
            google_sheet.update_acell('H94', deal_invoice_number_das_normal)
            google_sheet.update_acell('H209', deal_invoice_number_das_normal)
            google_sheet.update_acell('H240', deal_invoice_number_das_normal)
            google_sheet.update_acell('H243', deal_invoice_number_das_normal)
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품002(다스-일반)")

driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




## 자동화상품002(다스-당일) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품002(다스-당일)")
cell_data = google_sheet.acell('H51').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
# google_sheet.update_acell('I51', 'OK')
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
        # print("배송 현황 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 6: # 배송 상태
            cell_data = google_sheet.acell('H52').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I52', 'Pass')
            else:
                google_sheet.update_acell('I52', 'Failed')        
        
        if deal_list_count == 7: # 송장번호
            deal_test_result_check = td.get_attribute("innerText")
            # google_sheet.update_acell('I53', 'OK')
            print("53 PASS")
            deal_invoice_number_das_today  = deal_test_result_check
            google_sheet.update_acell('H53', deal_invoice_number_das_today)  # 딜리버드 송장 번호 -> 구글 시트에 업데이트
            google_sheet.update_acell('H100', deal_invoice_number_das_today)
            google_sheet.update_acell('H227', deal_invoice_number_das_today)
            google_sheet.update_acell('H248', deal_invoice_number_das_today)
            google_sheet.update_acell('H251', deal_invoice_number_das_today)
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품002(다스-당일)")

driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)





## 자동화상품003(개별-일반) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품003(개별-일반)")
cell_data = google_sheet.acell('H54').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
# google_sheet.update_acell('I54', 'OK')
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
        # print("배송 현황 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 6: # 배송 상태
            cell_data = google_sheet.acell('H55').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I55', 'Pass')
            else:
                google_sheet.update_acell('I55', 'Failed')        
        
        if deal_list_count == 7: # 송장번호
            deal_test_result_check = td.get_attribute("innerText")
            # google_sheet.update_acell('I56', 'OK')
            print("56 PASS")
            deal_invoice_number_each_normal  = deal_test_result_check
            google_sheet.update_acell('H56', deal_invoice_number_each_normal)  # 딜리버드 송장 번호 -> 구글 시트에 업데이트
            google_sheet.update_acell('H106', deal_invoice_number_each_normal)
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품003(개별-일반)")

driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




## 자동화상품003(개별-당일) 검색
print("배송 요청 번호 검색 시작 : # 자동화상품003(개별-당일)")
cell_data = google_sheet.acell('H57').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.NAME,'searchWord').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.ENTER)
# google_sheet.update_acell('I57', 'OK')
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
        
        if deal_list_count == 6: # 배송 상태
            cell_data = google_sheet.acell('H58').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I58', 'Pass')
            else:
                google_sheet.update_acell('I58', 'Failed')        
        
        if deal_list_count == 7: # 송장번호
            deal_test_result_check = td.get_attribute("innerText")
            # google_sheet.update_acell('I59', 'OK')
            print("59 PASS")
            deal_invoice_number_each_today  = deal_test_result_check
            google_sheet.update_acell('H59', deal_invoice_number_each_today)  # 딜리버드 송장 번호 -> 구글 시트에 업데이트
            google_sheet.update_acell('H112', deal_invoice_number_each_today)
            break
        
        deal_list_count += 1

print("배송 현황 리스트(테이블) 체크 완료 : 자동화상품003(개별-당일)")

driver.find_element(By.NAME,'searchWord').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.NAME,'searchWord').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(2)




#########################################################################################################################
#### 딜리버드 -> 상품 및 재고 #####
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[8]/a').click()

# google_sheet.update_acell('I60', 'OK')
print("60 PASS")
time.sleep(5)
print("##########")
print("상품 및 재고 시작 : 배송 요청 완료 후")


# 도매 상품명_자동화001 상품 및 재고 확인
cell_data = google_sheet.acell('H61').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
# google_sheet.update_acell('I61', 'OK')
print("61 PASS")
time.sleep(2)

# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="productList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 도매 상품명_자동화001 상품 및 재고 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        # print("재고 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
       
        if deal_list_count == 13: # 총재고
            cell_data = google_sheet.acell('H62').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I62', 'Pass')
            else:
                google_sheet.update_acell('I62', 'Failed')
            
        if deal_list_count == 14: # 정상재고
            cell_data = google_sheet.acell('H63').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I63', 'Pass')
            else:
                google_sheet.update_acell('I63', 'Failed')
            
        if deal_list_count == 20: # 배송 상품 포장중
            cell_data = google_sheet.acell('H64').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I64', 'Pass')
            else:
                google_sheet.update_acell('I64', 'Failed')
        
        deal_list_count += 1

driver.find_element(By.ID,'search_text').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 도매 상품명_자동화002 상품 및 재고 확인
cell_data = google_sheet.acell('H65').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
# google_sheet.update_acell('I65', 'OK')
print("65 PASS")
time.sleep(2)

# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="productList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)

# 도매 상품명_자동화002 상품 및 재고 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        #print("재고 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 13: # 총재고
            cell_data = google_sheet.acell('H66').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I66', 'Pass')
            else:
                google_sheet.update_acell('I66', 'Failed')
            
        if deal_list_count == 14: # 정상재고
            cell_data = google_sheet.acell('H67').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I67', 'Pass')
            else:
                google_sheet.update_acell('I67', 'Failed')
            
        if deal_list_count == 20: # 배송 상품 포장중
            cell_data = google_sheet.acell('H68').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I68', 'Pass')
            else:
                google_sheet.update_acell('I68', 'Failed')
        
        deal_list_count += 1

driver.find_element(By.ID,'search_text').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)




# 도매 상품명_자동화003 상품 및 재고 확인
cell_data = google_sheet.acell('H69').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
# google_sheet.update_acell('I69', 'OK')
print("69 PASS")
time.sleep(2)

# 재고 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="productList_wrapper"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
time.sleep(3)


# 도매 상품명_자동화002 상품 및 재고 확인
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        #print("재고 리스트(테이블)", deal_list_count,"번", deal_test_result_check)
        
        if deal_list_count == 13: # 총재고
            cell_data = google_sheet.acell('H70').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I70', 'Pass')
            else:
                google_sheet.update_acell('I70', 'Failed')
            
        if deal_list_count == 14: # 정상재고
            cell_data = google_sheet.acell('H71').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I71', 'Pass')
            else:
                google_sheet.update_acell('I71', 'Failed')
            
        if deal_list_count == 20: # 배송 상품 포장중
            cell_data = google_sheet.acell('H72').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I72', 'Pass')
            else:
                google_sheet.update_acell('I72', 'Failed')
        
        deal_list_count += 1


# hihida 230206
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



wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 대상 리스트 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(10)
print("google_sheet API 제한으로 10초 딜레이")


print("출고 관리 - 출고 대상 리스트 조회 이동")
# google_sheet.update_acell('I113', 'OK')
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

# driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
# time.sleep(5)


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
# google_sheet.update_acell('I114', 'OK')
print("114 PASS")
time.sleep(15)



# 테이블(리스트) -> 송장번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_now_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(5)


# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# 230223 / aria-colindex : 7 -> aria-colindex : 9로 변경
# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="9"]') # 배송요청번호 : 9
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
# google_sheet.update_acell('I116', 'OK')
print("116 PASS")
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

# 230223 / aria-colindex : 7 -> aria-colindex : 9로 변경
# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="9"]') # 배송요청번호 : 9
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
# google_sheet.update_acell('I118', 'OK')
print("118 PASS")
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

# 230223 / aria-colindex : 7 -> aria-colindex : 9로 변경
# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="9"]') # 배송요청번호 : 9
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
# google_sheet.update_acell('I120', 'OK')
print("120 PASS")
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

# 230223 / aria-colindex : 7 -> aria-colindex : 9로 변경
# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="9"]') # 배송요청번호 : 9
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
# google_sheet.update_acell('I122', 'OK')
print("122 PASS")
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
            # google_sheet.update_acell('I123', 'OK')
            print("123 PASS")
            break
            #print("클릭 완료\n")

time.sleep(10)


alert = driver.switch_to.alert 
alert.accept() # 얼럿 확인
# google_sheet.update_acell('I124', 'OK')
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
            # google_sheet.update_acell('I125', 'OK')
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
    wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
    wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
    driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합
    
    print("일반, DAS 의 선택 출고 지시 이 후 전체 체크박스 선택 되어 있는 부분 체크 해제 완료")
except:
    pass

time.sleep(15)

# 테이블(리스트) -> 송장번호 검색 자동화상품003(개별-일반)-> 
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='송장번호 필터 입력']").send_keys(deal_invoice_number_each_normal) 
print("송장번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품003(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# 230223 / aria-colindex : 7 -> aria-colindex : 9로 변경
# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="9"]') # 배송요청번호 : 9
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
# google_sheet.update_acell('I128', 'OK')
print("128 PASS")
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

# 230223 / aria-colindex : 7 -> aria-colindex : 9로 변경
# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="7"]') # 배송요청번호 : 7
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="9"]') # 배송요청번호 : 9
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
# google_sheet.update_acell('I130', 'OK')
print("130 PASS")
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
            # google_sheet.update_acell('I131', 'OK')
            print("131 PASS")
            break
            #print("클릭 완료\n")

time.sleep(10)


# 1번째 alert 처리
alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
alert_text = alert.text
alert.accept()

time.sleep(5)
# 2번째 alert 처리
alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
alert_text = alert.text
alert.accept()


'''
time.sleep(2)

try:    
    # esc 클릭해서 확인 얼럿 종료
    print("[개별출고지시] 버튼 클릭 - 확인 -except 시작")
    actions = ActionChains(driver)
    actions.send_keys(Keys.ESCAPE)
    actions.perform()
except:
    alert = driver.switch_to.alert
    print("[개별출고지시] 버튼 클릭 - 확인 - 시작")
    alert.accept() # 얼럿 확인
    print("[개별출고지시] 버튼 클릭 - 확인 - 종료")
    time.sleep(3)
'''
    

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

time.sleep(30)
print("google_sheet API 제한으로 30초 딜레이")

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

# driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='전체 선택']").click()#전체 선택 버튼 다시 클릭해 일단 전체 컬럼 선택 해제
# time.sleep(5)


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
# google_sheet.update_acell('I78', 'OK')
print("78 PASS")# SKU 기준 데이터 : 첫 조회
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
# google_sheet.update_acell('I91', 'OK')
print("91 PASS")# SKU 기준 데이터 : 첫 조회
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
# google_sheet.update_acell('I102', 'OK')
print("102 PASS")# SKU 기준 데이터 : 첫 조회
google_sheet.update_acell('H102', wms_sku_barcode_3)
google_sheet.update_acell('H108', wms_sku_barcode_3)
google_sheet.update_acell('H269', wms_sku_barcode_3)
google_sheet.update_acell('H272', wms_sku_barcode_3)
google_sheet.update_acell('H277', wms_sku_barcode_3)
google_sheet.update_acell('H280', wms_sku_barcode_3)
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

time.sleep(30)
print("google_sheet API 제한으로 30초 딜레이")

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
# google_sheet.update_acell('I74', 'OK')print(" PASS")

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
google_sheet.update_acell('H284', wms_out_round_number_now_normal)


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
google_sheet.update_acell('H292', wms_out_round_number_now_today)
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
google_sheet.update_acell('H300', wms_out_round_number_das_normal)

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
google_sheet.update_acell('H308', wms_out_round_number_das_today)

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
google_sheet.update_acell('H316', wms_out_round_number_each_normal)

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
google_sheet.update_acell('H324', wms_out_round_number_each_today)


# 출고테이블 - 테스트 제외
# google_sheet.update_acell('I138', "N/A")
# google_sheet.update_acell('I287', "N/A")
# google_sheet.update_acell('I149', "N/A")
# google_sheet.update_acell('I295', "N/A")
# google_sheet.update_acell('I160', "N/A")
# google_sheet.update_acell('I303', "N/A")
# google_sheet.update_acell('I171', "N/A")
# google_sheet.update_acell('I311', "N/A")
# google_sheet.update_acell('I182', "N/A")
# google_sheet.update_acell('I319', "N/A")
# google_sheet.update_acell('I193', "N/A")
# google_sheet.update_acell('I327', "N/A")



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

time.sleep(30)
print("google_sheet API 제한으로 30초 딜레이")

# wms_test_result = driver.find_element(By.CSS_SELECTOR,'input.PrivateSwitchBase-input[value="7"]')
# wms_test_result.click()


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
# google_sheet.update_acell('I133', 'OK')
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
# google_sheet.update_acell('I134', 'OK')
print("134 PASS")
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
        # google_sheet.update_acell('I141', 'OK')
        print("141 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        # google_sheet.update_acell('I142', 'OK')
        print("142 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_now_normal = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H143', wms_picking_barcode_now_normal) 
google_sheet.update_acell('H254', wms_picking_barcode_now_normal) 
# google_sheet.update_acell('I143', 'OK')
print("143 PASS")

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
        # google_sheet.update_acell('I152', 'OK')
        print("152 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        # google_sheet.update_acell('I153', 'OK')
        print("153 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_now_today = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H154', wms_picking_barcode_now_today) 
google_sheet.update_acell('H260', wms_picking_barcode_now_today) 
# google_sheet.update_acell('I154', 'OK')
print("154 PASS")

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
        # google_sheet.update_acell('I163', 'OK')
        print("163 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        # google_sheet.update_acell('I164', 'OK')
        print("164 PASS")
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
# google_sheet.update_acell('I165', 'OK')
print("165 PASS")

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
        # google_sheet.update_acell('I174', 'OK')
        print("174 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        # google_sheet.update_acell('I175', 'OK')
        print("175 PASS")
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
# google_sheet.update_acell('I176', 'OK')
print("176 PASS")

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
        # google_sheet.update_acell('I185', 'OK')
        print("185 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        # google_sheet.update_acell('I186', 'OK')
        print("186 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_each_normal = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H187', wms_picking_barcode_each_normal) 
# google_sheet.update_acell('I187', 'OK')
print("187 PASS")

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
        # google_sheet.update_acell('I196', 'OK')
        print("196 PASS")

        alert = driver.switch_to.alert 
        alert.accept() # 얼럿 확인
        # google_sheet.update_acell('I197', 'OK')
        print("197 PASS")
        time.sleep(3)
        break

# 피킹리스트 내 바코드번호 가져오기
wms_test_result = driver.find_element(By.CSS_SELECTOR,'svg text')
wms_test_result = wms_test_result.text
wms_picking_barcode_each_today = wms_test_result

# 피킹 바코드 기초 데이터 - 구글 업데이트
google_sheet.update_acell('H198', wms_picking_barcode_each_today) 
# google_sheet.update_acell('I198', 'OK')
print("198 PASS")

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

time.sleep(30)
print("google_sheet API 제한으로 30초 딜레이")

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
# google_sheet.update_acell('I201', 'OK')
print("201 PASS")


cell_data = google_sheet.acell('H202').value 

print("출고 관리 - DAS 피킹 자동화상품002(다스-일반) 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I202', 'OK')
print("202 PASS")
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
    # google_sheet.update_acell('I211', 'OK')
    print("211 PASS")
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    # google_sheet.update_acell('I213', 'OK')
    print("213 PASS")
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
# google_sheet.update_acell('I216', 'OK')
print("216 PASS")
# google_sheet.update_acell('I217', 'OK')
print("217 PASS")
time.sleep(2)

print("출고 관리 - DAS 피킹 자동화상품002(다스-일반) 완료")



cell_data = google_sheet.acell('H219').value 

print("출고 관리 - DAS 피킹 자동화상품002(다스-당일) 시작")
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I219', 'OK')
print("219 PASS")
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
    # google_sheet.update_acell('I229', 'OK')
    print("229 PASS")
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
    # google_sheet.update_acell('I231', 'OK')
    print("213 PASS")
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
# google_sheet.update_acell('I234', 'OK')
print("234 PASS")
# google_sheet.update_acell('I235', 'OK')
print("235 PASS")
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
    if wms_str_loop.text == "바로 출고 처리":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break

time.sleep(2)

# google_sheet.update_acell('I236', 'OK')
print("236 PASS")



print("##########")
print("출고 관리 - 출고 처리 - DAS 출고 처리 이동")


print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(DAS-일반) 시작")
cell_data = google_sheet.acell('H237').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I219', 'OK')
print("219 PASS")
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
# google_sheet.update_acell('I240', 'OK')
print("240 PASS")
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
# google_sheet.update_acell('I243', 'OK')
print("243 PASS")
time.sleep(2)




#########################################################################################################################
#########################################################################################################################
#########################################################################################################################










print("출고 관리 - 출고 처리 - DAS 출고 처리 자동화상품002(DAS-당일) 시작")
cell_data = google_sheet.acell('H245').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I245', 'OK')
print("245 PASS")
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
# google_sheet.update_acell('I248', 'OK')
print("248 PASS")
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
# google_sheet.update_acell('I251', 'OK')
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

# google_sheet.update_acell('I253', 'OK')
print("253 PASS")



print("##########")
print("출고 관리 - 출고 처리 - 바로 출고 처리 이동")


print("출고 관리 - 출고 처리 - 바로 출고 처리 자동화상품002(바로-일반) 시작")
cell_data = google_sheet.acell('H254').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I254', 'OK')
print("254 PASS")
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
# google_sheet.update_acell('I256', 'OK')
print("256 PASS")
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
# # google_sheet.update_acell('I259', 'OK')
print("259 PASS")
time.sleep(2)






print("출고 관리 - 출고 처리 - 바로 출고 처리 자동화상품002(바로-당일) 시작")
cell_data = google_sheet.acell('H260').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='task_group_id']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I260', 'OK')
print("260 PASS")
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
# google_sheet.update_acell('I262', 'OK')
print("262 PASS")
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
# google_sheet.update_acell('I265', 'OK')
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

# google_sheet.update_acell('I266', 'OK')
print("266 PASS")



print("##########")
print("출고 관리 - 출고 처리 - 개별 출고 처리 이동")


print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-일반) 시작")

cell_data = google_sheet.acell('H267').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I267', 'OK')
print("267 PASS")
time.sleep(2)


# 셀러명(대표자) 가져오기
cell_data = google_sheet.acell('H268').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 셀러명(대표자) 확인
    google_sheet.update_acell('I268', 'Pass')            
else:
    google_sheet.update_acell('I268', 'Failed')

time.sleep(2)

cell_data = google_sheet.acell('H269').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I269', 'OK')
print("269 PASS")
time.sleep(2)




wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="stock"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H270').value 
    if cell_data == element.text: # 주문수량
        google_sheet.update_acell('I270', 'Pass')            
    else:
        google_sheet.update_acell('I270', 'Failed')   

time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="delivery_stock"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H271').value 
    if cell_data == element.text: # 출고수량
        google_sheet.update_acell('I271', 'Pass')            
    else:
        google_sheet.update_acell('I271', 'Failed')    



# SKU 코드 입력하기
try:
    cell_data = google_sheet.acell('H272').value 
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    # google_sheet.update_acell('I269', 'OK')
    print("272 PASS")
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
    time.sleep(2)
    
    # google_sheet.update_acell('I231', 'OK')
    print("272 PASS")
    time.sleep(2)

except:
    print("#SKU 코드 입력하기 오류 발생 또는 횟수 모자람")
    pass

time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="delivery_stock"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H273').value 
    if cell_data == element.text: # 출고수량
        google_sheet.update_acell('I273', 'Pass')            
    else:
        google_sheet.update_acell('I273', 'Failed')    



cell_data = google_sheet.acell('H274').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I274', 'OK')
print("274 PASS")
print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-일반) 종료")
time.sleep(2)





print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-당일) 시작")

cell_data = google_sheet.acell('H275').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I267', 'OK')
print("275 PASS")
time.sleep(2)


# 셀러명(대표자) 가져오기
cell_data = google_sheet.acell('H276').value 
wms_test_result = driver.find_elements(By.CSS_SELECTOR, 'td.MuiTableCell-root.MuiTableCell-body.MuiTableCell-sizeSmall')
for wms_str_loop in wms_test_result:
    wms_str_loop_text = wms_str_loop.text
    if cell_data == wms_str_loop_text:
        wms_str_loop_result = wms_str_loop_text
        break
        
        
if cell_data == wms_str_loop_result: # 셀러명(대표자) 확인
    google_sheet.update_acell('I276', 'Pass')            
else:
    google_sheet.update_acell('I276', 'Failed')

time.sleep(2)


cell_data = google_sheet.acell('H277').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I269', 'OK')
print("277 PASS")
time.sleep(2)




wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="stock"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H278').value 
    if cell_data == element.text: # 주문수량
        google_sheet.update_acell('I278', 'Pass')            
    else:
        google_sheet.update_acell('I278', 'Failed')   

time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="delivery_stock"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H279').value 
    if cell_data == element.text: # 출고수량
        google_sheet.update_acell('I279', 'Pass')            
    else:
        google_sheet.update_acell('I279', 'Failed')    



# SKU 코드 입력하기
try:
    cell_data = google_sheet.acell('H280').value 
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(cell_data) #  피킹리스트 바코드번호 입력
    driver.find_element(By.CSS_SELECTOR, "div>input[name='product_barcode']").send_keys(Keys.ENTER) # 검색 적용
    # google_sheet.update_acell('I280', 'OK')
    print("280 PASS")
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
    time.sleep(2)
    
    # google_sheet.update_acell('I280', 'OK')
    print("280 PASS")
    time.sleep(2)

except:
    print("#SKU 코드 입력하기 오류 발생 또는 횟수 모자람")
    pass

time.sleep(2)

wms_test_result = driver.find_elements(By.CSS_SELECTOR, '.ag-cell.ag-cell-not-inline-editing.ag-cell-normal-height.ag-cell-value[col-id="delivery_stock"]')
for element in wms_test_result:
    cell_data = google_sheet.acell('H281').value 
    if cell_data == element.text: # 출고수량
        google_sheet.update_acell('I281', 'Pass')            
    else:
        google_sheet.update_acell('I281', 'Failed')    



cell_data = google_sheet.acell('H282').value 
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(cell_data) #  송장번호 입력
driver.find_element(By.CSS_SELECTOR, "div>input[name='delivery_number']").send_keys(Keys.ENTER) # 검색 적용
# google_sheet.update_acell('I282', 'OK')
print("282 PASS")
time.sleep(2)


print("출고 관리 - 출고 처리 - 개별 출고 처리 자동화상품003(개별-당일) 종료")



#########################################################################################################################
##### 출고 관리 - 출고 회차별 조회 진행 #####
time.sleep(2)

# 출고 관리 -> 출고 회차별 조회 이동(230213)
wms_test_result = driver.find_elements(By.CSS_SELECTOR,'span.MuiBox-root')
for wms_str_loop in wms_test_result:
    #print(wms_str_loop.text,"\n")
    if wms_str_loop.text == "출고 회차별 조회":
        #print("클릭 시도\n")
        wms_str_loop.click()
        #print("클릭 완료\n")
        break


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
# google_sheet.update_acell('I133', 'OK')
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


# 총주문수량 컬럼 선택
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("총주문수량") # 컬럼 검색 필드 - 총주문수량력
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 총주문수량 체크박스 xPATH값 조합


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제



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




# 테이블(리스트) -> 지시자ID
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='지시자ID 필터 입력']").send_keys(wms_login_id) 
print("지시자ID 필터 검색")
# google_sheet.update_acell('I183', 'OK')
print("183 PASS")
time.sleep(2)



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(바로-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_now_normal) 
print("출고회차번호 검색 -> 자동화상품001(바로-일반)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="user_id"]') # 배송타입 : user_id
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H285').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I285', 'Pass')            
else:
    google_sheet.update_acell('I285', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="pick_type"]') # 출고방식 : pick_type
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H286').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I286', 'Pass')            
else:
    google_sheet.update_acell('I286', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="12"]') # 총주문수량 : 12
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="out_stock_count"]') # 총주문수량 : out_stock_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H288').value 
if cell_data == wms_search_list_row_result_check_1: # 총주문수량 확인
    google_sheet.update_acell('I288', 'Pass')            
else:
    google_sheet.update_acell('I288', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="working_count"]') # 출고진행 : working_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H289').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I289', 'Pass')            
else:
    google_sheet.update_acell('I289', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="completed_count"]') # 출고완료 : completed_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H290').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I290', 'Pass')            
else:
    google_sheet.update_acell('I290', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)

print("자동화상품001(바로-일반) 데이터 확인 완료")




# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(바로-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_now_today) 
print("출고회차번호 검색 -> 자동화상품001(바로-당일)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(바로-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="user_id"]') # 배송타입 : user_id
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H293').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I293', 'Pass')            
else:
    google_sheet.update_acell('I293', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="pick_type"]') # 출고방식 : pick_type
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H294').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I294', 'Pass')            
else:
    google_sheet.update_acell('I294', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="12"]') # 총주문수량 : 12
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="out_stock_count"]') # 총주문수량 : out_stock_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H296').value 
if cell_data == wms_search_list_row_result_check_1: # 총주문수량 확인
    google_sheet.update_acell('I296', 'Pass')            
else:
    google_sheet.update_acell('I296', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="working_count"]') # 출고진행 : working_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H297').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I297', 'Pass')            
else:
    google_sheet.update_acell('I297', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="completed_count"]') # 출고완료 : completed_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H298').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I298', 'Pass')            
else:
    google_sheet.update_acell('I298', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)

print("자동화상품001(바로-당일) 데이터 확인 완료")




# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(다스-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_das_normal) 
print("출고회차번호 검색 -> 자동화상품001(다스-일반)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(다스-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="user_id"]') # 배송타입 : user_id
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H301').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I301', 'Pass')            
else:
    google_sheet.update_acell('I301', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="pick_type"]') # 출고방식 : pick_type
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H302').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I302', 'Pass')            
else:
    google_sheet.update_acell('I302', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="12"]') # 총주문수량 : 12
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="out_stock_count"]') # 총주문수량 : out_stock_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H304').value 
if cell_data == wms_search_list_row_result_check_1: # 총주문수량 확인
    google_sheet.update_acell('I304', 'Pass')            
else:
    google_sheet.update_acell('I304', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="working_count"]') # 출고진행 : working_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H305').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I305', 'Pass')            
else:
    google_sheet.update_acell('I305', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="completed_count"]') # 출고완료 : completed_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H306').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I306', 'Pass')            
else:
    google_sheet.update_acell('I306', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)

print("자동화상품001(다스-일반) 데이터 확인 완료")



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(다스-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_das_today) 
print("출고회차번호 검색 -> 자동화상품001(다스-당일)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(다스-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="user_id"]') # 배송타입 : user_id
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H309').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I309', 'Pass')            
else:
    google_sheet.update_acell('I309', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="pick_type"]') # 출고방식 : pick_type
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H310').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I310', 'Pass')            
else:
    google_sheet.update_acell('I310', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="12"]') # 총주문수량 : 12
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="out_stock_count"]') # 총주문수량 : out_stock_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H312').value 
if cell_data == wms_search_list_row_result_check_1: # 총주문수량 확인
    google_sheet.update_acell('I312', 'Pass')            
else:
    google_sheet.update_acell('I312', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="working_count"]') # 출고진행 : working_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H313').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I313', 'Pass')            
else:
    google_sheet.update_acell('I313', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="completed_count"]') # 출고완료 : completed_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H314').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I314', 'Pass')            
else:
    google_sheet.update_acell('I314', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)

print("자동화상품001(다스-당일) 데이터 확인 완료")



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(개별-일반)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_each_normal) 
print("출고회차번호 검색 -> 자동화상품001(개별-일반)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(개별-일반))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="user_id"]') # 배송타입 : user_id
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H317').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I317', 'Pass')            
else:
    google_sheet.update_acell('I317', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="pick_type"]') # 출고방식 : pick_type
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H318').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I318', 'Pass')            
else:
    google_sheet.update_acell('I318', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="12"]') # 총주문수량 : 12
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="out_stock_count"]') # 총주문수량 : out_stock_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H320').value 
if cell_data == wms_search_list_row_result_check_1: # 총주문수량 확인
    google_sheet.update_acell('I320', 'Pass')            
else:
    google_sheet.update_acell('I320', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="working_count"]') # 출고진행 : working_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H321').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I321', 'Pass')            
else:
    google_sheet.update_acell('I321', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="completed_count"]') # 출고완료 : completed_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H322').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I322', 'Pass')            
else:
    google_sheet.update_acell('I322', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)

print("자동화상품001(개별-일반) 데이터 확인 완료")



# 테이블(리스트) -> 출고회차번호 검색 -> 자동화상품001(개별-당일)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(wms_out_round_number_each_today) 
print("출고회차번호 검색 -> 자동화상품001(개별-당일)")
time.sleep(2)



# 출고 대상 리스트 조회 데이터 확인(테이블 - 리스트 - 자동화상품001(개별-당일))
wms_search_list_row_result = driver.find_element(By.CSS_SELECTOR,'div[class="ag-center-cols-container"]')

# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="5"]') # 배송타입 : 5
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="user_id"]') # 배송타입 : user_id
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H325').value 
if cell_data == wms_search_list_row_result_check_1: # 배송타입 확인
    google_sheet.update_acell('I325', 'Pass')            
else:
    google_sheet.update_acell('I325', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="6"]') # 출고방식 : 6
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="pick_type"]') # 출고방식 : pick_type
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H326').value 
if cell_data == wms_search_list_row_result_check_1: # 출고방식 확인
    google_sheet.update_acell('I326', 'Pass')            
else:
    google_sheet.update_acell('I326', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="12"]') # 총주문수량 : 12
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="out_stock_count"]') # 총주문수량 : out_stock_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H328').value 
if cell_data == wms_search_list_row_result_check_1: # 총주문수량 확인
    google_sheet.update_acell('I328', 'Pass')            
else:
    google_sheet.update_acell('I328', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="15"]') # 출고진행 : 15
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="working_count"]') # 출고진행 : working_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H329').value 
if cell_data == wms_search_list_row_result_check_1: # 출고진행 확인
    google_sheet.update_acell('I329', 'Pass')            
else:
    google_sheet.update_acell('I329', 'Failed')



# wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[aria-colindex="16"]') # 출고완료 : 16
wms_search_list_row_result_check_1 = wms_search_list_row_result.find_element(By.CSS_SELECTOR, 'div[col-id="completed_count"]') # 출고완료 : completed_count
wms_search_list_row_result_check_1 = wms_search_list_row_result_check_1.text
print("출고 회차별 조회 - wms_search_list_row_result_check_1   ", wms_search_list_row_result_check_1)

cell_data = google_sheet.acell('H330').value 
if cell_data == wms_search_list_row_result_check_1: # 출고완료 확인
    google_sheet.update_acell('I330', 'Pass')            
else:
    google_sheet.update_acell('I330', 'Failed')


driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='출고회차번호 필터 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)

print("자동화상품001(개별-당일) 데이터 확인 완료")



print("##########")
print("출고 관리 - 출고 회차별 조회 완료")











print("출고 일부 테스트 완료")



while(True):
    	pass
