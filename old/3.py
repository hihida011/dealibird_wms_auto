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


#########################################################################################################################
# 사입 요청 파일 다운로드
# https://docs.google.com/spreadsheets/d/19X1duCg7N2npHQHGu_pPcDaji_9pDWdI/edit#gid=1830395100
# 해당 엑셀 파일을 c:\test\ 에 저장

deal_test_saip_excel_upload = 'C:\\test\\덕규_자동화_사입요청.xlsx' # 사입 요청 파일 정보


#########################################################################################################################
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

''' # hihida 221230
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


#########################################################################################################################
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

cell_data = google_sheet.acell('H8').value # 사입 요청 수량(리스트 수량) / 사입 요청 수량을 확인한다

if cell_data == deal_test_result_check:
    google_sheet.update_acell('I8', 'Pass') 
else:
    google_sheet.update_acell('I8', 'Failed')


#요청 가능 수량 체크
deal_test_result = driver.find_element(By.XPATH,'//*[@id="purchase_able_to_count"]') # 페이지 중간 왼쪽 -> 요청 가능 : X 값
deal_test_result_check = deal_test_result.text

cell_data = google_sheet.acell('H9').value # 요청 가능 수량(리스트 수량) / 요청 가능 수량을 확인한다.

if cell_data == deal_test_result_check:
    google_sheet.update_acell('I9', 'Pass') 
else:
    google_sheet.update_acell('I9', 'Failed')


#요청 불가능 수량 체크
deal_test_result = driver.find_element(By.XPATH,'//*[@id="purchase_unable_to_count"]') # 페이지 중간 왼쪽 -> 요청 불가능 : X 값
deal_test_result_check = deal_test_result.text

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

#########################################################################################################################
# 결제 완료 후 사입 요청 페이지
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



'''


#########################################################################################################################
driver.switch_to.window(tabs[0])

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



#########################################################################################################################
#### 딜리버드 -> 상품 및 재고 #####
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[8]/a').click()
time.sleep(5)
print("##########")
print("상품 및 재고 시작")


# 도매 상품명_자동화001 상품 및 재고 확인
cell_data = google_sheet.acell('H126').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
time.sleep(2)


deal_table = driver.find_element(By.XPATH, '//*[@id="productList"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0


for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        
        if deal_list_count == 8:
            cell_data = google_sheet.acell('H127').value
            deal_test_result_check = td.get_attribute("innerText")
            deal_test_result_check = deal_test_result_check.replace('도매 매장 변경','') # 도매 매장명 필드에 [버튼 내용]과 줄바꿈이 있음
            deal_test_result_check = deal_test_result_check.replace("\n",'')
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I127', 'Pass')
            else:
                google_sheet.update_acell('I127', 'Failed')           
            
        if deal_list_count == 13:
            cell_data = google_sheet.acell('H128').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I128', 'Pass')
            else:
                google_sheet.update_acell('I128', 'Failed')           


        if deal_list_count == 14:
            cell_data = google_sheet.acell('H129').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I129', 'Pass')
            else:
                google_sheet.update_acell('I129', 'Failed')           
   
        if deal_list_count == 15:
            deal_list_count = int(38)

        deal_list_count += 1
        
driver.find_element(By.ID,'search_text').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)



# 도매 상품명_자동화002 상품 및 재고 확인
cell_data = google_sheet.acell('H130').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
time.sleep(2)


deal_table = driver.find_element(By.XPATH, '//*[@id="productList"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0


for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        
        if deal_list_count == 8:
            cell_data = google_sheet.acell('H131').value
            deal_test_result_check = td.get_attribute("innerText")
            deal_test_result_check = deal_test_result_check.replace('도매 매장 변경','') # 도매 매장명 필드에 [버튼 내용]과 줄바꿈이 있음
            deal_test_result_check = deal_test_result_check.replace("\n",'')
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I131', 'Pass')
            else:
                google_sheet.update_acell('I131', 'Failed')           
            
        if deal_list_count == 13:
            cell_data = google_sheet.acell('H132').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I132', 'Pass')
            else:
                google_sheet.update_acell('I132', 'Failed')           


        if deal_list_count == 14:
            cell_data = google_sheet.acell('H133').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I133', 'Pass')
            else:
                google_sheet.update_acell('I133', 'Failed')           
   
        if deal_list_count == 15:
            deal_list_count = int(38)

        deal_list_count += 1
        
driver.find_element(By.ID,'search_text').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)





# 도매 상품명_자동화003 상품 및 재고 확인
cell_data = google_sheet.acell('H134').value # 구글 시트 -> 딜리버드 상품코드
driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)
time.sleep(2)


deal_table = driver.find_element(By.XPATH, '//*[@id="productList"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0


for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        
        if deal_list_count == 8:
            cell_data = google_sheet.acell('H135').value
            deal_test_result_check = td.get_attribute("innerText")
            deal_test_result_check = deal_test_result_check.replace('도매 매장 변경','') # 도매 매장명 필드에 [버튼 내용]과 줄바꿈이 있음
            deal_test_result_check = deal_test_result_check.replace("\n",'')
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I135', 'Pass')
            else:
                google_sheet.update_acell('I135', 'Failed')           
            
        if deal_list_count == 13:
            cell_data = google_sheet.acell('H136').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I136', 'Pass')
            else:
                google_sheet.update_acell('I136', 'Failed')           


        if deal_list_count == 14:
            cell_data = google_sheet.acell('H137').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I137', 'Pass')
            else:
                google_sheet.update_acell('I137', 'Failed')           
   
        if deal_list_count == 15:
            deal_list_count = int(38)

        deal_list_count += 1
        
driver.find_element(By.ID,'search_text').send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
time.sleep(5)






'''
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        deal_test_result_check = td.get_attribute("innerText")
        print(deal_list_count, "번 ", deal_test_result_check)
        
        deal_list_count += 1


driver.find_element(By.ID,'search_text').send_keys(cell_data) # 검색어 창에 "딜리버드 상품 코드" 입력
time.sleep(2)
driver.find_element(By.ID,'search_text').send_keys(Keys.ENTER)


wms_test_result = "//*[@id=\"" + cell_data + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
element = driver.find_element(By.XPATH, wms_test_result)

testaaaaa = len(element)
print(testaaaaa)

testaaaaa = int(testaaaaa)
print(testaaaaa)
aaaaaaa = int(0)

while testaaaaa >= aaaaaaa : 
    bbbb = element[aaaaaaa]
    print(aaaaaaa, "번 !", bbbb)
    aaaaaaa = aaaaaaa + 1

'''






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
 