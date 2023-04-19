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
deal_wms_purchase_number = "19590"

# wms 각 항목 -> 조회 -> 리스트의 스크롤이 생기면서 데이터 로드의 어려움(현재 보여지는 화면의 데이터만 가져옴)
# pyautogui을 사용 -> wms 로그인 화면에서 글꼴 축소 -> 항목 조회 -> 리스트의 모든 데이터가 보이게 됨
pyautogui_count = 0
while pyautogui_count < 8:
    pyautogui.hotkey('ctrl', '-') # 화면 축소
    pyautogui_count = pyautogui_count + 1



#WMS 로그인 진행
driver.find_element(By.ID, 'login').send_keys(wms_login_id)
driver.find_element(By.ID, 'password').send_keys(wms_login_passWord)
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/form/div[2]/button').click()

time.sleep(5)



##### 입고 관리 - 입고 검수진행 #####


driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[1]/div/a[3]').click() # 입고 관리 -> 입고 검수진행 이동
time.sleep(2)

cell_date = "느낌표"
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매처를 입력해 주세요']").send_keys(cell_date) # 도매처를 입력해주세요에 도매 명 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매처를 입력해 주세요']").send_keys(Keys.ENTER) # 입력 필드 하단에 도매 드룹다운 메뉴 출력
time.sleep(2)

element = driver.find_element(By.XPATH, '/html/body/div[3]/div[3]/ul/li') # 해당 도매를 클릭
element.send_keys(Keys.ENTER)



wms_test_result = driver.find_element(By.CLASS_NAME, 'MuiButtonBase-root.MuiTab-root.MuiTab-textColorPrimary.Mui-selected.css-1fs0d0o')
wms_test_result = wms_test_result.get_attribute("ID")
print(wms_test_result)
wms_search_list_row_index_str = "//*[@id=\"" + wms_test_result + "\"]/span[1]"


wms_test_result = driver.find_element(By.XPATH, wms_search_list_row_index_str) 
wms_test_result_check = wms_test_result.text
#wms_test_result_check = str(wms_test_result_check)
print(wms_test_result_check)
wms_test_result_check = wms_test_result_check.replace('신규 주문 (','')
print(wms_test_result_check)
wms_test_result_check = wms_test_result_check.replace('건)','')

#총 도매단가, 총 장끼 수량 가져오기
#wms_test_result = driver.find_element(By.CLASS_NAME, 'MuiButtonBase-root.MuiTab-root.MuiTab-textColorPrimary.Mui-selected.css-1fs0d0o')

wms_test_result_check = wms_test_result.text
time.sleep(5)
#wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[4]/h2[2]') # 총 도매단가 / 위치 강제
wms_test_result_check = driver.find_element(By.CLASS_NAME, 'MuiTypography-root.MuiTypography-h2.css-i8nswv').text
#wms_test_result_check = wms_test_result.text
print("wms_test_result ->!", wms_test_result_check,"!")

####### hihida 1227

cell_date = google_sheet.acell('H48').value # 총 도매단가 확인한다.
print("cell_date ->!", cell_date,"!")

if cell_date == wms_test_result_check:
    google_sheet.update_acell('I48', 'Pass')
else:
    google_sheet.update_acell('I48', 'Failed')


wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[4]/h2[4]') # 총 장끼 수량 / 위치 강제
wms_test_result_check = wms_test_result.text

cell_date = google_sheet.acell('H49').value # 총 장끼수량 확인한다.
if cell_date == wms_test_result_check:
    google_sheet.update_acell('I49', 'Pass')
else:
    google_sheet.update_acell('I49', 'Failed')




"""
cell_date = google_sheet.acell('H43').value 

driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매처를 입력해 주세요']").send_keys(cell_date) # 도매처를 입력해주세요에 도매 명 입력
driver.find_element(By.CSS_SELECTOR, "div>input[placeholder='도매처를 입력해 주세요']").send_keys(Keys.ENTER) # 입력 필드 하단에 도매 드룹다운 메뉴 출력
time.sleep(2)

element = driver.find_element(By.XPATH, '/html/body/div[3]/div[3]/ul/li') # 해당 도매를 클릭
element.send_keys(Keys.ENTER)


#driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='셀러ID 필터 입력']").send_keys(deal_seller_login_id) # 테이블(리스트) -> 셀러ID 입력
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='주문번호 필터 입력']").send_keys(deal_wms_purchase_number) # 테이블(리스트) -> 주문번호
time.sleep(2)


#총 도매단가, 총 장끼 수량 가져오기
wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[4]/h2[2]') # 총 도매단가 / 위치 강제
wms_test_result_check = wms_test_result.text

cell_date = google_sheet.acell('H48').value # 총 도매단가 확인한다.

if cell_date == wms_test_result_check:
    google_sheet.update_acell('I48', 'Pass')
else:
    google_sheet.update_acell('I48', 'Failed')


wms_test_result = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[4]/h2[4]') # 총 장끼 수량 / 위치 강제
wms_test_result_check = wms_test_result.text

cell_date = google_sheet.acell('H49').value # 총 장끼수량 확인한다.
if cell_date == wms_test_result_check:
    google_sheet.update_acell('I49', 'Pass')
else:
    google_sheet.update_acell('I49', 'Failed')



#driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[9]/div[2]/button[2]').click() # 하단 -> [입고 검수 진행 처리 (소봉 바코드 출력) ] 버튼
driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/div/div[8]/div[2]/button[2]').click() # 하단 -> [입고 검수 진행 처리 (소봉 바코드 출력) ] 버튼
time.sleep(2)

alert = driver.switch_to.alert 
alert.accept() # 얼럿 확인
# alert.dismiss()# 취소
time.sleep(2)
"""


















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
 