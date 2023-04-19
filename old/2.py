import selenium.webdriver.support.ui as ui

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from datetime import datetime, timedelta


import time
import sys
import json

import gspread
from oauth2client.service_account import ServiceAccountCredentials


#########################
# 사입 요청 파일 다운로드
# https://docs.google.com/spreadsheets/d/19X1duCg7N2npHQHGu_pPcDaji_9pDWdI/edit#gid=1830395100
# 해당 엑셀 파일을 c:\test\ 에 저장
# 사입 요청 파일 정보
deal_test_sku1_price = '100'
deal_test_sku1_count = '3'
deal_test_sku2_price = '100'
deal_test_sku2_count = '3'
deal_test_sku3_price = '100'
deal_test_sku3_count = '3'

#########################
# 딜리버드 테스트 기본 설정
deal_admin_login_id = 'hihida@deali.net'
deal_admin_login_password = '!incasys0'
deal_admin_url = 'https://dealibird.qa.sinsang.market/ssm_admins/sign_in'
deal_seller_login_id = 'chy_soqa10'
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
#google_url = 'https://docs.google.com/spreadsheets/d/1Ote9hWcCPiOFyvcWCpSzybukKgytf0HfVvxSurW2Yns/edit?usp=sharing' # 구글 스프레트시트 URL
google_url = 'https://docs.google.com/spreadsheets/d/1fMf-pNUosGPMJ6evQJLTRRhAGpqZwEIhmC48wxLXHMQ/edit#gid=651282265'

google_doc = gc.open_by_url(google_url)

#google_sheet = google_doc.worksheet('시트1') # 구글 시트
google_sheet = google_doc.worksheet('TC')
google_email = 'client_email: fulfillment-test@fulfillment-371610.iam.gserviceaccount.com'
#
#########################
# 크롭 탭 2개 실행
chrome_options = Options()
chrome_options.add_argument('--start-maximized')

driver = webdriver.Chrome(chrome_options=chrome_options)
driver.execute_script('window.open("about:blank", "_blank");')
driver.execute_script('window.open("about:blank", "_blank");')

tabs = driver.window_handles

driver.switch_to.window(tabs[0])
#driver.get(deal_admin_url) #딜리버드->어드민->셀러 진입
driver.get(deal_seller_url) #신상마켓 소매 -> 딜리버드 진입

driver.maximize_window()

driver.switch_to.window(tabs[1])
driver.get(deal_admin_url)

driver.switch_to.window(tabs[2])
driver.get(wms_url)

action = ActionChains(driver)


#########################
driver.switch_to.window(tabs[2])


# 테스트를 위한 임시 저장 hihida
# deal_wms_purchase_number = "19590"

# wms 각 항목 -> 조회 -> 리스트의 스크롤이 생기면서 데이터 로드의 어려움(현재 보여지는 화면의 데이터만 가져옴)
# pyautogui을 사용 -> wms 로그인 화면에서 글꼴 축소 -> 항목 조회 -> 리스트의 모든 데이터가 보이게 됨
#pyautogui_count = 0
#while pyautogui_count < 8:
#    pyautogui.hotkey('ctrl', '-') # 화면 축소
#    pyautogui_count = pyautogui_count + 1

#time.sleep(80) # 사입 마감 처리 시간 후 로그인 시도 hihida
time.sleep(2) # 테스트 임시


#WMS 로그인 진행
driver.find_element(By.ID, 'login').send_keys(wms_login_id)
driver.find_element(By.ID, 'password').send_keys(wms_login_passWord)
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/form/div[2]/button').click()
time.sleep(2)
print("##########")
print("wms 로그인 완료")


time.sleep(5)
#########################################################################################################################
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



wms_test_result = driver.find_element(By.CSS_SELECTOR, "div>[aria-posinset='2']") # 사입성공수량 체크박스 xPATH값 조합
print("aria-posinset='2'",wms_test_result)
time.sleep(5)

wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
print("aria-describedby",wms_test_result)
time.sleep(5)


wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() 
time.sleep(5)



driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.CONTROL + "a") # 컬럼 검색 필드 - 입력값 전체 선택
time.sleep(2)
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys(Keys.DELETE) # 컬럼 검색 필드 - 입력값 삭제
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
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
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
driver.find_element(By.CSS_SELECTOR, "div>input[aria-label='필터 컬럼 입력']").send_keys("주문번호") # 컬럼 검색 필드 - 주문번호
time.sleep(5)

wms_test_result = driver.find_element(By.CLASS_NAME, 'ag-virtual-list-item.ag-column-select-virtual-list-item') # 체크박스의 상위 그룹의 클래스 네임으로 정보 취득
wms_test_result = wms_test_result.get_attribute("aria-describedby") # 동적으로 변화되는 aria-describedbyD값 취득
wms_test_result = "//*[@id=\"" + wms_test_result + "\"]""" # 컬럼 선택을 위한 각 체크박스 xPATH값 조합
driver.find_element(By.XPATH, wms_test_result).click() # 주문번호 체크박스 xPATH값 조합



















while(True):
    	pass
 