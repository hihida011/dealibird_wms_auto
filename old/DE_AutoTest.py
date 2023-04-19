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

deal_test_saip_excel_upload = 'C:\\test\\동은_자동화_사입요청.xlsx' # 사입 요청 파일 정보


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
google_url = 'https://docs.google.com/spreadsheets/d/1fMf-pNUosGPMJ6evQJLTRRhAGpqZwEIhmC48wxLXHMQ/edit#gid=651282265' # 테스트 시나리오 엑셀 주소

google_doc = gc.open_by_url(google_url)
google_sheet = ''

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
##어드민 -> 딜리버드 셀러 이동
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

#딜리버드 바로가기 클릭
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/div[1]/div/ul/li[1]/div/span').click()


#########################
#딜리버드 -> 사입 요청
time.sleep(3)

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

time.sleep(3)

#사입 요청 목록 체크
deal_test_result = driver.find_element(By.XPATH,'//*[@id="purchase_totalCount"]') # 페이지 중간 왼쪽 -> 사입 요청 : X 값
deal_test_result_check = deal_test_result.text

print(deal_test_result_check)
cell_data = google_sheet.acell('H8').value # 사입 요청 수량(리스트 수량)

if cell_data == deal_test_result_check:
    google_sheet.update_acell('I8', 'Pass') 
else:
    google_sheet.update_acell('I8', 'Failed')

# 사입 요청 리스트(테이블) 체크
deal_table = driver.find_element(By.XPATH, '//*[@id="purchasesList"]') # 리스트(테이블) 전체 경로
deal_tbody = deal_table.find_element(By.TAG_NAME,'tbody')
deal_list_count = 0
for tr in deal_tbody.find_elements(By.TAG_NAME,'tr'):
    for td in tr.find_elements(By.TAG_NAME,'td'):
        
        if deal_list_count == 11:
            cell_data = google_sheet.acell('H9').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I9', 'Pass')
            else:
                google_sheet.update_acell('I9', 'Failed')           
            
        if deal_list_count == 12:
            cell_data = google_sheet.acell('H10').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I10', 'Pass')
            else:
                google_sheet.update_acell('I10', 'Failed')           
                
        if deal_list_count == 37:
            cell_data = google_sheet.acell('H11').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I11', 'Pass')
            else:
                google_sheet.update_acell('I11', 'Failed')           
                
        if deal_list_count == 38:
            cell_data = google_sheet.acell('H12').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I12', 'Pass')
            else:
                google_sheet.update_acell('I12', 'Failed')           
                
        if deal_list_count == 63:
            cell_data = google_sheet.acell('H13').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I13', 'Pass')
            else:
                google_sheet.update_acell('I13', 'Failed')           
                
        if deal_list_count == 64:
            cell_data = google_sheet.acell('H14').value
            deal_test_result_check = td.get_attribute("innerText")
            if cell_data == deal_test_result_check:
                google_sheet.update_acell('I14', 'Pass')
            else:
                google_sheet.update_acell('I14', 'Failed')           
   
        deal_list_count += 1

driver.find_element(By.XPATH, '//*[@id="purchasesList_wrapper"]/div[1]/div/div/button[9]').click() # 페이지 중간 오른쪽 [사입 요청하기] 버튼

#driver.execute_script("arguments[0].send_keys();", element)

time.sleep(2)

# 사입 요청 버튼
driver.find_element(By.XPATH, '/html/body/div[5]/div/div[3]/button[3]').click() # 얼럿 / X건의 상품을~~ 진행하시겠습니까? -> [네] 버튼

time.sleep(2)
# 결제 정보 모달
driver.find_element(By.ID, 'method_SINSANGPOINT').click() # 모달 / 결제 수단 선택 -> 신상캐시 버튼
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="confirmCollapse"]/div[2]/div/label').click() # 모달 -> 이용 약관 -> 전체 동의합니다. 버튼
#driver.find_element(By.XPATH, '//*[@id="payment_button"]').click() # 모달 -> [결제하기] 버튼



##어드민 사입마감 처리
driver.switch_to.window(tabs[1])


#어드민 로그인 진행
driver.find_element(By.ID, 'ssm_admin_email').send_keys(deal_admin_login_id)
driver.find_element(By.ID, 'ssm_admin_password').send_keys(deal_admin_login_password)
driver.find_element(By.NAME, 'commit').click()
time.sleep(2)

driver.find_element(By.XPATH,'//*[@id="navbarSupportedContent"]/ul/li[8]/a').click() # 사입 마감 시간 설정
time.sleep(2)

driver.find_element(By.XPATH,'//*[@id="scheduleList_wrapper"]/div[1]/div[2]/div/button').click() # [기본 사입 일시 설정] 버튼
time.sleep(2)

test_now = datetime.now() # 테스트 시점 날짜 시간 구하기
test_date = test_now.date() # 날짜만
test_time = test_now + timedelta(minutes=3) # 현재 시간에서 3분 더하기

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



"""
table = driver.find_element(By.XPATH, '//*[@id="purchasesList"]')
rows = table.find_elements(By.TAG_NAME, 'tr')
print(rows)
print("###########################")

for index, value in enumerate(rows):
    print("index=",index)
    print("!!!!!!")
    print("value=", value)
    print("!!!!!!")
    #body = value.find_elements(By.ID, 'Row_1441a7909c087dbbe7ce59881b9df8b9_0')[0]
    body = value.find_element(By.XPATH, '//*[@id="Row_1441a7909c087dbbe7ce59881b9df8b9_0"]')
    
    #print(body.text)
    body_cost = body.text
    print(body_cost[8])
    #print(body[8])
   """ 

#for tr in table.find_element(By.TAG_NAME,'tr'):
#    td = tr.find
    


"""
deal_test_result = driver.find_element(By.CSS_SELECTOR,'div[class=dataTables_scrollBody]')
print(deal_test_result)
deal_test_result_index = deal_test_result.find_element(By.CSS_SELECTOR,'data-dt-row[0]')

print(deal_test_result_index)
"""
#element = driver.find_element(By.XPATH,'//*[@id="excel_file"]')
#driver.execute_script("arguments[0].send_keys();", element)


#driver.find_element(By.ID,'excel_file').click()
#driver.find_element(By.XPATH, '//*[@id="excel_file"]').click()



#########################
#########################
#########################
#########################
#########################
#########################
#########################
#########################
#########################


"""
#WMS 로그인 진행
driver.find_element(By.ID, 'login').send_keys(wms_login_id)
driver.find_element(By.ID, 'password').send_keys(wms_login_passWord)
driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/form/div[2]/button').click()

driver.implicitly_wait(5)

#WMS 로그인 정상 실행 여부
#WMS 오른쪽 상단의 로그인 계정 정보 가져오기 ex hihida 님
wms_test_result = driver.find_element(By.XPATH,'//*[@id="root"]/div/div/div/div/div[1]/div/div/div/div[2]/div[2]')
wms_test_result_check = wms_test_result.text


#로그인시도 후 정상으로 로그인 되었는지 체크
#로그인 ID와 WMS 상단에서 가져온 str값이 포함되었는지 확인
if wms_login_id in wms_test_result_check: 
    #구글 시트의 기대 결과값 가져오기
    cell_data = google_sheet.acell('A2').value
    
    #구글 시트 기대결과와 테스트 결과가 같은지 체트
    if cell_data == wms_login_id: 
        #엑셀에 결과값 입력
        google_sheet.update_acell('B2', 'Pass') 
    else:
        google_sheet.update_acell('B2', 'Fa')
  
else:
    google_sheet.update_acell('B2', 'Fa')

"""
while(True):
    	pass
 