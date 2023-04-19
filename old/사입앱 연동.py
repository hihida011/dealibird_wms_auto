import requests
from bs4 import BeautifulSoup
import json

buyer_wsIdx_name = '23124' # 사입 요청한 도내 wsIdx 값


#########################################################################################################################
########### 사입앱 로그인 ##########
#########################################################################################################################
buyer_login_url = 'https://buyer.qa.sinsang.market/api/v1/session' # 사입 로그인 URL
buyer_login_header = {'Content-Type' : 'application/json', "User-Agent" : "Mozilla/5.0"} # 로그인 시 헤더
buyer_login_data = {
    'password':'1234',
    'user':'qa_smkim'
} # 로그인 시 Body 정보 : 로그인 계정

buyer_response = requests.post(url=buyer_login_url, headers=buyer_login_header, params=buyer_login_data) # 로그인 시도
print("사입앱 로그인 성공\n", buyer_response)

buyer_login_content = buyer_response.content # 로그인 후 리턴되는 값(여러 정보가 있음)
buyer_login_content_data = json.loads(buyer_login_content) # JSON 문자열을 Python 객체로 변환
buyer_login_content_accesstoken = buyer_login_content_data["content"]["accessToken"] # accessToken 저장




#########################################################################################################################
########### 사입 리스트 상세 : 사입 예정 SKU의 ID 가져오기 ##########

buyer_login_accesstoken_header = {'Content-Type': 'application/json',
#'access-token': login_accesstoken,
'access-token': buyer_login_content_accesstoken,
'User-Agent': 'Mozilla/5.0',
'Cache-Control': 'no-cache',
'Accept': '*/*',
'Host': 'buyer.qa.sinsang.market',
'Accept-Encoding': 'gzip, deflate, br',
'Connection': 'keep-alive'} # accessToken을 가지고 사입 리스트 상세 조회


# id_search_url = 'https://buyer.qa.sinsang.market/buyer/api/dealibird/buying/order_detail?wsIdx=23124&orderType=purchase&warehouse=B1'

buyer_id_search_url = 'https://buyer.qa.sinsang.market/buyer/api/dealibird/buying/order_detail?wsIdx='+ buyer_wsIdx_name + '&orderType=purchase&warehouse=B1' # 사입 리스트 상세 조회 URL

print(buyer_wsIdx_name)

buyer_response = requests.get(url=buyer_id_search_url, headers=buyer_login_accesstoken_header) # 사입 리스트 상세 조회 시도
print("사입 리스트 상세 조회 성공\n", buyer_response)

buyer_id_search_content = buyer_response.content# 조회 후 리턴되는 값(여러 정보가 있음)

buyer_id_search_content_data = json.loads(buyer_id_search_content) # JSON 문자열을 Python 객체로 변환

buyer_id_search_content_ID_data = [] # 여러개의 ID 정보 저장을 위한 배열
buyer_id_search_int = int(0) # 배열 Len 체크
for product in buyer_id_search_content_data["content"]["products"]:
    buyer_id_search_content_ID_data.append(product["id"]) # id 정보를 배열(id_search_content_ID_data)에 저장
    print(buyer_id_search_int, "번쨰 ID는", buyer_id_search_content_ID_data[buyer_id_search_int])
    buyer_id_search_int = buyer_id_search_int +1

buyer_id_search_int = buyer_id_search_int -1 # 최종 배열 길이 체크


#########################################################################################################################
########### 사입 상품 옵션 저장 : SKU의 ID로 모두 사입으로 전송 ##########

buyer_order_status_url = 'https://buyer.qa.sinsang.market/buyer/api/dealibird/buying/order_status' # 사입 상태 전송 URL
                    
buyer_order_status_data = {
    "orderType" : "purchase", # 신규 주문
	"items": [
		{
            "id": buyer_id_search_content_ID_data[0],
            "purchasedCount" : 5
		},
        {
            "id": buyer_id_search_content_ID_data[1],
            "purchasedCount" : 8
		},
        {
            "id": buyer_id_search_content_ID_data[2],
            "purchasedCount" : 10
		}
	]
} # 로그인 시 Body 정보 : 로그인 계정


buyer_response = requests.post(url=buyer_order_status_url, headers=buyer_login_accesstoken_header, data=json.dumps(buyer_order_status_data)) # 사입 상태 전송 시도

print("사입앱 사입 성공 전달 성공\n", buyer_response)


while(True):
    	pass