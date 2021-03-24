import pandas as pd
import requests
from bs4 import BeautifulSoup

code_data = pd.read_html('http://kind.krx.co.kr/corpgeneral/corpList.do?method=download&searchType=13', header=0)[0]

def make_code(x):
    x = str(x)
    return '0' * (6 - len(x)) + x

code_data = code_data['종목코드']
#['종목코드'] 요것만 하면 판다스에 종목코드가 리스트로 저장될까? ㅇㅇ 가능하네
code_data = code_data.apply(make_code)
#print(code_data)

code_num = code_data[1571]
#(임시 테스트)랜덤 법인 종목코드 가져오기

fnguide_url = "http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A" + code_num + "&cID=&MenuYn=Y&ReportGB=D&NewMenuID=Y&stkGb=701"
#종목코드를 이용해서 fnquide 접속
webpage = requests.get(fnguide_url)
web_data = BeautifulSoup(webpage.content, 'html.parser')
#웹페이지 정보 가져오기
corp_name = web_data.find('h1', {'id': 'giName'})
print(corp_name)
#주식이름, 태그 하나만 가져와도 {}를 붙여야하네
#ifrs_D_A = web_data.find('div', {'class': 'um_table', 'id': 'highlight_D_A'})
#ifrs표 데이터 불러오기
#corp_name = web_data.find('h1', 'id': 'giName')


# print(webpage.text)