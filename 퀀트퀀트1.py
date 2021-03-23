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

webpage = requests.get("http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=D&NewMenuID=Y&stkGb=701")
web_data = BeautifulSoup(webpage.content, 'html.parser')
target = web_data.find('div', {'class': 'um_table', 'id': 'highlight_D_A'})
print(target)
# print(webpage.text)

#<table class="us_table_ty1 h_fix zigbg_no">