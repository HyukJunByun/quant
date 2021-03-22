import pandas as pd
import requests
from bs4 import BeautifulSoup

code_data = pd.read_html('http://kind.krx.co.kr/corpgeneral/corpList.do?method=download&searchType=13', header=0)[0]

code_data = code_data['종목코드']
# print(code_data)
#['종목코드'] 요것만 하면 판다스에 종목코드가 리스트로 저장될까?

def make_code(x):
    x = str(x)
    return '0' * (6 - len(x)) + x

code_data['종목코드'] = code_data['종목코드'].apply(make_code)

webpage = requests.get("http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=D&NewMenuID=Y&stkGb=701")
# print(webpage.text)
