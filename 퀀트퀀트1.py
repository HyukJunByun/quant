import pandas as pd

code_data = pd.read_html('http://kind.krx.co.kr/corpgeneral/corpList.do?method=download&searchType=13', header=0)[0]

code_data = code_data['종목코드']
print(code_data)