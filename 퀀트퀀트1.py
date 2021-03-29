import pandas as pd
import requests
from bs4 import BeautifulSoup, SoupStrainer
import xlwings as xw
import time
import gc

kisrating_url = 'https://www.kisrating.com/ratingsStatistics/statics_spread.do'
webbpage = requests.get(kisrating_url)
# 회사채 수익률 웹페이지 가져오기
only_table = SoupStrainer('div', {'class': 'table_ty1'})
webb_data = BeautifulSoup(webbpage.content, 'lxml', parse_only=only_table)
# 회사채 수익률 표 가져오기
bbb_table = webb_data.find_all('td')
# 회사채 수익률 표 숫자만 전부 가져오기
bbb_data = bbb_table[98].text
# 회사채 BBB- 5년 수익률 가져오기 = 요구수익률
webb_data.decompose()
# 정보 빼고 나면 웹페이지는 삭제(스피드업)
webbpage.close()
webbpage = None
del(webbpage)


code_data = pd.read_html('http://kind.krx.co.kr/corpgeneral/corpList.do?method=download&searchType=13', header=0)[0]
start = time.time()
# 시작 시간

def make_code(x):
    x = str(x)
    return '0' * (6 - len(x)) + x


def hms(s):
    hours = s // 3600
    s = s - hours*3600
    mu = s // 60
    ss = s - mu*60
    print('소요시간 = ', hours, '시간 ', mu, '분 ', ss, '초')


buy_zoo = []
# 매수할 주식 목록
buy_zoo_price = []
# 매수할 주식 적정가 대비%
buy_zoo_low_price = []
# 매수가격
buy_zoo_good_price = []
# 적정가
buy_zoo_high_price = []
# 매도가격
buy_zoo_code = []
# 종목코드
DA_row = []
DQ_row = []
BQ_row = []
ebi_row = []
# 매출액 등 위치에 따른 셀 위치 찾기 위한 리스트

code_data = code_data['종목코드']
# ['종목코드'] 요것만 하면 판다스에 종목코드가 데이터프레임으로 저장될까? ㅇㅇ 가능하네
code_data = code_data.apply(make_code)

main_data = SoupStrainer('div', {'class': 'um_table'})
fics_filter = SoupStrainer('span', {'class': "stxt stxt2"})

fnguide_url = 'http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=D&NewMenuID=Y&stkGb=701'
webpage = requests.get(fnguide_url)
web_data = BeautifulSoup(webpage.content, 'lxml', parse_only=main_data)
ifrs_D_A = web_data.find('div', {'class': 'um_table', 'id': 'highlight_D_Y'})
# ifrs(연결-연간) 데이터 불러오기
ifrs_D_Q = web_data.find('div', {'class': 'um_table', 'id': 'highlight_D_Q'})
# ifrs(연결-분기) 데이터 불러오기
ifrs_B_Q = web_data.find('div', {'class': 'um_table', 'id': 'highlight_B_Q'})
# ifrs(개별-분기) 데이터 불러오기
ev_ebita = web_data.find('div', {'class': 'um_table', 'id': 'svdMainGrid10D'})
#ev/ebita 있는 표 불러오기
web_data.decompose()
webpage.close()
webpage = None
fnguide_url = None
ifrs_DA_row = ifrs_D_A.find_all('th', {'scope': 'row'})
# ifrs(연결-연간) 열 데이터
ifrs_DQ_row = ifrs_D_Q.find_all('th', {'scope': 'row'})
# ifrs(연결-분기) 열 데이터
ifrs_BQ_row = ifrs_B_Q.find_all('th', {'scope': 'row'})
ev_ebita_row = ev_ebita.find_all('th', {'scope': 'row'})
for i in ifrs_DA_row:
    DA_row.append(i.text)
for i in ifrs_DQ_row:
    DQ_row.append(i.text)
for i in ifrs_BQ_row:
    BQ_row.append(i.text)
for i in ev_ebita_row:
    ebi_row.append(i.text)
where_ebi = ebi_row.index('EV/EBITDA')
ifrs_D_A = None
ifrs_D_Q = None
ifrs_B_Q = None
ev_ebita = None
ifrs_DA_row = None
ifrs_DQ_row = None
ifrs_BQ_row = None
ev_ebita_row = None
# 매출액 등 이름에 따라 셀 위치 찾기위한 인덱스 리스트


# len(code_data)
for a in range(0, len(code_data)):
    code_num = code_data[a]
    if(code_num[0] != '9'):
        #국내상장 해외기업 제외
        fnguide_url = "http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A" + code_num + "&cID=&MenuYn=Y&ReportGB=D&NewMenuID=Y&stkGb=701"
        # 종목코드 표에 있는 모든 법인 접속
        table = pd.read_html(fnguide_url)

















        webpage = requests.get(fnguide_url)
        web_data = BeautifulSoup(webpage.content, 'lxml', parse_only=main_data)
        fics_data = BeautifulSoup(webpage.content, 'lxml', parse_only=fics_filter)
        # 웹페이지 정보 가져오기
        try:
            corp_name = BeautifulSoup(webpage.content, 'lxml', parse_only=SoupStrainer('h1', {'id': 'giName'})).find_all('h1', {'id': 'giName'})[0].text
            # 주식이름
            price_and_num = web_data.find('div', {'class': 'um_table', 'id': 'svdMainGrid1'})
            # 시세현황표
            price_and_num_data = price_and_num.find_all('td')
            # 시세현황표 안에 있는 숫자들 전부 가져오기
            my_zoo = web_data.find('div', {'class': 'um_table', 'id': 'svdMainGrid5'})
            # 주주구분 현황
            my_zoo_data = my_zoo.find_all('td')
            # 주주구분 현황 안에 있는 숫자들 전부 가져오기
            ifrs_D_A = web_data.find('div', {'class': 'um_table', 'id': 'highlight_D_Y'})
            # ifrs(연결-연간) 데이터 불러오기
            ifrs_D_Q = web_data.find('div', {'class': 'um_table', 'id': 'highlight_D_Q'})
            # ifrs(연결-분기) 데이터 불러오기
            ifrs_B_Q = web_data.find('div', {'class': 'um_table', 'id': 'highlight_B_Q'})
            # ifrs(개별-분기) 데이터 불러오기
            ev_ebita = web_data.find('div', {'class': 'um_table', 'id': 'svdMainGrid10D'})
            ev_ebita_data = ev_ebita.find_all('td')[where_ebi * 3].text
            # ev/ebita 데이터 있는 표 불러오기, 18번째 숫자
            web_data.decompose()
            # 분석 끝낸 웹페이지는 삭제
            webpage.close()
            webpage = None
            my_zoo = None
            price_and_num = None
            #웹페이지 삭제

            ifrs_DA = ifrs_D_A.find_all('td')
            ifrs_DQ = ifrs_D_Q.find_all('td')
            ifrs_BQ = ifrs_B_Q.find_all('td')
            # ifrs 숫자들 전부 가져오기
            ifrs_DA_date = ifrs_D_A.find_all('th', {'scope': 'col'})
            # ifrs(연결-연간) 날짜 데이터
            ifrs_DQ_date = ifrs_D_Q.find_all('th', {'scope': 'col'})
            # ifrs(연결-분기) 날짜 데이터
            ifrs_BQ_date = ifrs_B_Q.find_all('th', {'scope': 'col'})
            ifrs_D_A = None
            ifrs_D_Q = None
            ifrs_B_Q = None
            #불필요한 변수들 삭제

            # ifrs(개별-분기) 날짜 데이터
            if(ifrs_DA[138].text != ifrs_DA[199].text and fics_data.find_all('span', {'class': "stxt stxt2"})[0].text != 'FICS  창업투자 및 종금'):
                # (마지막 사업보고서 기준) 2년전 roe 정보 없으면 취급하지 않음
                # 199는 연결-연간 표에서 가장 먼 컨센서스 배당수익률. 항상 빈칸이리 카더라
                # 기업인수목적회사 필터링
                fics_data.decompose()   
                for i in ifrs_DA_row:

                wb = xw.Book('G:\Hyuk_Rim_v6.xlsx')                                          # 엑셀 이름 업데이트!!
                wb_result = wb.sheets['Result']
                wb_data = wb.sheets['Data']
                wb_result.range('D11').value = bbb_data
                # 요구수익률 넣기
                wb_data.range('B4').value = corp_name
                # 주식이름 넣기
                wb_data.range('i5').value = price_and_num_data[0].text
                # 주식 현재가 넣기
                wb_data.range('i7').value = price_and_num_data[10].text
                # 발행주식수 넣기
                wb_data.range('i8').value = my_zoo_data[17].text
                # 자기주식수 넣기
                wb_data.range('b47').value = ev_ebita_data
                # ev/ebita 넣기
                wb_data.range('f45').value = ifrs_DA[180].text
                # 최근 pbr 넣기

                line = 0 # ifrs 표 안에 데이터 순번. 표 안의 모든 데이터를 추출하기 위함
                row = 27 # 엑셀 ifrs(연결-연간)표 첫번째 셀의 행 번호
                for b in range(0, 12):
                    # range 12는 ifrs 표의 행(매출액~자본금) 개수(세로 길이)
                    samsung = 66                                                      # ifrs(연결-연간) 숫자 데이터(매출액~자본금) 엑셀에 넣기
                    # samsung -> chr(66)이 대문자 알파벳 B를 뜻함, ifrs 표의 첫번째 셀 열 번호
                    for c in range(0, 8):
                        # range 8은 ifrs 표의 열 개수(가로 길이)
                        rrow = str(row)
                        col = chr(samsung)
                        cell = col + rrow
                        wb_data.range(cell).value = ifrs_DA[line].text
                        line += 1
                        samsung += 1
                    row += 1

                line = line + 4 * 8
                # ifrs 표에서 중간 내용 건너뛰고 roa 정보 순번
                for bbbb in range(0, 5):
                    samsung = 66                                                      # ifrs(연결-연간) 숫자 데이터(ROA~DPS) 엑셀에 넣기
                    for c in range(0, 8):
                        rrow = str(row)
                        col = chr(samsung)
                        cell = col + rrow
                        wb_data.range(cell).value = ifrs_DA[line].text
                        line += 1
                        samsung += 1
                    row += 1

                line = line + 3 * 8
                samsung = 66
                for cc in range(0, 8):
                        rrow = str(row)                                               # frs(연결-연간) 숫자 데이터(배당수익률) 엑셀에 넣기
                        col = chr(samsung)
                        cell = col + rrow
                        wb_data.range(cell).value = ifrs_DA[line].text
                        line += 1
                        samsung += 1


                for d in range(0, 5):                                                 # ifrs(연결-연간) 날짜 데이터 엑셀에 넣기
                    # range 5 -> 날짜 넣는칸의 총 개수
                    e = d + 66
                    # 날짜 넣는칸의 열 위치(알파벳)
                    g = chr(e) + str(25)
                    # str(25) -> 날짜 넣는칸의 행 위치
                    wb_data.range(g).value = ifrs_DA_date[d + 2].text

                line = 0
                row = 51
                for b in range(0, 12):                                                # ifrs(연결-분기) 숫자 데이터 엑셀에 넣기
                    samsung = 66
                    for c in range(0, 8):
                        rrow = str(row)
                        col = chr(samsung)
                        cell = col + rrow
                        wb_data.range(cell).value = ifrs_DQ[line].text
                        line += 1
                        samsung += 1
                    row += 1


                for d in range(0, 5):                                                 # ifrs(연결-분기) 날짜 데이터 엑셀에 넣기
                    e = d + 66
                    g = chr(e) + str(50)
                    wb_data.range(g).value = ifrs_DQ_date[d + 2].text

                line = 0
                row = 69
                for b in range(0, 8):                                                 # ifrs(개별-분기) 숫자 데이터 엑셀에 넣기
                    samsung = 66
                    for c in range(0, 8):
                        rrow = str(row)
                        col = chr(samsung)
                        cell = col + rrow
                        wb_data.range(cell).value = ifrs_BQ[line].text
                        line += 1
                        samsung += 1
                    row += 1

                for d in range(0, 5):                                                 # ifrs(개별-분기) 날짜 데이터 엑셀에 넣기
                    e = d + 66
                    g = chr(e) + str(68)
                    wb_data.range(g).value = ifrs_BQ_date[d + 2].text

                if wb_result.range('C26').value >= wb_result.range('C23').value:
                    # 매수가격 >= 현재가격
                    if wb_result.range('I31').value >= 0.01:
                        # 배당수익률 1% 이상
                        if wb_result.range('F27').value != '역배열':
                            # roe > 요구수익률
                            buy_zoo_code.append(code_num)
                            buy_zoo.append(wb_data.range('B4').value)
                            buy_zoo_price.append(wb_result.range('D24').value)
                            buy_zoo_low_price.append(wb_result.range('C26').value)
                            buy_zoo_good_price.append(wb_result.range('C24').value)
                            buy_zoo_high_price.append(wb_result.range('C25').value)
                print(a)
        except AttributeError:
            print('에러 = ', a)
        except IndexError:
            print('index 에러 = ', a)
        gc.collect()
        #불필요한 데이터 전부 삭제
code_data = None
wb2 = xw.Book('G:\RESULT.xlsx')
# 결과 기록 엑셀 파일 불러오기
for z in range(0, len(buy_zoo)):
    # b열에 매수할 종목들 하나씩 기록
    wb2.sheets[0].range('B' + str(z + 3)).value = buy_zoo_code[z]
    wb2.sheets[0].range('C' + str(z + 3)).value = buy_zoo[z]
    wb2.sheets[0].range('D' + str(z + 3)).value = buy_zoo_price[z]
    wb2.sheets[0].range('E' + str(z + 3)).value = buy_zoo_low_price[z]
    wb2.sheets[0].range('F' + str(z + 3)).value = buy_zoo_good_price[z]
    wb2.sheets[0].range('G' + str(z + 3)).value = buy_zoo_high_price[z]

# wb2.save('G:\RESULT.xlsx')
hms(time.time() - start)
# 엑셀 파일 저장, 경로까지 적으면 원하는 위치 저장 가능. 디폴트 경로는 파이썬 코드가 있는 곳
# G:\Hyuk_Rim.xlsx
