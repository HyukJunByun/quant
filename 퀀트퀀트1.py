import pandas as pd
import requests
from bs4 import BeautifulSoup, SoupStrainer
import xlwings as xw
import time
import gc
import math
import numpy as np

bbb_web_data = pd.read_html('https://www.kisrating.com/ratingsStatistics/statics_spread.do', match='국고채', header=0, index_col=0)[0]
# 회사채 수익률 웹페이지 가져오기
bbb_data = bbb_web_data['5년']['BBB-']
# 회사채 BBB- 5년 수익률 가져오기 = 요구수익률
bbb_web_data = None
# 정보 빼고 나면 웹페이지는 삭제(스피드업)


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
pbr_list = []
boon_per_list = []
psr_list = []
por_list = []
ev_ebita_list = []
pcr_list = []
pfcr_list = []
gpa_list = []
profit_list = [] # 영업이익 & 당기순이익 YOY, QOQ
bay_trend_list = [] # 배당성향
asset_growth_list  = [] # 자산증가율(최근 연도)


code_name = code_data['회사명']
code_data = code_data['종목코드']
# ['종목코드'] 요것만 하면 판다스에 종목코드가 데이터프레임으로 저장될까? ㅇㅇ 가능하네
code_data = code_data.apply(make_code)
column = [0, 1, 2, 3, 4, 5, 6, 7]
np.seterr(divide='raise')

# len(code_data)
for a in range(0, len(code_data)):
    code_num = code_data[a]
    # print(code_name[a])
    if(code_num[0] != '9'):
        # 국내상장 해외기업 제외
        fnguide_url = "http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A" + code_num + "&cID=&MenuYn=Y&ReportGB=D&NewMenuID=Y&stkGb=701"
        # 종목코드 표에 있는 모든 법인 접속
        webpage = requests.get(fnguide_url)
        fics_data = BeautifulSoup(webpage.content, 'lxml', parse_only=SoupStrainer('span', {'class' : "stxt stxt2"}))
        webpage.close()
        webpage = None
        no_fics = fics_data.find_all('span', {'class': "stxt stxt2"})[0].text
        if(no_fics != 'FICS  창업투자 및 종금'):
            # 기업인수목적회사 필터링
            fics_data.decompose()
            try:
                ifrs_table = pd.read_html(fnguide_url, match='Financial Highlight', flavor='lxml', header=0, index_col=0, attrs={'class': 'us_table_ty1 h_fix zigbg_no'})
                # ifrs 표 전부 가져옴 (attrs는 무조건 table class만 적어)
                # 연결-전체, 연결-연간, 연결-분기, 개별-전체 순...
                # 열 인덱스=(연간 = Annual, Annual.1~, 분기 = Net Quarter, Net Quarter.1~)
                DA = ifrs_table[1]
                if(len(DA.columns) == 8 and list(DA.index.values)[1] == '매출액'):  # 금융주 및 정보부족 회사 필터링
                    DA.columns = column
                    DQ = ifrs_table[2]
                    BQ = ifrs_table[5]
                    ifrs_table = None
                    if(DA.loc['ROE', [2, 3, 4]].isna().sum() == 0 and DA.loc['배당수익률', [3, 4]].isna().sum() == 0
                    and len(DQ.columns) == 8 and len(BQ.columns) == 8):
                        # 최근 3년 roe 있음 + 최근 2년 배당 함
                        DQ.columns = column
                        BQ.columns = column
                        for i in column:
                            DA[i] = DA[i].apply(pd.to_numeric, errors = 'ignore')
                            DQ[i] = DQ[i].apply(pd.to_numeric, errors = 'ignore')
                            BQ[i] = BQ[i].apply(pd.to_numeric, errors = 'ignore')
                        price_and_num = pd.read_html(fnguide_url, match='시세현황', flavor='lxml', index_col=0, attrs={'class': 'us_table_ty1 table-hb thbg_g h_fix zigbg_no'})[0]
                        # 시세현황 표
                        my_zoo_table = pd.read_html(fnguide_url, match='주주현황', flavor='lxml', header=0, index_col=0, attrs={'class': 'us_table_ty1 h_fix zigbg_no notres'})[1]
                        # 주주구분현황 표
                        siga = pd.to_numeric(price_and_num[1]['시가총액(보통주,억원)'], errors = 'ignore')
                        # 시가총액
                        pbr = siga / DQ[4]['지배주주지분']
                        boon_per = siga / (DQ[4]['지배주주순이익'] * 4)
                        psr = siga / (DQ[4]['매출액'] + DQ[3]['매출액'] + DQ[2]['매출액'] + DQ[1]['매출액'])
                        por = siga / (DQ[4]['영업이익(발표기준)'] + DQ[3]['영업이익(발표기준)'] + DQ[2]['영업이익(발표기준)'] + DQ[1]['영업이익(발표기준)'])
                        # 각종 가치지표(with 현재 시총 & 최근 4개분기 실적)
                        asset_growth = 100 * (DA[4]['자산총계'] - DA[3]['자산총계']) / DA[3]['자산총계']
                        # 자산증가율(1년)
                        if(pbr > 0.2 and boon_per > 2 and psr > 0.1 and por > 2 and asset_growth > 0):
                            # 1차 가치지표 필터링
                            price = pd.to_numeric(price_and_num[1]['종가/ 전일대비'], errors = 'ignore')
                            # 주식가격
                            all_zoo = pd.to_numeric(price_and_num[1]['발행주식수(보통주/ 우선주)'], errors = 'ignore')
                            # 발행주식수
                            my_zoo = pd.to_numeric(my_zoo_table.loc['자기주식 (자사주+자사주신탁)', '보통주'], errors = 'ignore')
                            # 자기주식수
                            my_zoo_table = None
                            price_and_num = None
                            fnguide_url = 'http://comp.fnguide.com/SVO2/ASP/SVD_Finance.asp?pGB=1&gicode=A' + code_num + '&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701'
                            # 재무제표
                            recent_DA = DA.iloc[0, 4]  # 최신 사업보고서 (연도/12월) str
                            recent_DQ = DQ.iloc[0, 4]
                            rrecent_DQ = DQ.iloc[0, 3]
                            rrrecent_DQ = DQ.iloc[0, 2]
                            rrrrecent_DQ = DQ.iloc[0, 1]  # 최근 4개분기 연도/월
                            earn = pd.read_html(fnguide_url, match='포괄손익계산서', flavor='lxml', header=0, index_col=0, attrs={'class': 'us_table_ty1 h_fix zigbg_no'})[1]
                            earn_all = (pd.to_numeric(earn.loc['매출총이익', recent_DQ], errors = 'ignore') +
                                        pd.to_numeric(earn.loc['매출총이익', rrecent_DQ], errors = 'ignore') +
                                        pd.to_numeric(earn.loc['매출총이익', rrrecent_DQ], errors = 'ignore') +
                                        pd.to_numeric(earn.loc['매출총이익', rrrrecent_DQ], errors = 'ignore'))
                            # 매출총이익
                            earn = None
                            gpa = earn_all / DQ[4]['자산총계']
                            # GP/A
                            cash_flow_table = pd.read_html(fnguide_url, match='현금흐름표', flavor='lxml', header=0, index_col=0, attrs={'class': 'us_table_ty1 h_fix zigbg_no'})[1]
                            cash_flow = (pd.to_numeric(cash_flow_table.loc['영업활동으로인한현금흐름', recent_DQ], errors = 'ignore') +
                                         pd.to_numeric(cash_flow_table.loc['영업활동으로인한현금흐름', rrecent_DQ], errors = 'ignore') +
                                         pd.to_numeric(cash_flow_table.loc['영업활동으로인한현금흐름', rrrecent_DQ], errors = 'ignore') +
                                         pd.to_numeric(cash_flow_table.loc['영업활동으로인한현금흐름', rrrrecent_DQ], errors = 'ignore'))
                            # 영업현금흐름
                            pcr = siga / cash_flow
                            # PCR
                            naver = (DQ[4]['영업이익(발표기준)'] + DQ[3]['영업이익(발표기준)'] + DQ[2]['영업이익(발표기준)'] + DQ[1]['영업이익(발표기준)'])
                            # 영업이익
                            kakao = (DQ[4]['지배주주순이익'] + DQ[3]['지배주주순이익'] + DQ[2]['지배주주순이익'] + DQ[1]['지배주주순이익'])
                            # 당기순이익(지배)
                            cash_flow_table = None
                            if(pcr > 1 and  naver > 0 and cash_flow > 0 and kakao > 0):
                                # pcr, 신f-score(영업이익, 영업현금흐름, 유상증자), 당기순이익 필터링
                                fnguide_url = 'http://comp.fnguide.com/SVO2/ASP/SVD_FinanceRatio.asp?pGB=1&gicode=A' + code_num + '&cID=&MenuYn=Y&ReportGB=&NewMenuID=104&stkGb=701'
                                # 재무비율
                                mon_ratio = pd.read_html(fnguide_url, match='재무비율', flavor='lxml', header=0, index_col=0, attrs={'class': 'us_table_ty1 h_fix zigbg_no'})[0]
                                owe = pd.to_numeric(mon_ratio.loc['순차입금비율계산에 참여한 계정 펼치기', recent_DA], errors = 'ignore')
                                # 순차입금비율
                                mon_ratio = None
                                if(math.isnan(owe) == True or owe < 200):
                                    # 순차입금비율 < 200%
                                    fnguide_url = 'http://comp.fnguide.com/SVO2/ASP/SVD_Invest.asp?pGB=1&gicode=A' + code_num + '&cID=&MenuYn=Y&ReportGB=&NewMenuID=105&stkGb=701'
                                    # 투자지표
                                    invest_idea = pd.read_html(fnguide_url, match='기업가치 지표', flavor='lxml', header=0, index_col=0, attrs={'class': 'us_table_ty1 h_fix zigbg_no'})[0]
                                    bay_trend = pd.to_numeric(invest_idea.loc['배당성향(현금)(%)계산에 참여한 계정 펼치기', recent_DA], errors = 'ignore')
                                    # 배당성향
                                    ev_ebita = pd.to_numeric(invest_idea.loc['EV/EBITDA계산에 참여한 계정 펼치기', recent_DA], errors = 'ignore')
                                    # EV/EBITA
                                    fcff = pd.to_numeric(invest_idea.loc['FCFF', recent_DA], errors = 'ignore')
                                    pfcr = siga / fcff
                                    invest_idea = None
                                    # PFCR
                                    if(pfcr > 1):
                                        wb = xw.Book('G:\Hyuk_Rim_v7.xlsx')                                         # 엑셀 이름 업데이트!!
                                        wb_result = wb.sheets['Result']
                                        wb_data = wb.sheets['Data']
                                        wb_result.range('D11').value = bbb_data
                                        # 요구수익률 넣기
                                        wb_data.range('B4').value = code_name[a]
                                        # 주식이름 넣기
                                        wb_data.range('i5').value = price
                                        # 주식 현재가 넣기
                                        wb_data.range('i7').value = all_zoo
                                        # 발행주식수 넣기
                                        wb_data.range('i8').value = my_zoo
                                        # 자기주식수 넣기
                                        cell_col = 66 # 현재 셀의 열, chr(66)이 대문자 알파벳 B를 뜻함
                                        for i in range(0, 5):
                                            wb_data.range(chr(cell_col) + '25').value = DA.iloc[0, i]
                                            wb_data.range(chr(cell_col) + '50').value = DQ.iloc[0, i]
                                            wb_data.range(chr(cell_col) + '68').value = BQ.iloc[0, i]
                                            # 날짜 데이터
                                            cell_col += 1
                                        k = 0
                                        table_content = ('매출액', '영업이익', '영업이익(발표기준)', '당기순이익', '지배주주순이익', '비지배주주순이익',
                                                         '자산총계', '부채총계', '자본총계', '지배주주지분', '비지배주주지분', '자본금', 'ROA', 'ROE')
                                        for i in range(0, 14):
                                            # 연결-연간 데이터 엑셀에 넣기
                                            cell_col = 66
                                            for t in range(0, 8):
                                                wb_data.range(chr(cell_col) + str(i + 27)).value = DA[t][table_content[i]]
                                                cell_col += 1
                                        for i in range(0, 12):
                                            # 연결-분기 데이터 엑셀에 넣기
                                            cell_col = 66
                                            for t in range(0, 8):
                                                wb_data.range(chr(cell_col) + str(i + 51)).value = DQ[t][table_content[i]]
                                                cell_col += 1
                                        for i in range(0, 4):
                                            # 개별-분기 데이터 엑셀에 넣기(매출액~당기순이익)
                                            cell_col = 66
                                            for t in range(0, 8):
                                                wb_data.range(chr(cell_col) + str(i + 69)).value = BQ[t][table_content[i]]
                                                cell_col += 1
                                        for i in range(0, 3):
                                            # 개별-분기 데이터 엑셀에 넣기(자산총계~자본총계)
                                            cell_col = 66
                                            for t in range(0, 8):
                                                wb_data.range(chr(cell_col) + str(i + 73)).value = BQ[t][table_content[i + 6]]
                                                cell_col += 1
                                        for t in range(0, 8):
                                            # 개별-분기 데이터 엑셀에 넣기(자본금)
                                            wb_data.range(chr(t + 66) + str(76)).value = BQ[t][table_content[11]]
                                        if wb_result.range('C26').value >= wb_result.range('C23').value:
                                            # 매수가격 >= 현재가격
                                            if wb_result.range('F27').value != '역배열':
                                                # roe > 요구수익률
                                                buy_zoo_code.append(code_num)
                                                buy_zoo.append(code_name[a])
                                                buy_zoo_price.append(wb_result.range('D24').value)
                                                buy_zoo_low_price.append(wb_result.range('C26').value)
                                                buy_zoo_good_price.append(wb_result.range('C24').value)
                                                buy_zoo_high_price.append(wb_result.range('C25').value)
                                                pbr_list.append(pbr)
                                                boon_per_list.append(boon_per)
                                                psr_list.append(psr)
                                                por_list.append(por)
                                                ev_ebita_list.append(ev_ebita)
                                                pcr_list.append(pcr)
                                                pfcr_list.append(pfcr)
                                                gpa_list.append(gpa)
                                                bay_trend_list.append(bay_trend)
                                                asset_growth_list.append(asset_growth)
                                                profit_lily = []
                                                for i in ['영업이익(발표기준)', '지배주주순이익']:
                                                    # 영업이익 & 당기순이익(지배) YOY QOQ '역수' -> 낮을수록 좋음
                                                    profit_lily.append(DA[3][i] / ((DA[4][i] - DA[3][i]) * 100))
                                                    profit_lily.append(DQ[3][i] / ((DQ[4][i] - DQ[3][i]) * 100))
                                                profit_list.append(profit_lily)
                                                DA = None
                                                DQ = None
                                                BQ = None
            except ValueError as v:
                print(a, '->', v)
            except FloatingPointError:
                print('분모가 0인 에러 = ', a)
    print(a + 1, '  /  ', len(code_data))
    gc.collect()
    # 불필요한 데이터 전부 삭제
        
code_data = None
wb2 = xw.Book('G:\RESULT.xlsx')
# 결과 기록 엑셀 파일 불러오기
for z in range(0, len(buy_zoo)):
    # b열에 매수할 종목들 하나씩 기록
    wb2.sheets[0].range('C' + str(z + 3)).value = buy_zoo_code[z]
    wb2.sheets[1].range('C' + str(z + 6)).value = buy_zoo[z]
    wb2.sheets[0].range('E' + str(z + 3)).value = buy_zoo_price[z]
    wb2.sheets[0].range('F' + str(z + 3)).value = buy_zoo_low_price[z]
    wb2.sheets[0].range('G' + str(z + 3)).value = buy_zoo_good_price[z]
    wb2.sheets[0].range('H' + str(z + 3)).value = buy_zoo_high_price[z]
    wb2.sheets[1].range('D' + str(z + 6)).value = pbr_list[z]
    wb2.sheets[1].range('E' + str(z + 6)).value = boon_per_list[z]
    wb2.sheets[1].range('F' + str(z + 6)).value = pcr_list[z]
    wb2.sheets[1].range('G' + str(z + 6)).value = psr_list[z]
    wb2.sheets[1].range('H' + str(z + 6)).value = pfcr_list[z]
    wb2.sheets[1].range('I' + str(z + 6)).value = por_list[z]
    wb2.sheets[1].range('J' + str(z + 6)).value = ev_ebita_list[z]
    wb2.sheets[2].range('D' + str(z + 6)).value = gpa_list[z]
    wb2.sheets[2].range('E' + str(z + 6)).value = asset_growth_list[z]
    wb2.sheets[2].range('F' + str(z + 6)).value = bay_trend_list[z]
    wb2.sheets[3].range('D' + str(z + 6)).value = profit_list[z][0]
    wb2.sheets[3].range('E' + str(z + 6)).value = profit_list[z][1]
    wb2.sheets[3].range('F' + str(z + 6)).value = profit_list[z][2]
    wb2.sheets[3].range('G' + str(z + 6)).value = profit_list[z][3]
# wb2.save('G:\RESULT.xlsx')
hms(time.time() - start)
# 엑셀 파일 저장, 경로까지 적으면 원하는 위치 저장 가능. 디폴트 경로는 파이썬 코드가 있는 곳
# G:\Hyuk_Rim.xlsx
