import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
from slacker import Slacker
import time, calendar
from bs4 import BeautifulSoup
from urllib.request import urlopen
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import requests
import json
from pandas import json_normalize

slack = Slacker('****')
def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    slack.chat.post_message('#stock', strbuf)

def printlog(message, *args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)
 
# 크레온 플러스 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  

def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        return False
 
    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        printlog('check_creon_system() : connect to server -> FAILED')
        return False
 
    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True

def get_current_price(code):
    """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # 현재가
    item['ask'] =  cpStock.GetHeaderValue(16)        # 매수호가
    item['bid'] =  cpStock.GetHeaderValue(17)        # 매도호가    
    return item['cur_price'], item['ask'], item['bid']

def get_ohlc(code, qty):
    """인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다."""
    cpOhlc.SetInputValue(0, code)           # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))        # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)             # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5]) # 0:날짜, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))        # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))        # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)   # 3:수신개수
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count): 
        index.append(cpOhlc.GetDataValue(0, i)) 
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
            cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)]) 
    df = pd.DataFrame(rows, columns=columns, index=index) 
    return df

def get_stock_balance(code):
    """인자로 받은 종목의 종목명과 수량을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    cpBalance.BlockRequest()     
    
    if code == 'BUY_COUNT':
        return cpBalance.GetHeaderValue(7)
    if code == 'GAIN':
        return cpBalance.GetHeaderValue(4), cpBalance.GetHeaderValue(3)
    if code == 'ALL':
        #dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        #dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
        if cpBalance.GetHeaderValue(3) != 0:
            dbgout('(%): ' + str(cpBalance.GetHeaderValue(4)/cpBalance.GetHeaderValue(3)))
        dbgout('종목수: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        if code == 'ALL':
            dbgout(str(i+1) + ' ' + stock_code + '(' + stock_name + ')' 
                + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name, 
                'qty': stock_qty})
        if stock_code == code:  
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0

def get_current_cash():
    """증거금 100% 주문 가능 금액을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)              # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])      # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest() 
    return cpCash.GetHeaderValue(9) # 증거금 100% 주문 가능 금액

# 변동성 돌파 전략: 매수 목표 확인
# 딥러닝으로 목표가 설정할 경우 이 함수를 변형하여 사용
def get_target_price(code):
    """매수 목표가를 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open 
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]                                      
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]

        # target_price: 지난날의 고가와 저가의 차이를 계산해서 K(0.5)만큼 곱한 값을 오늘 시가에 더해 목표값으로 지정
        target_price = today_open + (lastday_high - lastday_low) * 0.5
        return target_price
    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None
    
def get_movingaverage(code, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()         
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        dbgout('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None    

#★★★★★ 변동성 돌파 전략
def buy_etf(code):
    """인자로 받은 종목을 최유리 지정가 FOK 조건으로 매수한다."""
    try:
        global bought_list      # 함수 내에서 값 변경을 하기 위해 global로 지정
        if code in bought_list: # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
            return False
        
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code) # 현재가격

        '''
        주식을 잘 하는 사람들은 한가지 전략만 사용하지 않는다.
        안전한 운용을 위해 여기서 사용하는 3가지 조건은 아래와 같음
        (1) 변동성 돌파 전략
        (2) 이동평균선 5일
        (3) 이동평균선 10일
        '''

        # (1) 매수 목표가(변동성 돌파 전략)
        target_price = get_target_price(code)    
        
        # (2,3) 이동 평균선: 주가의 이동평균을 구해서 평균값을 이은 선
        '''
        현재 주가가 이동평균선 보다 낮을때는 하락하는 추세를 보이고
        이동평균선보다 위에 있을때는 상승하는 추세를 보임
        100%는 아니지만 어느정도 추세는 맞기 때문에 이를 주식차트를 볼때 보조지표로 많이 활용
        '''
        ma5_price = get_movingaverage(code, 5)   # 5일 이동평균가
        ma10_price = get_movingaverage(code, 10) # 10일 이동평균가
        
        buy_qty = 0        # 매수할 수량 초기화
        if ask_price > 0:  # 매수호가가 존재하면   
            buy_qty = buy_amount // ask_price  
        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회    
        
        # 4가지 조건을 모두 만족하는지 확인 (마지막 조건은 55000원 미만의 ETF만 구매하는 조건)
        if current_price > target_price and current_price > ma5_price \
            and current_price > ma10_price and current_price < 55000 :  
            
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
            accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                
            
            '''
            1. 주문 호가: 주문할 때 주문 가격을 결정하는 방식 (매도는 반대)
               (1) 최유리 방식 - 당장 가장 유리하게 매매할 수 있는 가격
               (2) 최우선 방식 - 우선 대기하는 가격
            
            2. 주문 조건
               (1) IOC 방식 - 체결 후 남은 수량 취소
               (2) FOK 방식 - 전량 체결되지 않으면 주문 자체를 취소
            '''
            
            #★★★★★ 최유리 FOK 매수 주문 설정 - 가용 자금이 적은 경우 추천
            cpOrder.SetInputValue(0, "2")        # 2: 매수
            cpOrder.SetInputValue(1, acc)        # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
            cpOrder.SetInputValue(3, code)       # 종목코드
            cpOrder.SetInputValue(4, buy_qty)    # 매수할 수량
            cpOrder.SetInputValue(7, "2")        # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")       # 주문호가 1:보통, 3:시장가
                                                 # 5:조건부, 12:최유리, 13:최우선 
            # 매수 주문 요청
            ret = cpOrder.BlockRequest() 
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
                time.sleep(remain_time/1000) 
                return False
            time.sleep(2)
            stock_name, bought_qty = get_stock_balance(code)
            
            if bought_qty > 0: # 구매했으면(구매수량이 0보다 크면)
                bought_list.append(code)
                # 현재 어떤걸 얼마나 샀는지 slack메시지를 보내줌
                dbgout_msg = str(stock_name) + '(' + str(code) + ') ' + str(buy_qty) + 'EA : ' + str(current_price) + ' meets the buy condition!`'
                dbgout(dbgout_msg)
                dbgout("`buy_etf("+ str(stock_name) + ' : ' + str(code) + 
                    ") -> " + str(bought_qty) + "EA bought!" + "`")
                dbgout("---------------------------")
    except Exception as ex:
        dbgout("`buy_etf("+ str(code) + ") -> exception! " + str(ex) + "`")
        dbgout("---------------------------")

def sell_all():
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]       # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션   
        while True:    
            stocks = get_stock_balance('ALL')
            dbgout("---------------------------") 
            total_qty = 0 
            for s in stocks:
                total_qty += s['qty'] 
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:                  
                    cpOrder.SetInputValue(0, "1")         # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)         # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])   # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])    # 매도수량
                    cpOrder.SetInputValue(7, "1")   # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선 
                    # 최유리 IOC 매도 주문 요청
                    ret = cpOrder.BlockRequest()

                    dbgout_msg = '최유리 IOC 매도 ' + str(s['code']) + ' ' + str(s['name']) + ' ' + str(s['qty']) + ' -> cpOrder.BlockRequest() -> returned ' + str(ret)
                    dbgout(dbgout_msg)
                    dbgout("---------------------------")

                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        dbgout_msg = '주의: 연속 주문 제한, 대기시간: ' + str(remain_time/1000)
                        dbgout(dbgout_msg)
                        dbgout("---------------------------")
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))

def autoETF():
    """네이버 증권에서 ETF정보를 읽어 거래량이 가장높은 30개의 종목을 추출한다."""
    try:
        url = 'https://finance.naver.com/api/sise/etfItemList.nhn'
        json_data = json.loads(requests.get(url).text)
        df = json_normalize(json_data['result']['etfItemList'])

        # ETF.xlsx 생성
        df = df.sort_values(by=['quant'], ascending=False) # 거래량 기준 내림차순 정렬
        df.to_excel('ETF.xlsx', index=False) 

        # symbol_list.txt 생성
        from openpyxl import load_workbook
        load_wb = load_workbook("ETF.xlsx", data_only=True)
        load_ws = load_wb['Sheet1']
        
        # symbol_list
        f = open("symbol_list.txt", 'w')
        get_cells = load_ws['A2':'A31']
        idx = 1
        for row in get_cells:
            for cell in row:
                if idx != 30:
                    f.write('A' + str(cell.value) + '\n')
                else:
                    f.write('A' + str(cell.value))
                idx += 1
        f.close()

        # symbol_list_itemname
        f = open("symbol_list_itemname.txt", 'w')
        get_cells = load_ws['C2':'C31']
        idx = 1
        for row in get_cells:
            for cell in row:
                if idx != 30:
                    f.write(str(cell.value) + '\n')
                else:
                    f.write(str(cell.value))
                idx += 1
        f.close()
        
    except Exception as ex:
        dbgout("autoETF() -> exception! " + str(ex))

if __name__ == '__main__': 
    try:
        #★★★★★ symbol_list: 자동매매를 원하는 종목
        autoETF()
        symbol_list = open('symbol_list.txt', 'r').read().split('\n')
        dbgout(str(symbol_list))
        bought_list = []     # 금일 매수 완료된 종목 리스트
        buy_count = get_stock_balance('BUY_COUNT')
        target_buy_count = 5 - buy_count # 매수할 종목 수 (symbol_list 중에서 매수할 최대 개수 설정)
        if target_buy_count != 0:
            buy_percent = 0.95 / target_buy_count # 전체 가용 자금에서 각 매수 종목을 몇 퍼센트 살건지
        else :
            buy_percent = 0.95
        
        printlog('check_creon_system() :', check_creon_system())  # 크레온 접속 점검
        stocks = get_stock_balance('ALL')      # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())   # 100% 증거금 주문 가능 금액 조회
        buy_amount = total_cash * buy_percent  # 종목별 주문 금액 계산
        dbgout("---------------------------")
        dbgout_msg = '100% 증거금 주문 가능 금액 :' + str(total_cash)
        dbgout(dbgout_msg)
        dbgout_msg = '금일 목표 종목 수:' + str(target_buy_count)
        dbgout(dbgout_msg)
        dbgout_msg = '종목별 주문 비율 :' + str(buy_percent)
        dbgout(dbgout_msg)
        dbgout_msg = '종목별 주문 금액 :' + str(buy_amount)
        dbgout(dbgout_msg)
        dbgout("---------------------------")

        autoETF_hour = 0
        alarm_minute = 0

        while True:
            #★★★★★ t_start, t_sell, t_exit: 자동매매 시간 설정
            '''
            (1) 주식 시장 정규 시간 09:00~15:30
            (2) LP(유동성공급자) 활동 시간 09:05~15:20
                -> 자동매매 시간 09:05~15:20(종료)
            '''
            t_now = datetime.now()
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0) 
            t_end = t_now.replace(hour=15, minute=20, second=0,microsecond=0)
            today = datetime.today().weekday()

            # 토요일이나 일요일이면 자동 종료(주말은 주식시장 닫힘)
            if today == 5 or today == 6: 
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                sys.exit(0)

            # AM 09:05 ~ PM 03:20 : 매수 및 매도
            if t_start < t_now < t_end :
                for sym in symbol_list:
                    
                    # 목표한 종목 수보다 아직 덜 샀으면
                    if len(bought_list) < target_buy_count:
                        # 변동성 돌파 전략: 종목의 가격이 매수할 타이밍이 맞는지 검사
                        buy_etf(sym)
                        time.sleep(1)
                
                # 약 'xx시:30분'마다 ETF 종목 정보 갱신
                if 30 <= t_now.minute < 55: 
                    if t_now.hour != autoETF_hour:
                        autoETF_hour = t_now.hour
                        alarm_minute = t_now.minute
                        get_stock_balance('ALL')
                        autoETF() # 현재 ETF 종목 정보 불러오기
                        symbol_list = open('symbol_list.txt', 'r').read().split('\n')
                        dbgout("---------------------------")
                        time.sleep(1)
                
                # 5분마다 체크: 평가손익이 0.001(0.1%) 이상일 경우 일괄매도
                if (t_now.minute % 5 == 0):
                    if alarm_minute != t_now.minute:
                        alarm_minute = t_now.minute
                        G_gain, G_total = get_stock_balance('GAIN')
                        dbgout('평가손익: ' + str(G_gain))
                        if G_total != 0:
                            if (G_gain/G_total) >= 0.001:
                                dbgout('`(G_gain/G_total) > 0.001`')
                                if sell_all() == True:
                                    bought_list = []
                                    target_buy_count = 5
                                    total_cash = int(get_current_cash())
                                    buy_amount = total_cash * 0.19
                                    dbgout('`sell_all() returned True!`')
                        dbgout("---------------------------")
                        time.sleep(3)

            # PM 03:20 ~ :프로그램 종료
            if t_end < t_now :  
                dbgout("---------------------------")
                autoETF() # 현재 ETF 종목 정보 저장
                get_stock_balance('ALL')
                dbgout("---------------------------")
                dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(1)

    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')
