import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
import time, calendar
import requests

myToken = "xoxb-2237722976004-2255239591280-aZYlOPpIMINDk9BlLalORN6E"

def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(myToken, '#stock', strbuf)

def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
                             headers={"Authorization": "Bearer " + token},
                             data={"channel": channel, "text": text}
                            )


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
        dbgout('check_creon_system() : admin user -> FAILED')
        return False

    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        dbgout('check_creon_system() : connect to server -> FAILED')
        return False

    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        dbgout('check_creon_system() : init trade -> FAILED')
        return False
    return True


def get_current_price(code):
    """인자로 받은 종목의 현재가, 매도호가, 매수호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)  # 현재가
    item['ask'] = cpStock.GetHeaderValue(16)  # 매도호가
    item['bid'] = cpStock.GetHeaderValue(17)  # 매수호가
    return item['cur_price'], item['ask'], item['bid']


def get_ohlc(code, qty):
    """인자로 받은 종목의 OHLC 가격 정보를 qty 기간 동안 만큼의 개수를 반환한다."""
    cpOhlc.SetInputValue(0, code)  # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))  # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)  # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5])  # 0:날짜, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))  # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))  # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)  # 3:수신개수
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
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)  # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)  # 요청 건수(최대 50)
    cpBalance.BlockRequest()
    if code == 'ALL':
        dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)) +
               ' / 결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)) +
               ' / 평가금액: ' + str(cpBalance.GetHeaderValue(3)) +
               ' 원 / 평가손익: ' + str(cpBalance.GetHeaderValue(4)) +
               ' 원 / 종목수: ' + str(cpBalance.GetHeaderValue(7)) +
               ' 개')
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)  # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)  # 수량
        if code == 'ALL':
            dbgout(str(i + 1) + ' ' + stock_code + '(' + stock_name + ')'
                   + ' : ' + str(stock_qty) + 'EA')
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
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)  # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest()
    return cpCash.GetHeaderValue(9)  # 증거금 100% 주문 가능 금액


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
        target_price = today_open + (lastday_high - lastday_low) * 0.2
        return target_price
    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None


def get_movingaverage(code, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 40)
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


def buy_etf(code):
    """인자로 받은 종목을 최유리 지정가 FOK 조건으로 매수한다."""
    try:
        global bought_list  # 함수 내에서 값 변경을 하기 위해 global로 지정
        if code in bought_list:  # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code)
        target_price = get_target_price(code)  # 매수 목표가
        ma5_price = get_movingaverage(code, 5)  # 5일 이동평균가
        ma10_price = get_movingaverage(code, 10)  # 10일 이동평균가
        printlog('코드, 현재, 타겟, 5일, 10일  ->', code, ' / ', str(current_price), ' / ', str(target_price), ' / ',
                 str(ma5_price), ' / ', str(ma10_price))
        dbgout('종목코드: ' + code +
               ' / 현재가: ' + str(current_price) +
               ' / 목표가: ' + str(target_price) +
               ' / 5일 이동평균가: ' + str(ma5_price) +
               ' / 10일 이동평균가: ' + str(ma10_price))
        buy_qty = 0  # 매수할 수량 초기화
        if ask_price > 0:  # 매도호가가 존재하면
            buy_qty = buy_amount // ask_price
        if buy_qty < 1:
            dbgout('종목코드: ' + code + ' / 매수할 수량이 없습니다. buy_qty :  ' + str(buy_qty))
            return False
        if current_price > target_price and current_price > ma5_price \
                and current_price > ma10_price:
            stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회
            dbgout(stock_name + '(' + code + ') ' + str(buy_qty) +
                   'EA : ' + str(current_price) + ' meets the buy condition!`')
            dbgout('stock_name: ' + stock_name + '(' + code + ')'
                   ' / current_price: ' + str(current_price) +
                   ' / target_price: ' + str(target_price) +
                   ' / ma5_price: ' + str(ma5_price) +
                   ' / ma10_price: ' + str(ma10_price))
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
            dbgout('### 계좌번호: ' + acc + ' ###')
            accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체,1:주식,2:선물/옵션
            # 최유리 FOK 매수 주문 설정
            cpOrder.SetInputValue(0, "2")  # 2: 매수
            cpOrder.SetInputValue(1, acc)  # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
            cpOrder.SetInputValue(3, code)  # 종목코드
            cpOrder.SetInputValue(4, buy_qty)  # 매수할 수량
            cpOrder.SetInputValue(7, "0")  # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "03")  # 주문호가 01:보통, 03:시장가, 05:조건부, 12:최유리, 13:최우선
            # 매수 주문 요청
            ret = cpOrder.BlockRequest()
            printlog('시장가 기본 매수 요청 ->', stock_name, code, buy_qty, '->', ret)
            dbgout('시장가 기본 매수 요청 -> 종목명: ' + stock_name +
                   ' / 종목코드: ' + code +
                   ' / 보유한 금액 : ' + str(buy_amount) +
                   ' 원 / 매수 수량: ' + str(buy_qty) +
                   'EA / 코드: ' + str(ret))
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time / 1000)
                time.sleep(remain_time / 1000)
                return False
            time.sleep(2)
            # 매수 주문 후 상태 확인 코드
            rqStatus = cpOrder.GetDibStatus()
            errMsg = cpOrder.GetDibMsg1()
            if rqStatus != 0:
                printlog('주문 실패: ', rqStatus, errMsg)
                dbgout('주문 실패 상태 코드 : ' + str(rqStatus) + ' / 에러메시지 : ' + str(errMsg))
                return False
            stock_name, bought_qty = get_stock_balance(code)
            printlog('get_stock_balance :', stock_name, bought_qty)
            dbgout('code : ' + code + ' / get_stock_balance after 주문 후 : ' + stock_name + ' / ' + str(bought_qty) + 'EA')
            if bought_qty > 0:
                bought_list.append(code)
                dbgout("`buy_etf(" + str(stock_name) + ' : ' + code +
                       ") -> " + str(bought_qty) + "EA bought!" + "`")
    except Exception as ex:
        dbgout("`buy_etf(" + code + ") -> exception! " + str(ex) + "`")


def sell_all():
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
        while True:
            stocks = get_stock_balance('ALL')
            total_qty = 0
            for s in stocks:
                total_qty += s['qty']
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:
                    cpOrder.SetInputValue(0, "1")  # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)  # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])  # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])  # 매도수량
                    cpOrder.SetInputValue(7, "1")  # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선
                    # 최유리 IOC 매도 주문 요청
                    ret = cpOrder.BlockRequest()
                    printlog('최유리 IOC 매도', s['code'], s['name'], s['qty'],
                             '-> cpOrder.BlockRequest() -> returned', ret)
                    dbgout('최유리 IOC 매도 요청 -> 종목명: ' + s['name'] +
                           ' / 종목코드: ' + s['code'] +
                           ' / 매도 수량: ' + str(s['qty']) +
                           'EA / 코드: ' + str(ret))
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        printlog('주의: 연속 주문 제한, 대기시간:', remain_time / 1000)
                    # 매수 주문 후 상태 확인 코드
                    rqStatus = cpOrder.GetDibStatus()
                    errMsg = cpOrder.GetDibMsg1()
                    if rqStatus != 0:
                        printlog('주문 실패: ', rqStatus, errMsg)
                        dbgout('주문 실패 상태 코드 : ' + str(rqStatus) + ' / 에러메시지 : ' + str(errMsg))
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))


if __name__ == '__main__':
    try:
        symbol_list = [
                    'A225040', # 미국S&P500레버리지(합성 H)
                    'A122630', 'A123320',  # 코스피200레버리지
                    'A267770',  # 코스피200선물레버리지
                    'A233740', 'A233160',  # 코스닥150레버리지
                    'A278240',  # 코스닥150선물레버리지
                    'A252670', 'A252710', 'A252420', 'A253230', 'A253160',  # 코스피200곱버스
                    'A143850', 'A219480',  # 미국S&P500선물(H)
                    'A225030',  # 미국S&P500선물인버스
                    'A360750', 'A360200', 'A379780', # 미국S&P500
                    'A379800', # 미국S&P500TR
                    'A133690', 'A367380', 'A368590', # 미국나스닥100
                    'A379810', # 미국나스닥100TR
                    'A381180', # 미국필라델피아반도체
                    'A069500', 'A102110', 'A148020', 'A105190', # 코스피200
                    'A114800', 'A123310', # 코스피200인버스
                    'A229200', 'A232080', # 코스닥150
                    'A251340', 'A250780' # 코스닥150선물인버스
                    ]
        bought_list = []  # 매수 완료된 종목 리스트
        target_buy_count = 5  # 매수할 종목 수
        buy_percent = 0.18  # 각각의 매수 종목을 전체 가용 자금 중 몇 퍼센트를 살 건지 정하는 것
        printlog('check_creon_system() :', check_creon_system())  # 크레온 접속 점검
        stocks = get_stock_balance('ALL')  # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())  # 100% 증거금 주문 가능 금액 조회
        buy_amount = total_cash * buy_percent  # 종목별 주문 금액 계산
        printlog('100% 증거금 주문 가능 금액 : ', total_cash)
        printlog('종목별 주문 비율 : ', buy_percent)
        printlog('종목별 주문 금액 : ', buy_amount)
        dbgout('100% 증거금 주문 가능 금액 : ' + str(total_cash) +
               ' 원 / 매수할 종목 수 : ' + str(target_buy_count) +
               ' 개 / 전체 가용 금액 중 종목별 비중 : ' + str(buy_percent) +
               ' / 종목별 주문 금액 : ' + str(buy_amount) + ' 원')
        printlog('시작 시간 :', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldOut = False

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=0, second=15, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                dbgout('Today is Saturday or Sunday.')
                sys.exit(0)
            if t_9 < t_now < t_start and soldOut == False:
                soldOut = True
                dbgout('장 시작 전 팔지 않은 것이 있으면 모두 팔기')
                sell_all()
            if t_start < t_now < t_sell:  # AM 09:00:15 ~ PM 03:15 : 매수
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy_etf(sym)
                        time.sleep(1)
                if t_now.minute == 30 and 0 <= t_now.second <= 20:
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : 일괄 매도
                if sell_all() == True:
                    dbgout('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
                dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')
