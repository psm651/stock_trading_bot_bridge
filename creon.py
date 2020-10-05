import os
import time

import win32com.client
from pywinauto import application

import constants
import util


class Creon:
    def __init__(self):
        self.obj_CpUtil_CpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        self.obj_CpUtil_CpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        self.obj_CpSysDib_StockChart = win32com.client.Dispatch('CpSysDib.StockChart')
        self.obj_CpTrade_CpTdUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
        self.obj_CpSysDib_MarketEye = win32com.client.Dispatch('CpSysDib.MarketEye')
        self.obj_CpUtil_CpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        self.obj_CpSysDib_CpSvr7238 = win32com.client.Dispatch('CpSysDib.CpSvr7238')

    # HTS연결
    def connect(self, id_, pwd, pwdcert, trycnt=300):
        if not self.connected():
            self.disconnect()
            self.kill_client()
            app = application.Application()
            app.start(
                'C:\\CREON\\STARTER\\coStarter.exe /prj:cp /id:{id} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(
                    id=id_, pwd=pwd, pwdcert=pwdcert
                )
            )

        cnt = 0
        while not self.connected():
            if cnt > trycnt:
                return False
            time.sleep(1)
            cnt += 1
        return True

    # HTS연결 돼있는지 확인
    def connected(self):
        b_connected = self.obj_CpUtil_CpCybos.IsConnect
        if b_connected == 0:
            return False
        return True

    # HTS연결 해제
    def disconnect(self):
        if self.connected():
            self.obj_CpUtil_CpCybos.PlusDisconnect()
            return True
        return False

    # 프로그램 종료
    def kill_client(self):
        os.system('taskkill /IM coStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')
        os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')

    # API요청제한 검사
    def avoid_reqlimitwarning(self):
        remainTime = self.obj_CpUtil_CpCybos.LimitRequestRemainTime
        remainCount = self.obj_CpUtil_CpCybos.GetLimitRemainCount(1)  # 시세 제한
        if remainCount <= 3:
            time.sleep(remainTime / 1000)

    # 종목코드획득 함수
    def get_stockcodes(self, code):
        if code == constants.MARKET_CODE_KOSPI:
            code = 1
        elif code == constants.MARKET_CODE_KOSDAQ:
            code = 2
        res = self.obj_CpUtil_CpCodeMgr.GetStockListByMarket(code)
        return res

    #종목상태 확인
    '''
    controller : 0 정상, 1 주의, 2 경고, 3 위험예고, 4 위험
    supervision : 0 일반종목, 1 관리
    status : 0 정상, 1 거래정지, 2 거래중단
    tip 세 값이 0이 아니면 투자시 유의해야 함.
    '''
    def get_stockstatus(self, code):
        if not code.startswith('A'):
            code = 'A' + code
        return {
            'control': self.obj_CpUtil_CpCodeMgr.GetStockControlKind(code),
            'supervision': self.obj_CpUtil_CpCodeMgr.GetStockSupervisionKind(code),
            'status': self.obj_CpUtil_CpCodeMgr.GetStockStatusKind(code),
        }

    #주식 종목의 다양한 특징
    def get_stockfeatures(self, code):
        if not code.startswith('A'):
            code = 'A' + code
        stock = {
            'name': self.obj_CpUtil_CpCodeMgr.CodeToName(code),
            'marginrate': self.obj_CpUtil_CpCodeMgr.GetStockMarginRate(code),
            'unit': self.obj_CpUtil_CpCodeMgr.GetStockMemeMin(code),
            'industry': self.obj_CpUtil_CpCodeMgr.GetStockIndustryCode(code),
            'market': self.obj_CpUtil_CpCodeMgr.GetStockMarketKind(code),
            'control': self.obj_CpUtil_CpCodeMgr.GetStockControlKind(code),
            'supervision': self.obj_CpUtil_CpCodeMgr.GetStockSupervisionKind(code),
            'status': self.obj_CpUtil_CpCodeMgr.GetStockStatusKind(code),
            'capital': self.obj_CpUtil_CpCodeMgr.GetStockCapital(code),
            'fiscalmonth': self.obj_CpUtil_CpCodeMgr.GetStockFiscalMonth(code),
            'groupcode': self.obj_CpUtil_CpCodeMgr.GetStockGroupCode(code),
            'kospi200kind': self.obj_CpUtil_CpCodeMgr.GetStockKospi200Kind(code),
            'section': self.obj_CpUtil_CpCodeMgr.GetStockSectionKind(code),
            'off': self.obj_CpUtil_CpCodeMgr.GetStockLacKind(code),
            'listeddate': self.obj_CpUtil_CpCodeMgr.GetStockListedDate(code),
            'maxprice': self.obj_CpUtil_CpCodeMgr.GetStockMaxPrice(code),
            'minprice': self.obj_CpUtil_CpCodeMgr.GetStockMinPrice(code),
            'ydopen': self.obj_CpUtil_CpCodeMgr.GetStockYdOpenPrice(code),
            'ydhigh': self.obj_CpUtil_CpCodeMgr.GetStockYdHighPrice(code),
            'ydlow': self.obj_CpUtil_CpCodeMgr.GetStockYdLowPrice(code),
            'ydclose': self.obj_CpUtil_CpCodeMgr.GetStockYdClosePrice(code),
            'creditenabled': self.obj_CpUtil_CpCodeMgr.IsStockCreditEnable(code),
            'parpricechangetype': self.obj_CpUtil_CpCodeMgr.GetStockParPriceChageType(code),
            'spac': self.obj_CpUtil_CpCodeMgr.IsSPAC(code),
            'biglisting': self.obj_CpUtil_CpCodeMgr.IsBigListingStock(code),
            'groupname': self.obj_CpUtil_CpCodeMgr.GetGroupName(code),
            'industryname': self.obj_CpUtil_CpCodeMgr.GetIndustryName(code),
            'membername': self.obj_CpUtil_CpCodeMgr.GetMemberName(code),
        }

        _fields = [67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 116, 118, 120, 123, 124, 125, 127, 156]
        _keys = ['PER', '시간외매수잔량', '시간외매도잔량', 'EPS', '자본금', '액면가', '배당률', '배당수익률', '부채비율', '유보율', '자기자본이익률', '매출액증가율', '경상이익증가율', '순이익증가율', '투자심리', 'VR', '5일회전율', '4일종가합', '9일종가합', '매출액', '경상이익', '당기순이익', 'BPS', '영업이익증가율', '영업이익', '매출액영업이익률', '매출액경상이익률', '이자보상비율', '분기BPS', '분기매출액증가율', '분기영업이액증가율', '분기경상이익증가율', '분기순이익증가율', '분기매출액', '분기영업이익', '분기경상이익', '분기당기순이익', '분개매출액영업이익률', '분기매출액경상이익률', '분기ROE', '분기이자보상비율', '분기유보율', '분기부채비율', '프로그램순매수', '당일외국인순매수', '당일기관순매수', 'SPS', 'CFPS', 'EBITDA', '공매도수량', '당일개인순매수']
        self.obj_CpSysDib_MarketEye.SetInputValue(0, _fields)
        self.obj_CpSysDib_MarketEye.SetInputValue(1, 'A'+code)
        self.obj_CpSysDib_MarketEye.BlockRequest()

        cnt_field = self.obj_CpSysDib_MarketEye.GetHeaderValue(0)
        if cnt_field > 0:
            for i in range(cnt_field):
                stock[_keys[i]] = self.obj_CpSysDib_MarketEye.GetDataValue(i, 0)
        return stock

    #주식 시장이나, 종목의 차트데이터 호출 (기간이 길면 연속조회로 수행)
    def get_chart(self, code, target='A', unit='D', n=None, date_from=None, date_to=None):
        _fields = []
        _keys = []
        if unit == 'm':
            _fields = [0, 1, 2, 3, 4, 5, 6, 8, 9, 37]
            _keys = ['date', 'time', 'open', 'high', 'low', 'close', 'diff', 'volume', 'price', 'diffsign']
        else:
            _fields = [0, 2, 3, 4, 5, 6, 8, 9, 37]
            _keys = ['date', 'open', 'high', 'low', 'close', 'diff', 'volume', 'price', 'diffsign']

        if date_to is None:
            date_to = util.get_str_today()

        self.obj_CpSysDib_StockChart.SetInputValue(0, target+code) # 주식코드: A, 업종코드: U
        if n is not None:
            self.obj_CpSysDib_StockChart.SetInputValue(1, ord('2'))  # 0: ?, 1: 기간, 2: 개수
            self.obj_CpSysDib_StockChart.SetInputValue(4, n)  # 요청 개수
        if date_from is not None or date_to is not None:
            if date_from is not None and date_to is not None:
                self.obj_CpSysDib_StockChart.SetInputValue(1, ord('1'))  # 0: ?, 1: 기간, 2: 개수
            if date_from is not None:
                self.obj_CpSysDib_StockChart.SetInputValue(3, date_from)  # 시작일
            if date_to is not None:
                self.obj_CpSysDib_StockChart.SetInputValue(2, date_to)  # 종료일
        self.obj_CpSysDib_StockChart.SetInputValue(5, _fields)  # 필드
        self.obj_CpSysDib_StockChart.SetInputValue(6, ord(unit))
        self.obj_CpSysDib_StockChart.SetInputValue(9, ord('1')) # 0: 무수정주가, 1: 수정주가

        def req(prev_result):
            self.obj_CpSysDib_StockChart.BlockRequest()

            status = self.obj_CpSysDib_StockChart.GetDibStatus()
            msg = self.obj_CpSysDib_StockChart.GetDibMsg1()
            if status != 0:
                return None

            cnt = self.obj_CpSysDib_StockChart.GetHeaderValue(3)
            list_item = []
            for i in range(cnt):
                dict_item = {k: self.obj_CpSysDib_StockChart.GetDataValue(j, cnt-1-i) for j, k in enumerate(_keys)}

                # type conversion
                dict_item['diffsign'] = chr(dict_item['diffsign'])
                for k in ['open', 'high', 'low', 'close', 'diff']:
                    dict_item[k] = float(dict_item[k])
                for k in ['volume', 'price']:
                    dict_item[k] = int(dict_item[k])

                # additional fields
                dict_item['diffratio'] = (dict_item['diff'] / (dict_item['close'] - dict_item['diff'])) * 100
                list_item.append(dict_item)
            return list_item

        # 연속조회 처리
        result = req([])
        while self.obj_CpSysDib_StockChart.Continue:
            self.avoid_reqlimitwarning()
            _list_item = req(result)
            if len(_list_item) > 0:
                result = _list_item + result
                if n is not None and n <= len(result):
                    break
            else:
                break
        return result
