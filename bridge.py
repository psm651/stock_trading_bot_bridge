#creon.py를 외부와 연결해주는 브리지 서버

from flask import Flask, request, jsonify
from creon import Creon
import constants


app = Flask(__name__)
c = Creon()

#HTS 연결 인터페이스
@app.route('/connection', methods=['GET', 'POST', 'DELETE'])
def handle_connect():
    c = Creon()
    if request.method == 'GET':
        # check connection status
        return jsonify(c.connected())
    elif request.method == 'POST':
        # make connection
        data = request.get_json()
        _id = data['id']
        _pwd = data['pwd']
        _pwdcert = data['pwdcert']
        return jsonify(c.connect(_id, _pwd, _pwdcert))
    elif request.method == 'DELETE':
        # disconnect
        res = c.disconnect()
        c.kill_client()
        return jsonify(res)

# 종목 코드를 획득
@app.route('/stockcodes', methods=['GET'])
def handle_stockcodes():
    c = Creon()
    c.avoid_reqlimitwarning()
    market = request.args.get('market')
    if market == 'kospi':
        return jsonify(c.get_stockcodes(constants.MARKET_CODE_KOSPI))
    elif market == 'kosdaq':
        return jsonify(c.get_stockcodes(constants.MARKET_CODE_KOSDAQ))
    else:
        return '"market" should be one of "kospi" and "kosdaq".', 400

# 종목코드 상태조회
@app.route('/stockstatus', methods=['GET'])
def handle_stockstatus():
    c = Creon()
    c.avoid_reqlimitwarning()
    stockcode = request.args.get('code')
    if not stockcode:
        return '', 400
    status = c.get_stockstatus(stockcode)
    return jsonify(status)

# 종목의 차트 데이터 조회
'''
code : 종목 코드(값 필수),
n : 최근n개의 데이터조회
date_from : 해당 날짜부터 최근 날짜까지 조회
date_to : 데이터의 마지막 날짜 지정

return
[
  {
    "close": 1482.4599609375,
    "date": 20200323,
    "diff": -83.69000244140625,
    "diffratio": -5.3436774509669815,
    "diffsign": "0",
    "high": 1516.75,
    "low": 1458.4100341796875,
    "open": 1474.449951171875,
    "price": 9645271000000,
    "volume": 647528300
  }
]
'''
def handle_stockcandles():
    c = Creon()
    c.avoid_reqlimitwarning()
    stockcode = request.args.get('code')
    n = request.args.get('n')
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    if not (n or date_from):
        return 'Need to provide "n" or "date_from" argument.', 400
    stockcandles = c.get_chart(stockcode, target='A', unit='D', n=n, date_from=date_from, date_to=date_to)
    return jsonify(stockcandles)

#종목코드로 주식상세정보조
@app.route('/stockfeatures', methods=['GET'])
def handle_stockfeatures():
    c = Creon()
    c.avoid_reqlimitwarning()
    stockcode = request.args.get('code')
    if not stockcode:
        return '', 400
    stockfeatures = c.get_stockfeatures(stockcode)
    return jsonify(stockfeatures)

#공매도 추
@app.route('/short', methods=['GET'])
def handle_short():
    c = Creon()
    c.avoid_reqlimitwarning()
    stockcode = request.args.get('code')
    n = request.args.get('n')
    if not stockcode:
        return '', 400
    stockfeatures = c.get_shortstockselling(stockcode, n=n)
    return jsonify(stockfeatures)

if __name__ == "__main__":
    app.run()
