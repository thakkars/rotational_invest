import enum
import logging
import re
import requests
import json
import random
import string
from datetime import datetime, date
import pandas as pd

logger = logging.getLogger(__name__)


class Interval(enum.Enum):
    in_1_minute = '1'
    in_3_minute = '3'
    in_5_minute = '5'
    in_15_minute = '15'
    in_30_minute = '30'
    in_45_minute = '45'
    in_1_hour = '1H'
    in_2_hour = '2H'
    in_3_hour = '3H'
    in_4_hour = '4H'
    in_daily = '1D'
    in_weekly = '1W'
    in_monthly = '1M'


class TvDatafeed:
    __sign_in_url = 'https://www.tradingview.com/accounts/signin/'
    __search_url = 'https://symbol-search.tradingview.com/symbol_search/?text={}&hl=1&exchange={}&lang=en&type=&domain=production'
    __ws_headers = json.dumps({"Origin": "https://data.tradingview.com"})
    __signin_headers = {'Referer': 'https://www.tradingview.com'}
    __ws_timeout = 5

    def __init__(self, username=None, password=None):
        self.ws_debug = False
        self.token = self.__auth(username, password)
        if self.token is None:
            self.token = 'unauthorized_user_token'
            logger.warning('using unauthorized token')
        self.ws = None
        self.session = self.__generate_session()
        self.chart_session = self.__generate_chart_session()

    def __auth(self, username, password):
        if username is None or password is None:
            return None
        data = {"username": username, "password": password,
                "remember": "on"}
        try:
            response = requests.post(url=self.__sign_in_url, data=data,
                                     headers=self.__signin_headers)
            return response.json()['user']['auth_token']
        except Exception as e:
            logger.error('error during auth: ' + str(e))
            return None

    def __generate_session(self):
        stringLength = 12
        letters = string.ascii_lowercase
        random_string = ''.join(random.choice(letters) for _ in range(stringLength))
        return 'qs_' + random_string

    def __generate_chart_session(self):
        stringLength = 12
        letters = string.ascii_lowercase
        random_string = ''.join(random.choice(letters) for _ in range(stringLength))
        return 'cs_' + random_string

    def __prepend_header(self, st):
        return '~m~' + str(len(st)) + '~m~' + st

    def __construct_message(self, func, paramList):
        return json.dumps({'m': func, 'p': paramList}, separators=(',', ':'))

    def __create_message(self, func, paramList):
        return self.__prepend_header(self.__construct_message(func, paramList))

    def __send_message(self, ws, func, args):
        ws.send(self.__create_message(func, args))

    def __create_df(self, raw_data, symbol):
        try:
            out = re.search('"s":\[(.+?)\]', raw_data).group(1)
            x = out.split(',{"i"')
            x[0] = '[{"i"' + x[0][1:]
            data = []
            for xi in x:
                xi = re.split(r'\[|\]', xi)
                ts = datetime.fromtimestamp(float(xi[1].split(',')[0]))
                row = [ts]
                for i in xi[1].split(',')[1:]:
                    try:
                        row.append(float(i))
                    except ValueError:
                        row.append(0)
                data.append(row)
            df = pd.DataFrame(data, columns=['datetime', 'open', 'high', 'low', 'close', 'volume']).set_index('datetime')
            df.index = pd.to_datetime(df.index)
            df['symbol'] = symbol
            return df
        except AttributeError:
            logger.error('no data found')
            return None

    def __fetch_data(self, symbol, exchange, interval, n_bars, fut_contract=None, extended_session=False):
        import websocket
        ws = websocket.WebSocket()
        ws.settimeout(self.__ws_timeout)
        try:
            ws.connect('wss://data.tradingview.com/socket.io/websocket',
                       headers=self.__ws_headers)
        except Exception as e:
            logger.error('websocket connect error: ' + str(e))
            return None

        session = self.__generate_session()
        chart_session = self.__generate_chart_session()

        symbol_id = 'symbol_' + session

        if fut_contract is None:
            symbol_str = json.dumps({'symbol': symbol if exchange == '' else f'{exchange}:{symbol}',
                                     'adjustment': 'splits'})
        else:
            symbol_str = json.dumps({'symbol': f'{exchange}:{symbol}{fut_contract}',
                                     'adjustment': 'splits'})

        interval_val = interval.value

        self.__send_message(ws, 'set_auth_token', [self.token])
        self.__send_message(ws, 'chart_create_session', [chart_session, ''])
        self.__send_message(ws, 'quote_create_session', [session])
        self.__send_message(ws, 'quote_set_fields', [session,
                            'ch', 'chp', 'current_session', 'description',
                            'local_description', 'language', 'exchange',
                            'fractional', 'is_tradable', 'lp', 'lp_time',
                            'minmov', 'minmov2', 'original_name', 'pricescale',
                            'pro_name', 'short_name', 'type', 'update_mode',
                            'volume', 'currency_code', 'rchp', 'rtc'])
        self.__send_message(ws, 'quote_add_symbols', [session, f'{exchange}:{symbol}',
                            {'flags': ['force_permission']}])
        self.__send_message(ws, 'resolve_symbol', [chart_session,
                            symbol_id, symbol_str])
        self.__send_message(ws, 'create_series', [chart_session, 's1', 's1',
                            symbol_id, interval_val, n_bars])

        raw_data = ''
        while True:
            try:
                result = ws.recv()
                raw_data += result + '\n'
            except Exception:
                break

        ws.close()
        return self.__create_df(raw_data, symbol)

    def get_hist(self, symbol: str, exchange: str = 'NSE',
                 interval: Interval = Interval.in_daily,
                 n_bars: int = 5000, fut_contract: int = None,
                 extended_session: bool = False) -> pd.DataFrame:
        return self.__fetch_data(symbol, exchange, interval, n_bars,
                                 fut_contract, extended_session)
