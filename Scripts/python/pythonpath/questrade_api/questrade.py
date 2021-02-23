import os
import json
from datetime import datetime, timedelta
import configparser
import urllib
from questrade_api.auth import Auth

CONFIG_PATH = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'questrade.cfg')

class Questrade:
    def __init__(self, **kwargs):
        if 'config' in kwargs:
            self.config = self.__read_config(kwargs['config'])
        else:
            self.config = self.__read_config(CONFIG_PATH)

        auth_kwargs = {x: y for x, y in kwargs.items() if x in
            ['token_path', 'refresh_token', 'storage_adaptor', 'logger']}

        __remaining = None
        __reset = None

        self.auth = Auth(**auth_kwargs, config=self.config)

    def __read_config(self, fpath):
        config = {
			"Settings" : {
				"Version": "v1"
			},
			"Auth" : {
				"RefreshURL": "https://login.questrade.com/oauth2/token?grant_type=refresh_token&refresh_token={}"
			},
			"API" : {
				"Time": "/time",
				"Accounts": "/accounts",
				"AccountPositions": "/accounts/{}/positions",
				"AccountBalances": "/accounts/{}/balances",
				"AccountExecutions": "/accounts/{}/executions",
				"AccountOrders": "/accounts/{}/orders",
				"AccountOrder": "/accounts/{}/orders/{}",
				"AccountActivities": "/accounts/{}/activities",
				"Symbol": "/symbols/{}",
				"SymbolOptions": "/symbols/{}/options",
				"Symbols": "/symbols",
				"SymbolsSearch": "/symbols/search",
				"Markets": "/markets",
				"MarketsQuote": "/markets/quotes/{}",
				"MarketsQuotes": "/markets/quotes",
				"MarketsOptions": "/markets/quotes/options",
				"MarketsStrategies": "/markets/quotes/strategies",
				"MarketsCandles": "/markets/candles/{}"
			}
		}
        return config

    @property
    def __base_url(self):
        return self.auth.token['api_server'] + self.config['Settings']['Version']

    def __build_get_req(self, url, params):
        if params:
            url = self.__base_url + url + '?' + urllib.parse.urlencode(params)
            return urllib.request.Request(url)
        else:
            return urllib.request.Request(self.__base_url + url)

    def __get(self, url, params=None):
        req = self.__build_get_req(url, params)
        req.add_header(
            'Authorization',
            self.auth.token['token_type'] + ' ' +
            self.auth.token['access_token']
        )
        try:
            r = urllib.request.urlopen(req)
            h = r.getheaders()
            self.__remaining = next( ( int( t[1] ) for t in h if t[0] == 'X-RateLimit-Remaining'), None )
            self.__reset = next( ( int( t[1] ) for t in h if t[0] == 'X-RateLimit-Reset'), None )
            return json.loads(r.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            return json.loads(e.read().decode('utf-8'))

    def __build_post_req(self, url, params):
        url = self.__base_url + url
        return urllib.request.Request(url, data=json.dumps(params).encode('utf8'))

    def __post(self, url, params):
        req = self.__build_post_req(url, params)
        req.add_header(
            'Authorization',
            self.auth.token['token_type'] + ' ' +
            self.auth.token['access_token']
        )
        try:
            r = urllib.request.urlopen(req)
            return json.loads(r.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            return json.loads(e.read().decode('utf-8'))

    @property
    def ratelimit(self):
        return ( self.__remaining, self.__reset, )

    @property
    def __now(self):
        return datetime.now().astimezone().isoformat('T')

    def __days_ago(self, d):
        now = datetime.now().astimezone()
        return (now - timedelta(days=d)).isoformat('T')

    @property
    def time(self):
        return self.__get(self.config['API']['Time'])

    @property
    def accounts(self):
        return self.__get(self.config['API']['Accounts'])

    def account_positions(self, id):
        return self.__get(self.config['API']['AccountPositions'].format(id))

    def account_balances(self, id):
        return self.__get(self.config['API']['AccountBalances'].format(id))

    def account_executions(self, id, **kwargs):
        return self.__get(self.config['API']['AccountExecutions'].format(id), kwargs)

    def account_orders(self, id, **kwargs):
        if 'ids' in kwargs:
            kwargs['ids'] = kwargs['ids'].replace(' ', '')
        return self.__get(self.config['API']['AccountOrders'].format(id), kwargs)

    def account_order(self, id, order_id):
        return self.__get(self.config['API']['AccountOrder'].format(id, order_id))

    def account_activities(self, id, **kwargs):
        if 'startTime' not in kwargs:
            kwargs['startTime'] = self.__days_ago(1)
        if 'endTime' not in kwargs:
            kwargs['endTime'] = self.__now
        return self.__get(self.config['API']['AccountActivities'].format(id), kwargs)

    def symbol(self, id):
        return self.__get(self.config['API']['Symbol'].format(id))

    def symbols(self, **kwargs):
        if 'ids' in kwargs:
            kwargs['ids'] = kwargs['ids'].replace(' ', '')
        return self.__get(self.config['API']['Symbols'].format(id), kwargs)

    def symbols_search(self, **kwargs):
        return self.__get(self.config['API']['SymbolsSearch'].format(id), kwargs)

    def symbol_options(self, id):
        return self.__get(self.config['API']['SymbolOptions'].format(id))

    @property
    def markets(self):
        return self.__get(self.config['API']['Markets'])

    def markets_quote(self, id):
        return self.__get(self.config['API']['MarketsQuote'].format(id))

    def markets_quotes(self, **kwargs):
        if 'ids' in kwargs:
            kwargs['ids'] = kwargs['ids'].replace(' ', '')
        return self.__get(self.config['API']['MarketsQuotes'], kwargs)

    def markets_options(self, **kwargs):
        return self.__post(self.config['API']['MarketsOptions'], kwargs)

    def markets_strategies(self, **kwargs):
        return self.__post(self.config['API']['MarketsStrategies'], kwargs)

    def markets_candles(self, id, **kwargs):
        if 'startTime' not in kwargs:
            kwargs['startTime'] = self.__days_ago(1)
        if 'endTime' not in kwargs:
            kwargs['endTime'] = self.__now
        return self.__get(self.config['API']['MarketsCandles'].format(id), kwargs)
