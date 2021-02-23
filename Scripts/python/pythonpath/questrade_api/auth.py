import os
import json
import time
from urllib import request

TOKEN_PATH = os.path.expanduser('~/.questrade.json')

class Auth:
    def __init__(self, **kwargs):
        if 'logger' in kwargs:
            self.logger = kwargs['logger']
        if 'config' in kwargs:
            self.config = kwargs['config']
        else:
            raise Exception('No config supplied')
        if 'token_path' in kwargs:
            self.token_path = kwargs['token_path']
        else:
            self.token_path = TOKEN_PATH
        if 'storage_adaptor' in kwargs:
            ( self.read_adaptor, self.write_adaptor) = kwargs['storage_adaptor']
        if 'refresh_token' in kwargs:
            self.__refresh_token(kwargs['refresh_token'])

    def __logger(self, s):
        if hasattr(self, 'logger'):
            self.logger(s)

    def __read_token(self):
        try:
            if hasattr( self, 'read_adaptor'):
                s = self.read_adaptor( )
                self.__logger( '__read_token::read_adaptor( ) = {}'.format( s ) )
            else:
                with open(self.token_path) as f:
                    s = f.read()
                self.__logger( '__read_token::default( ) = {}'.format( s ) )
            return json.loads(s)
        except:
            raise RuntimeError('No token provided and none found at {}'.format(TOKEN_PATH))

    def __write_token(self, token):
        try:
            s = json.dumps(token, indent=4, sort_keys=True)
            if hasattr( self, 'write_adaptor'):
                self.write_adaptor(s)
                self.__logger( '__write_token::write_adaptor( ) = {}'.format( s ) )
            else:
                with open(self.token_path, 'w') as f:
                    f.write(s)
                os.chmod(self.token_path, 0o600)
                self.__logger( '__write_token::default( ) = {}'.format( s ) )
        except:
            raise RuntimeError('Failed to cache token file.')

    def __refresh_token(self, token):
        req_time = int(time.time())
        r = request.urlopen(self.config['Auth']['RefreshURL'].format(token))
        if r.getcode() == 200:
            token = json.loads(r.read().decode('utf-8'))
            token['expires_at'] = str(req_time + token['expires_in'])
            self.__write_token(token)

    def __get_valid_token(self):
        try:
            self.token_data
        except AttributeError:
            self.token_data = self.__read_token()
        finally:
            if time.time() + 60 < int(self.token_data['expires_at']):
                return self.token_data
            else:
                self.__refresh_token(self.token_data['refresh_token'])
                self.token_data = self.__read_token()
                return self.token_data

    @property
    def token(self):
        return self.__get_valid_token()
