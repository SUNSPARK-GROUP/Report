from urllib.parse import quote
from hashlib import md5
from http import client
import random
import json

appid = '20180210000122225'
secret_key = 'lSVlEXTpSk5Ihppi3IpS'

def get_sign(salt, qurey):
    sign = appid + qurey + str(salt) + secret_key
    m = md5()
    m.update(sign.encode('utf-8'))
    return m.hexdigest()

def trans(qurey, from_lang, to_lang):
    http_client = None
    salt = random.randint(12345, 67890)
    sign = get_sign(salt, qurey)
    myurl = '/api/trans/vip/translate' + '?appid=' + appid + '&q=' + quote(
        qurey) + '&from=' + from_lang + '&to=' + to_lang + '&salt=' + str(salt) + '&sign=' + sign
    try:
        http_client = client.HTTPConnection('api.fanyi.baidu.com')
        http_client.request('GET', myurl)
        response = http_client.getresponse()
        content = json.loads(response.read())
        return content['trans_result'][0]['dst']
    except Exception as e:
        return e
    finally:
        if http_client:
            http_client.close()


if __name__ == '__main__':
    print(trans('翻译工具'))
