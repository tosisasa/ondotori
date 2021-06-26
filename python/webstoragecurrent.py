# おんどとりWeb Storage
# 指定期間・件数によるデータの取得
import requests

import json
import datetime


url = "https://api.webstorage.jp/v1/devices/current"

api_key = "APIキー"
login_id ="閲覧用ID"
password = "閲覧用アカウントのパスワード"



paylord = {'api-key':api_key,"login-id":login_id,'login-pass':password}

header = {'Host':'api.webstrage.js:443','Content-Type': 'application/json','X-HTTP-Method-Override':'GET'}

def main():
    response = requests.post(url,json.dumps(paylord).encode('utf-8'),headers=header).json()


    for davices in response['devices']:

        dt = datetime.datetime.fromtimestamp(int(davices['unixtime']))

        print(davices['group']['name'] + ',' + 
              davices['channel'][0]['name'] + ',' + 
              davices['channel'][0]['value'] + ',' + 
              davices['channel'][1]['name'] + ',' + 
              davices['channel'][1]['value'] + ',' +
              dt.strftime('%Y/%m/%d %H:%M:%S'))


if __name__ =="__main__":
    main()