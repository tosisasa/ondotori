# おんどとりWeb Storage
# 指定期間・件数によるデータの取得
import requests

import json
import datetime


url = "https://api.webstorage.jp/v1/devices/data"

api_key = "APIキー"
login_id ="閲覧用ID"
password = "閲覧用アカウントのパスワード"
sereal = "機器のシリアル番号"


dtfrom = datetime.datetime(2021, 6, 20, 0, 0, 0, 0)
timefrom = int(dtfrom.timestamp())

dtto = datetime.datetime(2021, 6, 25, 0, 0, 0, 0)
timeto = int(dtto.timestamp())

paylord = {'api-key':api_key,"login-id":login_id,'login-pass':password,
'remote-serial':sereal, 'unixtime-from':timefrom, 'unixtime-to':timeto}

header = {'Host':'api.webstrage.js:443','Content-Type': 'application/json','X-HTTP-Method-Override':'GET'}

def main():
    response = requests.post(url,json.dumps(paylord).encode('utf-8'),headers=header).json()

    print(response['serial'] + ',' + response['model'] + ',' + response['name'])

    print(response['channel'][0]['num'] + ',' + response['channel'][0]['name'] + ',' + response['channel'][0]['unit']
     + '       ' + response['channel'][1]['num'] + ',' + response['channel'][1]['name'] + ',' + response['channel'][1]['unit'])


    for data in response['data']:

        dt = datetime.datetime.fromtimestamp(int(data['unixtime']))
        print(dt.strftime('%Y/%m/%d %H:%M:%S') + ',' + data['ch1'] + ',' + data['ch2'])


if __name__ =="__main__":
    main()