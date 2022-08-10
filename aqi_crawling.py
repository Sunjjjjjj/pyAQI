import json
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
from pandas import json_normalize

# city list
city_list = 'https://raw.githubusercontent.com/modood/Administrative-divisions-of-China/master/dist/cities.json'
cities = json.loads(requests.get(city_list).text)  # 发起http Post请求

headers = {
    'Accept': '*/*',
    'Accept-Encoding': 'gzip',
    'Connection': 'keep-alive',
    'Content-Type': 'application/json',
    'Host': 'epapi.moji.com',
    'User-Agent': 'AirMonitoring/4.3.8 (iPhone iOS 15.6 Scale/3.00)',
    'Accept-Language': 'zh-Hans-CN;q=1,en-CN;q=0.9'
}  


#%% hourly by city
city_hourly_url = 'http://epapi.moji.com/json/epa/stationList'  
city_hourly_request_body = {
    "common": {
        "cityid": "320100",
        "app_version": "4504030802",
        "device": "iPhone14,2",
        "apnsisopen": "1",
        "platform": "iPhone",
        "uid": 2125537239493246070,
        "language": "CN",
        "identifier": "00000000-0000-0000-0000-000000000000",
        "token": "{length = 32, bytes = 0xb2acc743 18bc3b51 f2fde485 30cc944b ... c1db0d9c 91adffc6 }"
    },
    "params": {
        "measureLevel": 4,
        "longitude": "107.2668609619141",
        "latitude": "38.22311019897461"
    }
}  

response = requests.post(city_hourly_url, headers=headers,data=json.dumps(city_hourly_request_body)) #发起http Post请求
response_text = response.text 
aqi_json = json.loads(response_text) 
aqi_df = json_normalize(aqi_json['list']) 
filename =time.strftime('%Y_%m_%d_%H_%M_%S', time.localtime()) + '_aqi_data.csv'
aqi_df.to_csv(filename, index=False)
#%% hourly by site 
site_hourly_url = 'http://epapi.moji.com/json/epa/cityStationList'
site_hourly_request_body = {
    "common": {
        "cityid": "320100",
        "app_version": "4504030802",
        "device": "iPhone14,2",
        "apnsisopen": "1",
        "platform": "iPhone",
        "uid": 2125537239493246070,
        "language": "CN",
        "identifier": "00000000-0000-0000-0000-000000000000",
        "token": "{length = 32, bytes = 0xb2acc743 18bc3b51 f2fde485 30cc944b ... c1db0d9c 91adffc6 }"
    },
    "params": {
        "latitude": 31.97385036892361,
        "longitude": 118.77205810546874,
        "cityId": ""
    }
}  

for city in cities[:10]:
    site_hourly_request_body["params"]["cityId"] = city["code"] + '00'
    response = requests.post(
        site_hourly_url, headers=headers, data=json.dumps(site_hourly_request_body))  # 发起http Post请求

    response_text = response.text 
    aqi_json = json.loads(response_text)  
    aqi_df = json_normalize(aqi_json['list'])
    filename = city["name"] + '_' + time.strftime(
        '%Y_%m_%d_%H_%M_%S', time.localtime()) + '_aqi_data.csv'
    aqi_df.to_csv(filename, index=False)

#%% calendar by city 
city_daily_url = 'http://epapi.moji.com/json/aqi/calendar/query' 
city_daily_request_body = {
    "common": {
        "cityid": "320100",
        "app_version": "4504030802",
        "device": "iPhone14,2",
        "apnsisopen": "1",
        "platform": "iPhone",
        "uid": 2125537239493246070,
        "language": "CN",
        "identifier": "00000000-0000-0000-0000-000000000000",
        "token": "{length = 32, bytes = 0xb2acc743 18bc3b51 f2fde485 30cc944b ... c1db0d9c 91adffc6 }"
    },
    "params": {
        "latitude": 31.97385036892361,
        "longitude": 118.77205810546874,
        "cityId": "320100",
        "startTime":"2022-07-01",
        "endTime":"2022-07-31"
    }
}

response = requests.post(city_daily_url, headers=headers,data=json.dumps(city_daily_request_body)) #发起http Post请求
response_text = response.text #返回报文
aqi_json = json.loads(response_text) #序列化json
aqi_df = json_normalize(aqi_json['data']['list']) #json转dataframe
filename =time.strftime('%Y_%m_%d_%H_%M_%S', time.localtime()) + '_aqi_calendar.csv'
aqi_df.to_csv(filename, index=False)

