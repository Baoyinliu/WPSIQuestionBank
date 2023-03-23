
# encoding:utf-8
import requests 

# client_id 为官网获取的AK， client_secret 为官网获取的SK
#host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=8THFpveNZN1rp6mBwH8qX1ES&client_secret=Lk4wMvzEGG2zhnRYLGk1OKwy3RQ06NmT'
host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=KPmoHDDpOguXBBzgVE4xo910&client_secret=FLKdfjcSDMptF1vmHXMQuTfDMY4CE5nQ'
response = requests.get(host)
if response:
    print(response.json())