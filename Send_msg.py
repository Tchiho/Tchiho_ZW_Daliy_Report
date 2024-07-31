import requests
import ssl
import urllib3
# import certifi

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

"""
# 创建自定义证书存储
cert_file = 'msg/certificate.pem'
key_file = 'msg/private.key'
cert_store = ssl.create_default_context(ssl.Purpose.CLIENT_AUTH)
cert_store.load_cert_chain(cert_file, key_file, password='true')
# 使用自定义证书存储创建PoolManager
http_pool_manager = urllib3.PoolManager(
    ssl_context=cert_store,
    # cert_reqs='CERT_REQUIRED',
    cert_reqs='CERT_NONE',
    ca_certs=certifi.where()
)
"""

host = "https://10.37.3.250/"
headers = {"User-Agent":"test request headers"}
token  = None

def get_token_from_api():
    endpoint = "login"
    url = ''.join([host,endpoint])
    data = {
        'email': 'Tchiho@outlook.com',
        'password': '201214'
    }
    r = requests.post(url, headers=headers, json=data, verify=False)
    token = r.json().get('access_token')
    print(token)
    return token


def post_img_api(file_path = "D:\Program\ZY_Daily_Reprt\日报表\装维工单日报模板\幻灯片1.png"):
    global token
    endpoint = "file"
    url = ''.join([host,endpoint])
    if token is None:
        token = get_token_from_api()
    headers['Authorization'] = 'Bearer ' + token
    r = None
    with open(file_path, "rb") as file:
        files = {"file": (file_path, file)}
        res = requests.post(url, headers=headers, files=files, verify=False)
        r = res.json()
    return r


def post_chat_api(seq, to="发微信测试"):
    global token
    endpoint = "chat"
    url = ''.join([host,endpoint])
    if token is None:
        token = get_token_from_api()
    headers['Authorization'] = 'Bearer ' + token
    data = {
        'img': seq,
        'to': to,
    }
    r = requests.post(url, headers=headers, json=data, verify=False)
    return r


seq = post_img_api().get('seq')

post_chat_api(seq)

