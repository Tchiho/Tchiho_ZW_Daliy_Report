# -*- coding: UTF-8 -*-

from PIL import ImageGrab
import os
import sys
import time
import requests
import json
import schedule
import xlwings
import urllib3
urllib3.disable_warnings()

cwd = os.getcwd()
sys.path.append(cwd)

# 测试环境用True，到生产环境用False
is_test = False
# 需要发送报表的时间列表
send_msg_times = ('08:30', '16:30')
cur_dir = os.path.split(os.path.abspath(sys.argv[0]))[0] + '\\files'


class WechatInfo:
    """ 企业微信消息发送接口 """
    def __init__(self, corp_id, secret, agent_id):
        self.corpID = corp_id
        self.secret = secret
        self.agentID = agent_id
        self.token = None
        self.token_time = 0

    def __get_token(self):
        """获得令牌"""
        url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken"
        data = {
            "corpid": self.corpID,
            "corpsecret": self.secret
        }
        r = requests.get(url=url, params=data)
        token = r.json()['access_token']
        return token

    def _get_token(self):
        """令牌刷新机制"""
        if time.time() - self.token_time > 3600:
            self.token = self.__get_token()
            self.token_time = time.time()
        return self.token

    def send_message(self, message):
        """发送文字消息"""
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?" \
              f"access_token={self._get_token()}"
        data = {
            "touser": "@all",
            "msgtype": "text",
            "agentid": self.agentID,
            "text": {
                "content": message
            },
            "safe": "0"
        }
        result = requests.post(url=url, data=json.dumps(data))
        re = json.loads(result.text)
        print(re)
        if re['errcode'] == 40014:
            self.__get_token()
        if re['errcode'] != 0:
            print(re)
            return False
        return True

    def __get_media_url(self, path, mtype='image'):  # image 、 file
        """上传图片或文件，生成微信可用资源"""
        token = self._get_token()
        img_url = f"https://qyapi.weixin.qq.com/cgi-bin/media/upload?" \
                  f"access_token={token}&type={mtype}"
        files = {'media': open(path, 'rb')}
        result = requests.post(img_url, files=files)
        re = json.loads(result.text)
        print('__get_media_url:', re)
        return re['media_id'] if re['errcode'] == 0 else 0

    def __send_media(self, data):
        """将资源信息发送给用户"""
        token = self._get_token()
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?" \
              f"access_token={token}"
        result = requests.post(url=url, data=json.dumps(data))
        re = json.loads(result.text)
        if re['errcode'] == 40014:
            self.__get_token()
        if re['errcode'] != 0:
            print(re)
            return False
        return True

    def send_pic(self, pic_path):
        """发送图片资源"""
        img_url = self.__get_media_url(pic_path, "image")
        data = {
            "msgtype": "image",
            "agentid": self.agentID,
            "touser": "@all",
            "image": {
                "media_id": img_url
            },
            "safe": "0"
        }
        return self.__send_media(data)

    def send_file(self, file_path):
        """发送文件资源"""
        img_url = self.__get_media_url(file_path, "file")
        data = {
            "msgtype": "file",
            "agentid": self.agentID,
            "touser": "@all",
            "file": {
                "media_id": img_url
            },
            "safe": "0"
        }
        return self.__send_media(data)


def update_xls(template_file, data, xls_file, flag):
    """ 操作 Excel 我一般使用 openpyxl，通报模板使用了文本框，openpyxl 好像不能操作，就临时改成了 xlwings """
    os.system('taskkill /F /IM wps.exe')
    app = xlwings.App(visible=False, add_book=False)
    wb = app.books.open(f'{cur_dir}\\{template_file}')

    if flag == 1:
        sh = wb.sheets['报表']
        for shap in sh.shapes:
            if shap.type == 'text_box':
                tm = time.localtime()
                if tm.tm_min == 0:
                    shap.text = f'全渠道战报（{tm.tm_mon:02}月{tm.tm_mday:02}日{tm.tm_hour:02}时）'
                else:
                    shap.text = f'全渠道战报（{tm.tm_mon:02}月{tm.tm_mday:02}日{tm.tm_hour:02}时{tm.tm_min:02}分）'
                break

        for index, item in enumerate(data, 6):
            if item[0]:
                sh.range(f'A{index}').value = item[1:]
    elif flag == 2:
        sh = wb.sheets['拆机']
        for index, item in enumerate(data, 2):
            if item[0]:
                sh.range(f'A{index}').value = item[0:]

    wb.save(f'{cur_dir}\\{xls_file}')
    wb.close()
    app.quit()


def create_png():
    app = xlwings.App(visible=False, add_book=False)
    wb = app.books.open(f'{cur_dir}\\report.xlsx')
    sheet = wb.sheets('报表')
    _all = sheet.used_range
    # print(all.value)
    _all.api.CopyPicture()
    sheet.api.Paste()
    pic = sheet.pictures[1]
    pic.api.Copy()

    time.sleep(3)
    img = ImageGrab.grabclipboard()
    # 以下代码为去透明
    _l, _h = img.size
    for i in range(_h):
        for j in range(_l):
            d = (j, i)
            c = img.getpixel(d)
            if c == (0, 0, 0, 0):
                img.putpixel(d, (255, 255, 255, 255))
    img.save(f'{cur_dir}\\report.png')
    wb.save()
    wb.close()
    app.quit()


def send_msg():
    # 渠道发展
    msg1 = get_msg_from_api("msg1").get("data")[0][0]
    print(msg1)
    # wx = WechatInfo('wx50bf03cc4ce3f2a4', '0zhvUXfz15SdjhYarPpVvY3cJvuAlcOVMClFfaApjB0', '1000004')
    wx = WechatInfo('ww5a9ef257928f9273','VCzVz4DI2ac-Fy_017Q2zWFK83mg7FosKv-TzAR0F7c','1000004')
    wx.send_message(msg1)

    data1 = get_msg_from_api("data1").get("data")
    if data1 is not None:
        print("生成渠道发展xls")
        update_xls('template.xlsx', data1, 'report.xlsx', 1)
        print("生成图片")
        create_png()
        print("发送消息")
        wx.send_pic(f'{cur_dir}\\report.png')
        wx.send_file(f'{cur_dir}\\report.xlsx')

    # 拆机
    msg2 = get_msg_from_api("msg2").get("data")[0][0]
    print(msg2)
    wx.send_message(msg2)

    data2 = get_msg_from_api("data2").get("data")
    if data2 is not None:
        print("生成拆机xls")
        update_xls('start.xlsx', data2, 'chai_ji.xlsx', 2)
        wx.send_file(f'{cur_dir}\\chai_ji.xlsx')

    # wx0 = WechatInfo('ww5a9ef257928f9273','MAsRvMAMQWcJIxgZe-Rg4dioJmFBBXuQTWBi4c422bQ','1000002')
    # wx1 = WechatInfo('ww5a9ef257928f9273','HTjabdAxhvhc3MEOTzn9ZNCjLEGzpIcNDfbPFw0hE9w','1000003')
    # wx2 = WechatInfo('ww5a9ef257928f9273','hbIlhO3m3zlb7L750PPg7IhrMFx4-6TRV4ke6NiSjXk','1000005')
    # wx3 = WechatInfo('ww5a9ef257928f9273','VCzVz4DI2ac-Fy_017Q2zWFK83mg7FosKv-TzAR0F7c','1000004')



host = "https://10.37.3.194:9002/"
headers = {"User-Agent": "test request headers"}
api_token = None


def get_token_from_api():
    endpoint = "login"
 
    url = ''.join([host, endpoint])
    data = {
        'email': 'test@test.com',
        'password': 'test'
    }
    r = requests.post(url, headers=headers, json=data, verify=False)
    token = r.json().get('access_token')
    return token


def get_msg_from_api(sql_id):
    global api_token
    if api_token is None:
        api_token = get_token_from_api()
    headers['Authorization'] = 'Bearer ' + api_token
    url = ''.join([host, 'data/', sql_id])
    r = requests.get(url, headers=headers, verify=False)
    return r.json()


if __name__ == "__main__":
    print(f'系统开始工作，将于 {send_msg_times} 向企业微信发送报表...')
    for _time in send_msg_times:
        schedule.every().day.at(_time).do(send_msg)
    while True:
        schedule.run_pending()
        time.sleep(1)
    # send_msg()
