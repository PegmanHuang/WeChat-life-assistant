from wxpy import*
import pygame
from openpyxl import load_workbook
import requests
import datetime
import time
from datetime import datetime, timedelta
import win32com.client as win
import win32com.client as wc

speak = win.Dispatch("SAPI.SpVoice")
bot = Bot()
girl_friend=bot.search('YQD')[0]
self_friend=bot.search('XHR')[0]

global zhuangtai
zhuangtai = '尚未设置状态信息'

def music():
    file = r'‪D:/Loveless.mp3'
    pygame.mixer.init()
    print("正在播放音乐")
    track = pygame.mixer.music.load(file)

    pygame.mixer.music.play()
    time.sleep(60)
    pygame.mixer.music.stop()

@bot.register()
def recv_send_msg(recv_msg):
    global zhuangtai
    print(recv_msg) # recv_msg.text取得文本
    if recv_msg.sender == girl_friend:
        msg = str(recv_msg)
        if "你好" in msg:
            girl_friend.send('你好啊')
        elif "叫醒" in msg:
            girl_friend.send('正在播放音乐')
            import pygame
            pygame.mixer.init()
            print("正在播放音乐")
            track = pygame.mixer.music.load(r'‪‪D:/Loveless.mp3')
            pygame.mixer.music.play()
            time.sleep(60)
            pygame.mixer.music.stop()
            girl_friend.send('播放结束')
            print("播放结束")
        elif "叫小呆起床" in msg:
            girl_friend.send('正在播放音乐')
            import pygame
            pygame.mixer.init()
            print("正在播放音乐")
            track = pygame.mixer.music.load(r'D:/Loveless.mp3')
            pygame.mixer.music.play()
            time.sleep(60)
            pygame.mixer.music.stop()
            girl_friend.send('播放结束')
            print("播放结束")
            
        elif "睡觉" in msg:
            speak.Speak('小乙说时间比较晚了该睡觉了')
            girl_friend.send('提醒发送成功')
        elif "小呆在干嘛" in msg:
            girl_friend.send('小呆在'+zhuangtai)
            speak.Speak('小呆，小乙有点想你，在问你在干嘛')
        elif "早上好" in msg:
            pre = datetime(2019, 4, 13)  # 年月日时分秒 微秒
            now = datetime.now()
            dst = now - pre
            dst = str(dst)
            dst = dst[0:3]
            msg0 = '早上好呀，你们已经相爱' + dst + '天了哦。小呆在' + zhuangtai + '。'

            req1 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320505&extensions=all")
            req2 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320923&extensions=all")
            info1 = dict(req1.json())
            info2 = dict(req2.json())
            msg1 = '苏州今日' + info1['forecasts'][0]['casts'][0][
                'dayweather'] + '，气温:' + ' ' + info1['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info1['forecasts'][0]['casts'][0]['daytemp'] + '度。'
            msg2 = '阜宁今日' + info2['forecasts'][0]['casts'][0][
                'dayweather'] + '，气温:' + ' ' + info2['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info2['forecasts'][0]['casts'][0]['daytemp'] + '度。'
            print(msg0)
            print(msg1)
            print(msg2)
            girl_friend.send(msg0 + msg2 + msg1 + '新的一天要元气满满哦。')

        elif "今日天气" in msg:
            req1 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320505&extensions=all")
            req2 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320923&extensions=all")
            info1 = dict(req1.json())
            info2 = dict(req2.json())
            msg1 = '苏州今日天气:' + ' ' + info1['forecasts'][0]['casts'][0][
                'dayweather'] + ' ' + ':气温:' + ' ' + info1['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info1['forecasts'][0]['casts'][0]['daytemp'] + '度'
            msg2 = '阜宁今日天气:' + ' ' + info2['forecasts'][0]['casts'][0][
                'dayweather'] + ' ' + ':气温:' + ' ' + info2['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info2['forecasts'][0]['casts'][0]['daytemp'] + '度'
            print(msg1)
            print(msg2)
            girl_friend.send(msg1)
            girl_friend.send(msg2)
            
        elif "明日天气" in msg:
            req1 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320505&extensions=all")
            req2 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320923&extensions=all")
            info1 = dict(req1.json())
            info2 = dict(req2.json())
            msg1 = '苏州' + '明日天气:' + ' ' + info1['forecasts'][0]['casts'][1][
                'dayweather'] + ' ' + ':气温:' + ' ' + info1['forecasts'][0]['casts'][1]['nighttemp'] + '~' + \
                   info1['forecasts'][0]['casts'][1]['daytemp'] + '度'
            msg2 = '阜宁' + '明日天气:' + ' ' + info2['forecasts'][0]['casts'][1][
                'dayweather'] + ' ' + ':气温' + ' ' + info2['forecasts'][0]['casts'][1]['nighttemp'] + '~' + \
                   info2['forecasts'][0]['casts'][1]['daytemp'] + '度'
            print(msg1)
            print(msg2)
            girl_friend.send(msg1)
            girl_friend.send(msg2)
            
        else:
            msg = '包子还没有学会哦，已经通知小呆教我啦，可先尝试其他指令哦'
            girl_friend.send(msg)

    if recv_msg.sender == self_friend:
        msg = str(recv_msg)
        if "当前状态" in msg:
            zhuangtai = msg[10:-6]
            speak.Speak('当前状态已设置为:' + zhuangtai)
        elif "叫小呆起床" in msg:
            self_friend.send('正在播放音乐')
            import pygame
            pygame.mixer.init()
            print("正在播放音乐")
            track = pygame.mixer.music.load(r'D:/Loveless.mp3')
            
            pygame.mixer.music.play()
            time.sleep(60)
            pygame.mixer.music.stop()
            self_friend.send('播放结束')
            print("播放结束")
        elif "早上好" in msg:
            pre = datetime(2019, 4, 13)  # 年月日时分秒 微秒
            now = datetime.now()
            dst = now - pre
            dst = str(dst)
            dst = dst[0:3]
            msg0 = '早上好，你们已经相爱' + dst + '天。'

            req1 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320505&extensions=all")
            req2 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320923&extensions=all")
            info1 = dict(req1.json())
            info2 = dict(req2.json())
            msg1 = '苏州今日' + info1['forecasts'][0]['casts'][0][
                'dayweather'] + '，气温:' + ' ' + info1['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info1['forecasts'][0]['casts'][0]['daytemp'] + '度。'
            msg2 = '阜宁今日' + info2['forecasts'][0]['casts'][0][
                'dayweather'] + '，气温:' + ' ' + info2['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info2['forecasts'][0]['casts'][0]['daytemp'] + '度。'
            print(msg0)
            print(msg1)
            print(msg2)
            self_friend.send(msg0+msg2+msg1)
            speak.Speak(msg0+msg2+msg1)

            excel_app = wc.Dispatch('Excel.Application')
            workbook = excel_app.Workbooks.Open(r'D:\\richengbiao.xlsx')
            deadtime = str(time.strftime("%m%d", time.localtime()))
            i = 2
            while True:
                code1 = workbook.Worksheets('Sheet1').Cells(i, 4).Value
                if code1 == None:
                    break
                code1 = int(code1)
                #print(code1)
                if code1 == int(deadtime):
                    print(workbook.Worksheets('Sheet1').Cells(i, 3).Value)
                    self_friend.send('记得'+workbook.Worksheets('Sheet1').Cells(i, 3).Value)
                    speak.Speak('记得'+workbook.Worksheets('Sheet1').Cells(i, 3).Value)
                # print(i)
                i += 1
            excel_app.Quit()

        elif "晚上好" in msg:
            req1 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320505&extensions=all")
            req2 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320923&extensions=all")
            info1 = dict(req1.json())
            info2 = dict(req2.json())
            msg1 = '苏州' + '明日天气:' + ' ' + info1['forecasts'][0]['casts'][1][
                'dayweather'] + ' ' + ':气温:' + ' ' + info1['forecasts'][0]['casts'][1]['nighttemp'] + '~' + \
                   info1['forecasts'][0]['casts'][1]['daytemp'] + '度'
            msg2 = '阜宁' + '明日天气:' + ' ' + info2['forecasts'][0]['casts'][1][
                'dayweather'] + ' ' + ':气温' + ' ' + info2['forecasts'][0]['casts'][1]['nighttemp'] + '~' + \
                   info2['forecasts'][0]['casts'][1]['daytemp'] + '度'
            print(msg1)
            print(msg2)
            self_friend.send(msg1)
            self_friend.send(msg2)
            speak.Speak(msg1)
            speak.Speak(msg2)

        elif "今日天气" in msg:
            req1 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320505&extensions=all")
            req2 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320923&extensions=all")
            info1 = dict(req1.json())
            info2 = dict(req2.json())
            msg1 = '苏州今日天气:' + ' ' + info1['forecasts'][0]['casts'][0][
                'dayweather'] + ' ' + ':气温:' + ' ' + info1['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info1['forecasts'][0]['casts'][0]['daytemp'] + '度'
            msg2 = '阜宁今日天气:' + ' ' + info2['forecasts'][0]['casts'][0][
                'dayweather'] + ' ' + ':气温:' + ' ' + info2['forecasts'][0]['casts'][0]['nighttemp'] + '~' + \
                   info2['forecasts'][0]['casts'][0]['daytemp'] + '度'
            print(msg1)
            print(msg2)
            self_friend.send(msg1)
            self_friend.send(msg2)
            speak.Speak(msg1)
            speak.Speak(msg2)
        elif "明日天气" in msg:
            req1 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320505&extensions=all")
            req2 = requests.get(
                "http://restapi.amap.com/v3/weather/weatherInfo?key=**************************&city=320923&extensions=all")
            info1 = dict(req1.json())
            info2 = dict(req2.json())
            msg1 = '苏州' + '明日天气:' + ' ' + info1['forecasts'][0]['casts'][1][
                'dayweather'] + ' ' + ':气温:' + ' ' + info1['forecasts'][0]['casts'][1]['nighttemp'] + '~' + \
                   info1['forecasts'][0]['casts'][1]['daytemp'] + '度'
            msg2 = '阜宁' + '明日天气:' + ' ' + info2['forecasts'][0]['casts'][1][
                'dayweather'] + ' ' + ':气温' + ' ' + info2['forecasts'][0]['casts'][1]['nighttemp'] + '~' + \
                   info2['forecasts'][0]['casts'][1]['daytemp'] + '度'
            print(msg1)
            print(msg2)
            self_friend.send(msg1)
            self_friend.send(msg2)
            speak.Speak(msg1)
            speak.Speak(msg2)
        elif "提醒我" in msg:
            deadtime = str(time.strftime("%Y%m%d %H:%M:%S", time.localtime()))
            wb = load_workbook('D:\\richengbiao.xlsx')  # 创建/读取文件对象
            ws = wb.active  # 获取第一个sheet
            # 字段截取
            msg1 = msg[0:3]
            msg2 = msg[6:9]
            msg3 = msg[10:-12]
            msg4 = msg[-11:-6]

            # print(msg1)  # 截取用户名
            # print(msg2)  # 截取日程类型
            # print(msg3)  # 截取日程内容
            # print(msg4)  # 截取提醒时间

            ws.append([msg1, msg2, msg3, msg4, deadtime])  # 写入多个单元格
            # Save the file
            wb.save('D:\\richengbiao.xlsx')

            speak.Speak('新日程已建立，将会于' + msg4 + '日提醒你' + msg3)
            self_friend.send('新日程已建立，将会于' + msg4 + '日提醒你' + msg3)
            print('新日程已建立，将会于' + msg4 + '日提醒你' + msg3)
        else:
            msg = '包子还没有学会哦，可尝试其他指令'
            self_friend.send(msg)
            
embed()









