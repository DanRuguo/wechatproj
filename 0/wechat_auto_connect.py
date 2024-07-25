from wxauto import WeChat
from contextlib import redirect_stdout
import io
import time
import pyautogui

# 不断尝试连接微信
while True:
    try:
        output = io.StringIO()
        with redirect_stdout(output):
            wx = WeChat()  # 尝试创建 WeChat 对象，输出会被捕获
        output_value = output.getvalue()  # 获取输出内容
        print(output_value)  # 可以打印查看全部输出

        # 解析输出中的微信昵称
        if "获取到已登录窗口：" in output_value:
            nickname_index = output_value.index("获取到已登录窗口：") + len("获取到已登录窗口：")
            wechat_nickname = output_value[nickname_index:].strip()
            print(f"已连接微信，昵称为：{wechat_nickname}")
            # 尝试调出文件传输助手的聊天界面
            who = '文件传输助手'
            chatname = wx.ChatWith(who)  # 调用ChatWith函数
            print("初始化微信成功")
            time.sleep(1)  # 等待1秒
            pyautogui.press('esc')  # 发送ESC键
            break
    except Exception as e:
        print(f"尝试连接微信失败：{e}")
    time.sleep(0.8)  # 每0.8秒尝试一次