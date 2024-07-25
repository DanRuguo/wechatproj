import time
from PIL import ImageGrab
import os
import io
import win32clipboard
import winreg
import cv2
import numpy as np
import pygetwindow as gw
import pyautogui
import pytesseract


# Configure tesseract path if needed
# pytesseract.pytesseract.tesseract_cmd = r'YOUR_TESSERACT_PATH'

def send_to_clipboard(image):
    output = io.BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def find_qr_code(image):
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_BGR2GRAY)
    qr_code_detector = cv2.QRCodeDetector()
    data, points, _ = qr_code_detector.detectAndDecode(gray)
    if points is not None:
        points = points[0]
        left = int(min(points[:, 0]))
        top = int(min(points[:, 1]))
        right = int(max(points[:, 0]))
        bottom = int(max(points[:, 1]))
        return (left, top, right, bottom)
    return None

def click_switch_account():
    screenshot = ImageGrab.grab()
    text_data = pytesseract.image_to_data(screenshot, lang='chi_sim', output_type=pytesseract.Output.DICT)
    
    for i, text in enumerate(text_data['text']):
        if "切换" in text and "账号" in text_data['text'][i+1]:
            x1 = text_data['left'][i] + text_data['width'][i] // 2
            y1 = text_data['top'][i] + text_data['height'][i] // 2
            x2 = text_data['left'][i+1] + text_data['width'][i+1] // 2
            y2 = text_data['top'][i+1] + text_data['height'][i+1] // 2
            x = (x1 + x2) // 2
            y = (y1 + y2) // 2
            pyautogui.click(x, y)
            time.sleep(1)
            return True
    return False

def start_wechat_and_screenshot():
    # 获取脚本目录和二维码图片的路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    qr_code_path = os.path.join(script_dir, 'qr_code_screenshot.png')

    # 如果存在旧的二维码图片，则删除它
    if os.path.exists(qr_code_path):
        os.remove(qr_code_path)
        print("Previous QR code screenshot removed.")

    # 等待微信启动并显示二维码
    time.sleep(3)  # 根据实际情况调整等待时间

    # 尝试两次截图
    for _ in range(2):
        # 将微信窗口置顶
        bring_wechat_to_foreground()

        # 捕获整个屏幕截图
        screenshot = ImageGrab.grab()

        # 自动定位二维码位置
        qr_code_region_coords = find_qr_code(screenshot)
        if qr_code_region_coords:
            qr_code_region = screenshot.crop(qr_code_region_coords)

            # 保存截图到同一目录
            script_dir = os.path.dirname(os.path.abspath(__file__))
            qr_code_path = os.path.join(script_dir, 'qr_code_screenshot.png')
            qr_code_region.save(qr_code_path)

            # 将截图保存到剪贴板
            send_to_clipboard(qr_code_region)
            print("QR code screenshot saved and copied to clipboard.")

            # 最小化微信窗口
            minimize_wechat()
            return
        else:
            print("QR code not found. Trying to switch account...")
            if click_switch_account():
                print("Clicked on '切换账号'. Retrying...")
            time.sleep(1)  # 等待一秒再次尝试

def bring_wechat_to_foreground():
    wechat_window = gw.getWindowsWithTitle('微信')[0]  # 假设窗口标题包含"微信"
    if wechat_window:
        wechat_window.activate()
        wechat_window.maximize()  # 最大化微信窗口

def minimize_wechat():
    wechat_window = gw.getWindowsWithTitle('微信')[0]
    if wechat_window:
        wechat_window.minimize()

def find_wechat_path():
    registry_paths = [
        r"SOFTWARE\\WOW6432Node\\Tencent\\WeChat",
        r"SOFTWARE\\Tencent\\WeChat",
        r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\WeChat"
    ]
    for registry_path in registry_paths:
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path)
            path, _ = winreg.QueryValueEx(key, "InstallLocation")
            winreg.CloseKey(key)
            wechat_exe_path = os.path.join(path, "WeChat.exe")
            if os.path.exists(wechat_exe_path):
                return wechat_exe_path
        except FileNotFoundError:
            continue
        except Exception as e:
            print(f"Error finding WeChat path in {registry_path}: {e}")
    return None

# 找到微信路径
wechat_path = find_wechat_path()

if wechat_path:
    # 执行微信启动
    os.system(f'start "" "{wechat_path}"')

    # 执行截图函数
    start_wechat_and_screenshot()

else:
    print("WeChat is not installed or the path could not be found.")
