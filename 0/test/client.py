import requests
from requests.auth import HTTPBasicAuth

# 电脑B的IP地址和端口
url = 'http://192.168.1.16:8001/run-script'

try:
    # 发送带有基本认证的GET请求执行脚本并获取图片
    response = requests.get(url, auth=HTTPBasicAuth('username', 'password'))

    # 检查是否成功获取到图片
    if response.status_code == 200:
        # 保存图片
        with open('received_qr_code_screenshot.png', 'wb') as file:
            file.write(response.content)
        print('Image saved successfully.')
    else:
        print(f'Failed to retrieve the image. Status code: {response.status_code}')
except requests.exceptions.RequestException as e:
    print(f'HTTP Request failed: {e}')
