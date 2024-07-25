import http.server
import socketserver
import os
import base64
import logging
import threading

class RequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_AUTHHEAD(self):
        self.send_response(401)
        self.send_header('WWW-Authenticate', 'Basic realm=\"Test\"')
        self.send_header('Content-type', 'text/html')
        self.end_headers()

    def do_GET(self):
        # 基本认证
        if self.headers.get('Authorization') == 'Basic ' + str(base64.b64encode(b'username:password'), 'utf-8'):
            # 判断请求路径
            if self.path == '/run-script':
                try:
                    # 执行截图脚本
                    screenshot_script_path = '..\\screenshot_qr_code.py'
                    file_path = '..\\qr_code_screenshot.png'

                    # 删除已有的图片文件，避免发送旧图片
                    if os.path.exists(file_path):
                        os.remove(file_path)

                    # 运行截图脚本
                    result = os.system(f'python "{screenshot_script_path}"')
                    if result != 0:
                        raise Exception("Screenshot script execution failed.")

                    # 检查图片文件是否存在
                    if not os.path.exists(file_path):
                        raise FileNotFoundError("QR code screenshot file not found.")

                    # 发送图片文件回客户端
                    self.send_response(200)
                    self.send_header('Content-type', 'image/png')
                    self.end_headers()
                    with open(file_path, 'rb') as file:
                        self.wfile.write(file.read())
                    self.wfile.flush()  # 确保图片已发送

                    # 异步执行微信连接脚本
                    wechat_connect_script_path = '..\\wechat_auto_connect.py'
                    threading.Thread(target=lambda: os.system(f'python "{wechat_connect_script_path}"')).start()

                except Exception as e:
                    logging.error(f"Failed to run script or send file: {e}")
                    self.send_error(500, "Internal Server Error")
                return
            else:
                self.send_error(404, "File not found.")
        else:
            self.do_AUTHHEAD()
            self.wfile.write(b'Authentication required.')


# 设置服务器监听端口
PORT = 8001
logging.basicConfig(level=logging.INFO)
with socketserver.TCPServer(("", PORT), RequestHandler) as httpd:
    print("Server at PORT", PORT)
    httpd.serve_forever()