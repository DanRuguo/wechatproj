import requests
import json
import tkinter as tk
from tkinter import messagebox, simpledialog
import socket

def get_ip_address():
    # 获取本机的IPv4地址
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        IP = s.getsockname()[0]
    finally:
        s.close()
    return IP

def send_request(user_params):
    response_messages = []
    for params in user_params:
        try:
            # 发送POST请求
            start_response = requests.post(
                'https://crm.xinlanrenli.com/xsdrec-recruit-api/rec/robot/start',
                json=params
            )

            # 检查响应并保存请求参数到配置文件
            if start_response.status_code == 200:
                response_data = start_response.json()
                if response_data.get('code') == 200:
                    config_filename = f'config_{params["userId"]}.json'  # 为每个用户生成单独的配置文件
                    with open(config_filename, 'w') as config_file:
                        json.dump(params, config_file)
                    response_messages.append(f"Configuration for user {params['userId']} saved successfully.")
                else:
                    response_messages.append(f"Server returned an error for user {params['userId']}: {response_data.get('msg')}")
            else:
                response_messages.append(f"HTTP Error for user {params['userId']}: Received response code {start_response.status_code}")

        except requests.RequestException as e:
            response_messages.append(f"Request failed for user {params['userId']}: {e}")
        except IOError as e:
            response_messages.append(f"File I/O error for user {params['userId']}: {e}")
        except Exception as e:
            response_messages.append(f"An unexpected error occurred for user {params['userId']}: {e}")

    return "\n".join(response_messages)

def add_user():
    # 添加用户配置
    userId = simpledialog.askstring("Input", "Please enter userId:")
    token = simpledialog.askstring("Input", "Please enter token:")
    ip = get_ip_address()  # 自动获取IP地址
    user_params.append({"userId": userId, "token": token, "ip": ip})
    messagebox.showinfo("Info", f"Added User: {userId}, Token: {token}, IP: {ip}")

def run_requests():
    # 运行请求并获取结果
    response_messages = send_request(user_params)
    messagebox.showinfo("Result", response_messages)

# 初始化用户参数列表
user_params = []

# 创建主窗口
root = tk.Tk()
root.title("User Configuration")

# 添加按钮和事件
add_button = tk.Button(root, text="Add User", command=add_user)
add_button.pack(pady=20)

run_button = tk.Button(root, text="Run Requests", command=run_requests)
run_button.pack(pady=20)

# 启动GUI
root.mainloop()
