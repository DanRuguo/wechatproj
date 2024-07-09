import requests
import json
import glob
import re
import os
import tkinter as tk
from tkinter import messagebox

def load_config(config_filename):
    """从JSON文件中加载配置并返回一个字典。"""
    try:
        with open(config_filename, 'r', encoding='utf-8') as file:
            return json.load(file)
    except FileNotFoundError:
        return f"Error: The configuration file '{config_filename}' does not exist."
    except json.JSONDecodeError:
        return "Error: The configuration file is not a valid JSON file."
    except Exception as e:
        return f"Error: An unexpected error occurred while reading the configuration file: {e}"

def load_wechat_results(filename):
    """从文件中加载并返回不合法的微信号列表及其他相关信息。"""
    wechat_data = []
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            for line in file:
                data = json.loads(line.strip())
                if data["verificationCode"] != 1:
                    filtered_data = {
                        "resumeId": data["resumeId"],
                        "commissionerId": data["commissionerId"],
                        "wechat": data["wechat"],
                        "commissionerName": data["commissionerName"]
                    }
                    wechat_data.append(filtered_data)
    except FileNotFoundError:
        return f"Error: The file '{filename}' does not exist."
    except Exception as e:
        return f"Error: An unexpected error occurred while reading the file: {e}"
    return wechat_data

def post_wechat_data(wechat_data_list, config):
    """向API发送微信号信息列表，并处理响应，使用从JSON配置文件加载的参数。"""
    check_params = {
        "userId": config["userId"],
        "token": config["token"],
        "ip": config["ip"],
        "wxRobotVoList": wechat_data_list
    }
    try:
        check_response = requests.post(
            'https://crm.xinlanrenli.com/xsdrec-recruit-api/rec/robot/check',
            json=check_params
        )
        if check_response.status_code == 200:
            check_data = check_response.json()
            if check_data.get('code') == 200:
                return f"成功提交用户ID: {config['userId']}的非微信好友微信号信息，提交条数: {len(wechat_data_list)}"
            else:
                return f"提交失败，错误信息：{check_data.get('msg')}"
        else:
            return f"HTTP请求失败，状态码：{check_response.status_code}"
    except requests.RequestException as e:
        return f"网络请求异常: {e}"

def run_script():
    """主函数，负责整体流程控制，并在GUI中显示结果。"""
    messages = []
    config_files = glob.glob('config_*.json')
    valid_ids = [re.search(r'config_(\d+)\.json', filename).group(1) for filename in config_files if re.search(r'config_(\d+)\.json', filename)]

    for config_id in valid_ids:
        config_result = load_config(f'config_{config_id}.json')
        if isinstance(config_result, str):
            messages.append(config_result)
            continue
        config = config_result

        filename = f'微信号_config_{config_id}.json-结果.txt'
        wechat_data_result = load_wechat_results(filename)
        if isinstance(wechat_data_result, str):
            messages.append(wechat_data_result)
        elif wechat_data_result:
            result = post_wechat_data(wechat_data_result, config)
            messages.append(result)

    messagebox.showinfo("Script Results", "\n".join(messages))

# 创建主窗口
root = tk.Tk()
root.title("WeChat Verification Script")

# 添加按钮和事件
run_button = tk.Button(root, text="Run Script", command=run_script)
run_button.pack(pady=20)

# 启动GUI
root.mainloop()
