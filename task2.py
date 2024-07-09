from wxauto import WeChat
import requests
import json
import os
import tkinter as tk
from tkinter import filedialog
import pyautogui
import tkinter.scrolledtext as scrolledtext

# 全局变量用于存储最近处理的微信号文件名
processed_files = []

def setup_ui(root):
    text_area = scrolledtext.ScrolledText(root, width=100, height=20)
    text_area.pack(pady=20)
    return text_area

def load_and_process_files(text_area):
    filenames = filedialog.askopenfilenames(
        title="Select configuration files",
        filetypes=(("JSON files", "*.json"), ("All files", "*.*"))
    )
    if filenames:
        result = process_config_files(filenames)
        text_area.insert(tk.END, result + "\n")

def process_config_files(config_files):
    response_messages = []
    global processed_files
    processed_files = []  # 清空之前的记录

    for config_file in config_files:
        try:
            with open(config_file, 'r') as file:
                valid_params = json.load(file)
            valid_response = requests.post(
                'https://crm.xinlanrenli.com/xsdrec-recruit-api/rec/robot/wxValid',
                json=valid_params
            )
            if valid_response.status_code == 200:
                valid_data = valid_response.json()
                if valid_data.get('code') == 200:
                    if 'rows' in valid_data:
                        filename = f'微信号_{os.path.basename(config_file)}.txt'
                        with open(filename, 'w', encoding='utf-8') as file:
                            for row in valid_data['rows']:
                                entry = json.dumps(row, ensure_ascii=False)
                                file.write(entry + '\n')
                        processed_files.append(filename)
                        response_messages.append(f"File {filename} has been created.")
                    else:
                        response_messages.append(f"Error: 'rows' field is missing in the response for {config_file}")
                else:
                    response_messages.append(f"Error from {config_file}: {valid_data.get('msg')}")
            else:
                response_messages.append(f"HTTP Error from {config_file}: Received response code {valid_response.status_code}")
        except FileNotFoundError as e:
            response_messages.append(f"File not found: {e}")
        except requests.RequestException as e:
            response_messages.append(f"Request failed: {e}")
        except Exception as e:
            response_messages.append(f"An error occurred: {e}")

    if processed_files:
        response_messages.append(f"已成功加载配置文件：{', '.join(processed_files)}的微信号数据。")
    return "\n".join(response_messages)

def search_wechat_ids(text_area):
    wx = WeChat()
    all_results = []

    for wechat_file in processed_files:
        text_area.insert(tk.END, f"Processing file: {wechat_file}\n")
        with open(wechat_file, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            results = []

            for line in lines:
                try:
                    data = json.loads(line.strip())
                    who = data.get('wechat', '')  # 获取微信号
                    if who is None:  # 如果获取到的微信号为None，使用空字符串
                        who = ''
                    who = who.strip()  # 确保获取到的微信号已经去除两端空格

                    if not who or who.lower() in ["null", "none"]:  # 检查微信号是否为空或无效
                        text_area.insert(tk.END, f"无效微信号跳过: {who}\n")
                        continue

                    chatname = wx.ChatWith(who)
                    code = ''
                    message = ''

                    if isinstance(chatname, bool) and not chatname:
                        message = f"查无此友: '{who}'"
                        code = 0
                    elif chatname == who:
                        message = f"查询到备注为'{who}'的好友"
                        code = 3
                    elif '微信号:' in chatname:
                        wechat_id = chatname.replace('<em>', '').replace('</em>', '').split('微信号:')[-1].strip()
                        if wechat_id == who:
                            message = f"查询到微信号为'{who}'的好友"
                            code = 1
                        elif who in wechat_id:
                            message = f"查询到微信号包含'{who}'的好友"
                            code = 4
                    elif '昵称:' in chatname:
                        nickname = chatname.replace('<em>', '').replace('</em>', '').split('昵称:')[-1].strip()
                        if nickname == who:
                            message = f"查询到昵称为'{who}'的好友"
                            code = 2
                        elif who in nickname:
                            message = f"查询到昵称包含'{who}'的好友"
                            code = 4
                    else:
                        message = f"查询到备注包含'{who}'的好友"
                        code = 4

                    text_area.insert(tk.END, message + "\n")

                    updated_line = data
                    updated_line['verificationCode'] = code
                    results.append(json.dumps(updated_line, ensure_ascii=False))

                except json.JSONDecodeError as e:
                    text_area.insert(tk.END, f"JSON Decode Error: {e}\n")
                    continue

            result_file_name = wechat_file.replace('.txt', '-结果.txt')
            with open(result_file_name, 'w', encoding='utf-8') as outputFile:
                outputFile.write('\n'.join(results))
            all_results.extend(results)

    text_area.insert(tk.END, "所有微信号处理完毕。\n")
    pyautogui.press('esc')

root = tk.Tk()
root.title("微信配置与验证")

text_area = setup_ui(root)

select_button = tk.Button(root, text="Load and Process Files", command=lambda: load_and_process_files(text_area))
select_button.pack(pady=20)

search_button = tk.Button(root, text="Search WeChat IDs", command=lambda: search_wechat_ids(text_area))
search_button.pack(pady=20)

root.mainloop()