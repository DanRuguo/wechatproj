from wxauto import WeChat
import requests
import json
import os
import tkinter as tk
import pyautogui
import tkinter.scrolledtext as scrolledtext
import re
from tkinter import messagebox, Button, Label
import socket
import io
import sys
import datetime
from threading import Timer

# 全局变量用于存储最近处理的微信号文件名和结果文件名
processed_files = []
result_files = []
user_params = []
timers = []

def on_closing():
    """在关闭窗口时清理资源并退出程序"""
    for timer in timers:
        timer.cancel()
    root.destroy()
    sys.exit()

def load_time_settings():
    try:
        with open('.timeset', 'r', encoding='utf-8') as file:
            lines = file.readlines()
        hourly_schedule = None
        specific_time_schedule = None

        for line in lines:
            # 检查是否为注释行
            if line.strip().startswith('//'):
                continue

            # 匹配每隔【】小时运行一次的设置
            match_hourly = re.search(r"每隔\【(\d+\.?\d*)\】小时运行一次", line)
            if match_hourly:
                hourly = float(match_hourly.group(1))
                if 0 <= hourly <= 24:
                    hourly_schedule = hourly
                else:
                    messagebox.showerror("错误", "小时数必须在0到24之间")
                    return

            # 匹配每天【】时刻开始运行一次的设置
            match_specific = re.search(r"每天\【(\d+\.?\d*)\】时刻开始运行一次", line)
            if match_specific:
                specific = float(match_specific.group(1))
                if 0 <= specific <= 24:
                    specific_time_schedule = specific
                else:
                    messagebox.showerror("错误", "时刻必须在0到24之间")
                    return

        # 立即运行一次
        scheduled_operations()

        # 设定任务
        if hourly_schedule is not None:
            schedule_task(hourly_schedule, None)
        elif specific_time_schedule is not None:
            schedule_task(None, specific_time_schedule)

    except FileNotFoundError:
        messagebox.showinfo("错误", "找不到配置文件.timeset")
    except Exception as e:
        messagebox.showinfo("错误", f"读取配置文件.timeset时发生错误：{e}")

def schedule_task(hourly, specific):
    global timers

    def run_periodic():
        """周期性执行任务"""
        try:
            scheduled_operations()
        finally:
            # 无论任务成功与否，都重新设置定时器
            if hourly is not None:
                timer = Timer(hourly * 3600, run_periodic)
                timers.append(timer)
                timer.start()

    def run_daily():
        """在特定时刻执行任务"""
        try:
            scheduled_operations()
        finally:
            # 设置次日同一时刻执行任务
            set_daily_timer(specific)

    def set_daily_timer(specific_time):
        now = datetime.datetime.now()
        target_time = now.replace(hour=int(specific_time), minute=int((specific_time - int(specific_time)) * 60), second=0, microsecond=0)
        if now >= target_time:
            target_time += datetime.timedelta(days=1)
        delay = (target_time - now).total_seconds()
        timer = Timer(delay, run_daily)
        timers.append(timer)
        timer.start()

    # 根据设置选择任务
    if hourly is not None:
        run_periodic()
    elif specific is not None:
        set_daily_timer(specific)
    else:
        # 如果没有有效的设置，程序在1分钟后退出
        timer = Timer(60, sys.exit)
        timers.append(timer)
        timer.start()

def get_ip_address():
    # 获取本机的IPv4地址
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        IP = s.getsockname()[0]
    finally:
        s.close()
    return IP

def load_users_config():
    # 从.config文件读取专员信息
    try:
        with open('.config', 'r', encoding='utf-8') as file:
            for line in file:
                parts = line.strip().split()
                userId = parts[0]
                token = ' '.join(parts[1:3])  # 合并日期和时间部分作为 token
                ip = get_ip_address()
                user_params.append({"userId": userId, "token": token, "ip": ip})
    except FileNotFoundError:
        messagebox.showinfo("错误", "找不到配置文件.config")
    except Exception as e:
        messagebox.showinfo("错误", f"读取配置文件.config时发生错误：{e}")

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
                    response_messages.append(f"专员 {params['userId']} 的配置保存成功。")
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

def run_requests():
    # 检查 user_params 是否为空
    if not user_params:
        messagebox.showinfo("结果", "专员信息为空", parent=root)
        return

    # 清空上一次的结果内容
    text_area.delete('1.0', tk.END)

    # 运行请求并获取结果
    response_messages = send_request(user_params)

    # 检查每一条消息是否包含"配置保存成功"
    all_success = True
    for message in response_messages.split("\n"):
        if "配置保存成功" not in message:
            all_success = False
            break

    # 检查是否存在所有配置文件
    for param in user_params:
        config_filename = f'config_{param["userId"]}.json'
        if not os.path.isfile(config_filename):
            all_success = False
            break

    # 如果所有消息都表示成功且所有配置文件存在，不显示消息框
    if not all_success:
        # 如果有任何错误或者文件缺失
        messagebox.showinfo("结果", response_messages, parent=root)

def resource_path(relative_path):
    """获取资源文件的路径"""
    try:
        # PyInstaller 创建临时文件时
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def setup_ui(root):
    icon_path = resource_path("favicon.ico")
    root.iconbitmap(icon_path)
    description_text = "本程序用于自动化处理微信号验证。程序启动后，会自动加载配置文件，连接服务器，并定时执行任务。"
    desc_label = Label(root, text=description_text, wraplength=400, justify="left", font=("华文中宋", 12))
    desc_label.pack(pady=(10, 0))

    # 创建一个frame用于居中text_area
    frame = tk.Frame(root)
    frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

    text_area = scrolledtext.ScrolledText(frame, width=100, height=20, font=("微软雅黑", 10))
    text_area.pack(fill=tk.BOTH, expand=True, padx=(frame.winfo_width() * 0.05, 0), pady=20, anchor='w')  # 只留左边5%的空白

    # 定义tags
    text_area.tag_configure("even", background="#f0f0f0")
    text_area.tag_configure("odd", background="#ffffff")

    start_button = Button(root, text="启动", command=lambda: load_time_settings(), width=20, bg='#4CAF50', fg='white', font=("黑体", 12, "bold"))
    start_button.pack(pady=20)

    return text_area

def insert_text(text_area, text):
    """插入带有明暗交替效果的文本"""
    current_line = int(text_area.index('end-1c').split('.')[0])
    tag = 'even' if current_line % 2 == 0 else 'odd'
    text_area.insert(tk.END, text, tag)
    text_area.insert(tk.END, '\n')

def insert_multiline_text(text_area, text):
    for line in text.split('\n'):
        insert_text(text_area, line)

def auto_load_and_process_files():
    # 自动从已保存的配置文件加载
    config_files = [f'config_{param["userId"]}.json' for param in user_params if os.path.exists(f'config_{param["userId"]}.json')]
    if config_files:
        result = process_config_files(config_files)
        insert_multiline_text(text_area, result)
    else:
        insert_text(text_area, "没有找到任何配置文件。")

def process_config_files(config_files):
    response_messages = []
    global processed_files
    processed_files = []  # 清空之前的记录

    for config_file in config_files:
        try:
            with open(config_file, 'r', encoding='utf-8') as file:
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
    # 重定向标准输出以捕获WeChat初始化的输出
    old_stdout = sys.stdout
    sys.stdout = mystdout = io.StringIO()

    try:
        wx = WeChat()
    finally:
        # 恢复标准输出
        sys.stdout = old_stdout

    # 获取初始化输出内容
    init_output = mystdout.getvalue()

    # 从初始化输出中筛选微信昵称
    nickname = None
    match = re.search(r"初始化成功，获取到已登录窗口：(.+)", init_output)
    if match:
        nickname = match.group(1)
        insert_text(text_area, f"正在检索微信昵称：{nickname}")
    else:
        messagebox.showerror("错误", "无法获取微信昵称")

    all_results = []
    global result_files
    result_files = []  # 清空之前的记录

    for wechat_file in processed_files:
        insert_text(text_area, f"正在处理: {wechat_file}文件")
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
                        insert_text(text_area, f"无效微信号跳过: {who}")
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

                    insert_text(text_area, message)

                    updated_line = data
                    updated_line['verificationCode'] = code
                    results.append(json.dumps(updated_line, ensure_ascii=False))

                except json.JSONDecodeError as e:
                    text_area.insert(tk.END, f"JSON Decode Error: {e}\n")
                    continue

            result_file_name = wechat_file.replace('.txt', '-结果.txt')
            with open(result_file_name, 'w', encoding='utf-8') as outputFile:
                outputFile.write('\n'.join(results))
            result_files.append(result_file_name)  # 将结果文件名存储到全局变量
            all_results.extend(results)

    insert_text(text_area, "所有微信号处理完毕。")
    pyautogui.press('esc')

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

def load_config(config_filename):
    """从JSON文件中加载配置并返回一个字典。"""
    try:
        with open(config_filename, 'r', encoding='utf-8') as file:
            return json.load(file)
    except FileNotFoundError:
        print(f"Error: The configuration file '{config_filename}' does not exist.")
        return None
    except json.JSONDecodeError:
        print("Error: The configuration file is not a valid JSON file.")
        return None
    except Exception as e:
        print(f"Error: An unexpected error occurred while reading the configuration file: {e}")
        return None

def run_script():
    """主函数，负责整体流程控制，并在GUI中显示结果。"""
    if not result_files:
        messagebox.showinfo("出错", "结果文件不存在")
        return

    messages = []
    for result_file in result_files:
        wechat_data_result = load_wechat_results(result_file)
        if isinstance(wechat_data_result, str):
            messages.append(wechat_data_result)
        elif wechat_data_result:
            config_id = re.search(r'微信号_config_(\d+)\.json-结果\.txt', result_file).group(1)
            config = load_config(f'config_{config_id}.json')
            if config:
                result = post_wechat_data(wechat_data_result, config)
                messages.append(result)

    insert_multiline_text(text_area, "提交结果：\n" + "\n".join(messages))

def scheduled_operations():
    """执行定期安排的操作集合，并在结束时清空全局变量和保存日志"""
    global processed_files, result_files, user_params
    try:
        text_area.delete('1.0', tk.END)  # 清空文本区
        load_users_config()  # 加载用户配置
        run_requests()  # 发送请求并处理结果
        auto_load_and_process_files()  # 自动加载并处理文件
        search_wechat_ids(text_area)  # 搜索微信ID
        run_script()  # 运行脚本并处理结果
    finally:
        save_log(text_area.get("1.0", "end-1c"))  # 保存日志
        processed_files = []
        result_files = []
        user_params = []

def save_log(content):
    """将当前的text_area内容保存到日志文件"""
    now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    log_filename = f'logs/log_{now}.txt'
    os.makedirs(os.path.dirname(log_filename), exist_ok=True)
    with open(log_filename, 'w', encoding='utf-8') as file:
        file.write(content)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("CRM微信开发")
    root.protocol("WM_DELETE_WINDOW", on_closing)  # 捕获关闭窗口事件
    text_area = setup_ui(root)
    root.mainloop()
