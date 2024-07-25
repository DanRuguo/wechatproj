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
from threading import Timer
import pytz
from datetime import datetime, timedelta
from PIL import Image, ImageTk, ImageFont
import tkinter.font as tkFont

# 设置东八区时区
tz = pytz.timezone('Asia/Shanghai')

# 全局变量用于存储最近处理的微信号文件名和结果文件名
processed_files = []
result_files = []
user_params = []
timers = []
next_run_time_label = None
next_next_run_time_label = None

def on_closing():
    """在关闭窗口时清理资源并退出程序"""
    for timer in timers:
        timer.cancel()
    root.destroy()
    sys.exit()

def load_time_settings():
    global timers, user_params
    # 清理旧定时器
    for timer in timers:
        timer.cancel()
    timers.clear()

    load_users_config()  # 加载用户配置
    try:
        for user in user_params:
            timetype = user['timetype']
            times = user['times']

            if timetype == 1:  # 周期性定时任务
                interval = int(times) * 60
                if interval <= 0:
                    raise ValueError("时间间隔必须为正整数")
                # 设置周期性定时器，这里直接在定时器中使用具体的interval值
                timer = Timer(interval, scheduled_recurring, [interval])
                timers.append(timer)
                timer.start()
                update_run_times(interval, interval)

            elif timetype == 2:  # 固定时刻执行的任务
                schedule_fixed_times(times)

            elif timetype == 0:
                # 无定时任务，仅在启动时执行一次
                continue

        # 立即执行一次任务
        scheduled_operations()
    except Exception as e:
        messagebox.showerror("错误", f"读取定时设置时发生错误：{e}")

def schedule_fixed_times(times):
    global timers
    """安排一天中多个固定时刻的任务"""
    times_list = sorted(set(times.split(';')))
    now = datetime.now(tz)
    next_run_time = None
    next_next_run_time = None

    upcoming_times = []

    for time_str in times_list:
        hour, minute = map(int, time_str.split(':'))
        target_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        upcoming_times.append(target_time)

    upcoming_times = [time for time in upcoming_times if time >= now]

    if len(upcoming_times) == 0:
        next_run_time = now.replace(hour=int(times_list[0].split(':')[0]),
                                    minute=int(times_list[0].split(':')[1]),
                                    second=0, microsecond=0) + timedelta(days=1)
        if len(times_list) == 1:
            next_next_run_time = next_run_time + timedelta(days=1)
        else:
            next_next_run_time = now.replace(hour=int(times_list[1].split(':')[0]),
                                            minute=int(times_list[1].split(':')[1]),
                                            second=0, microsecond=0) + timedelta(days=1)
    elif len(upcoming_times) == 1:
        next_run_time = upcoming_times[0]
        next_next_run_time = now.replace(hour=int(times_list[0].split(':')[0]),
                                         minute=int(times_list[0].split(':')[1]),
                                         second=0, microsecond=0) + timedelta(days=1)
    else:
        next_run_time = upcoming_times[0]
        next_next_run_time = upcoming_times[1]

    delay = (next_run_time - now).total_seconds()
    timer = Timer(delay, scheduled_operations_fixed, [times])
    timers.append(timer)
    timer.start()

    update_run_times((next_run_time - now).total_seconds(), (next_next_run_time - now).total_seconds())

def scheduled_operations_fixed(times):
    """执行固定时刻的操作，并重新设置定时器"""
    try:
        scheduled_operations()
    except Exception as e:
        log_error(f"执行固定时刻的操作时发生错误: {e}")
    finally:
        schedule_fixed_times(times)

def scheduled_recurring(interval):
    global timers
    """执行周期性任务，并重新设置定时器"""
    try:
        scheduled_operations()
    except Exception as e:
        log_error(f"执行周期性任务时发生错误: {e}")
    finally:
        if interval > 0:
            timer = Timer(interval, scheduled_recurring, [interval])
            timers.append(timer)
            timer.start()
            update_run_times(interval, interval * 2)
        else:
            log_error("时间间隔必须为正整数，无法设置周期性任务。")

def update_run_times(next_interval, next_next_interval):
    """更新下次运行和下下次运行的时间标签"""
    next_run_time = datetime.now(tz) + timedelta(seconds=next_interval)
    next_next_run_time = datetime.now(tz) + timedelta(seconds=next_next_interval)
    next_run_time_label.config(text=f"下次运行时间: {next_run_time.strftime('%Y-%m-%d %H:%M:%S %Z%z')}")
    next_next_run_time_label.config(text=f"下下次运行时间: {next_next_run_time.strftime('%Y-%m-%d %H:%M:%S %Z%z')}")

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
    global user_params
    try:
        with open('.config', 'r', encoding='utf-8') as file:
            config_data = json.load(file)

        new_user = {
            "userId": config_data["userid"],
            "token": config_data["token"],
            "url": config_data["url"],
            "ip": get_ip_address(),
            "timetype": config_data["timetype"],
            "times": config_data["times"]
        }

        # 删除旧的相同 userId 数据
        user_params = [user for user in user_params if user["userId"] != new_user["userId"]]

        # 添加新的用户数据
        user_params.append(new_user)
    except json.JSONDecodeError as e:
        messagebox.showerror("配置错误", "配置文件格式不正确: " + str(e))
    except FileNotFoundError:
        messagebox.showerror("配置错误", "找不到配置文件.config")
    except Exception as e:
        messagebox.showerror("配置错误", "无法读取配置文件: " + str(e))

def send_request(user_params):
    response_messages = []
    for params in user_params:
        try:
            # 构建完整的URL
            request_url = f"{params['url']}/rec/robot/start"
            # 发送POST请求
            start_response = requests.post(
                request_url,
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
    global user_params
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

def load_font(font_path, size, weight=tkFont.NORMAL):
    """加载字体"""
    pil_font = ImageFont.truetype(resource_path(font_path), size)
    return tkFont.Font(family=pil_font.getname(), size=size, weight=weight)

def setup_ui(root):
    global next_run_time_label, next_next_run_time_label

    icon_path = resource_path("favicon.ico")
    root.iconbitmap(icon_path)

    # 使用 Label 来显示背景图片
    bg_image_path = resource_path("background.jpg")
    bg_image = Image.open(bg_image_path)
    bg_image = bg_image.resize((200, 150), Image.LANCZOS)  # 使用 LANCZOS 代替 ANTIALIAS
    bg_image = ImageTk.PhotoImage(bg_image)

    bg_label = Label(root, image=bg_image)
    bg_label.image = bg_image  # 保持对图片的引用
    bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    # 动态加载字体
    stzhongs_font_b = load_font("STZHONGS.TTF", 14, tkFont.BOLD)
    stzhongs_font = load_font("STZHONGS.TTF", 12)
    simhei_font_b = load_font("simhei.ttf", 16, tkFont.BOLD)
    msyh_font = load_font("msyh.ttc", 10)

    description_text = "本程序用于自动化处理微信号验证。程序启动后，会自动加载配置文件，连接服务器，并定时执行任务。"
    desc_label = Label(root, text=description_text, wraplength=400, justify="left", font=stzhongs_font)
    desc_label.pack(pady=(16, 0))

    # 创建一个frame用于居中text_area
    frame = tk.Frame(root)
    frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

    text_area = scrolledtext.ScrolledText(frame, width=70, height=15, font=msyh_font)
    text_area.pack(fill=tk.BOTH, expand=True, padx=(frame.winfo_width() * 0.05, 0), pady=20, anchor='w')  # 只留左边5%的空白

    # 定义tags
    text_area.tag_configure("even", background="#f0f0f0")
    text_area.tag_configure("odd", background="#ffffff")

    start_button = Button(root, text="启动", command=lambda: load_time_settings(), width=20, bg='#4CAF50', fg='white', font=simhei_font_b)
    start_button.pack(pady=20)

    # 显示下次运行和下下次运行的时间标签
    next_run_time_label = Label(root, text="下次运行时间: 未知", font=stzhongs_font_b)
    next_run_time_label.pack(pady=(0, 10))

    next_next_run_time_label = Label(root, text="下下次运行时间: 未知", font=stzhongs_font)
    next_next_run_time_label.pack(pady=(0, 20))

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
            # 构建完整的URL
            valid_url = f"{valid_params['url']}/rec/robot/wxValid"
            valid_response = requests.post(
                valid_url,
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
                        response_messages.append(f"已经创建 {filename} 文件.")
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
        insert_text(text_area, f"正在检索专员微信，昵称为：{nickname}")
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
                        # insert_text(text_area, f"无效微信号跳过: {who}")
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
                    elif '<em>' in chatname and '</em>' in chatname:
                        nickname_in_group = chatname.replace('<em>', '').replace('</em>', '').split(':')[-1].strip()
                        message = f"查询到群中有昵称'{nickname_in_group}'，但此人不是你的好友"
                        code = 5
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
    all_amount = 0
    commissioner_name = ""
    original_data = []
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            for line in file:
                data = json.loads(line.strip())
                original_data.append(data)
                all_amount += 1
                if data["verificationCode"] != 1:
                    filtered_data = {
                        "resumeId": data["resumeId"],
                        "commissionerId": data["commissionerId"],
                        "wechat": data["wechat"],
                        "commissionerName": data["commissionerName"]
                    }
                    wechat_data.append(filtered_data)
                    if not commissioner_name:
                        commissioner_name = data["commissionerName"]
    except FileNotFoundError:
        return f"Error: The file '{filename}' does not exist."
    except Exception as e:
        return f"Error: An unexpected error occurred while reading the file: {e}"
    return wechat_data, commissioner_name, all_amount, original_data

def post_wechat_data(wechat_data_list, config, user_name, all_amount):
    """向API发送微信号信息列表，并处理响应，使用从JSON配置文件加载的参数。"""
    check_params = {
        "userId": config["userId"],
        "userName": user_name,  # 添加 userName
        "token": config["token"],
        "ip": config["ip"],
        "wxRobotVoList": wechat_data_list,
        "allAmount": all_amount  # 添加 allAmount
    }
    # print(check_params)

    try:
        # 构建完整的URL
        check_url = f"{config['url']}/rec/robot/check"
        check_response = requests.post(
            check_url,
            json=check_params
        )
        if check_response.status_code == 200:
            check_data = check_response.json()
            if check_data.get('code') == 200:
                return f"专员ID: {config['userId']}的非微信好友微信号信息成功提交，提交条数: {len(wechat_data_list)}"
            else:
                return f"非微信好友微信号信息提交失败，错误信息：{check_data.get('msg')}"
        else:
            return f"HTTP请求失败，状态码：{check_response.status_code}"
    except requests.RequestException as e:
        return f"网络请求异常: {e}"

def post_right_wechat_data(config, user_id, token, resume_ids):
    """向指定API发送满足条件的微信号信息。"""
    url = f"{config['url']}/rec/robot/right"
    data = {
        "userId": user_id,
        "token": token,
        "wxRobotVoList": [{"resumeId": rid} for rid in resume_ids]
    }
    # print(data)

    try:
        response = requests.post(url, json=data)
        if response.status_code == 200:
            result = response.json()
            if result.get('code') == 200:
                return f"专员ID: {config['userId']}的所有已经验证为正确添加微信好友的数据的简历ID提交成功，提交条数: {len(resume_ids)}"
            else:
                return f"简历ID提交失败，错误信息：{result.get('msg')}"
        else:
            return f"HTTP请求失败，状态码：{response.status_code}"
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
        wechat_data_result, commissioner_name, all_amount, original_data = load_wechat_results(result_file)
        if isinstance(wechat_data_result, str):
            messages.append(wechat_data_result)
        elif wechat_data_result:
            config_id = re.search(r'微信号_config_(\d+)\.json-结果\.txt', result_file).group(1)
            config = load_config(f'config_{config_id}.json')
            if config:
                result = post_wechat_data(wechat_data_result, config, commissioner_name, all_amount)
                messages.append(result)

                # 筛选 verificationCode 为 1, 2, 3 的 resumeId
                resume_ids = [data["resumeId"] for data in original_data if data["verificationCode"] in [1, 2, 3]]
                # print(resume_ids)
                if resume_ids:
                    right_result = post_right_wechat_data(config, config["userId"], config["token"], resume_ids)
                    messages.append(right_result)
                else:
                    messages.append("本次查询未发现符合要求的微信好友")

    insert_multiline_text(text_area, "提交结果：\n" + "\n".join(messages))

def scheduled_operations():
    """执行定期安排的操作集合，并在结束时清空全局变量和保存日志"""
    global processed_files, result_files
    try:
        text_area.delete('1.0', tk.END)  # 清空文本区
        run_requests()  # 发送请求并处理结果
        auto_load_and_process_files()  # 自动加载并处理文件
        search_wechat_ids(text_area)  # 搜索微信ID
        run_script()  # 运行脚本并处理结果
        current_time = datetime.now(tz)
        text_to_insert = f"本次任务在: {current_time}结束"
        insert_multiline_text(text_area, text_to_insert)
    finally:
        save_log(text_area.get("1.0", "end-1c"))  # 保存日志
        processed_files = []
        result_files = []
        user_params = []

def save_log(content):
    """将当前的text_area内容保存到日志文件"""
    now = datetime.now(tz).strftime("%Y-%m-%d_%H-%M")
    log_filename = f'logs/log_{now}.txt'
    os.makedirs(os.path.dirname(log_filename), exist_ok=True)
    with open(log_filename, 'w', encoding='utf-8') as file:
        file.write(content)

def log_error(message):
    """记录错误到日志文件中"""
    now = datetime.now(tz).strftime("%Y-%m-%d_%H-%M-%S")
    log_filename = f'logs/error_log_{now}.txt'
    os.makedirs(os.path.dirname(log_filename), exist_ok=True)
    with open(log_filename, 'a', encoding='utf-8') as log_file:
        log_file.write(f"{now} - ERROR - {message}\n")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        root.title("CRM微信端开发")
        text_area = setup_ui(root)
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
    except Exception as e:
        log_error("程序启动失败: " + str(e))
        raise e  # 可以选择重新抛出异常或者处理异常