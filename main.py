import tkinter as tk
from tkinter import Label, messagebox
from PIL import Image, ImageTk, ImageGrab
import pygetwindow as gw
import pyautogui
import time
import os
import threading
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from comtypes import CoInitialize, CoUninitialize
import numpy as np
import keyboard  # 用于监听F10键
import json  # 用于保存和读取颜色数据

# 全局变量
clicked_position = None
lock_bobber_color = None
window_x, window_y, window_width, window_height = 0, 0, 0, 0
stop_fishing = False  # 控制循环终止
sound_triggered = False  # 用于记录是否检测到巨大的声音
sound_threshold = 0.3  # 设置音量电平阈值为0.3
center_x, center_y = 0, 0  # 截图的中心区域坐标

# 等待时间变量，默认值
wait_after_switch = 1  # 切换窗口后等待时间
wait_after_cast = 3  # 抛竿后等待时间
wait_after_catch = 3  # 收杆后等待时间
wait_after_color_detection = 1  # 等待颜色识别后的时间
max_timer = 30  # 最大等待时间

# 默认颜色
default_color = (72, 42, 42)
color_file = "bobber_color.json"

# 读取或设置默认颜色
def load_bobber_color():
    global lock_bobber_color
    if os.path.exists(color_file):
        with open(color_file, "r") as file:
            try:
                data = json.load(file)
                lock_bobber_color = tuple(data["color"])
            except (json.JSONDecodeError, KeyError, TypeError):
                lock_bobber_color = default_color
    else:
        lock_bobber_color = default_color

def save_bobber_color(color):
    with open(color_file, "w") as file:
        json.dump({"color": color}, file)

# 更新状态标签的函数
def update_status(message):
    status_label.config(text=f"当前状态: {message}")
    root.update()  # 更新GUI

# 捕捉特定应用的系统音量变化并判断是否超过阈值
def listen_for_system_audio(app_name="Wow.exe", duration=40):
    global sound_triggered, stop_fishing
    sound_triggered = False
    start_time = time.time()

    # 初始化COM库
    CoInitialize()

    try:
        while time.time() - start_time < duration:
            if stop_fishing:
                return  # 如果停止钓鱼标志为真，立刻停止音量监听

            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                # 只捕捉魔兽世界的声音
                if session.Process and session.Process.name().lower() == app_name.lower():
                    audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
                    current_volume = audio_meter.GetPeakValue()  # 获取当前音量电平

                    # 如果音量电平超过设置的阈值则触发
                    if current_volume > sound_threshold:
                        sound_triggered = True
                        return
            time.sleep(0.1)
    finally:
        # 释放COM库
        CoUninitialize()

# 获取截图的函数
def pre_cast_setup():
    global screenshot, img, window_x, window_y, window_width, window_height, center_x, center_y

    # 获取所有窗口
    windows = gw.getWindowsWithTitle('魔兽世界')

    if not windows:
        messagebox.showerror("错误", "未找到名称为'魔兽世界'的窗口")
        return

    window = windows[0]
    window.activate()  # 切换到该窗口

    # 获取窗口位置和大小
    window_x, window_y, window_width, window_height = window.left, window.top, window.width, window.height

    # 计算截图的中心区域（300x200）
    center_x = window_x + window_width // 2 - 150  # 300px的中心点位置
    center_y = window_y + window_height // 2 - 100  # 200px的中心点位置
    right_x = center_x + 300
    bottom_y = center_y + 200

    # 截取窗口中心区域（300x200）的截图
    screenshot = ImageGrab.grab(bbox=(center_x, center_y, right_x, bottom_y))

    # 保存截图
    save_path = os.path.join(os.getcwd(), "pre_cast_screenshot.png")
    screenshot.save(save_path)

    # 在GUI上显示截图
    img = ImageTk.PhotoImage(screenshot)
    img_label.config(image=img)
    img_label.image = img

# 颜色比较函数，允许插值误差
def is_color_similar(color1, color2, tolerance=10):
    return all(abs(c1 - c2) <= tolerance for c1, c2 in zip(color1, color2))

# 钓鱼过程循环
def fishing_loop():
    global lock_bobber_color, window_x, window_y, window_width, window_height, stop_fishing, sound_triggered, sound_threshold, center_x, center_y

    if lock_bobber_color is None:
        messagebox.showerror("错误", "尚未识别鱼漂颜色")
        return

    # 获取魔兽世界窗口
    windows = gw.getWindowsWithTitle('魔兽世界')

    if not windows:
        messagebox.showerror("错误", "未找到名称为'魔兽世界'的窗口")
        return

    window = windows[0]

    while not stop_fishing:
        # 切换到魔兽世界窗口
        window.activate()
        update_status("已切换到魔兽世界窗口")
        time.sleep(wait_after_switch)

        if stop_fishing:
            return  # 如果停止钓鱼标志为真，立刻停止

        # 更新状态并等待1秒，按下Z键
        update_status("正在按下Z键抛竿")
        time.sleep(1)
        pyautogui.press('z')

        if stop_fishing:
            return  # 如果停止钓鱼标志为真，立刻停止

        # 抛竿后等待
        update_status("等待抛竿后")
        time.sleep(wait_after_cast)

        # 扫描窗口中心区域（300x200）寻找鱼漂颜色
        update_status("正在识别鱼漂颜色")
        current_screenshot = ImageGrab.grab(bbox=(center_x, center_y, center_x + 300, center_y + 200))
        found = False
        width, height = current_screenshot.size
        for x in range(width):
            for y in range(height):
                if stop_fishing:
                    return  # 如果停止钓鱼标志为真，立刻停止

                pixel_color = current_screenshot.getpixel((x, y))
                if is_color_similar(pixel_color, lock_bobber_color):
                    # 移动鼠标到匹配的颜色位置，并向右下偏移8x8像素
                    pyautogui.moveTo(center_x + x + 8, center_y + y + 8)
                    update_status("找到鱼漂，准备监听水花声音")
                    found = True
                    break
            if found:
                break

        if not found:
            update_status("未找到鱼漂颜色，重新开始")
            continue  # 如果未找到鱼漂，跳过当前循环重新开始

        # 开启声音捕捉系统，监听水花声音
        listen_thread = threading.Thread(target=listen_for_system_audio, args=("Wow.exe", max_timer))
        listen_thread.start()

        # 等待声音触发或超时
        start_time = time.time()
        while time.time() - start_time < max_timer:
            if stop_fishing:
                return  # 如果停止钓鱼标志为真，立刻停止

            if sound_triggered:
                update_status("检测到巨大的声音，准备右键收杆")
                time.sleep(wait_after_color_detection)
                pyautogui.rightClick()  # 收杆
                time.sleep(wait_after_catch)  # 收杆后等待
                break

        # 如果超过时间未检测到声音，重新开始循环
        if time.time() - start_time >= max_timer and not sound_triggered:
            update_status("超过30秒未检测到声音，重新抛竿")
            continue

# 监听停止钓鱼
def stop_fishing_on_keypress():
    global stop_fishing
    stop_fishing = True
    update_status("钓鱼已停止")

# 监听F10按键
def listen_for_f10():
    while True:
        if keyboard.is_pressed("F10"):
            stop_fishing_on_keypress()
            break
        time.sleep(0.1)

# 获取并更新声音阈值
def update_sound_threshold():
    global sound_threshold
    try:
        sound_threshold = float(sound_threshold_entry.get())
        update_status(f"声音阈值已更新为: {sound_threshold}")
    except ValueError:
        messagebox.showerror("错误", "请输入有效的阈值")

# 获取并更新等待时间
def update_wait_times():
    global wait_after_switch, wait_after_cast, wait_after_catch, wait_after_color_detection, max_timer
    try:
        wait_after_switch = float(switch_wait_entry.get())
        wait_after_cast = float(cast_wait_entry.get())
        wait_after_catch = float(catch_wait_entry.get())
        wait_after_color_detection = float(color_detect_wait_entry.get())
        max_timer = float(timer_entry.get())
        update_status("等待时间已更新")
    except ValueError:
        messagebox.showerror("错误", "请输入有效的等待时间")

# 开始钓鱼的函数
def start_fishing():
    global stop_fishing
    stop_fishing = False  # 重置停止标志
    update_status("开始钓鱼")

    # 使用线程启动钓鱼循环，防止阻塞主线程
    fishing_thread = threading.Thread(target=fishing_loop)
    fishing_thread.daemon = True
    fishing_thread.start()

    # 启动F10监听线程
    f10_listener_thread = threading.Thread(target=listen_for_f10)
    f10_listener_thread.daemon = True
    f10_listener_thread.start()

# 创建主窗口
root = tk.Tk()
root.title("阿黄真菜")

# 创建按钮
btn_pre_cast = tk.Button(root, text="预抛竿设置", command=pre_cast_setup)
btn_pre_cast.pack(pady=10)

# 创建开始钓鱼按钮
btn_start_fishing = tk.Button(root, text="开始钓鱼", command=start_fishing)
btn_start_fishing.pack(pady=10)

# 创建停止钓鱼按钮
btn_stop_fishing = tk.Button(root, text="停止钓鱼", command=stop_fishing_on_keypress)
btn_stop_fishing.pack(pady=10)

# 显示截图的标签
img_label = tk.Label(root)
img_label.pack(pady=10)

# 显示鱼漂锁定颜色的标签
bobber_color_label = Label(root, text=f"鱼漂锁定颜色: {lock_bobber_color}")
bobber_color_label.pack(pady=10)

# 显示当前捕捉颜色的区域，初始颜色设置为 (72, 42, 42)
color_display = Label(root, text="捕捉颜色显示", width=20, height=2, bg=f'#{default_color[0]:02x}{default_color[1]:02x}{default_color[2]:02x}')
color_display.pack(pady=5)

# 显示钓鱼状态的标签
status_label = Label(root, text="当前状态: 无")
status_label.pack(pady=10)

# 创建声音阈值输入框和设置按钮
sound_threshold_label = Label(root, text="声音阈值 (默认0.3):")
sound_threshold_label.pack(pady=5)
sound_threshold_entry = tk.Entry(root)
sound_threshold_entry.insert(0, "0.3")  # 默认阈值
sound_threshold_entry.pack(pady=5)
btn_set_threshold = tk.Button(root, text="设置声音阈值", command=update_sound_threshold)
btn_set_threshold.pack(pady=5)

# 创建等待时间的输入框和设置按钮
switch_wait_label = Label(root, text="切换窗口等待时间 (秒):")
switch_wait_label.pack(pady=5)
switch_wait_entry = tk.Entry(root)
switch_wait_entry.insert(0, "1")  # 默认1秒
switch_wait_entry.pack(pady=5)

cast_wait_label = Label(root, text="抛竿后等待时间 (秒):")
cast_wait_label.pack(pady=5)
cast_wait_entry = tk.Entry(root)
cast_wait_entry.insert(0, "3")  # 默认3秒
cast_wait_entry.pack(pady=5)

color_detect_wait_label = Label(root, text="颜色识别后等待时间 (秒):")
color_detect_wait_label.pack(pady=5)
color_detect_wait_entry = tk.Entry(root)
color_detect_wait_entry.insert(0, "1")  # 默认1秒
color_detect_wait_entry.pack(pady=5)

catch_wait_label = Label(root, text="收杆后等待时间 (秒):")
catch_wait_label.pack(pady=5)
catch_wait_entry = tk.Entry(root)
catch_wait_entry.insert(0, "3")  # 默认3秒
catch_wait_entry.pack(pady=5)

timer_label = Label(root, text="最大计时器 (秒):")
timer_label.pack(pady=5)
timer_entry = tk.Entry(root)
timer_entry.insert(0, "30")  # 默认30秒
timer_entry.pack(pady=5)

btn_set_wait_times = tk.Button(root, text="设置等待时间", command=update_wait_times)
btn_set_wait_times.pack(pady=5)

# 加载鱼漂颜色
load_bobber_color()

# 运行GUI主循环
root.mainloop()
