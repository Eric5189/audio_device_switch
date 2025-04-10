import sys
import ctypes
import subprocess
import json
import os
import pystray
import keyboard
import threading
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

# 隐藏控制台窗口
if sys.platform == 'win32':
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    
    def get_startup_info():
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        return startupinfo
else:
    def get_startup_info():
        return None

# 配置文件路径
CONFIG_FILE = "config.json"

class HotkeyManager:
    def __init__(self):
        self.hotkey = None
        self.running = False
        self.listener = None
        self.callback = None

    def start_listener(self):
        def listen():
            while self.running:
                keyboard.wait()
        self.running = True
        threading.Thread(target=listen, daemon=True).start()

    def register_hotkey(self, hotkey, callback):
        if self.listener:
            keyboard.remove_hotkey(self.listener)
        try:
            self.listener = keyboard.add_hotkey(hotkey, callback)
            self.hotkey = hotkey
            self.callback = callback
            return True
        except Exception as e:
            print(f"注册失败: {str(e)}")
            return False

    def reload_hotkey(self, new_hotkey):
        if self.callback:
            return self.register_hotkey(new_hotkey, self.callback)
        return False


def ensure_audio_module_installed():
    # 检查是否已经安装
    check_command = [
        "powershell", "-Command",
        "Get-Module -ListAvailable -Name AudioDeviceCmdlets"
    ]
    result = subprocess.run(check_command, capture_output=True, text=True)
    if "AudioDeviceCmdlets" not in result.stdout:
        print("AudioDeviceCmdlets 模块未安装，正在尝试安装...")

        # 设置执行策略（某些系统需要）
        subprocess.run([
            "powershell", "-Command",
            "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force"
        ], creationflags=subprocess.CREATE_NO_WINDOW)

        # 安装模块
        install_command = [
            "powershell", "-Command",
            "Install-Module -Name AudioDeviceCmdlets -Scope CurrentUser -Force"
        ]
        try:
            subprocess.run(install_command, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            print("AudioDeviceCmdlets 安装完成。")
        except subprocess.CalledProcessError as e:
            print("安装 AudioDeviceCmdlets 失败：", e)

def load_config():
    default_config = {
        "hotkey": "ctrl+alt+s",
        "DEVICE1": "",
        "DEVICE2": "",
        "DEVICE1_ID": "",
        "DEVICE2_ID": ""
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return {**default_config, **json.load(f)}
        except Exception:
            return default_config
    return default_config

def save_config(config):
    keys_to_save = ["hotkey", "DEVICE1", "DEVICE2", "DEVICE1_ID", "DEVICE2_ID"]
    filtered_config = {k: config.get(k, "") for k in keys_to_save}
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(filtered_config, f, indent=4, ensure_ascii=False)
    except Exception:
        pass

def get_all_audio_devices():
    ps_command = (
        "Import-Module AudioDeviceCmdlets; "
        "Get-AudioDevice -List | Where-Object { $_.Type -eq 'Playback' } "
        "| Select-Object Name, ID | ConvertTo-Json"
    )
    try:
        result = subprocess.check_output(
            ['powershell', '-Command', ps_command],
            stderr=subprocess.STDOUT,
            startupinfo=get_startup_info()
        )
        return json.loads(result.strip().decode('gbk'))
    except Exception:
        return []

def get_current_audio_device_id():
    ps_command = "Import-Module AudioDeviceCmdlets; (Get-AudioDevice -Playback).ID"
    try:
        result = subprocess.check_output(
            ['powershell', '-Command', ps_command],
            stderr=subprocess.STDOUT,
            startupinfo=get_startup_info()
        )
        return result.strip().decode('gbk')
    except Exception:
        return None

def switch_audio_device():
    config = load_config()
    device1_id = config.get("DEVICE1_ID")
    device2_id = config.get("DEVICE2_ID")
    
    if not device1_id or not device2_id:
        return

    current_id = get_current_audio_device_id()
    target_id = device2_id if current_id == device1_id else device1_id
    ps_command = f"Import-Module AudioDeviceCmdlets; Set-AudioDevice -ID '{target_id}'"
    
    try:
        subprocess.run(
            ['powershell', '-Command', ps_command],
            check=True,
            startupinfo=get_startup_info()
        )
    except Exception:
        pass

def get_device_name(device_id):
    for d in get_all_audio_devices():
        if d['ID'] == device_id:
            return d['Name']
    return "未知设备"

def select_devices(hotkey_mgr=None):
    devices = get_all_audio_devices()
    if not devices:
        messagebox.showerror("错误", "无法获取音频设备列表")
        return

    root = tk.Tk()
    root.title("设备设置")
    root.geometry("475x300")
    
    config = load_config()
    current_device1 = config.get("DEVICE1", "")
    current_device2 = config.get("DEVICE2", "")
    current_hotkey = config.get("hotkey", "ctrl+alt+s")
    
    id_mapping = {d['Name']: d['ID'] for d in devices}
    device_names = [d['Name'] for d in devices]
    
    # 设备选择部分
    ttk.Label(root, text="设备1:").pack(pady=5)
    combo1 = ttk.Combobox(root, values=device_names, state="readonly", width=60)
    combo1.pack(pady=5)
    
    ttk.Label(root, text="设备2:").pack(pady=5)
    combo2 = ttk.Combobox(root, values=device_names, state="readonly", width=60)
    combo2.pack(pady=5)

    # 设置默认选中项
    try:
        if current_device1 in device_names:
            combo1.current(device_names.index(current_device1))
        if current_device2 in device_names:
            combo2.current(device_names.index(current_device2))
    except ValueError:
        if devices:
            combo1.current(0)
            if len(devices) > 1:
                combo2.current(1)

    # 快捷键设置
    ttk.Label(root, text="快捷键：").pack(pady=5)
    entry_hotkey = ttk.Entry(root, width=20)
    entry_hotkey.pack(pady=5)
    entry_hotkey.insert(0, current_hotkey)

    def on_ok():
        sel1 = combo1.get().strip()
        sel2 = combo2.get().strip()
        hotkey = entry_hotkey.get().strip().lower()
        
        if not all([sel1, sel2, hotkey]) or sel1 == sel2:
            messagebox.showerror("错误", "无效配置")
            return
            
        new_config = {
            "DEVICE1": sel1,
            "DEVICE2": sel2,
            "DEVICE1_ID": id_mapping[sel1],
            "DEVICE2_ID": id_mapping[sel2],
            "hotkey": hotkey
        }
        save_config(new_config)
        
        # 即时更新快捷键
        if hotkey_mgr and hotkey != hotkey_mgr.hotkey:
            if hotkey_mgr.reload_hotkey(hotkey):
                messagebox.showinfo("成功", f"快捷键已更新为：{hotkey}")
            else:
                messagebox.showerror("失败", "快捷键注册失败，请检查组合键格式")
        
        root.destroy()

    ttk.Button(root, text="保存", command=on_ok).pack(pady=20)
    root.mainloop()

def create_image():
    image = Image.new('RGB', (64, 64), "white")
    draw = ImageDraw.Draw(image)
    draw.ellipse((8, 8, 56, 56), fill="blue")
    return image

def on_switch(icon, item):
    switch_audio_device()
    current_id = get_current_audio_device_id()
    icon.notify(f"当前设备：{get_device_name(current_id)}", "切换成功")

def on_get_current(icon, item):
    current_id = get_current_audio_device_id()
    config = load_config()
    msg = f"快捷键：{config.get("hotkey", "ctrl+alt+s")}\n预设1：{config.get('DEVICE1', '')}\n预设2：{config.get('DEVICE2', '')}"
    icon.notify(msg, "设备状态")

def on_exit(icon, item):
    icon.stop()
    os._exit(0)

def main():
    hotkey_mgr = HotkeyManager()
    config = load_config()
    
    ensure_audio_module_installed()

    if not (config.get("DEVICE1_ID") and config.get("DEVICE2_ID")):
        select_devices(hotkey_mgr)  # 传递实例
        config = load_config()

    hotkey = config.get("hotkey", "ctrl+alt+s")
    hotkey_mgr.register_hotkey(hotkey, switch_audio_device)
    hotkey_mgr.start_listener()

    icon = pystray.Icon(
        "AudioSwitcher",
        create_image(),
        "音频切换器",
        menu = pystray.Menu(
            pystray.MenuItem("设备设置", lambda: select_devices(hotkey_mgr)),  # 传递实例
            pystray.MenuItem("切换设备", on_switch),
            pystray.MenuItem("当前状态", on_get_current),
            pystray.MenuItem("退出", on_exit)
        )
    )
    icon.run()

if __name__ == '__main__':
    if sys.platform == 'win32':
        if ctypes.windll.shell32.IsUserAnAdmin() == 0:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
            sys.exit()
    main()