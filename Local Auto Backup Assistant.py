import os
import shutil
import time
import threading
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import configparser
import keyboard
import sys
from PIL import Image, ImageTk
import pystray
import win32api
import win32con
import win32gui
from collections import deque
import logging
from pathlib import Path

logging.basicConfig(
    filename='backup_tool.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 语言支持
LANGUAGES = {
    'zh': {
        'app_title': '文件自动备份小工具',
        'tab_backup': '自动备份',
        'tab_manage': '备份管理',
        'frame_settings': '备份设置',
        'source_paths': '源路径:',
        'add_file': '添加文件',
        'add_dir': '添加目录',
        'remove': '移除',
        'open_location': '打开位置',
        'dest_dir': '备份目录:',
        'browse': '浏览...',
        'interval': '自动保存间隔(秒):',
        'max_backups': '最大备份数量:',
        'hotkey': '快捷键:',
        'enable_hotkey': '启用快捷键',
        'start_backup': '开始自动备份',
        'stop_backup': '停止自动备份',
        'manual_backup': '手动备份',
        'backup_time': '备份时间',
        'backup_path': '备份路径',
        'source_path': '源路径',
        'refresh_list': '刷新列表',
        'delete_selected': '删除选中',
        'restore_selected': '还原选中',
        'open_backup': '打开位置',
        'status_ready': '就绪',
        'hotkey_enabled': '全局快捷键已启用: {hotkey}',
        'hotkey_disabled': '全局快捷键已禁用',
        'backup_started': '自动备份已启动，间隔: {interval}秒',
        'backup_stopped': '自动备份已停止',
        'backup_running': '自动备份运行中',
        'manual_backup_running': '正在执行手动备份...',
        'manual_backup_done': '手动备份完成',
        'backup_success': '备份成功: {timestamp}',
        'deleted_backup': '已删除备份: {name}',
        'restored_backup': '已成功还原备份: {name}',
        'select_backup': '请先选择一个备份',
        'select_one_backup': '一次只能还原一个备份',
        'confirm_restore': '确定要还原备份 {name} 吗？这将覆盖当前文件！',
        'restore_success': '备份已还原！',
        'error_invalid_number': '请输入有效的数字',
        'error_no_source': '请添加至少一个源路径',
        'error_no_dest': '请选择备份目录',
        'error_invalid_hotkey': '无效的快捷键格式，请使用如\'ctrl+F1\'的格式',
        'error_hotkey_reg': '注册快捷键失败: {error}',
        'error_backup': '备份过程中出错: {error}',
        'error_cleanup': '清理旧备份时出错: {error}',
        'error_delete': '删除失败: {error}',
        'error_restore': '还原失败: {error}',
        'error_no_source_path': '找不到源路径: {name}',
        'error_backup_content': '备份内容不存在或已损坏',
        'tray_show': '显示窗口',
        'tray_exit': '退出',
        'tray_toggle_backup': '启用自动备份',
        'error_path_same': '错误: 源路径和备份路径不能相同',
        'error_path_contains': '错误: 源路径和备份路径不能存在包含关系',
        'language': '语言:'
    },
    'en': {
        'app_title': 'File Auto Backup Tool',
        'tab_backup': 'Auto Backup',
        'tab_manage': 'Backup Management',
        'frame_settings': 'Backup Settings',
        'source_paths': 'Source Paths:',
        'add_file': 'Add File',
        'add_dir': 'Add Folder',
        'remove': 'Remove',
        'open_location': 'Open Location',
        'dest_dir': 'Backup Directory:',
        'browse': 'Browse...',
        'interval': 'Auto-save Interval (seconds):',
        'max_backups': 'Max Backup Count:',
        'hotkey': 'Hotkey:',
        'enable_hotkey': 'Enable Hotkey',
        'start_backup': 'Start Auto Backup',
        'stop_backup': 'Stop Auto Backup',
        'manual_backup': 'Manual Backup',
        'backup_time': 'Backup Time',
        'backup_path': 'Backup Path',
        'source_path': 'Source Path',
        'refresh_list': 'Refresh List',
        'delete_selected': 'Delete Selected',
        'restore_selected': 'Restore Selected',
        'open_backup': 'Open Location',
        'status_ready': 'Ready',
        'hotkey_enabled': 'Global hotkey enabled: {hotkey}',
        'hotkey_disabled': 'Global hotkey disabled',
        'backup_started': 'Auto backup started, interval: {interval}s',
        'backup_stopped': 'Auto backup stopped',
        'backup_running': 'Auto backup running',
        'manual_backup_running': 'Performing manual backup...',
        'manual_backup_done': 'Manual backup completed',
        'backup_success': 'Backup successful: {timestamp}',
        'deleted_backup': 'Deleted backup: {name}',
        'restored_backup': 'Successfully restored backup: {name}',
        'select_backup': 'Please select a backup first',
        'select_one_backup': 'Only one backup can be restored at a time',
        'confirm_restore': 'Are you sure you want to restore backup {name}? This will overwrite current files!',
        'restore_success': 'Backup restored successfully!',
        'error_invalid_number': 'Please enter a valid number',
        'error_no_source': 'Please add at least one source path',
        'error_no_dest': 'Please select a backup directory',
        'error_invalid_hotkey': 'Invalid hotkey format, please use format like \'ctrl+F1\'',
        'error_hotkey_reg': 'Failed to register hotkey: {error}',
        'error_backup': 'Error during backup: {error}',
        'error_cleanup': 'Error while cleaning up old backups: {error}',
        'error_delete': 'Deletion failed: {error}',
        'error_restore': 'Restore failed: {error}',
        'error_no_source_path': 'Source path not found: {name}',
        'error_backup_content': 'Backup content does not exist or is corrupted',
        'tray_show': 'Show Window',
        'tray_exit': 'Exit',
        'tray_toggle_backup': 'Enable Auto Backup',
        'error_path_same': 'Error: Source path and backup path cannot be the same',
        'error_path_contains': 'Error: Source path and backup path cannot contain each other',
        'language': 'Language:'
    }
}

class AutoBackupTool:
    def __init__(self, root):
        self.root = root
        self.current_lang = 'zh'  # 默认中文
        self.lang_strings = LANGUAGES[self.current_lang]
        
        self.root.title(self.lang_strings['app_title'])
        self.root.geometry("800x550")
        self.root.minsize(600, 480)
        
        self.source_paths = []
        self.dest_dir = ""
        self.interval = 300
        self.max_backups = 10
        self.running = False
        self.backup_thread = None
        self.hotkey_registered = False
        self.config_file = "backup_config.ini"
        self.tray_icon = None
        self.tray_thread = None  # 用于跟踪托盘图标的线程
        
        self.normal_icon = "folder-sync.png"
        self.active_icon = "folder-sync-start.png"
        
        self.hotkey = "ctrl+F1"
        
        self.load_config()
        
        self.create_widgets()
        self.update_backup_list()
        
        self.set_icon()
        
        self.create_tray_icon()
        
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        
        self.root.bind("<Destroy>", self.on_destroy)
        
        # 线程事件，用于优雅地停止线程
        self.stop_event = threading.Event()
    
    def set_language(self, lang):
        """设置界面语言"""
        if lang in LANGUAGES and lang != self.current_lang:
            self.current_lang = lang
            self.lang_strings = LANGUAGES[self.current_lang]
            self.update_ui_texts()
            self.update_lang_button_state()  # 确保语言按钮状态更新
            self.save_config()
    
    def update_ui_texts(self):
        """更新界面所有文本"""
        self.root.title(self.lang_strings['app_title'])
        self.tab_control.tab(0, text=self.lang_strings['tab_backup'])
        self.tab_control.tab(1, text=self.lang_strings['tab_manage'])
        self.frame_settings.config(text=self.lang_strings['frame_settings'])
        self.source_label.config(text=self.lang_strings['source_paths'])
        self.btn_add_file.config(text=self.lang_strings['add_file'])
        self.btn_add_dir.config(text=self.lang_strings['add_dir'])
        self.btn_remove_source.config(text=self.lang_strings['remove'])
        self.btn_open_source.config(text=self.lang_strings['open_location'])
        self.dest_label.config(text=self.lang_strings['dest_dir'])
        self.btn_browse_dest.config(text=self.lang_strings['browse'])
        self.btn_open_dest.config(text=self.lang_strings['open_location'])
        self.interval_label.config(text=self.lang_strings['interval'])
        self.max_backups_label.config(text=self.lang_strings['max_backups'])
        self.hotkey_label.config(text=self.lang_strings['hotkey'])
        self.check_hotkey.config(text=self.lang_strings['enable_hotkey'])
        
        if self.running:
            self.start_btn.config(text=self.lang_strings['stop_backup'])
        else:
            self.start_btn.config(text=self.lang_strings['start_backup'])
        
        self.btn_manual_backup.config(text=self.lang_strings['manual_backup'])
        self.backup_list.heading("date", text=self.lang_strings['backup_time'])
        self.backup_list.heading("path", text=self.lang_strings['backup_path'])
        self.backup_list.heading("source", text=self.lang_strings['source_path'])
        self.btn_refresh.config(text=self.lang_strings['refresh_list'])
        self.btn_delete.config(text=self.lang_strings['delete_selected'])
        self.btn_restore.config(text=self.lang_strings['restore_selected'])
        self.btn_open_backup.config(text=self.lang_strings['open_backup'])
        
        self.status_var.set(self.lang_strings['status_ready'])
        
        # 更新托盘图标菜单
        self.create_tray_icon(active=self.running)
    
    def check_path_relationship(self, path1, path2):
        """检查两个路径是否相同或存在包含关系"""
        if not path1 or not path2:
            return False
            
        # 标准化路径
        path1 = os.path.abspath(os.path.normpath(path1))
        path2 = os.path.abspath(os.path.normpath(path2))
        
        # 检查是否相同
        if path1 == path2:
            return True
            
        # 检查是否存在包含关系
        path1_parts = Path(path1).parts
        path2_parts = Path(path2).parts
        
        min_length = min(len(path1_parts), len(path2_parts))
        for i in range(min_length):
            if path1_parts[i] != path2_parts[i]:
                return False
        
        return True
    
    def set_icon(self, active=False):
        try:
            icon_path = self.active_icon if active else self.normal_icon
            
            if os.path.exists(icon_path):
                self.root.iconphoto(True, tk.PhotoImage(file=icon_path))
            else:
                self.root.iconbitmap(default="")
        except Exception as e:
            logging.error(f"设置图标失败: {e}")
    
    def create_tray_icon(self, active=False):
        """创建或更新托盘图标，确保只有一个实例"""
        try:
            # 停止并清理现有托盘图标
            if self.tray_icon:
                self.tray_icon.stop()
                self.tray_icon = None
            
            if self.tray_thread and self.tray_thread.is_alive():
                # 等待线程结束
                self.tray_thread.join(timeout=1.0)
            
            icon_path = self.active_icon if active else self.normal_icon
            
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
            else:
                color = 'green' if active else 'gray'
                image = Image.new('RGB', (16, 16), color=color)
            
            # 托盘菜单
            toggle_text = self.lang_strings['tray_toggle_backup']
            menu = (
                pystray.MenuItem(self.lang_strings['tray_show'], self.restore_from_tray),
                pystray.MenuItem(toggle_text, self.toggle_backup_from_tray, 
                                checked=lambda item: self.running),
                pystray.MenuItem(self.lang_strings['tray_exit'], self.exit_application)
            )
            
            self.tray_icon = pystray.Icon(self.lang_strings['app_title'], image, self.lang_strings['app_title'], menu)
            
            # 使用新线程运行托盘图标
            self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
            self.tray_thread.start()
        except Exception as e:
            logging.error(f"创建托盘图标失败: {e}")
    
    def toggle_backup_from_tray(self):
        """从托盘菜单切换自动备份状态"""
        self.root.after(0, self.toggle_backup)
    
    def minimize_to_tray(self):
        self.root.withdraw()
        if self.tray_icon:
            self.tray_icon.visible = True
    
    def restore_from_tray(self):
        self.root.deiconify()
        self.root.lift()
        if self.tray_icon:
            self.tray_icon.visible = False
    
    def exit_application(self):
        # 确保在主线程中执行退出操作
        def do_exit():
            # 优雅地停止线程
            self.running = False
            self.stop_event.set()
            
            if self.hotkey_registered:
                try:
                    keyboard.remove_hotkey(self.hotkey)
                except:
                    pass
            
            # 停止托盘图标
            if self.tray_icon:
                self.tray_icon.stop()
                self.tray_icon = None
            
            if self.tray_thread and self.tray_thread.is_alive():
                self.tray_thread.join(timeout=1.0)
            
            # 销毁主窗口
            self.root.destroy()
            
            # 强制退出所有线程
            os._exit(0)
        
        # 如果当前在主线程，直接执行；否则通过after调度到主线程
        if threading.current_thread() is threading.main_thread():
            do_exit()
        else:
            self.root.after(0, do_exit)
    
    def on_destroy(self, event):
        if event.widget == self.root:
            self.exit_application()
    
    def load_config(self):
        config = configparser.ConfigParser()
        if os.path.exists(self.config_file):
            try:
                config.read(self.config_file)
                
                if config.has_option('Settings', 'SourcePaths'):
                    self.source_paths = eval(config.get('Settings', 'SourcePaths'))
                
                if config.has_option('Settings', 'DestDir'):
                    self.dest_dir = config.get('Settings', 'DestDir')
                if config.has_option('Settings', 'Interval'):
                    self.interval = config.getint('Settings', 'Interval')
                if config.has_option('Settings', 'MaxBackups'):
                    self.max_backups = config.getint('Settings', 'MaxBackups')
                if config.has_option('Settings', 'HotkeyEnabled'):
                    hotkey_enabled = config.getboolean('Settings', 'HotkeyEnabled')
                    self.hotkey_var.set(hotkey_enabled)
                
                if config.has_option('Settings', 'Hotkey'):
                    self.hotkey = config.get('Settings', 'Hotkey')
                else:
                    self.hotkey = "ctrl+F1"
                
                # 加载语言设置
                if config.has_option('Settings', 'Language'):
                    lang = config.get('Settings', 'Language')
                    if lang in LANGUAGES:
                        self.current_lang = lang
                        self.lang_strings = LANGUAGES[self.current_lang]
                
                if self.hotkey_var.get():
                    self.toggle_hotkey()
                
                # 确保语言按钮状态正确
                self.update_lang_button_state()
                    
            except Exception as e:
                logging.error(f"加载配置文件失败: {e}")
    
    def save_config(self):
        config = configparser.ConfigParser()
        config['Settings'] = {
            'SourcePaths': str(self.source_paths),
            'DestDir': self.dest_dir,
            'Interval': str(self.interval),
            'MaxBackups': str(self.max_backups),
            'HotkeyEnabled': str(self.hotkey_var.get()),
            'Hotkey': self.hotkey,
            'Language': self.current_lang
        }
        
        try:
            with open(self.config_file, 'w') as configfile:
                config.write(configfile)
        except Exception as e:
            logging.error(f"保存配置文件失败: {e}")
    
    def create_widgets(self):
        # 创建顶部框架，用于放置语言切换按钮
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        # 语言切换按钮放在右上角
        lang_btn_frame = ttk.Frame(top_frame)
        lang_btn_frame.pack(side="right")
        
        self.btn_zh = ttk.Button(lang_btn_frame, text="中文", command=lambda: self.set_language('zh'))
        self.btn_zh.pack(side="left", padx=2)
        self.btn_en = ttk.Button(lang_btn_frame, text="English", command=lambda: self.set_language('en'))
        self.btn_en.pack(side="left", padx=2)
        
        # 根据当前语言设置按钮状态
        self.update_lang_button_state()
        
        self.tab_control = ttk.Notebook(self.root)
        tab_backup = ttk.Frame(self.tab_control)
        tab_manage = ttk.Frame(self.tab_control)
        self.tab_control.add(tab_backup, text=self.lang_strings['tab_backup'])
        self.tab_control.add(tab_manage, text=self.lang_strings['tab_manage'])
        self.tab_control.pack(expand=1, fill="both", padx=10, pady=5)
        
        self.frame_settings = ttk.LabelFrame(tab_backup, text=self.lang_strings['frame_settings'])
        self.frame_settings.pack(fill="x", padx=10, pady=5)
        
        # 配置网格权重，使界面可缩放
        self.frame_settings.columnconfigure(1, weight=1)
        
        self.source_label = ttk.Label(self.frame_settings, text=self.lang_strings['source_paths'])
        self.source_label.grid(row=0, column=0, padx=5, pady=5, sticky="ne")
        
        self.source_listbox = tk.Listbox(self.frame_settings, width=50, height=4)
        self.source_listbox.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        
        btn_frame = ttk.Frame(self.frame_settings)
        btn_frame.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        self.btn_add_file = ttk.Button(btn_frame, text=self.lang_strings['add_file'], command=lambda: self.add_source('file'))
        self.btn_add_file.pack(fill="x", padx=2, pady=2)
        self.btn_add_dir = ttk.Button(btn_frame, text=self.lang_strings['add_dir'], command=lambda: self.add_source('dir'))
        self.btn_add_dir.pack(fill="x", padx=2, pady=2)
        self.btn_remove_source = ttk.Button(btn_frame, text=self.lang_strings['remove'], command=self.remove_source)
        self.btn_remove_source.pack(fill="x", padx=2, pady=2)
        self.btn_open_source = ttk.Button(btn_frame, text=self.lang_strings['open_location'], command=self.open_source_location)
        self.btn_open_source.pack(fill="x", padx=2, pady=2)
        
        self.dest_label = ttk.Label(self.frame_settings, text=self.lang_strings['dest_dir'])
        self.dest_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.dest_entry = ttk.Entry(self.frame_settings)
        self.dest_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")
        self.dest_entry.insert(0, self.dest_dir)
        
        dest_btn_frame = ttk.Frame(self.frame_settings)
        dest_btn_frame.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.btn_browse_dest = ttk.Button(dest_btn_frame, text=self.lang_strings['browse'], command=self.select_dest_dir)
        self.btn_browse_dest.pack(fill="x", padx=2, pady=2)
        self.btn_open_dest = ttk.Button(dest_btn_frame, text=self.lang_strings['open_location'], command=self.open_dest_location)
        self.btn_open_dest.pack(fill="x", padx=2, pady=2)
        
        self.interval_label = ttk.Label(self.frame_settings, text=self.lang_strings['interval'])
        self.interval_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.interval_entry = ttk.Entry(self.frame_settings, width=15)
        self.interval_entry.insert(0, str(self.interval))
        self.interval_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        self.max_backups_label = ttk.Label(self.frame_settings, text=self.lang_strings['max_backups'])
        self.max_backups_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.max_backups_entry = ttk.Entry(self.frame_settings, width=15)
        self.max_backups_entry.insert(0, str(self.max_backups))
        self.max_backups_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        self.hotkey_var = tk.BooleanVar(value=False)
        self.hotkey_label = ttk.Label(self.frame_settings, text=self.lang_strings['hotkey'])
        self.hotkey_label.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.hotkey_entry = ttk.Entry(self.frame_settings, width=15)
        self.hotkey_entry.insert(0, self.hotkey)
        self.hotkey_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.check_hotkey = ttk.Checkbutton(self.frame_settings, text=self.lang_strings['enable_hotkey'], 
                        variable=self.hotkey_var, command=self.toggle_hotkey)
        self.check_hotkey.grid(row=4, column=2, padx=5, pady=5, sticky="w")
        
        btn_frame = ttk.Frame(tab_backup)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        self.start_btn = ttk.Button(btn_frame, text=self.lang_strings['start_backup'], command=self.toggle_backup)
        self.start_btn.pack(side="left", padx=5, pady=5)
        
        self.btn_manual_backup = ttk.Button(btn_frame, text=self.lang_strings['manual_backup'], command=self.manual_backup)
        self.btn_manual_backup.pack(side="left", padx=5, pady=5)
        
        manage_frame = ttk.Frame(tab_manage)
        manage_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.backup_list = ttk.Treeview(manage_frame, columns=("date", "path", "source"), show="headings")
        self.backup_list.heading("date", text=self.lang_strings['backup_time'])
        self.backup_list.heading("path", text=self.lang_strings['backup_path'])
        self.backup_list.heading("source", text=self.lang_strings['source_path'])
        self.backup_list.column("date", width=150)
        self.backup_list.column("path", width=300, stretch=True)
        self.backup_list.column("source", width=250, stretch=True)
        self.backup_list.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(manage_frame, orient="vertical", command=self.backup_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.backup_list.configure(yscrollcommand=scrollbar.set)
        
        btn_frame = ttk.Frame(tab_manage)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        self.btn_refresh = ttk.Button(btn_frame, text=self.lang_strings['refresh_list'], command=self.update_backup_list)
        self.btn_refresh.pack(side="left", padx=5, pady=5)
        self.btn_delete = ttk.Button(btn_frame, text=self.lang_strings['delete_selected'], command=self.delete_selected)
        self.btn_delete.pack(side="left", padx=5, pady=5)
        self.btn_restore = ttk.Button(btn_frame, text=self.lang_strings['restore_selected'], command=self.restore_selected)
        self.btn_restore.pack(side="left", padx=5, pady=5)
        self.btn_open_backup = ttk.Button(btn_frame, text=self.lang_strings['open_backup'], command=self.open_backup_location)
        self.btn_open_backup.pack(side="left", padx=5, pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set(self.lang_strings['status_ready'])
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")
        
        self.update_source_list()
    
    def update_lang_button_state(self):
        """更新语言按钮状态，当前语言按钮禁用"""
        if self.current_lang == 'zh':
            self.btn_zh.config(state='disabled')
            self.btn_en.config(state='normal')
        else:
            self.btn_zh.config(state='normal')
            self.btn_en.config(state='disabled')
    
    def update_source_list(self):
        self.source_listbox.delete(0, tk.END)
        for path in self.source_paths:
            self.source_listbox.insert(tk.END, path)
    
    def add_source(self, source_type):
        if source_type == 'file':
            files = filedialog.askopenfilenames(title=self.lang_strings['add_file'])
            if files:
                new_paths = []
                for file in files:
                    # 检查与备份目录的关系
                    if self.dest_dir and self.check_path_relationship(file, self.dest_dir):
                        messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                        continue
                    new_paths.append(file)
                self.source_paths.extend(new_paths)
        else:
            directory = filedialog.askdirectory(title=self.lang_strings['add_dir'])
            if directory:
                # 检查与备份目录的关系
                if self.dest_dir and self.check_path_relationship(directory, self.dest_dir):
                    messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                    return
                self.source_paths.append(directory)
        
        self.source_paths = list(set(self.source_paths))
        self.update_source_list()
        self.save_config()
    
    def remove_source(self):
        selected = self.source_listbox.curselection()
        if selected:
            index = selected[0]
            self.source_paths.pop(index)
            self.update_source_list()
            self.save_config()
    
    def open_source_location(self):
        selected = self.source_listbox.curselection()
        if not selected:
            return
        
        path = self.source_paths[selected[0]]
        if os.path.exists(path):
            if os.path.isfile(path):
                directory = os.path.dirname(path)
                os.startfile(directory)
            else:
                os.startfile(path)
    
    def open_dest_location(self):
        if self.dest_dir and os.path.exists(self.dest_dir):
            os.startfile(self.dest_dir)
    
    def open_backup_location(self):
        selected = self.backup_list.selection()
        if not selected:
            return
        
        item = selected[0]
        values = self.backup_list.item(item, "values")
        if values and len(values) >= 2:
            backup_path = values[1]
            if os.path.exists(backup_path):
                os.startfile(backup_path)
    
    def select_dest_dir(self):
        directory = filedialog.askdirectory(title=self.lang_strings['dest_dir'])
        if directory:
            # 检查与所有源路径的关系
            for source in self.source_paths:
                if self.check_path_relationship(source, directory):
                    messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                    return
            
            self.dest_dir = directory
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, directory)
            self.update_backup_list()
            self.save_config()
    
    def is_valid_hotkey(self, hotkey_str):
        if not hotkey_str:
            return False
        
        valid_modifiers = ["ctrl", "shift", "alt", "win"]
        
        keys = hotkey_str.split('+')
        
        if len(keys) < 1:
            return False
        
        for i, key in enumerate(keys):
            key_lower = key.lower()
            
            if i == len(keys) - 1:
                return len(key) > 0
                
            if key_lower not in valid_modifiers:
                return False
        
        return True

    def toggle_hotkey(self):
        new_hotkey = self.hotkey_entry.get().strip()
        
        if not self.is_valid_hotkey(new_hotkey):
            messagebox.showerror("Error", self.lang_strings['error_invalid_hotkey'])
            self.hotkey_var.set(False)
            return
        
        self.hotkey = new_hotkey
        
        if self.hotkey_var.get():
            try:
                if self.hotkey_registered:
                    try:
                        keyboard.remove_hotkey(self.hotkey)
                    except:
                        pass
                
                keyboard.add_hotkey(self.hotkey, self.manual_backup)
                self.hotkey_registered = True
                self.status_var.set(self.lang_strings['hotkey_enabled'].format(hotkey=self.hotkey))
                self.save_config()
            except Exception as e:
                logging.error(f"注册快捷键失败: {e}")
                messagebox.showerror("Error", self.lang_strings['error_hotkey_reg'].format(error=str(e)))
                self.hotkey_var.set(False)
        elif self.hotkey_registered:
            try:
                keyboard.remove_hotkey(self.hotkey)
                self.hotkey_registered = False
                self.status_var.set(self.lang_strings['hotkey_disabled'])
                self.save_config()
            except Exception as e:
                logging.error(f"移除快捷键失败: {e}")
                self.hotkey_registered = False
    
    def toggle_backup(self):
        if self.running:
            self.running = False
            self.stop_event.set()  # 触发事件，通知线程停止
            if self.backup_thread and self.backup_thread.is_alive():
                self.backup_thread.join(timeout=2.0)  # 等待线程结束
            self.start_btn.config(text=self.lang_strings['start_backup'])
            self.status_var.set(self.lang_strings['backup_stopped'])
            
            self.create_tray_icon(active=False)
            self.set_icon(active=False)
        else:
            try:
                self.interval = int(self.interval_entry.get())
                self.max_backups = int(self.max_backups_entry.get())
                
                if self.interval <= 0 or self.max_backups <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            
            if not self.source_paths:
                messagebox.showerror("Error", self.lang_strings['error_no_source'])
                return
            
            if not self.dest_dir:
                messagebox.showerror("Error", self.lang_strings['error_no_dest'])
                return
            
            # 检查所有源路径与备份目录的关系
            for source in self.source_paths:
                if self.check_path_relationship(source, self.dest_dir):
                    messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                    return
            
            self.running = True
            self.stop_event.clear()  # 重置事件
            self.start_btn.config(text=self.lang_strings['stop_backup'])
            self.status_var.set(self.lang_strings['backup_started'].format(interval=self.interval))
            self.save_config()
            
            self.create_tray_icon(active=True)
            self.set_icon(active=True)
            
            # 创建新线程执行自动备份
            self.backup_thread = threading.Thread(target=self.auto_backup, daemon=True)
            self.backup_thread.start()
    
    def auto_backup(self):
        last_backup_time = time.time()
        
        while self.running and not self.stop_event.is_set():
            current_time = time.time()
            if current_time - last_backup_time >= self.interval:
                # 使用主线程队列执行UI更新操作
                self.root.after(0, lambda: self.status_var.set(self.lang_strings['backup_running']))
                self.perform_backup()
                last_backup_time = current_time
                self.root.after(0, lambda: self.status_var.set(self.lang_strings['backup_started'].format(interval=self.interval)))
                self.root.after(0, self.update_backup_list)  # 刷新备份列表
            
            # 使用事件等待代替time.sleep，允许更快响应停止请求
            self.stop_event.wait(1)
    
    def manual_backup(self):
        if not self.source_paths:
            messagebox.showerror("Error", self.lang_strings['error_no_source'])
            return
        
        if not self.dest_dir:
            messagebox.showerror("Error", self.lang_strings['error_no_dest'])
            return
        
        # 检查所有源路径与备份目录的关系
        for source in self.source_paths:
            if self.check_path_relationship(source, self.dest_dir):
                messagebox.showerror("Error", self.lang_strings['error_path_contains'])
                return
        
        self.status_var.set(self.lang_strings['manual_backup_running'])
        self.root.update()
        
        self.perform_backup()
        
        self.status_var.set(self.lang_strings['manual_backup_done'])
        self.update_backup_list()
    
    def perform_backup(self):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        try:
            os.makedirs(self.dest_dir, exist_ok=True)
            
            for source_path in self.source_paths:
                # 再次检查路径关系，确保安全
                if self.check_path_relationship(source_path, self.dest_dir):
                    logging.error(f"路径关系不安全: {source_path} 和 {self.dest_dir}")
                    continue
                    
                source_name = os.path.basename(source_path)
                backup_name = f"{source_name}_{timestamp}"
                backup_path = os.path.join(self.dest_dir, backup_name)
                
                os.makedirs(backup_path, exist_ok=True)
                
                # 存储完整源路径信息
                with open(os.path.join(backup_path, ".source_path"), "w", encoding="utf-8") as f:
                    f.write(source_path)
                
                if os.path.isdir(source_path):
                    shutil.copytree(source_path, os.path.join(backup_path, source_name), dirs_exist_ok=True)
                else:
                    shutil.copy2(source_path, os.path.join(backup_path, source_name))
            
            self.cleanup_old_backups()
            
            self.status_var.set(self.lang_strings['backup_success'].format(timestamp=timestamp))
            logging.info(f"成功完成备份: {timestamp}")
        except Exception as e:
            error_msg = self.lang_strings['error_backup'].format(error=str(e))
            logging.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def cleanup_old_backups(self):
        try:
            backup_queue = deque()
            
            for item in os.listdir(self.dest_dir):
                item_path = os.path.join(self.dest_dir, item)
                if os.path.isdir(item_path):
                    try:
                        ctime = os.path.getctime(item_path)
                        backup_queue.append((ctime, item_path))
                    except:
                        continue
            
            backup_queue = deque(sorted(backup_queue, key=lambda x: x[0]))
            
            while len(backup_queue) > self.max_backups:
                _, oldest = backup_queue.popleft()
                try:
                    shutil.rmtree(oldest)
                    logging.info(f"清理旧备份: {os.path.basename(oldest)}")
                except Exception as e:
                    logging.error(f"清理旧备份失败: {e}")
        except Exception as e:
            error_msg = self.lang_strings['error_cleanup'].format(error=str(e))
            logging.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def update_backup_list(self):
        # 清空现有列表
        for item in self.backup_list.get_children():
            self.backup_list.delete(item)
        
        if not self.dest_dir or not os.path.exists(self.dest_dir):
            return
        
        # 遍历备份目录并添加到列表
        for item in os.listdir(self.dest_dir):
            item_path = os.path.join(self.dest_dir, item)
            if os.path.isdir(item_path):
                try:
                    # 获取创建时间
                    ctime = os.path.getctime(item_path)
                    date_str = datetime.datetime.fromtimestamp(ctime).strftime("%Y-%m-%d %H:%M:%S")
                    
                    # 尝试读取源路径信息
                    source_path = "Unknown"
                    source_file = os.path.join(item_path, ".source_path")
                    if os.path.exists(source_file):
                        with open(source_file, "r", encoding="utf-8") as f:
                            source_path = f.read().strip()
                    
                    self.backup_list.insert("", "end", values=(date_str, item_path, source_path))
                except Exception as e:
                    logging.warning(f"无法添加备份到列表: {item_path}, 错误: {e}")
    
    def delete_selected(self):
        selected = self.backup_list.selection()
        if not selected:
            messagebox.showinfo("Info", self.lang_strings['select_backup'])
            return
        
        for item in selected:
            values = self.backup_list.item(item, "values")
            if values and len(values) >= 2:
                backup_path = values[1]
                try:
                    shutil.rmtree(backup_path)
                    self.backup_list.delete(item)
                    logging.info(self.lang_strings['deleted_backup'].format(name=os.path.basename(backup_path)))
                except Exception as e:
                    error_msg = self.lang_strings['error_delete'].format(error=str(e))
                    logging.error(error_msg)
                    messagebox.showerror("Error", error_msg)
        
        self.update_backup_list()
    
    def restore_selected(self):
        selected = self.backup_list.selection()
        if not selected:
            messagebox.showinfo("Info", self.lang_strings['select_backup'])
            return
        
        if len(selected) > 1:
            messagebox.showinfo("Info", self.lang_strings['select_one_backup'])
            return
        
        item = selected[0]
        values = self.backup_list.item(item, "values")
        if not values or len(values) < 3:
            return
        
        backup_path = values[1]
        source_path = values[2]
        
        # 验证备份内容
        backup_content = None
        for item in os.listdir(backup_path):
            item_full = os.path.join(backup_path, item)
            if os.path.isdir(item_full) or os.path.isfile(item_full) and item != ".source_path":
                backup_content = item_full
                break
        
        if not backup_content:
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        
        # 确认还原操作
        if not messagebox.askyesno("Confirm", self.lang_strings['confirm_restore'].format(name=os.path.basename(backup_path))):
            return
        
        try:
            # 确保源路径存在
            if not os.path.exists(source_path):
                raise Exception(self.lang_strings['error_no_source_path'].format(name=source_path))
            
            # 执行还原操作
            backup_item_name = os.path.basename(backup_content)
            dest_path = os.path.join(os.path.dirname(source_path), backup_item_name)
            
            if os.path.isdir(backup_content):
                if os.path.exists(dest_path):
                    shutil.rmtree(dest_path)
                shutil.copytree(backup_content, dest_path)
            else:
                shutil.copy2(backup_content, dest_path)
            
            logging.info(self.lang_strings['restored_backup'].format(name=os.path.basename(backup_path)))
            messagebox.showinfo("Success", self.lang_strings['restore_success'])
        except Exception as e:
            error_msg = self.lang_strings['error_restore'].format(error=str(e))
            logging.error(error_msg)
            messagebox.showerror("Error", error_msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoBackupTool(root)
    root.mainloop()
