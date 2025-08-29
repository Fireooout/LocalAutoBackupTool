import os
import sys
import shutil
import datetime
import logging
import threading
import configparser
import keyboard
import re
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, Menu
import tkinter as tk
import win32gui
import win32con
import win32api

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 字体配置
FONT_CONFIG = {
    'base_size': 12,
    'title_size': 14,
    'scale_factor': 1.5  # 固定缩放因子
}

# 语言配置
LANGUAGES = {
    'zh': {
        'app_title': '本地自动备份助手',
        'tab_backup': '自动备份',
        'tab_manage': '备份管理',
        'tab_settings': '高级设置',
        'frame_settings': '备份设置',
        'source_paths': '源路径:',
        'dest_dir': '备份目录:',
        'interval': '时间间隔(秒):',
        'max_backups': '最大备份数:',
        'hotkey': '快捷键:',
        'enable_hotkey': '启用快捷键',
        'browse': '浏览',
        'add_file': '添加文件',
        'add_dir': '添加目录',
        'remove': '移除',
        'open_location': '打开位置',
        'start_backup': '开始自动备份',
        'stop_backup': '停止自动备份',
        'manual_backup': '手动备份',
        'backup_time': '备份时间',
        'backup_path': '备份路径',
        'source_path': '源路径',
        'refresh_list': '刷新列表',
        'delete_selected': '删除选中',
        'restore_selected': '还原选中',
        'open_backup': '打开备份位置',
        'rename_backup': '重命名备份',
        'appearance': '外观设置',
        'backup_options': '备份选项',
        'backup_prefix': '备份前缀:',
        'prefix_disabled': '禁用',
        'prefix_number': '序号',
        'prefix_custom': '自定义',
        'custom_prefix': '自定义前缀:',
        'suffix_type': '后缀类型:',
        'suffix_timestamp': '时间戳',
        'suffix_number': '版本号',
        'suffix_custom': '自定义',
        'custom_suffix': '自定义后缀:',
        'handle_duplicates': '重名文件处理:',
        'duplicates_overwrite': '覆盖',
        'duplicates_rename': '重命名',
        'duplicates_skip': '跳过',
        'backup_count': '当前备份序号: {count}',
        'set_start_number': '设置起始序号',
        'save_settings': '保存设置',
        'status_ready': '就绪',
        'backup_started': '自动备份已启动，间隔 {interval} 秒',
        'backup_stopped': '自动备份已停止',
        'backup_success': '备份成功完成 {timestamp}',
        'manual_backup_running': '正在执行手动备份...',
        'manual_backup_done': '手动备份完成',
        'hotkey_enabled': '快捷键已启用: {hotkey}',
        'hotkey_disabled': '快捷键已禁用',
        'error_no_source': '请添加源路径',
        'error_no_dest': '请选择备份目录',
        'error_invalid_number': '请输入有效的数字',
        'error_backup': '备份过程中出错: {error}',
        'error_hotkey_reg': '快捷键注册失败: {error}',
        'error_invalid_hotkey': '无效的快捷键格式',
        'error_path_contains': '源路径和备份目录不能相互包含',
        'select_backup': '请选择一个备份',
        'select_one_backup': '请选择一个备份进行操作',
        'error_backup_content': '备份内容无效或不存在',
        'confirm_restore': '确定要还原 {name} 吗？',
        'restore_success': '备份还原成功',
        'error_restore': '还原备份时出错: {error}',
        'error_delete': '删除备份时出错: {error}',
        'deleted_backup': '已删除备份: {name}',
        'error_no_source_path': '找不到源路径 {name}',
        'restored_backup': '已还原备份: {name}',
        'new_backup_name': '新备份名称:',
        'rename_success': '备份重命名成功',
        'error_rename': '重命名备份时出错: {error}',
        'error_prefix_empty': '自定义前缀不能为空',
        'error_suffix_empty': '自定义后缀不能为空',
        'confirm_overwrite': '文件已存在，是否覆盖？',
        'confirm_set_start_number': '请输入新的起始序号:',
        'tray_show': '显示窗口',
        'tray_enable_auto': '启用自动备份',
        'tray_disable_auto': '禁用自动备份',
        'tray_exit': '退出'
    },
    'en': {
        'app_title': 'Local Auto Backup Assistant',
        'tab_backup': 'Auto Backup',
        'tab_manage': 'Backup Management',
        'tab_settings': 'Advanced Settings',
        'frame_settings': 'Backup Settings',
        'source_paths': 'Source Paths:',
        'dest_dir': 'Backup Directory:',
        'interval': 'Interval (seconds):',
        'max_backups': 'Max Backups:',
        'hotkey': 'Hotkey:',
        'enable_hotkey': 'Enable Hotkey',
        'browse': 'Browse',
        'add_file': 'Add File',
        'add_dir': 'Add Directory',
        'remove': 'Remove',
        'open_location': 'Open Location',
        'start_backup': 'Start Auto Backup',
        'stop_backup': 'Stop Auto Backup',
        'manual_backup': 'Manual Backup',
        'backup_time': 'Backup Time',
        'backup_path': 'Backup Path',
        'source_path': 'Source Path',
        'refresh_list': 'Refresh List',
        'delete_selected': 'Delete Selected',
        'restore_selected': 'Restore Selected',
        'open_backup': 'Open Backup Location',
        'rename_backup': 'Rename Backup',
        'appearance': 'Appearance',
        'backup_options': 'Backup Options',
        'backup_prefix': 'Backup Prefix:',
        'prefix_disabled': 'Disabled',
        'prefix_number': 'Number',
        'prefix_custom': 'Custom',
        'custom_prefix': 'Custom Prefix:',
        'suffix_type': 'Suffix Type:',
        'suffix_timestamp': 'Timestamp',
        'suffix_number': 'Version Number',
        'suffix_custom': 'Custom',
        'custom_suffix': 'Custom Suffix:',
        'handle_duplicates': 'Handle Duplicates:',
        'duplicates_overwrite': 'Overwrite',
        'duplicates_rename': 'Rename',
        'duplicates_skip': 'Skip',
        'backup_count': 'Current Backup Number: {count}',
        'set_start_number': 'Set Start Number',
        'save_settings': 'Save Settings',
        'status_ready': 'Ready',
        'backup_started': 'Auto backup started, interval {interval} seconds',
        'backup_stopped': 'Auto backup stopped',
        'backup_success': 'Backup successfully completed {timestamp}',
        'manual_backup_running': 'Performing manual backup...',
        'manual_backup_done': 'Manual backup completed',
        'hotkey_enabled': 'Hotkey enabled: {hotkey}',
        'hotkey_disabled': 'Hotkey disabled',
        'error_no_source': 'Please add source paths',
        'error_no_dest': 'Please select backup directory',
        'error_invalid_number': 'Please enter a valid number',
        'error_backup': 'Error during backup: {error}',
        'error_hotkey_reg': 'Failed to register hotkey: {error}',
        'error_invalid_hotkey': 'Invalid hotkey format',
        'error_path_contains': 'Source paths and backup directory cannot contain each other',
        'select_backup': 'Please select a backup',
        'select_one_backup': 'Please select one backup for operation',
        'error_backup_content': 'Invalid or non-existent backup content',
        'confirm_restore': 'Are you sure you want to restore {name}?',
        'restore_success': 'Backup restored successfully',
        'error_restore': 'Error restoring backup: {error}',
        'error_delete': 'Error deleting backup: {error}',
        'deleted_backup': 'Deleted backup: {name}',
        'error_no_source_path': 'Source path not found {name}',
        'restored_backup': 'Restored backup: {name}',
        'new_backup_name': 'New Backup Name:',
        'rename_success': 'Backup renamed successfully',
        'error_rename': 'Error renaming backup: {error}',
        'error_prefix_empty': 'Custom prefix cannot be empty',
        'error_suffix_empty': 'Custom suffix cannot be empty',
        'confirm_overwrite': 'File already exists, overwrite?',
        'confirm_set_start_number': 'Please enter new start number:',
        'tray_show': 'Show Window',
        'tray_enable_auto': 'Enable Auto Backup',
        'tray_disable_auto': 'Disable Auto Backup',
        'tray_exit': 'Exit'
    }
}

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

class AutoBackupTool:
    def __init__(self, root):
        self.root = root
        self.root.title(LANGUAGES['zh']['app_title'])
        self.root.geometry("800x600")
        
        # 固定DPI设置
        self.set_dpi_awareness()
        
        # 应用字体配置
        self.apply_font_config()
        
        # 配置文件路径
        self.config_file = "backup_config.ini"
        
        # 初始化变量
        self.source_paths = []
        self.dest_dir = ""
        self.interval = 300
        self.max_backups = 10
        self.hotkey = "ctrl+F1"
        self.hotkey_var = tk.BooleanVar(value=False)
        self.hotkey_registered = False
        self.running = False
        self.stop_event = threading.Event()
        self.backup_thread = None
        
        # 新增功能变量
        self.current_lang = 'zh'
        self.lang_strings = LANGUAGES[self.current_lang]
        self.backup_prefix = ""
        self.prefix_type = "disabled"  # disabled, number, custom
        self.custom_prefix = ""
        self.suffix_type = "timestamp"  # timestamp, number, custom
        self.custom_suffix = ""
        self.duplicate_handling = "rename"  # overwrite, rename, skip
        self.backup_counter = 1
        self.start_number = 1
        
        # 图标路径
        self.normal_icon = get_resource_path("folder-sync.png")
        self.active_icon = get_resource_path("folder-sync-start.png")
        
        # 加载配置
        self.load_config()
        
        # 创建界面
        self.create_widgets()
        
        # 设置窗口图标
        self.set_icon()
        
        # 创建托盘图标
        self.root.after(100, self.create_tray_icon)
        
        # 绑定窗口关闭事件到最小化到托盘
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        
        # 更新界面文本
        self.update_ui_texts()
    
    def set_dpi_awareness(self):
        """设置DPI感知，使用固定缩放因子"""
        try:
            import ctypes
            # 设置DPI感知
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            # 使用固定的缩放因子
            self.root.tk.call('tk', 'scaling', FONT_CONFIG['scale_factor'])
        except Exception as e:
            logging.error(f"设置DPI感知失败: {e}")
    
    def apply_font_config(self):
        """应用字体配置到所有ttk控件"""
        try:
            # 定义字体
            default_font = ('Microsoft YaHei', FONT_CONFIG['base_size'])
            title_font = ('Microsoft YaHei', FONT_CONFIG['title_size'], 'bold')
            
            # 应用到ttk控件样式
            style = ttk.Style()
            
            # 配置标签样式
            style.configure('TLabel', font=default_font)
            style.configure('Title.TLabel', font=title_font)
            
            # 配置按钮样式
            style.configure('TButton', font=default_font)
            
            # 配置框架样式
            style.configure('TLabelframe.Label', font=title_font)
            
            # 配置输入框样式，确保最小字体大小
            min_entry_font_size = max(FONT_CONFIG['base_size'], 16)
            entry_font = ('Microsoft YaHei', min_entry_font_size)
            style.configure('TEntry', font=entry_font)
            
            # 配置下拉框样式，确保最小字体大小
            min_combobox_font_size = max(FONT_CONFIG['base_size'], 16)
            combobox_font = ('Microsoft YaHei', min_combobox_font_size)
            style.configure('TCombobox', font=combobox_font)
            
            # 配置复选框样式
            style.configure('TCheckbutton', font=default_font)
            
            # 配置单选按钮样式
            style.configure('TRadiobutton', font=default_font)
            
            # 配置表格样式
            style.configure('Treeview', font=default_font, rowheight=int(FONT_CONFIG['base_size'] * 2))
            style.configure('Treeview.Heading', font=('Microsoft YaHei', FONT_CONFIG['base_size'], 'bold'))
            
            # 配置列表框样式，确保字体大小与其他控件一致
            self.root.option_add('*Listbox.font', default_font)
            
        except Exception as e:
            logging.error(f"应用字体配置失败: {e}")
    
    def set_language(self, lang):
        """设置界面语言"""
        self.current_lang = lang
        self.lang_strings = LANGUAGES[self.current_lang]
        self.root.title(self.lang_strings['app_title'])
        self.update_ui_texts()
        self.update_lang_button_state()
        # 更新托盘图标
        self.create_tray_icon()
    
    def _apply_font_to_widget(self, widget, font):
        """递归应用字体到控件及其子控件"""
        try:
            if hasattr(widget, 'configure'):
                widget.configure(font=font)
            
            # 递归处理子控件
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    self._apply_font_to_widget(child, font)
        except Exception as e:
            logging.error(f"应用字体到控件失败: {e}")
    
    def update_ui_texts(self):
        """更新界面文本"""
        # 更新标签页文本
        for i, tab_text in enumerate([self.lang_strings['tab_backup'], self.lang_strings['tab_manage'], self.lang_strings['tab_settings']]):
            self.tab_control.tab(i, text=tab_text)
        
        # 更新备份设置标签页
        self.frame_settings.config(text=self.lang_strings['frame_settings'])
        self.source_frame.config(text=self.lang_strings['source_paths'])
        self.dest_frame.config(text=self.lang_strings['dest_dir'])
        self.hotkey_label.config(text=self.lang_strings['hotkey'])
        self.interval_label.config(text=self.lang_strings['interval'])
        self.max_backups_label.config(text=self.lang_strings['max_backups'])
        self.check_hotkey.config(text=self.lang_strings['enable_hotkey'])
        self.btn_add_file.config(text=self.lang_strings['add_file'])
        self.btn_add_dir.config(text=self.lang_strings['add_dir'])
        self.btn_remove_source.config(text=self.lang_strings['remove'])
        self.btn_browse_dest.config(text=self.lang_strings['browse'])
        self.btn_open_source.config(text=self.lang_strings['open_location'])
        self.btn_open_dest.config(text=self.lang_strings['open_location'])
        
        if self.running:
            self.start_btn.config(text=self.lang_strings['stop_backup'])
        else:
            self.start_btn.config(text=self.lang_strings['start_backup'])
        
        self.btn_manual_backup.config(text=self.lang_strings['manual_backup'])
        
        # 更新备份管理标签页
        self.backup_list.heading("date", text=self.lang_strings['backup_time'])
        self.backup_list.heading("path", text=self.lang_strings['backup_path'])
        self.backup_list.heading("source", text=self.lang_strings['source_path'])
        self.btn_refresh.config(text=self.lang_strings['refresh_list'])
        self.btn_delete.config(text=self.lang_strings['delete_selected'])
        self.btn_restore.config(text=self.lang_strings['restore_selected'])
        self.btn_open_backup.config(text=self.lang_strings['open_backup'])
        self.btn_rename.config(text=self.lang_strings['rename_backup'])
        
        # 更新高级设置标签页
        self.frame_appearance.config(text=self.lang_strings['appearance'])
        self.frame_backup_options.config(text=self.lang_strings['backup_options'])
        self.backup_prefix_label.config(text=self.lang_strings['backup_prefix'])
        self.prefix_disabled_radio.config(text=self.lang_strings['prefix_disabled'])
        self.prefix_number_radio.config(text=self.lang_strings['prefix_number'])
        self.prefix_custom_radio.config(text=self.lang_strings['prefix_custom'])
        self.custom_prefix_label.config(text=self.lang_strings['custom_prefix'])
        self.suffix_type_label.config(text=self.lang_strings['suffix_type'])
        self.suffix_timestamp_radio.config(text=self.lang_strings['suffix_timestamp'])
        self.suffix_number_radio.config(text=self.lang_strings['suffix_number'])
        self.suffix_custom_radio.config(text=self.lang_strings['suffix_custom'])
        self.custom_suffix_label.config(text=self.lang_strings['custom_suffix'])
        self.duplicates_label.config(text=self.lang_strings['handle_duplicates'])
        self.duplicates_overwrite_radio.config(text=self.lang_strings['duplicates_overwrite'])
        self.duplicates_rename_radio.config(text=self.lang_strings['duplicates_rename'])
        self.duplicates_skip_radio.config(text=self.lang_strings['duplicates_skip'])
        self.backup_count_label.config(text=self.lang_strings['backup_count'].format(count=self.backup_counter))
        self.btn_set_start_number.config(text=self.lang_strings['set_start_number'])
        self.btn_save_settings.config(text=self.lang_strings['save_settings'] if 'save_settings' in self.lang_strings else 'Save Settings')
        
        # 更新语言切换按钮文本和状态
        self.btn_zh.config(text="中文")
        self.btn_en.config(text="English")
        self.update_lang_button_state()
        
        # 更新状态栏
        if self.status_var.get() == LANGUAGES['zh']['status_ready'] or self.status_var.get() == LANGUAGES['en']['status_ready']:
            self.status_var.set(self.lang_strings['status_ready'])
    
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
            # 使用处理后的图标路径
            icon_path = self.active_icon if active else self.normal_icon
            
            if os.path.exists(icon_path):
                self.root.iconphoto(True, tk.PhotoImage(file=icon_path))
            else:
                self.root.iconbitmap(default="")
        except Exception as e:
            logging.error(f"设置图标失败: {e}")
    
    def create_tray_icon(self, active=False):
        """创建或更新托盘图标"""
        try:
            # 首先尝试使用pystray库（如果可用）
            try:
                import pystray
                from PIL import Image, ImageDraw
                
                # 首先停止并清理现有的托盘图标
                if hasattr(self, 'tray_icon') and self.tray_icon:
                    try:
                        self.tray_icon.stop()
                        logging.info("已停止并清理现有托盘图标")
                    except Exception as e:
                        logging.warning(f"清理现有托盘图标时出错: {e}")
                
                # 确保语言字符串存在
                app_title = self.lang_strings.get('app_title', 'Auto Backup Tool')
                tray_show_text = self.lang_strings.get('tray_show', 'Show Window')
                tray_disable_auto_text = self.lang_strings.get('tray_disable_auto', 'Disable Auto Backup')
                tray_enable_auto_text = self.lang_strings.get('tray_enable_auto', 'Enable Auto Backup')
                tray_exit_text = self.lang_strings.get('tray_exit', 'Exit')
                
                # 尝试从文件加载图标
                try:
                    icon_path = self.active_icon if active else self.normal_icon
                    if os.path.exists(icon_path):
                        icon_image = Image.open(icon_path)
                        logging.info(f"成功加载图标文件: {icon_path}")
                    else:
                        # 如果图标文件不存在，创建简单图标
                        def create_image(color):
                            width, height = 16, 16
                            image = Image.new('RGB', (width, height), color)
                            draw = ImageDraw.Draw(image)
                            # 添加简单的备份图标表示
                            draw.rectangle([(3, 3), (13, 13)], fill='white')
                            draw.rectangle([(6, 6), (10, 10)], fill=color)
                            return image
                        
                        color = 'green' if active else 'gray'
                        icon_image = create_image(color)
                        logging.warning(f"图标文件不存在，使用默认图标: {icon_path}")
                except Exception as e:
                    # 创建备用图标
                    def create_image(color):
                        width, height = 16, 16
                        image = Image.new('RGB', (width, height), color)
                        draw = ImageDraw.Draw(image)
                        # 添加简单的备份图标表示
                        draw.rectangle([(3, 3), (13, 13)], fill='white')
                        draw.rectangle([(6, 6), (10, 10)], fill=color)
                        return image
                    
                    color = 'green' if active else 'gray'
                    icon_image = create_image(color)
                    logging.error(f"加载图标文件失败，使用备用图标: {e}")
                
                # 定义处理函数
                def on_clicked(icon, item):
                    # 这个函数会在点击图标或菜单项时调用
                    if item is None:  # 双击托盘图标
                        self.root.after(0, self.restore_from_tray)
                
                # 创建菜单
                menu_items = [
                    pystray.MenuItem(tray_show_text, lambda: self.root.after(0, self.restore_from_tray)),
                    pystray.MenuItem(
                        tray_disable_auto_text if self.running else tray_enable_auto_text,
                        lambda: self.root.after(0, self.toggle_backup)
                    ),
                    pystray.Menu.SEPARATOR,
                    pystray.MenuItem(tray_exit_text, lambda: self.root.after(0, self.exit_application))
                ]
                
                # 创建托盘图标
                self.tray_icon = pystray.Icon(app_title, icon_image, app_title)
                self.tray_icon.menu = pystray.Menu(*menu_items)
                self.tray_icon.click = on_clicked  # 设置点击事件处理函数
                
                # 在单独线程中运行托盘图标
                import threading
                def run_tray():
                    try:
                        self.tray_icon.run()
                    except Exception as e:
                        logging.error(f"托盘图标运行错误: {e}")
                
                self.tray_thread = threading.Thread(target=run_tray, daemon=True)
                self.tray_thread.start()
                logging.info("使用pystray创建托盘图标成功")
                
            except ImportError:
                # 如果没有pystray库，回退到使用win32gui
                logging.warning("pystray库不可用，使用win32gui实现托盘图标")
                self._create_win32_tray_icon(active)
                
            except Exception as e:
                logging.error(f"创建pystray托盘图标失败: {e}")
                # 失败时回退到win32gui实现
                self._create_win32_tray_icon(active)
        except Exception as e:
            logging.error(f"创建/更新托盘图标失败: {e}")
    
    def _create_win32_tray_icon(self, active=False):
        """使用win32gui创建托盘图标"""
        try:
            import win32gui
            import win32con
            
            # 创建图标
            hicon = None
            try:
                # 尝试加载图标文件
                icon_path = self.active_icon if active else self.normal_icon
                if os.path.exists(icon_path):
                    hicon = win32gui.LoadImage(0, icon_path, win32con.IMAGE_ICON, 0, 0, 
                                             win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE)
            except Exception as e:
                logging.warning(f"加载图标文件失败: {e}")
            
            # 定义托盘图标常量
            self.TRAY_ICON_ID = 1001
            self.TRAY_CALLBACK_MESSAGE = win32con.WM_USER + 1
            
            # 确保语言字符串存在
            app_title = self.lang_strings.get('app_title', 'Auto Backup Tool')
            
            # 定义窗口类和回调函数
            wc = win32gui.WNDCLASS()
            wc.lpszClassName = "AutoBackupTrayIcon"
            wc.lpfnWndProc = self._tray_icon_callback
            wc.hInstance = win32gui.GetModuleHandle(None)
            
            try:
                # 尝试注册类，如果已存在则跳过错误
                self._class_atom = win32gui.RegisterClass(wc)
            except Exception:
                # 类已存在，获取类原子
                self._class_atom = win32gui.FindWindow(wc.lpszClassName, None)
                
            # 创建隐藏窗口
            self.tray_window = win32gui.CreateWindow(
                self._class_atom, "Auto Backup Tray Window", 0, 0, 0, 1, 1, 0, 0, wc.hInstance, None
            )
            
            # 添加托盘图标
            nid = (self.tray_window, self.TRAY_ICON_ID, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                  self.TRAY_CALLBACK_MESSAGE, hicon, app_title)
            
            # 检查是否已存在托盘图标
            if hasattr(self, 'tray_icon_info'):
                win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
                logging.info("托盘图标更新成功")
            else:
                win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
                logging.info("托盘图标创建成功")
            
            self.tray_icon_info = nid
        except Exception as e:
            logging.error(f"创建win32托盘图标失败: {e}")
    
    def _create_default_icon(self, active):
        """创建默认图标"""
        try:
            import win32gui
            import win32con
            
            # 创建一个简单的位图图标
            hdc = win32gui.GetDC(0)
            hdcMem = win32gui.CreateCompatibleDC(hdc)
            hbm = win32gui.CreateCompatibleBitmap(hdc, 16, 16)
            hbmOld = win32gui.SelectObject(hdcMem, hbm)
            
            # 填充背景
            color = 0x00FF00 if active else 0xAAAAAA  # 绿色或灰色
            brush = win32gui.CreateSolidBrush(color)
            win32gui.FillRect(hdcMem, (0, 0, 16, 16), brush)
            win32gui.DeleteObject(brush)
            
            # 创建图标信息结构
            import ctypes
            class ICONINFO(ctypes.Structure):
                _fields_ = [
                    ('fIcon', ctypes.c_bool),
                    ('xHotspot', ctypes.c_long),
                    ('yHotspot', ctypes.c_long),
                    ('hbmMask', ctypes.c_void_p),
                    ('hbmColor', ctypes.c_void_p)
                ]
            
            icon_info = ICONINFO()
            icon_info.fIcon = True
            icon_info.xHotspot = 0
            icon_info.yHotspot = 0
            icon_info.hbmMask = hbm
            icon_info.hbmColor = hbm
            
            # 创建图标
            hicon = win32gui.CreateIconIndirect(icon_info)
            
            # 清理
            win32gui.SelectObject(hdcMem, hbmOld)
            win32gui.DeleteObject(hbm)
            win32gui.DeleteDC(hdcMem)
            win32gui.ReleaseDC(0, hdc)
            
            return hicon
        except Exception as e:
            logging.error(f"创建默认图标失败: {e}")
            return None
    
    def _tray_icon_callback(self, hwnd, msg, wparam, lparam):
        """托盘图标消息处理回调函数"""
        import win32gui
        import win32con
        
        if msg == self.TRAY_CALLBACK_MESSAGE:
            if lparam == win32con.WM_LBUTTONDBLCLK:
                # 双击显示窗口
                self.root.after(0, self.restore_from_tray)
            elif lparam == win32con.WM_RBUTTONUP:
                # 右键显示菜单
                self.root.after(0, lambda: self._show_tray_menu(hwnd))
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    
    def _show_tray_menu(self, hwnd):
        """显示托盘右键菜单"""
        try:
            import win32gui
            import win32con
            
            # 创建弹出菜单
            menu = win32gui.CreatePopupMenu()
            
            # 添加菜单项
            show_id = 1001
            toggle_id = 1002
            exit_id = 1003
            
            # 确保语言字符串存在
            tray_show_text = self.lang_strings.get('tray_show', 'Show Window')
            tray_disable_auto_text = self.lang_strings.get('tray_disable_auto', 'Disable Auto Backup')
            tray_enable_auto_text = self.lang_strings.get('tray_enable_auto', 'Enable Auto Backup')
            tray_exit_text = self.lang_strings.get('tray_exit', 'Exit')
            
            # 添加菜单项
            win32gui.AppendMenu(menu, win32con.MF_STRING, show_id, tray_show_text)
            
            # 根据当前状态添加启用/禁用自动备份选项
            if self.running:
                win32gui.AppendMenu(menu, win32con.MF_STRING, toggle_id, tray_disable_auto_text)
            else:
                win32gui.AppendMenu(menu, win32con.MF_STRING, toggle_id, tray_enable_auto_text)
            
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, None)
            win32gui.AppendMenu(menu, win32con.MF_STRING, exit_id, tray_exit_text)
            
            # 显示菜单
            x, y = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(hwnd)
            
            # 显示菜单并获取选择结果
            cmd = win32gui.TrackPopupMenu(
                menu, win32con.TPM_LEFTALIGN | win32con.TPM_LEFTBUTTON | win32con.TPM_BOTTOMALIGN | win32con.TPM_RETURNCMD,
                x, y, 0, hwnd, None
            )
            
            # 处理选择结果
            if cmd == show_id:
                self.root.after(0, self.restore_from_tray)
            elif cmd == toggle_id:
                self.root.after(0, self.toggle_backup)
            elif cmd == exit_id:
                self.root.after(0, self.exit_application)
            
            # 确保菜单正确关闭
            win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
            
        except Exception as e:
            logging.error(f"显示托盘菜单失败: {e}")
    
    def minimize_to_tray(self):
        """最小化到托盘"""
        self.root.withdraw()
    
    def restore_from_tray(self):
        """从托盘恢复窗口"""
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_set()
            logging.info("窗口从托盘恢复成功")
        except Exception as e:
            logging.error(f"恢复窗口失败: {e}")
    
    def exit_application(self):
        """退出应用程序"""
        try:
            # 停止备份线程
            self.running = False
            self.stop_event.set()
            
            # 移除快捷键
            if self.hotkey_registered:
                try:
                    keyboard.remove_hotkey(self.hotkey)
                except:
                    pass
            
            # 移除托盘图标
            if hasattr(self, 'tray_icon'):
                try:
                    self.tray_icon.stop()
                except:
                    pass
            elif hasattr(self, 'tray_icon_info'):
                try:
                    win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, self.tray_icon_info)
                except:
                    pass
            
            # 销毁主窗口
            self.root.destroy()
        except Exception as e:
            logging.error(f"退出应用程序时出错: {e}")
            os._exit(0)
    
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
                
                # 加载新增功能设置
                if config.has_option('Settings', 'BackupPrefix'):
                    self.backup_prefix = config.get('Settings', 'BackupPrefix')
                
                if config.has_option('Settings', 'PrefixType'):
                    self.prefix_type = config.get('Settings', 'PrefixType')
                
                if config.has_option('Settings', 'CustomPrefix'):
                    self.custom_prefix = config.get('Settings', 'CustomPrefix')
                
                if config.has_option('Settings', 'SuffixType'):
                    self.suffix_type = config.get('Settings', 'SuffixType')
                
                if config.has_option('Settings', 'CustomSuffix'):
                    self.custom_suffix = config.get('Settings', 'CustomSuffix')
                
                if config.has_option('Settings', 'DuplicateHandling'):
                    self.duplicate_handling = config.get('Settings', 'DuplicateHandling')
                
                if config.has_option('Settings', 'BackupCounter'):
                    self.backup_counter = config.getint('Settings', 'BackupCounter')
                
                if config.has_option('Settings', 'StartNumber'):
                    self.start_number = config.getint('Settings', 'StartNumber')
                
                if self.hotkey_var.get():
                    self.toggle_hotkey()
                    
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
            'Language': self.current_lang,
            'BackupPrefix': self.backup_prefix,
            'PrefixType': self.prefix_type,
            'CustomPrefix': self.custom_prefix,
            'SuffixType': self.suffix_type,
            'CustomSuffix': self.custom_suffix,
            'DuplicateHandling': self.duplicate_handling,
            'BackupCounter': str(self.backup_counter),
            'StartNumber': str(self.start_number)
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
        self.root.after(100, self.update_lang_button_state)
        
        self.tab_control = ttk.Notebook(self.root)
        tab_backup = ttk.Frame(self.tab_control)
        tab_manage = ttk.Frame(self.tab_control)
        tab_settings = ttk.Frame(self.tab_control)  # 新增高级设置标签页
        self.tab_control.add(tab_backup, text=self.lang_strings['tab_backup'])
        self.tab_control.add(tab_manage, text=self.lang_strings['tab_manage'])
        self.tab_control.add(tab_settings, text=self.lang_strings['tab_settings'])
        self.tab_control.pack(expand=1, fill="both", padx=10, pady=5)
        
        # 备份设置标签页
        self.frame_settings = ttk.LabelFrame(tab_backup, text=self.lang_strings['frame_settings'])
        self.frame_settings.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 配置网格权重，使界面可缩放
        self.frame_settings.columnconfigure(1, weight=1)
        self.frame_settings.rowconfigure(0, weight=1)  # 使源路径列表可扩展
        
        # 创建主框架以更好地利用空间
        main_frame = ttk.Frame(self.frame_settings)
        main_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # 源路径设置
        self.source_frame = ttk.LabelFrame(main_frame, text=self.lang_strings['source_paths'])
        self.source_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        self.source_frame.columnconfigure(1, weight=1)
        self.source_frame.rowconfigure(0, weight=1)
        
        self.source_listbox = tk.Listbox(self.source_frame, width=50, height=8)  # 增加高度以利用空间
        self.source_listbox.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        self.source_btn_frame = ttk.Frame(self.source_frame)
        self.source_btn_frame.grid(row=0, column=2, padx=5, pady=5, sticky="nw")
        
        self.btn_add_file = ttk.Button(self.source_btn_frame, text=self.lang_strings['add_file'], command=lambda: self.add_source('file'))
        self.btn_add_file.pack(fill="x", padx=2, pady=2)
        self.btn_add_dir = ttk.Button(self.source_btn_frame, text=self.lang_strings['add_dir'], command=lambda: self.add_source('dir'))
        self.btn_add_dir.pack(fill="x", padx=2, pady=2)
        self.btn_remove_source = ttk.Button(self.source_btn_frame, text=self.lang_strings['remove'], command=self.remove_source)
        self.btn_remove_source.pack(fill="x", padx=2, pady=2)
        self.btn_open_source = ttk.Button(self.source_btn_frame, text=self.lang_strings['open_location'], command=self.open_source_location)
        self.btn_open_source.pack(fill="x", padx=2, pady=2)
        
        # 备份目录设置
        self.dest_frame = ttk.LabelFrame(main_frame, text=self.lang_strings['dest_dir'])
        self.dest_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        self.dest_frame.columnconfigure(1, weight=1)
        
        self.dest_entry = ttk.Entry(self.dest_frame)
        self.dest_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.dest_entry.insert(0, self.dest_dir)
        
        self.dest_btn_frame = ttk.Frame(self.dest_frame)
        self.dest_btn_frame.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.btn_browse_dest = ttk.Button(self.dest_btn_frame, text=self.lang_strings['browse'], command=self.select_dest_dir)
        self.btn_browse_dest.pack(fill="x", padx=2, pady=2)
        self.btn_open_dest = ttk.Button(self.dest_btn_frame, text=self.lang_strings['open_location'], command=self.open_dest_location)
        self.btn_open_dest.pack(fill="x", padx=2, pady=2)
        
        # 设置框架
        settings_frame = ttk.LabelFrame(main_frame, text=self.lang_strings['backup_options'])
        settings_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        settings_frame.columnconfigure(1, weight=1)
        
        # 快捷键设置
        hotkey_frame = ttk.Frame(settings_frame)
        hotkey_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        
        self.hotkey_label = ttk.Label(hotkey_frame, text=self.lang_strings['hotkey'])
        self.hotkey_label.pack(side="left", padx=5, pady=5)
        
        self.hotkey_entry = ttk.Entry(hotkey_frame, width=15)
        self.hotkey_entry.insert(0, self.hotkey)
        self.hotkey_entry.pack(side="left", padx=5, pady=5)
        
        self.check_hotkey = ttk.Checkbutton(hotkey_frame, text=self.lang_strings['enable_hotkey'], 
                        variable=self.hotkey_var, command=self.toggle_hotkey)
        self.check_hotkey.pack(side="left", padx=5, pady=5)
        
        # 时间间隔设置
        interval_frame = ttk.Frame(settings_frame)
        interval_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        
        self.interval_label = ttk.Label(interval_frame, text=self.lang_strings['interval'])
        self.interval_label.pack(side="left", padx=5, pady=5)
        
        self.interval_entry = ttk.Entry(interval_frame, width=15)
        self.interval_entry.insert(0, str(self.interval))
        self.interval_entry.pack(side="left", padx=5, pady=5)
        
        # 最大备份数设置
        max_backups_frame = ttk.Frame(settings_frame)
        max_backups_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        
        self.max_backups_label = ttk.Label(max_backups_frame, text=self.lang_strings['max_backups'])
        self.max_backups_label.pack(side="left", padx=5, pady=5)
        
        self.max_backups_entry = ttk.Entry(max_backups_frame, width=15)
        self.max_backups_entry.insert(0, str(self.max_backups))
        self.max_backups_entry.pack(side="left", padx=5, pady=5)
        
        # 备份按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=10, sticky="ew")
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        self.start_btn = ttk.Button(button_frame, text=self.lang_strings['start_backup'], command=self.toggle_backup)
        self.start_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.btn_manual_backup = ttk.Button(button_frame, text=self.lang_strings['manual_backup'], command=self.manual_backup)
        self.btn_manual_backup.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # 备份管理标签页
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
        
        # 应用字体缩放到表格
        self._apply_font_scaling_to_treeview()
        
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
        
        # 添加重命名按钮
        self.btn_rename = ttk.Button(btn_frame, text=self.lang_strings['rename_backup'], command=self.rename_selected_backup)
        self.btn_rename.pack(side="left", padx=5, pady=5)
        
        # 高级设置标签页
        settings_frame = ttk.Frame(tab_settings)
        settings_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 外观设置
        self.frame_appearance = ttk.LabelFrame(settings_frame, text=self.lang_strings['appearance'])
        self.frame_appearance.pack(fill="x", padx=10, pady=5)
        
        # 配置网格权重，使界面可缩放
        self.frame_appearance.columnconfigure(1, weight=1)
        
        # 备份选项设置
        self.frame_backup_options = ttk.LabelFrame(settings_frame, text=self.lang_strings['backup_options'])
        self.frame_backup_options.pack(fill="x", padx=10, pady=5)
        
        # 配置网格权重，使界面可缩放
        self.frame_backup_options.columnconfigure(1, weight=1)
        
        # 备份前缀设置
        self.backup_prefix_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['backup_prefix'])
        self.backup_prefix_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        prefix_frame = ttk.Frame(self.frame_backup_options)
        prefix_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        self.prefix_type_var = tk.StringVar(value=self.prefix_type)
        self.prefix_disabled_radio = ttk.Radiobutton(prefix_frame, text=self.lang_strings['prefix_disabled'], 
                                                      variable=self.prefix_type_var, value="disabled")
        self.prefix_disabled_radio.pack(side="left", padx=5)
        
        self.prefix_number_radio = ttk.Radiobutton(prefix_frame, text=self.lang_strings['prefix_number'], 
                                                   variable=self.prefix_type_var, value="number")
        self.prefix_number_radio.pack(side="left", padx=5)
        
        self.prefix_custom_radio = ttk.Radiobutton(prefix_frame, text=self.lang_strings['prefix_custom'], 
                                                   variable=self.prefix_type_var, value="custom")
        self.prefix_custom_radio.pack(side="left", padx=5)
        
        # 自定义前缀设置
        self.custom_prefix_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['custom_prefix'])
        self.custom_prefix_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.custom_prefix_entry = ttk.Entry(self.frame_backup_options)
        self.custom_prefix_entry.insert(0, self.custom_prefix)
        self.custom_prefix_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")
        
        # 后缀类型设置
        self.suffix_type_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['suffix_type'])
        self.suffix_type_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        
        suffix_frame = ttk.Frame(self.frame_backup_options)
        suffix_frame.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        self.suffix_type_var = tk.StringVar(value=self.suffix_type)
        self.suffix_timestamp_radio = ttk.Radiobutton(suffix_frame, text=self.lang_strings['suffix_timestamp'], 
                                                      variable=self.suffix_type_var, value="timestamp")
        self.suffix_timestamp_radio.pack(side="left", padx=5)
        
        self.suffix_number_radio = ttk.Radiobutton(suffix_frame, text=self.lang_strings['suffix_number'], 
                                                   variable=self.suffix_type_var, value="number")
        self.suffix_number_radio.pack(side="left", padx=5)
        
        self.suffix_custom_radio = ttk.Radiobutton(suffix_frame, text=self.lang_strings['suffix_custom'], 
                                                   variable=self.suffix_type_var, value="custom")
        self.suffix_custom_radio.pack(side="left", padx=5)
        
        # 自定义后缀设置
        self.custom_suffix_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['custom_suffix'])
        self.custom_suffix_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.custom_suffix_entry = ttk.Entry(self.frame_backup_options)
        self.custom_suffix_entry.insert(0, self.custom_suffix)
        self.custom_suffix_entry.grid(row=3, column=1, padx=5, pady=5, sticky="we")
        
        # 重名文件处理设置
        self.duplicates_label = ttk.Label(self.frame_backup_options, text=self.lang_strings['handle_duplicates'])
        self.duplicates_label.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        
        duplicates_frame = ttk.Frame(self.frame_backup_options)
        duplicates_frame.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        
        self.duplicates_var = tk.StringVar(value=self.duplicate_handling)
        self.duplicates_overwrite_radio = ttk.Radiobutton(duplicates_frame, text=self.lang_strings['duplicates_overwrite'], 
                                                          variable=self.duplicates_var, value="overwrite")
        self.duplicates_overwrite_radio.pack(side="left", padx=5)
        
        self.duplicates_rename_radio = ttk.Radiobutton(duplicates_frame, text=self.lang_strings['duplicates_rename'], 
                                                       variable=self.duplicates_var, value="rename")
        self.duplicates_rename_radio.pack(side="left", padx=5)
        
        self.duplicates_skip_radio = ttk.Radiobutton(duplicates_frame, text=self.lang_strings['duplicates_skip'], 
                                                     variable=self.duplicates_var, value="skip")
        self.duplicates_skip_radio.pack(side="left", padx=5)
        
        # 备份序号
        counter_frame = ttk.Frame(self.frame_backup_options)
        counter_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="we")
        
        self.backup_count_label = ttk.Label(counter_frame, text=self.lang_strings['backup_count'].format(count=self.backup_counter))
        self.backup_count_label.pack(side="left", padx=5)
        
        self.btn_set_start_number = ttk.Button(counter_frame, text=self.lang_strings['set_start_number'], command=self.set_start_number)
        self.btn_set_start_number.pack(side="left", padx=5)
        
        # 保存设置按钮
        save_btn_frame = ttk.Frame(tab_settings)
        save_btn_frame.pack(fill="x", padx=10, pady=10)
        
        self.btn_save_settings = ttk.Button(save_btn_frame, text="保存设置" if self.current_lang == 'zh' else "Save Settings", command=self.save_advanced_settings)
        self.btn_save_settings.pack(side="left", padx=5, pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set(self.lang_strings['status_ready'])
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")
        
        self.update_source_list()
    
    def _apply_font_scaling_to_treeview(self):
        """应用字体缩放到表格控件"""
        try:
            # 调整表格行高以适应字体大小
            row_height = int(FONT_CONFIG['base_size'] * 2)
            style = ttk.Style()
            style.configure('Treeview', rowheight=row_height)
        except Exception as e:
            logging.error(f"应用字体缩放到表格失败: {e}")
    
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
            except Exception as e:
                logging.error(f"注册快捷键失败: {e}")
                messagebox.showerror("Error", self.lang_strings['error_hotkey_reg'].format(error=str(e)))
                self.hotkey_var.set(False)
                self.hotkey_registered = False
        else:
            if self.hotkey_registered:
                try:
                    keyboard.remove_hotkey(self.hotkey)
                except:
                    pass
                self.hotkey_registered = False
                self.status_var.set(self.lang_strings['hotkey_disabled'])
    
    def toggle_backup(self):
        if self.running:
            # 停止备份
            self.running = False
            self.stop_event.set()
            self.start_btn.config(text=self.lang_strings['start_backup'])
            self.status_var.set(self.lang_strings['backup_stopped'])
            self.set_icon(active=False)
            self.create_tray_icon(active=False)
        else:
            # 开始备份
            # 检查源路径和目标路径
            if not self.source_paths:
                messagebox.showerror("Error", self.lang_strings['error_no_source'])
                return
            
            if not self.dest_dir:
                messagebox.showerror("Error", self.lang_strings['error_no_dest'])
                return
            
            # 检查时间间隔
            try:
                interval = int(self.interval_entry.get())
                if interval <= 0:
                    messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                    return
                self.interval = interval
            except ValueError:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            
            # 检查最大备份数
            try:
                max_backups = int(self.max_backups_entry.get())
                if max_backups <= 0:
                    messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                    return
                self.max_backups = max_backups
            except ValueError:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            
            # 保存配置
            self.save_config()
            
            # 开始备份线程
            self.running = True
            self.stop_event.clear()
            self.start_btn.config(text=self.lang_strings['stop_backup'])
            self.status_var.set(self.lang_strings['backup_started'].format(interval=self.interval))
            self.set_icon(active=True)
            self.create_tray_icon(active=True)
            
            self.backup_thread = threading.Thread(target=self.backup_loop)
            self.backup_thread.daemon = True
            self.backup_thread.start()
    
    def backup_loop(self):
        """自动备份循环"""
        while self.running and not self.stop_event.is_set():
            try:
                self.perform_backup()
                
                # 等待指定时间间隔或直到停止事件被触发
                if self.stop_event.wait(self.interval):
                    break
                    
            except Exception as e:
                logging.error(f"备份过程中出错: {e}")
                self.status_var.set(self.lang_strings['error_backup'].format(error=str(e)))
                
                # 出错后等待较短时间再继续
                if self.stop_event.wait(10):
                    break
        
        if self.running:
            self.running = False
            self.root.after(0, lambda: self.start_btn.config(text=self.lang_strings['start_backup']))
            self.root.after(0, lambda: self.status_var.set(self.lang_strings['backup_stopped']))
            self.root.after(0, lambda: self.set_icon(active=False))
            self.root.after(0, lambda: self.create_tray_icon(active=False))
    
    def get_backup_prefix(self):
        """根据设置获取备份前缀"""
        prefix = ""
        if self.prefix_type == "number":
            prefix = f"{self.backup_counter:03d}_"
        elif self.prefix_type == "custom" and self.custom_prefix:
            prefix = f"{self.custom_prefix}_"
        return prefix
    
    def get_backup_suffix(self):
        """根据设置获取备份后缀"""
        suffix = ""
        if self.suffix_type == "timestamp":
            suffix = datetime.datetime.now().strftime("_%Y%m%d_%H%M%S")
        elif self.suffix_type == "number":
            suffix = f"_v{self.backup_counter:03d}"
        elif self.suffix_type == "custom" and self.custom_suffix:
            suffix = f"_{self.custom_suffix}"
        return suffix
    
    def perform_backup(self):
        """执行备份操作"""
        if not self.source_paths or not self.dest_dir:
            return
        
        # 创建备份目录（如果不存在）
        if not os.path.exists(self.dest_dir):
            try:
                os.makedirs(self.dest_dir)
            except Exception as e:
                logging.error(f"创建备份目录失败: {e}")
                raise
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 获取前缀和后缀
        prefix = self.get_backup_prefix()
        suffix = self.get_backup_suffix()
        
        # 执行备份
        try:
            for source_path in self.source_paths:
                source_name = os.path.basename(source_path)
                
                # 构建备份文件名或目录名
                backup_name = f"{prefix}{source_name}{suffix}"
                backup_path = os.path.join(self.dest_dir, backup_name)
                
                # 检查备份是否已存在
                if os.path.exists(backup_path):
                    if self.duplicate_handling == "overwrite":
                        if os.path.isdir(backup_path):
                            shutil.rmtree(backup_path)
                        else:
                            os.remove(backup_path)
                    elif self.duplicate_handling == "rename":
                        # 添加额外的时间戳以避免覆盖
                        unique_suffix = datetime.datetime.now().strftime("_%Y%m%d_%H%M%S_%f")[:-3]
                        backup_name = f"{prefix}{source_name}{suffix}{unique_suffix}"
                        backup_path = os.path.join(self.dest_dir, backup_name)
                    elif self.duplicate_handling == "skip":
                        logging.info(f"跳过备份，文件已存在: {backup_path}")
                        continue
                
                # 执行复制操作
                if os.path.isdir(source_path):
                    shutil.copytree(source_path, backup_path)
                else:
                    shutil.copy2(source_path, backup_path)
                
            # 备份成功后递增计数器
            self.backup_counter += 1
            
            # 清理旧备份
            self.cleanup_old_backups()
            
            # 更新状态
            self.status_var.set(self.lang_strings['backup_success'].format(timestamp=timestamp))
            
            # 保存配置（包含更新后的计数器）
            self.save_config()
            
            # 更新备份列表
            self.root.after(0, self.update_backup_list)
            
        except Exception as e:
            logging.error(f"备份过程中出错: {e}")
            raise
    
    def manual_backup(self):
        """执行手动备份"""
        if not self.source_paths or not self.dest_dir:
            messagebox.showerror("Error", self.lang_strings['error_no_source'] if not self.source_paths else self.lang_strings['error_no_dest'])
            return
        
        try:
            self.status_var.set(self.lang_strings['manual_backup_running'])
            self.perform_backup()
            self.status_var.set(self.lang_strings['manual_backup_done'])
        except Exception as e:
            logging.error(f"手动备份失败: {e}")
            messagebox.showerror("Error", self.lang_strings['error_backup'].format(error=str(e)))
            self.status_var.set(self.lang_strings['status_ready'])
    
    def cleanup_old_backups(self):
        """清理超过最大备份数量的旧备份"""
        try:
            # 获取所有备份文件/目录，并按修改时间排序
            backup_items = []
            for item in os.listdir(self.dest_dir):
                item_path = os.path.join(self.dest_dir, item)
                if os.path.isfile(item_path) or os.path.isdir(item_path):
                    # 获取修改时间（使用文件的创建时间或修改时间）
                    mtime = os.path.getmtime(item_path)
                    backup_items.append((mtime, item_path))
            
            # 按时间排序（最新的在前）
            backup_items.sort(reverse=True, key=lambda x: x[0])
            
            # 删除超过最大备份数量的旧备份
            if len(backup_items) > self.max_backups:
                for _, item_path in backup_items[self.max_backups:]:
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                    else:
                        os.remove(item_path)
        except Exception as e:
            logging.error(f"清理旧备份时出错: {e}")
            raise
    
    def update_backup_list(self):
        """更新备份列表"""
        # 清空现有列表
        for item in self.backup_list.get_children():
            self.backup_list.delete(item)
        
        if not self.dest_dir or not os.path.exists(self.dest_dir):
            return
        
        try:
            backup_items = []
            
            # 遍历备份目录
            for item in os.listdir(self.dest_dir):
                item_path = os.path.join(self.dest_dir, item)
                if os.path.isfile(item_path) or os.path.isdir(item_path):
                    # 获取创建时间
                    ctime = os.path.getctime(item_path)
                    timestamp = datetime.datetime.fromtimestamp(ctime).strftime("%Y-%m-%d %H:%M:%S")
                    
                    # 尝试解析原始文件名
                    original_name = item
                    # 移除前缀（如果有）
                    if hasattr(self, 'prefix_type') and self.prefix_type == "number":
                        match = re.match(r'^\d+_(.+)', item)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'prefix_type') and self.prefix_type == "custom" and hasattr(self, 'custom_prefix') and self.custom_prefix:
                        prefix = f"{self.custom_prefix}_"
                        if item.startswith(prefix):
                            original_name = item[len(prefix):]
                    
                    # 移除后缀（如果有）
                    if hasattr(self, 'suffix_type') and self.suffix_type == "timestamp":
                        # 匹配时间戳格式 _YYYYMMDD_HHMMSS
                        match = re.match(r'(.+)_\d{8}_\d{6}$', original_name)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'suffix_type') and self.suffix_type == "number":
                        # 匹配版本号格式 _vXXX
                        match = re.match(r'(.+)_v\d+$', original_name)
                        if match:
                            original_name = match.group(1)
                    elif hasattr(self, 'suffix_type') and self.suffix_type == "custom" and hasattr(self, 'custom_suffix') and self.custom_suffix:
                        suffix = f"_{self.custom_suffix}"
                        if original_name.endswith(suffix):
                            original_name = original_name[:-len(suffix)]
                    
                    backup_items.append((-ctime, timestamp, item_path, original_name))  # 使用负时间戳以便升序排列时最新的在前
            
            # 按时间排序（最新的在前）
            backup_items.sort(key=lambda x: x[0])
            
            # 添加到列表
            for _, timestamp, path, original_name in backup_items:
                self.backup_list.insert("", "end", values=(timestamp, path, original_name))
        except Exception as e:
            logging.error(f"更新备份列表时出错: {e}")
    
    def delete_selected(self):
        """删除选中的备份"""
        selected = self.backup_list.selection()
        if not selected:
            messagebox.showinfo("Info", self.lang_strings['select_backup'])
            return
        
        try:
            for item in selected:
                values = self.backup_list.item(item, "values")
                if values and len(values) >= 2:
                    backup_path = values[1]
                    if os.path.exists(backup_path):
                        if os.path.isdir(backup_path):
                            shutil.rmtree(backup_path)
                        else:
                            os.remove(backup_path)
                        logging.info(f"已删除备份: {os.path.basename(backup_path)}")
                        self.status_var.set(self.lang_strings['deleted_backup'].format(name=os.path.basename(backup_path)))
            
            self.update_backup_list()
        except Exception as e:
            logging.error(f"删除备份时出错: {e}")
            messagebox.showerror("Error", self.lang_strings['error_delete'].format(error=str(e)))
    
    def restore_selected(self):
        """还原选中的备份"""
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
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        
        backup_path = values[1]
        original_name = values[2]
        
        if not os.path.exists(backup_path):
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        
        if not messagebox.askyesno("Confirm", self.lang_strings['confirm_restore'].format(name=original_name)):
            return
        
        try:
            # 查找原始文件位置
            source_found = False
            for source_path in self.source_paths:
                if os.path.basename(source_path) == original_name:
                    # 删除原始文件或目录
                    if os.path.exists(source_path):
                        if os.path.isdir(source_path):
                            shutil.rmtree(source_path)
                        else:
                            os.remove(source_path)
                    
                    # 还原备份
                    if os.path.isdir(backup_path):
                        shutil.copytree(backup_path, source_path)
                    else:
                        shutil.copy2(backup_path, source_path)
                    
                    source_found = True
                    break
            
            if not source_found:
                messagebox.showerror("Error", self.lang_strings['error_no_source_path'].format(name=original_name))
                return
            
            self.status_var.set(self.lang_strings['restored_backup'].format(name=original_name))
            messagebox.showinfo("Success", self.lang_strings['restore_success'])
        except Exception as e:
            logging.error(f"还原备份时出错: {e}")
            messagebox.showerror("Error", self.lang_strings['error_restore'].format(error=str(e)))
    
    def rename_selected_backup(self):
        """重命名选中的备份"""
        selected = self.backup_list.selection()
        if not selected:
            messagebox.showinfo("Info", self.lang_strings['select_backup'])
            return
        
        if len(selected) > 1:
            messagebox.showinfo("Info", self.lang_strings['select_one_backup'])
            return
        
        item = selected[0]
        values = self.backup_list.item(item, "values")
        if not values or len(values) < 2:
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        
        backup_path = values[1]
        if not os.path.exists(backup_path):
            messagebox.showerror("Error", self.lang_strings['error_backup_content'])
            return
        
        # 获取当前备份名称
        current_name = os.path.basename(backup_path)
        
        # 弹出输入对话框
        new_name = tk.simpledialog.askstring(self.lang_strings['rename_backup'], 
                                             self.lang_strings['new_backup_name'], 
                                             initialvalue=current_name)
        
        if not new_name or new_name == current_name:
            return
        
        try:
            # 检查新名称是否已存在
            new_path = os.path.join(os.path.dirname(backup_path), new_name)
            if os.path.exists(new_path):
                if not messagebox.askyesno("Confirm", self.lang_strings['confirm_overwrite']):
                    return
                
                # 删除已存在的文件/目录
                if os.path.isdir(new_path):
                    shutil.rmtree(new_path)
                else:
                    os.remove(new_path)
            
            # 重命名备份
            os.rename(backup_path, new_path)
            
            self.status_var.set(self.lang_strings['rename_success'])
            self.update_backup_list()
        except Exception as e:
            logging.error(f"重命名备份时出错: {e}")
            messagebox.showerror("Error", self.lang_strings['error_rename'].format(error=str(e)))
    
    def save_advanced_settings(self):
        """保存高级设置"""
        try:
            # 验证自定义前缀
            if self.prefix_type_var.get() == "custom" and not self.custom_prefix_entry.get().strip():
                messagebox.showerror("Error", self.lang_strings['error_prefix_empty'])
                return
            
            # 验证自定义后缀
            if self.suffix_type_var.get() == "custom" and not self.custom_suffix_entry.get().strip():
                messagebox.showerror("Error", self.lang_strings['error_suffix_empty'])
                return
            
            # 更新设置
            self.prefix_type = self.prefix_type_var.get()
            self.custom_prefix = self.custom_prefix_entry.get().strip()
            self.suffix_type = self.suffix_type_var.get()
            self.custom_suffix = self.custom_suffix_entry.get().strip()
            self.duplicate_handling = self.duplicates_var.get()
            
            # 保存配置
            self.save_config()
            
            # 更新界面文本
            self.update_ui_texts()
            
            messagebox.showinfo("Info", self.lang_strings['save_settings'])
        except Exception as e:
            logging.error(f"保存高级设置时出错: {e}")
            messagebox.showerror("Error", "保存设置时出错")
    
    def set_start_number(self):
        """设置起始序号"""
        try:
            # 弹出输入对话框
            new_number_str = tk.simpledialog.askstring(self.lang_strings['set_start_number'], 
                                                       self.lang_strings['confirm_set_start_number'], 
                                                       initialvalue=str(self.start_number))
            
            if not new_number_str:
                return
            
            new_number = int(new_number_str)
            if new_number < 1:
                messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
                return
            
            # 更新起始序号和当前计数器
            self.start_number = new_number
            self.backup_counter = new_number
            
            # 更新界面显示
            self.backup_count_label.config(text=self.lang_strings['backup_count'].format(count=self.backup_counter))
            
            # 保存配置
            self.save_config()
            
            messagebox.showinfo("Info", "起始序号已更新")
        except ValueError:
            messagebox.showerror("Error", self.lang_strings['error_invalid_number'])
        except Exception as e:
            logging.error(f"设置起始序号时出错: {e}")
            messagebox.showerror("Error", "设置起始序号时出错")

def main():
    root = tk.Tk()
    app = AutoBackupTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
