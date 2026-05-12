import json
import os
import shutil
import subprocess
import sys
import threading
import time
import uuid
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import keyboard  # type: ignore
except Exception:
    keyboard = None

try:
    from PIL import Image, ImageTk  # type: ignore
except Exception:
    Image = None
    ImageTk = None


APP_TITLE = "Local Auto Backup Tool - Modern"
DATA_FILE = "profiles_config.json"
APP_ROOT = Path(__file__).resolve().parent
RESOURCES_DIR = APP_ROOT / "resources"
DEFAULT_ICON_PATH = RESOURCES_DIR / "default.ico"
WINDOW_ICON_PATH = RESOURCES_DIR / "folder-sync.ico"


def format_size(num_bytes: int) -> str:
    size = max(0, num_bytes)
    units = ["B", "KB", "MB", "GB", "TB"]
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.2f} {unit}"
        size /= 1024


def open_path(path: Path) -> None:
    if not path.exists():
        return
    if os.name == "nt":
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def safe_name(value: str) -> str:
    invalid = '\\/:*?"<>|'
    cleaned = "".join("_" if c in invalid else c for c in value).strip()
    return cleaned or "backup"


def is_hidden(path: Path) -> bool:
    name = path.name
    if name.startswith("."):
        return True
    if os.name == "nt":
        try:
            import ctypes

            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
            if attrs == -1:
                return False
            return bool(attrs & 0x2)
        except Exception:
            return False
    return False


def copy_file(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def copy_dir(src: Path, dst: Path, skip_hidden: bool) -> None:
    if dst.exists():
        shutil.rmtree(dst, ignore_errors=True)
    for root, dirs, files in os.walk(src):
        root_path = Path(root)
        rel = root_path.relative_to(src)
        if skip_hidden and is_hidden(root_path) and rel != Path("."):
            dirs[:] = []
            continue
        target_root = dst / rel
        target_root.mkdir(parents=True, exist_ok=True)
        filtered_dirs = []
        for d in dirs:
            d_path = root_path / d
            if skip_hidden and is_hidden(d_path):
                continue
            filtered_dirs.append(d)
        dirs[:] = filtered_dirs
        for f in files:
            src_file = root_path / f
            if skip_hidden and is_hidden(src_file):
                continue
            copy_file(src_file, target_root / f)


def remove_path(path: Path) -> None:
    if path.is_dir():
        shutil.rmtree(path, ignore_errors=True)
    elif path.exists():
        path.unlink(missing_ok=True)


@dataclass
class Profile:
    id: str
    name: str
    icon_path: str
    enabled: bool
    source_paths: list
    dest_dir: str
    app_path: str
    interval_minutes: int
    max_backups: int
    skip_hidden: bool
    suffix_type: str
    custom_suffix: str
    number_mode: str
    start_number: int
    counter: int
    duplicate_handling: str

    @staticmethod
    def default() -> "Profile":
        return Profile(
            id=str(uuid.uuid4()),
            name="新建配置",
            icon_path=str(DEFAULT_ICON_PATH),
            enabled=False,
            source_paths=[],
            dest_dir="",
            app_path="",
            interval_minutes=10,
            max_backups=5,
            skip_hidden=False,
            suffix_type="number",
            custom_suffix="",
            number_mode="auto",
            start_number=1,
            counter=1,
            duplicate_handling="rename",
        )


class ConfigStore:
    def __init__(self, path: Path):
        self.path = path

    def load(self):
        if not self.path.exists():
            return {
                "global": {
                    "backup_hotkey": "ctrl+f1",
                    "backup_hotkey_enabled": False,
                    "restore_hotkey": "ctrl+f2",
                    "restore_hotkey_enabled": False,
                    "close_behavior": "minimize_to_tray",
                    "language": "zh",
                },
                "profiles": [asdict(Profile.default())],
            }
        with open(self.path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("global", {})
        data.setdefault("profiles", [])
        if not data["profiles"]:
            data["profiles"] = [asdict(Profile.default())]
        g = data["global"]
        g.setdefault("backup_hotkey", "ctrl+f1")
        g.setdefault("backup_hotkey_enabled", False)
        g.setdefault("restore_hotkey", "ctrl+f2")
        g.setdefault("restore_hotkey_enabled", False)
        g.setdefault("close_behavior", "minimize_to_tray")
        g.setdefault("language", "zh")
        return data

    def save(self, data):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


class ProfileRunner:
    def __init__(self, manager, profile_id: str):
        self.manager = manager
        self.profile_id = profile_id
        self.stop_event = threading.Event()
        self.thread = threading.Thread(target=self.run, daemon=True)

    def start(self):
        self.thread.start()

    def stop(self):
        self.stop_event.set()

    def run(self):
        while not self.stop_event.is_set():
            profile = self.manager.get_profile(self.profile_id)
            if not profile:
                return
            try:
                self.manager.perform_backup(profile)
                self.manager.set_profile_status(profile.id, "最近备份成功")
            except Exception as e:
                self.manager.set_profile_status(profile.id, f"备份失败: {e}")
            seconds = max(1, int(profile.interval_minutes) * 60)
            if self.stop_event.wait(seconds):
                break


class ProfileWindow:
    def __init__(self, manager, profile_id: str):
        self.manager = manager
        self.profile_id = profile_id
        self.window = tk.Toplevel(manager.root)
        self.window.title("配置管理")
        self.window.geometry("980x680")
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        self.profile = self.manager.get_profile(profile_id)
        if not self.profile:
            raise RuntimeError("配置不存在")
        self.create_widgets()
        self.refresh_from_profile()
        self.refresh_backup_list()

    def on_close(self):
        self.manager.profile_windows.pop(self.profile_id, None)
        self.window.destroy()

    def create_widgets(self):
        self.tabs = ttk.Notebook(self.window)
        self.tab_settings = ttk.Frame(self.tabs)
        self.tab_manage = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_settings, text="备份设置")
        self.tabs.add(self.tab_manage, text="备份管理")
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)

        s = self.tab_settings
        s.columnconfigure(0, weight=1)
        s.columnconfigure(1, weight=1)

        top = ttk.Frame(s)
        top.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=6)
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text="配置名称").grid(row=0, column=0, sticky="w", padx=4)
        self.name_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.name_var).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(top, text="应用名称", command=self.apply_name).grid(row=0, column=2, padx=4)

        src_frame = ttk.LabelFrame(s, text="源路径管理")
        src_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=8, pady=6)
        src_frame.columnconfigure(0, weight=1)
        self.source_list = tk.Listbox(src_frame, height=8, selectmode=tk.EXTENDED)
        self.source_list.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        src_btns = ttk.Frame(src_frame)
        src_btns.grid(row=0, column=1, sticky="ns", padx=4, pady=6)
        ttk.Button(src_btns, text="添加文件", command=self.add_file).pack(fill="x", pady=2)
        ttk.Button(src_btns, text="添加目录", command=self.add_dir).pack(fill="x", pady=2)
        ttk.Button(src_btns, text="移除", command=self.remove_sources).pack(fill="x", pady=2)
        ttk.Button(src_btns, text="打开位置", command=self.open_selected_source).pack(fill="x", pady=2)

        dest_frame = ttk.LabelFrame(s, text="备份目录")
        dest_frame.grid(row=2, column=0, sticky="ew", padx=8, pady=6)
        dest_frame.columnconfigure(0, weight=1)
        self.dest_var = tk.StringVar()
        ttk.Entry(dest_frame, textvariable=self.dest_var).grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        ttk.Button(dest_frame, text="设置备份路径", command=self.pick_dest).grid(row=0, column=1, padx=4)
        ttk.Button(dest_frame, text="打开位置", command=self.open_dest).grid(row=0, column=2, padx=4)

        app_frame = ttk.LabelFrame(s, text="关联应用程序")
        app_frame.grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        app_frame.columnconfigure(0, weight=1)
        self.app_var = tk.StringVar()
        ttk.Entry(app_frame, textvariable=self.app_var).grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        ttk.Button(app_frame, text="选择 EXE", command=self.pick_app).grid(row=0, column=1, padx=4)
        ttk.Button(app_frame, text="启动应用", command=self.launch_app).grid(row=0, column=2, padx=4)

        opts = ttk.LabelFrame(s, text="自动备份选项")
        opts.grid(row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=6)
        ttk.Label(opts, text="间隔(分钟)").grid(row=0, column=0, padx=4, pady=5, sticky="e")
        self.interval_var = tk.StringVar()
        ttk.Entry(opts, width=12, textvariable=self.interval_var).grid(row=0, column=1, padx=4, pady=5, sticky="w")
        ttk.Label(opts, text="最大备份数").grid(row=0, column=2, padx=4, pady=5, sticky="e")
        self.max_var = tk.StringVar()
        ttk.Entry(opts, width=12, textvariable=self.max_var).grid(row=0, column=3, padx=4, pady=5, sticky="w")
        self.skip_hidden_var = tk.BooleanVar(value=self.profile.skip_hidden)
        ttk.Checkbutton(opts, text="跳过隐藏文件/文件夹", variable=self.skip_hidden_var).grid(row=0, column=4, padx=4, pady=5, sticky="w")

        ttk.Label(opts, text="后缀类型").grid(row=1, column=0, padx=4, pady=5, sticky="e")
        self.suffix_var = tk.StringVar()
        suffix_box = ttk.Combobox(opts, width=12, textvariable=self.suffix_var, state="readonly", values=["timestamp", "number", "custom"])
        suffix_box.grid(row=1, column=1, padx=4, pady=5, sticky="w")
        suffix_box.bind("<<ComboboxSelected>>", lambda *_: self.update_custom_state())
        ttk.Label(opts, text="自定义后缀").grid(row=1, column=2, padx=4, pady=5, sticky="e")
        self.custom_var = tk.StringVar()
        self.custom_entry = ttk.Entry(opts, width=16, textvariable=self.custom_var)
        self.custom_entry.grid(row=1, column=3, padx=4, pady=5, sticky="w")

        action = ttk.Frame(s)
        action.grid(row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        for i in range(6):
            action.columnconfigure(i, weight=1)
        ttk.Button(action, text="应用设置", command=self.save_profile).grid(row=0, column=0, padx=4, sticky="ew")
        ttk.Button(action, text="恢复默认", command=self.restore_defaults).grid(row=0, column=1, padx=4, sticky="ew")
        ttk.Button(action, text="开始自动备份", command=self.start_auto).grid(row=0, column=2, padx=4, sticky="ew")
        ttk.Button(action, text="停止自动备份", command=self.stop_auto).grid(row=0, column=3, padx=4, sticky="ew")
        ttk.Button(action, text="手动备份", command=self.manual_backup).grid(row=0, column=4, padx=4, sticky="ew")
        ttk.Button(action, text="返回面板", command=self.on_close).grid(row=0, column=5, padx=4, sticky="ew")

        m = self.tab_manage
        m.columnconfigure(0, weight=1)
        m.rowconfigure(0, weight=1)
        self.backup_tree = ttk.Treeview(m, columns=("name", "time", "size"), show="headings")
        self.backup_tree.heading("name", text="备份名称")
        self.backup_tree.heading("time", text="备份时间")
        self.backup_tree.heading("size", text="占用空间")
        self.backup_tree.column("name", width=320)
        self.backup_tree.column("time", width=180)
        self.backup_tree.column("size", width=120)
        self.backup_tree.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        m_btn = ttk.Frame(m)
        m_btn.grid(row=1, column=0, sticky="ew", padx=8, pady=8)
        ttk.Button(m_btn, text="刷新列表", command=self.refresh_backup_list).pack(side="left", padx=4)
        ttk.Button(m_btn, text="删除选中", command=self.delete_selected_backup).pack(side="left", padx=4)
        ttk.Button(m_btn, text="还原选中", command=self.restore_selected_backup).pack(side="left", padx=4)
        ttk.Button(m_btn, text="打开备份位置", command=self.open_backup_root).pack(side="left", padx=4)
        ttk.Button(m_btn, text="重命名备份", command=self.rename_selected_backup).pack(side="left", padx=4)

    def update_custom_state(self):
        self.custom_entry.configure(state="normal" if self.suffix_var.get() == "custom" else "disabled")

    def refresh_from_profile(self):
        self.profile = self.manager.get_profile(self.profile_id)
        if not self.profile:
            return
        self.name_var.set(self.profile.name)
        self.dest_var.set(self.profile.dest_dir)
        self.app_var.set(self.profile.app_path)
        self.interval_var.set(str(self.profile.interval_minutes))
        self.max_var.set(str(self.profile.max_backups))
        self.skip_hidden_var.set(self.profile.skip_hidden)
        self.suffix_var.set(self.profile.suffix_type)
        self.custom_var.set(self.profile.custom_suffix)
        self.update_custom_state()
        self.source_list.delete(0, tk.END)
        for p in self.profile.source_paths:
            self.source_list.insert(tk.END, p)
        self.window.title(f"配置管理 - {self.profile.name}")

    def profile_backup_root(self) -> Path:
        p = self.manager.get_profile(self.profile_id)
        if not p or not p.dest_dir:
            return Path(".")
        return Path(p.dest_dir) / ".labt_profiles" / p.id

    def refresh_backup_list(self):
        for i in self.backup_tree.get_children():
            self.backup_tree.delete(i)
        root = self.profile_backup_root()
        if not root.exists():
            return
        for folder in sorted([x for x in root.iterdir() if x.is_dir()], key=lambda x: x.stat().st_mtime, reverse=True):
            t = datetime.fromtimestamp(folder.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            size = format_size(self.manager.calculate_dir_size(folder))
            self.backup_tree.insert("", tk.END, values=(folder.name, t, size))

    def collect_ui_to_profile(self):
        profile = self.manager.get_profile(self.profile_id)
        if not profile:
            return None
        try:
            interval = int(self.interval_var.get().strip())
            max_backups = int(self.max_var.get().strip())
        except ValueError:
            messagebox.showerror("错误", "间隔和最大备份数必须是整数")
            return None
        if interval <= 0 or max_backups <= 0:
            messagebox.showerror("错误", "间隔和最大备份数必须大于0")
            return None
        profile.name = self.name_var.get().strip() or "未命名配置"
        profile.dest_dir = self.dest_var.get().strip()
        profile.app_path = self.app_var.get().strip()
        profile.interval_minutes = interval
        profile.max_backups = max_backups
        profile.skip_hidden = self.skip_hidden_var.get()
        profile.suffix_type = self.suffix_var.get().strip() or "number"
        profile.custom_suffix = self.custom_var.get().strip()
        profile.source_paths = [self.source_list.get(i) for i in range(self.source_list.size())]
        if profile.suffix_type == "custom" and not profile.custom_suffix:
            messagebox.showerror("错误", "自定义后缀不能为空")
            return None
        return profile

    def save_profile(self):
        profile = self.collect_ui_to_profile()
        if not profile:
            return
        self.manager.save_all()
        self.manager.refresh_dashboard()
        messagebox.showinfo("提示", "设置已应用")

    def restore_defaults(self):
        default = Profile.default()
        p = self.manager.get_profile(self.profile_id)
        if not p:
            return
        p.interval_minutes = default.interval_minutes
        p.max_backups = default.max_backups
        p.skip_hidden = default.skip_hidden
        p.suffix_type = default.suffix_type
        p.custom_suffix = default.custom_suffix
        p.counter = default.counter
        self.refresh_from_profile()
        self.manager.save_all()
        self.manager.refresh_dashboard()

    def start_auto(self):
        self.save_profile()
        self.manager.start_profile(self.profile_id)
        self.manager.refresh_dashboard()

    def stop_auto(self):
        self.manager.stop_profile(self.profile_id)
        self.manager.refresh_dashboard()

    def manual_backup(self):
        self.save_profile()
        p = self.manager.get_profile(self.profile_id)
        if not p:
            return
        def run():
            try:
                self.manager.perform_backup(p)
                self.manager.set_profile_status(p.id, "手动备份完成")
                self.window.after(0, self.refresh_backup_list)
                self.window.after(0, self.manager.refresh_dashboard)
            except Exception as e:
                self.manager.set_profile_status(p.id, f"手动备份失败: {e}")
                self.window.after(0, lambda: messagebox.showerror("错误", f"手动备份失败: {e}"))
        threading.Thread(target=run, daemon=True).start()

    def add_file(self):
        path = filedialog.askopenfilename(title="选择文件")
        if path:
            self.source_list.insert(tk.END, path)

    def add_dir(self):
        path = filedialog.askdirectory(title="选择目录")
        if path:
            self.source_list.insert(tk.END, path)

    def remove_sources(self):
        selected = list(self.source_list.curselection())
        for idx in reversed(selected):
            self.source_list.delete(idx)

    def open_selected_source(self):
        selected = self.source_list.curselection()
        if not selected:
            return
        p = Path(self.source_list.get(selected[0]))
        open_path(p if p.is_dir() else p.parent)

    def pick_dest(self):
        path = filedialog.askdirectory(title="选择备份目录")
        if path:
            self.dest_var.set(path)

    def open_dest(self):
        dest = self.dest_var.get().strip()
        if dest:
            open_path(Path(dest))

    def pick_app(self):
        path = filedialog.askopenfilename(
            title="选择关联应用程序",
            filetypes=[("Executable", "*.exe"), ("All Files", "*.*")],
        )
        if path:
            self.app_var.set(path)

    def launch_app(self):
        p = self.app_var.get().strip()
        if not p:
            return
        self.manager.launch_profile_app(self.profile_id)

    def open_backup_root(self):
        root = self.profile_backup_root()
        root.mkdir(parents=True, exist_ok=True)
        open_path(root)

    def selected_backup_path(self):
        sel = self.backup_tree.selection()
        if not sel:
            return None
        values = self.backup_tree.item(sel[0], "values")
        if not values:
            return None
        return self.profile_backup_root() / values[0]

    def delete_selected_backup(self):
        path = self.selected_backup_path()
        if not path or not path.exists():
            return
        if not messagebox.askyesno("确认", f"确认删除备份：{path.name} ?"):
            return
        remove_path(path)
        self.refresh_backup_list()
        self.manager.refresh_dashboard()

    def restore_selected_backup(self):
        path = self.selected_backup_path()
        if not path or not path.exists():
            return
        p = self.manager.get_profile(self.profile_id)
        if not p:
            return
        try:
            self.manager.restore_backup(p, path)
            self.manager.set_profile_status(p.id, f"已还原: {path.name}")
            messagebox.showinfo("提示", "还原完成")
        except Exception as e:
            messagebox.showerror("错误", f"还原失败: {e}")

    def rename_selected_backup(self):
        path = self.selected_backup_path()
        if not path or not path.exists():
            return
        win = tk.Toplevel(self.window)
        win.title("重命名备份")
        win.geometry("380x120")
        win.transient(self.window)
        win.grab_set()
        var = tk.StringVar(value=path.name)
        ttk.Label(win, text="新备份名称").pack(padx=10, pady=8, anchor="w")
        ttk.Entry(win, textvariable=var).pack(fill="x", padx=10)

        def do_rename():
            new_name = safe_name(var.get().strip())
            if not new_name:
                return
            new_path = path.parent / new_name
            if new_path.exists():
                messagebox.showerror("错误", "目标名称已存在")
                return
            path.rename(new_path)
            win.destroy()
            self.refresh_backup_list()
            self.manager.refresh_dashboard()

        ttk.Button(win, text="确定", command=do_rename).pack(pady=10)

    def apply_name(self):
        p = self.manager.get_profile(self.profile_id)
        if not p:
            return
        p.name = self.name_var.get().strip() or "未命名配置"
        self.manager.save_all()
        self.manager.refresh_dashboard()
        self.refresh_from_profile()


class GlobalSettingsWindow:
    def __init__(self, manager):
        self.manager = manager
        self.win = tk.Toplevel(manager.root)
        self.win.title("全局设置")
        self.win.geometry("560x330")
        self.create_widgets()

    def create_widgets(self):
        g = self.manager.data["global"]
        frame = ttk.Frame(self.win, padding=14)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="备份快捷键").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.backup_hotkey = tk.StringVar(value=g.get("backup_hotkey", "ctrl+f1"))
        ttk.Entry(frame, textvariable=self.backup_hotkey).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        self.backup_enabled = tk.BooleanVar(value=g.get("backup_hotkey_enabled", False))
        ttk.Checkbutton(frame, text="启用", variable=self.backup_enabled).grid(row=0, column=2, sticky="w")

        ttk.Label(frame, text="还原快捷键").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.restore_hotkey = tk.StringVar(value=g.get("restore_hotkey", "ctrl+f2"))
        ttk.Entry(frame, textvariable=self.restore_hotkey).grid(row=1, column=1, sticky="ew", padx=6, pady=6)
        self.restore_enabled = tk.BooleanVar(value=g.get("restore_hotkey_enabled", False))
        ttk.Checkbutton(frame, text="启用", variable=self.restore_enabled).grid(row=1, column=2, sticky="w")

        ttk.Label(frame, text="关闭按钮行为").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        self.close_behavior = tk.StringVar(value=g.get("close_behavior", "minimize_to_tray"))
        ttk.Combobox(
            frame,
            textvariable=self.close_behavior,
            state="readonly",
            values=["minimize_to_tray", "exit"],
        ).grid(row=2, column=1, sticky="w", padx=6, pady=6)

        action = ttk.Frame(frame)
        action.grid(row=3, column=0, columnspan=3, sticky="ew", pady=14)
        action.columnconfigure(0, weight=1)
        action.columnconfigure(1, weight=1)
        action.columnconfigure(2, weight=1)
        action.columnconfigure(3, weight=1)
        ttk.Button(action, text="导出设置", command=self.export_settings).grid(row=0, column=0, padx=4, sticky="ew")
        ttk.Button(action, text="导入设置", command=self.import_settings).grid(row=0, column=1, padx=4, sticky="ew")
        ttk.Button(action, text="应用", command=self.apply).grid(row=0, column=2, padx=4, sticky="ew")
        ttk.Button(action, text="关闭", command=self.win.destroy).grid(row=0, column=3, padx=4, sticky="ew")

    def apply(self):
        g = self.manager.data["global"]
        g["backup_hotkey"] = self.backup_hotkey.get().strip() or "ctrl+f1"
        g["backup_hotkey_enabled"] = bool(self.backup_enabled.get())
        g["restore_hotkey"] = self.restore_hotkey.get().strip() or "ctrl+f2"
        g["restore_hotkey_enabled"] = bool(self.restore_enabled.get())
        g["close_behavior"] = self.close_behavior.get().strip() or "minimize_to_tray"
        self.manager.save_all()
        self.manager.apply_hotkeys()
        messagebox.showinfo("提示", "全局设置已应用")

    def export_settings(self):
        path = filedialog.asksaveasfilename(
            title="导出设置",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All Files", "*.*")],
            initialfile="profiles_config_export.json",
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.manager.data, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("提示", f"已导出设置: {path}")

    def import_settings(self):
        path = filedialog.askopenfilename(
            title="导入设置",
            filetypes=[("JSON", "*.json"), ("All Files", "*.*")],
        )
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "profiles" not in data or "global" not in data:
            messagebox.showerror("错误", "配置格式无效")
            return
        self.manager.data = data
        self.manager.profiles = [Profile(**p) for p in data["profiles"]]
        self.manager.save_all()
        self.manager.refresh_dashboard()
        self.manager.apply_hotkeys()
        messagebox.showinfo("提示", "导入成功")


class ModernBackupManagerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x760")
        self.root.minsize(980, 620)
        self.store = ConfigStore(APP_ROOT / DATA_FILE)
        self.data = self.store.load()
        self.profiles = [Profile(**p) for p in self.data["profiles"]]
        self.runners = {}
        self.profile_windows = {}
        self.status = {}
        self.selected_profile_id = self.profiles[0].id if self.profiles else None
        self._hotkeys = []
        self._card_images = {}
        self.apply_style()
        self.create_widgets()
        self.refresh_dashboard()
        self.apply_hotkeys()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        for p in self.profiles:
            if p.enabled:
                self.start_profile(p.id, save=False)

    def apply_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei", 16, "bold"))
        style.configure("Sub.TLabel", font=("Microsoft YaHei", 10))
        style.configure("CardName.TLabel", font=("Microsoft YaHei", 12, "bold"))
        style.configure("CardMeta.TLabel", font=("Microsoft YaHei", 10))

    def create_widgets(self):
        self.create_menu()
        container = ttk.Frame(self.root, padding=10)
        container.pack(fill="both", expand=True)
        header = ttk.Frame(container)
        header.pack(fill="x")
        ttk.Label(header, text="集成式本地自动备份管理面板", style="Title.TLabel").pack(side="left")
        ttk.Button(header, text="新建配置", command=self.add_profile).pack(side="right", padx=4)
        ttk.Button(header, text="全局设置", command=self.open_global_settings).pack(side="right", padx=4)

        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(container, textvariable=self.status_var, style="Sub.TLabel").pack(fill="x", pady=(4, 8))

        self.canvas = tk.Canvas(container, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        scroll.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.cards_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.cards_frame.bind("<Configure>", self.on_cards_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

    def create_menu(self):
        menu = tk.Menu(self.root)
        m_settings = tk.Menu(menu, tearoff=0)
        m_settings.add_command(label="全局设置", command=self.open_global_settings)
        m_settings.add_separator()
        m_settings.add_command(label="导出设置", command=self.export_settings)
        m_settings.add_command(label="导入设置", command=self.import_settings)
        m_settings.add_separator()
        m_settings.add_command(label="退出", command=self.on_close)
        menu.add_cascade(label="设置", menu=m_settings)
        self.root.config(menu=menu)

    def on_cards_configure(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def get_profile(self, profile_id: str):
        for p in self.profiles:
            if p.id == profile_id:
                return p
        return None

    def save_all(self):
        self.data["profiles"] = [asdict(p) for p in self.profiles]
        self.store.save(self.data)

    def export_settings(self):
        path = filedialog.asksaveasfilename(
            title="导出设置",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All Files", "*.*")],
            initialfile="profiles_config_export.json",
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("提示", f"已导出设置: {path}")

    def import_settings(self):
        path = filedialog.askopenfilename(
            title="导入设置",
            filetypes=[("JSON", "*.json"), ("All Files", "*.*")],
        )
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "profiles" not in data or "global" not in data:
            messagebox.showerror("错误", "配置格式无效")
            return
        for pid in list(self.runners.keys()):
            self.stop_profile(pid, save=False)
        self.data = data
        self.profiles = [Profile(**p) for p in data["profiles"]]
        self.selected_profile_id = self.profiles[0].id if self.profiles else None
        self.save_all()
        self.refresh_dashboard()
        self.apply_hotkeys()
        messagebox.showinfo("提示", "导入成功")

    def set_profile_status(self, profile_id: str, message: str):
        self.status[profile_id] = message
        selected = self.get_profile(self.selected_profile_id) if self.selected_profile_id else None
        if selected and selected.id == profile_id:
            self.root.after(0, lambda: self.status_var.set(f"[{selected.name}] {message}"))
        self.root.after(0, self.refresh_dashboard)
        win = self.profile_windows.get(profile_id)
        if win:
            self.root.after(0, win.refresh_backup_list)

    def profile_backup_root(self, profile: Profile) -> Path:
        return Path(profile.dest_dir) / ".labt_profiles" / profile.id

    def calculate_dir_size(self, path: Path) -> int:
        total = 0
        if not path.exists():
            return 0
        for p in path.rglob("*"):
            if p.is_file():
                try:
                    total += p.stat().st_size
                except Exception:
                    pass
        return total

    def profile_used_size(self, profile: Profile) -> int:
        if not profile.dest_dir:
            return 0
        root = self.profile_backup_root(profile)
        return self.calculate_dir_size(root)

    def backup_suffix(self, profile: Profile) -> str:
        if profile.suffix_type == "timestamp":
            return "_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        if profile.suffix_type == "custom":
            return "_" + safe_name(profile.custom_suffix or "custom")
        if profile.number_mode == "manual":
            s = "_" + str(profile.counter)
            profile.counter += 1
            return s
        root = self.profile_backup_root(profile)
        if not root.exists():
            return "_1"
        max_num = 0
        for p in root.iterdir():
            if not p.is_dir():
                continue
            parts = p.name.rsplit("_", 1)
            if len(parts) == 2 and parts[1].isdigit():
                max_num = max(max_num, int(parts[1]))
        return "_" + str(max_num + 1)

    def cleanup_backups(self, profile: Profile):
        root = self.profile_backup_root(profile)
        if not root.exists():
            return
        max_backups = max(1, int(profile.max_backups or 1))
        folders = sorted([x for x in root.iterdir() if x.is_dir()], key=lambda p: p.stat().st_mtime, reverse=True)
        for old in folders[max_backups:]:
            remove_path(old)

    def perform_backup(self, profile: Profile):
        if not profile.source_paths:
            raise RuntimeError("请先添加至少一个源路径")
        if not profile.dest_dir:
            raise RuntimeError("请先设置备份目录")
        backup_root = self.profile_backup_root(profile)
        backup_root.mkdir(parents=True, exist_ok=True)
        suffix = self.backup_suffix(profile)
        snap_name = safe_name(f"snapshot{suffix}")
        snap_dir = backup_root / snap_name
        if snap_dir.exists():
            if profile.duplicate_handling == "skip":
                return
            if profile.duplicate_handling == "overwrite":
                remove_path(snap_dir)
            else:
                idx = 1
                while (backup_root / f"{snap_name}_{idx}").exists():
                    idx += 1
                snap_dir = backup_root / f"{snap_name}_{idx}"
        items_dir = snap_dir / "items"
        items_dir.mkdir(parents=True, exist_ok=True)

        manifest = {"created_at": datetime.now().isoformat(), "items": []}
        for i, src_str in enumerate(profile.source_paths):
            src = Path(src_str)
            if not src.exists():
                continue
            if profile.skip_hidden and is_hidden(src):
                continue
            item_name = f"{i}_{safe_name(src.name)}"
            item_path = items_dir / item_name
            if src.is_file():
                copy_file(src, item_path)
                typ = "file"
            else:
                copy_dir(src, item_path, profile.skip_hidden)
                typ = "dir"
            manifest["items"].append({"source": str(src), "item": f"items/{item_name}", "type": typ})

        if not manifest["items"]:
            remove_path(snap_dir)
            raise RuntimeError("没有可备份内容")

        with open(snap_dir / "manifest.json", "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        self.cleanup_backups(profile)
        self.save_all()

    def restore_backup(self, profile: Profile, backup_dir: Path):
        manifest_path = backup_dir / "manifest.json"
        if not manifest_path.exists():
            raise RuntimeError("备份清单不存在")
        with open(manifest_path, "r", encoding="utf-8") as f:
            manifest = json.load(f)
        for item in manifest.get("items", []):
            src = Path(item["source"])
            stored = backup_dir / item["item"]
            typ = item.get("type", "file")
            if not stored.exists():
                continue
            if typ == "file":
                src.parent.mkdir(parents=True, exist_ok=True)
                if src.exists():
                    remove_path(src)
                copy_file(stored, src)
            else:
                src.parent.mkdir(parents=True, exist_ok=True)
                if src.exists():
                    remove_path(src)
                copy_dir(stored, src, skip_hidden=False)

    def restore_latest(self, profile: Profile):
        root = self.profile_backup_root(profile)
        if not root.exists():
            raise RuntimeError("无可用备份")
        folders = sorted([x for x in root.iterdir() if x.is_dir()], key=lambda p: p.stat().st_mtime, reverse=True)
        if not folders:
            raise RuntimeError("无可用备份")
        self.restore_backup(profile, folders[0])

    def start_profile(self, profile_id: str, save: bool = True):
        p = self.get_profile(profile_id)
        if not p:
            return
        if profile_id in self.runners:
            return
        p.enabled = True
        runner = ProfileRunner(self, profile_id)
        self.runners[profile_id] = runner
        runner.start()
        self.set_profile_status(profile_id, "自动备份已启动")
        if save:
            self.save_all()

    def stop_profile(self, profile_id: str, save: bool = True):
        p = self.get_profile(profile_id)
        if not p:
            return
        p.enabled = False
        runner = self.runners.pop(profile_id, None)
        if runner:
            runner.stop()
        self.set_profile_status(profile_id, "自动备份已停止")
        if save:
            self.save_all()

    def add_profile(self):
        p = Profile.default()
        p.name = f"配置 {len(self.profiles) + 1}"
        self.profiles.append(p)
        self.selected_profile_id = p.id
        self.save_all()
        self.refresh_dashboard()
        self.open_profile(p.id)

    def remove_profile(self, profile_id: str):
        p = self.get_profile(profile_id)
        if not p:
            return
        if not messagebox.askyesno("确认", f"确定删除配置 \"{p.name}\" ?"):
            return
        self.stop_profile(profile_id, save=False)
        self.profiles = [x for x in self.profiles if x.id != profile_id]
        self.profile_windows.pop(profile_id, None)
        if self.selected_profile_id == profile_id:
            self.selected_profile_id = self.profiles[0].id if self.profiles else None
        if not self.profiles:
            self.profiles = [Profile.default()]
            self.selected_profile_id = self.profiles[0].id
        self.save_all()
        self.refresh_dashboard()

    def open_profile(self, profile_id: str):
        self.selected_profile_id = profile_id
        win = self.profile_windows.get(profile_id)
        if win and win.window.winfo_exists():
            win.window.lift()
            return
        self.profile_windows[profile_id] = ProfileWindow(self, profile_id)
        self.refresh_dashboard()

    def launch_profile_app(self, profile_id: str):
        p = self.get_profile(profile_id)
        if not p or not p.app_path:
            messagebox.showinfo("提示", "尚未设置关联应用程序")
            return
        app = Path(p.app_path)
        if not app.exists():
            messagebox.showerror("错误", "关联应用路径无效")
            return
        try:
            if os.name == "nt":
                os.startfile(str(app))
            else:
                subprocess.Popen([str(app)])
        except Exception as e:
            messagebox.showerror("错误", f"启动失败: {e}")

    def load_card_icon(self, profile: Profile):
        icon_path = Path(profile.icon_path) if profile.icon_path else DEFAULT_ICON_PATH
        if profile.app_path and Path(profile.app_path).exists():
            icon_path = Path(profile.app_path)
        if not Image or not ImageTk:
            return None
        if not icon_path.exists():
            return None
        key = f"{profile.id}:{icon_path}"
        if key in self._card_images:
            return self._card_images[key]
        try:
            image = Image.open(str(icon_path)).resize((52, 52))
            tk_image = ImageTk.PhotoImage(image)
            self._card_images[key] = tk_image
            return tk_image
        except Exception:
            return None

    def refresh_dashboard(self):
        for c in self.cards_frame.winfo_children():
            c.destroy()
        if not self.profiles:
            ttk.Label(self.cards_frame, text="暂无配置，请点击“新建配置”").pack(pady=30)
            return
        card_width = 360
        available_width = max(720, self.canvas.winfo_width() or self.root.winfo_width())
        columns = max(1, min(4, available_width // card_width))
        for idx, profile in enumerate(self.profiles):
            row = idx // columns
            col = idx % columns
            card = tk.Frame(self.cards_frame, bd=1, relief="solid", bg="#f5f7fb")
            card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
            card.config(width=340, height=220)
            card.grid_propagate(False)
            self.cards_frame.columnconfigure(col, weight=1)
            selected = profile.id == self.selected_profile_id
            if selected:
                card.config(bg="#e9f2ff", highlightbackground="#4f8cff", highlightthickness=2)

            top = tk.Frame(card, bg=card.cget("bg"))
            top.pack(fill="x", padx=8, pady=6)
            enabled_var = tk.BooleanVar(value=profile.id in self.runners)

            def on_toggle(pid=profile.id, var=enabled_var):
                self.selected_profile_id = pid
                if var.get():
                    self.start_profile(pid)
                else:
                    self.stop_profile(pid)
                self.refresh_dashboard()

            ttk.Checkbutton(top, text="自动备份", variable=enabled_var, command=on_toggle).pack(side="left")
            ttk.Button(top, text="启动应用", command=lambda pid=profile.id: self.launch_profile_app(pid)).pack(side="right", padx=2)

            icon = self.load_card_icon(profile)
            icon_container = tk.Frame(card, bg=card.cget("bg"))
            icon_container.pack(pady=(2, 6))
            if icon:
                tk.Label(icon_container, image=icon, bg=card.cget("bg")).pack()
            else:
                tk.Label(icon_container, text="🗂", font=("Segoe UI Emoji", 30), bg=card.cget("bg")).pack()

            name_label = ttk.Label(card, text=profile.name, style="CardName.TLabel")
            name_label.pack()
            size_text = format_size(self.profile_used_size(profile))
            ttk.Label(card, text=f"备份占用: {size_text}", style="CardMeta.TLabel").pack(pady=(4, 0))
            status_text = self.status.get(profile.id, "就绪")
            ttk.Label(card, text=status_text, style="CardMeta.TLabel").pack(pady=(2, 0))

            btns = ttk.Frame(card)
            btns.pack(fill="x", padx=8, pady=8)
            ttk.Button(btns, text="打开配置", command=lambda pid=profile.id: self.open_profile(pid)).pack(side="left", fill="x", expand=True, padx=2)
            ttk.Button(btns, text="手动备份", command=lambda pid=profile.id: self.manual_backup_from_dashboard(pid)).pack(side="left", fill="x", expand=True, padx=2)
            ttk.Button(btns, text="删除", command=lambda pid=profile.id: self.remove_profile(pid)).pack(side="left", padx=2)

            for w in (card, top, icon_container, name_label):
                w.bind("<Button-1>", lambda _e, pid=profile.id: self.select_card(pid))

    def select_card(self, profile_id: str):
        self.selected_profile_id = profile_id
        profile = self.get_profile(profile_id)
        if profile:
            self.status_var.set(f"当前配置: {profile.name}")
        self.refresh_dashboard()

    def manual_backup_from_dashboard(self, profile_id: str):
        p = self.get_profile(profile_id)
        if not p:
            return
        self.selected_profile_id = profile_id
        def run():
            try:
                self.perform_backup(p)
                self.set_profile_status(profile_id, "手动备份完成")
            except Exception as e:
                self.set_profile_status(profile_id, f"手动备份失败: {e}")
        threading.Thread(target=run, daemon=True).start()
        self.refresh_dashboard()

    def open_global_settings(self):
        GlobalSettingsWindow(self)

    def apply_hotkeys(self):
        if keyboard is None:
            return
        for hk in self._hotkeys:
            try:
                keyboard.remove_hotkey(hk)
            except Exception:
                pass
        self._hotkeys = []
        g = self.data.get("global", {})
        backup_enabled = bool(g.get("backup_hotkey_enabled"))
        restore_enabled = bool(g.get("restore_hotkey_enabled"))
        backup_hotkey = g.get("backup_hotkey", "ctrl+f1")
        restore_hotkey = g.get("restore_hotkey", "ctrl+f2")

        if backup_enabled:
            try:
                hk = keyboard.add_hotkey(backup_hotkey, self.hotkey_backup_action)
                self._hotkeys.append(hk)
            except Exception as e:
                messagebox.showwarning("快捷键", f"备份快捷键注册失败: {e}")
        if restore_enabled:
            try:
                hk = keyboard.add_hotkey(restore_hotkey, self.hotkey_restore_action)
                self._hotkeys.append(hk)
            except Exception as e:
                messagebox.showwarning("快捷键", f"还原快捷键注册失败: {e}")

    def hotkey_backup_action(self):
        pid = self.selected_profile_id
        if not pid:
            self.root.after(0, lambda: self.status_var.set("未选中配置，无法执行快捷键备份"))
            return
        self.root.after(0, lambda: self.manual_backup_from_dashboard(pid))

    def hotkey_restore_action(self):
        pid = self.selected_profile_id
        if not pid:
            self.root.after(0, lambda: self.status_var.set("未选中配置，无法执行快捷键还原"))
            return
        p = self.get_profile(pid)
        if not p:
            return
        def run():
            try:
                self.restore_latest(p)
                self.set_profile_status(pid, "快捷键还原完成")
            except Exception as e:
                self.set_profile_status(pid, f"快捷键还原失败: {e}")
        threading.Thread(target=run, daemon=True).start()

    def on_close(self):
        behavior = self.data.get("global", {}).get("close_behavior", "minimize_to_tray")
        if behavior == "exit":
            self.shutdown()
            self.root.destroy()
        else:
            self.root.iconify()
            self.status_var.set("已最小化，可在任务栏恢复")

    def shutdown(self):
        for pid in list(self.runners.keys()):
            self.stop_profile(pid, save=False)
        self.save_all()
        if keyboard is not None:
            for hk in self._hotkeys:
                try:
                    keyboard.remove_hotkey(hk)
                except Exception:
                    pass


def run():
    root = tk.Tk()
    if WINDOW_ICON_PATH.exists():
        try:
            root.iconbitmap(str(WINDOW_ICON_PATH))
        except Exception:
            pass
    ModernBackupManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    run()
