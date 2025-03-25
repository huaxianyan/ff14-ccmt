import os
import ctypes
import sys
from tkinter import Text, messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import threading
import json
from typing import Literal

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def set_dpi_awareness():
    """设置 DPI 感知"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except AttributeError:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except AttributeError:
            pass

class ConfigManagerWindow:
    def __init__(self, parent, international_path, china_path, backup_path):
        # 创建新窗口
        self.window = ttk.Toplevel(parent)
        self.window.title("角色配置管理")
        
        # 保存参数
        self.parent = parent
        self.international_path = international_path
        self.china_path = china_path
        self.backup_path = backup_path
        
        # 初始化选择的文件夹
        self.selected_folder = None
        
        # 设置口大小
        window_width = 600
        window_height = 400
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置窗口最小大小
        self.window.minsize(500, 300)
        
        # 设置样式
        style = ttk.Style()
        style.configure("Config.TRadiobutton", background="#f0f0f0")
        style.configure("Dialog.TFrame", background="#f0f0f0")
        style.configure("Dialog.TLabel", background="#f0f0f0")
        
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # 创建服务器选择区域
        server_frame = ttk.Frame(main_frame)
        server_frame.pack(fill="x", pady=(0, 10))
        
        # 创建单选按钮变量
        self.server_var = ttk.StringVar(value="international")
        self.server_var.trace_add("write", self.on_server_change)
        
        # 创建单选按钮
        ttk.Radiobutton(
            server_frame,
            text="国际服",
            value="international",
            variable=self.server_var,
            style="Config.TRadiobutton"
        ).pack(side="left", padx=5)
        
        ttk.Radiobutton(
            server_frame,
            text="国服",
            value="china",
            variable=self.server_var,
            style="Config.TRadiobutton"
        ).pack(side="left", padx=5)
        
        # 创建列表框
        list_frame = ttk.LabelFrame(main_frame, text="角色列表", padding=5)
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        self.listbox = ttk.Treeview(list_frame, show="tree", selectmode="browse")
        self.listbox.pack(fill="both", expand=True)
        
        # 创建右侧操作区域
        operation_frame = ttk.LabelFrame(main_frame, text="操作", padding=5)
        operation_frame.pack(side="left", fill="both", padx=(5, 0))
        
        # 添加标记按钮
        self.mark_button = ttk.Button(
            operation_frame,
            text="标记",
            command=self.mark_folder,
            style="primary.TButton",
            width=15
        )
        self.mark_button.pack(pady=5)
        
        # 加载配置数据
        self.load_configs()
        
        # 并显示文件夹
        self.scan_folders()

    def load_configs(self):
        """加载所有配置数据"""
        # 确保 data 文件夹存在
        data_dir = "data"
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        
        # 加载国际服和国服的标记数据
        self.international_marks = self.load_marks("international")
        self.china_marks = self.load_marks("china")

    def load_marks(self, server_type):
        """加载指定服务器的标记数据"""
        config_file = os.path.normpath(os.path.join("data", f"{server_type}_marks.json"))
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def save_marks(self, server_type, marks):
        """保存指定服务器的标记数据"""
        config_file = os.path.join("data", f"{server_type}_marks.json")
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(marks, f, ensure_ascii=False, indent=2)

    def on_server_change(self, *args):
        """服务器选择改变时的处理"""
        self.scan_folders()

    def scan_folders(self):
        """扫描并显示文件夹"""
        # 保存当前选择
        current_selection = None
        selected = self.listbox.selection()
        if selected:
            current_selection = selected[0]
            self.selected_folder = current_selection  # 更新 selected_folder
        
        # 清空列表
        for item in self.listbox.get_children():
            self.listbox.delete(item)
        
        # 获取当前选择的服务器和对应的路径
        server_type = self.server_var.get()
        path_var = self.international_path if server_type == "international" else self.china_path
        marks = self.international_marks if server_type == "international" else self.china_marks
        
        # 获取路径
        base_path = path_var.get()
        
        if not base_path:
            messagebox.showwarning("警告", "请先在路径设置中设置对应的游戏路径！", parent=self.window)
            return
        
        # 扫描文件夹
        try:
            first_item = None
            for item in os.listdir(base_path):
                if "FFXIV_" in item and os.path.isdir(os.path.join(base_path, item)):
                    # 检查是否有标记，如果有标记则显示"标记名 (文件夹名)"，否则直接显示文件夹名
                    display_name = f"{marks.get(item)} ({item})" if item in marks else item
                    
                    self.listbox.insert("", "end", item, text=display_name)
                    if first_item is None:
                        first_item = item
            
            # 优先使用保存的选择
            if self.selected_folder and self.selected_folder in os.listdir(base_path):
                self.listbox.selection_set(self.selected_folder)
            # 其次使用当前选择
            elif current_selection and current_selection in os.listdir(base_path):
                self.listbox.selection_set(current_selection)
            # 最后才使用第一项
            elif first_item:
                self.listbox.selection_set(first_item)
                
        except Exception as e:
            messagebox.showerror("错误", f"扫描文件夹时出错：{str(e)}", parent=self.window)

    def mark_folder(self):
        """标记选中的文件夹"""
        # 获取选中的项
        selected = self.listbox.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个文件夹！")
            return
        
        folder_id = selected[0]
        current_name = self.listbox.item(folder_id, "text")
        # 如果当前名包含括号，提取括号前的部分作为当标
        if "(" in current_name:
            current_name = current_name.split(" (")[0]
        
        # 弹出输入对话框
        dialog = ttk.Toplevel(self.window)
        dialog.title("标记文件夹")
        dialog.transient(self.window)
        
        # 对话框居中
        dialog_width = 300
        dialog_height = 150  # 增加高度
        dialog_x = self.window.winfo_x() + (self.window.winfo_width() - dialog_width) // 2
        dialog_y = self.window.winfo_y() + (self.window.winfo_height() - dialog_height) // 2
        dialog.geometry(f"{dialog_width}x{dialog_height}+{dialog_x}+{dialog_y}")
        
        # 创建输入框
        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill="both", expand=True)
        
        # 设置框和背
        dialog.configure(background="#f0f0f0")
        frame.configure(style="Dialog.TFrame")
        
        # 创建标签
        ttk.Label(
            frame, 
            text="请输入标记名称：",
            style="Dialog.TLabel"
        ).pack(pady=(0, 5))
        
        # 创建输入框
        name_var = ttk.StringVar(value=current_name)
        entry = ttk.Entry(
            frame, 
            textvariable=name_var
        )
        entry.pack(fill="x", pady=(0, 15))
        
        # 定义确认函数（修复未定义错误）
        def confirm():
            new_name = name_var.get().strip()
            if new_name:
                server_type = self.server_var.get()
                marks = self.international_marks if server_type == "international" else self.china_marks
                marks[folder_id] = new_name
                self.save_marks(server_type, marks)
                display_name = f"{new_name} ({folder_id})"
                self.listbox.item(folder_id, text=display_name)
                dialog.destroy()
        
        # 确认按钮
        ttk.Button(
            frame, 
            text="确定", 
            command=confirm,
            style="primary.TButton"
        ).pack()
        
        # 设置焦点并绑定回车键
        entry.focus()
        entry.bind("<Return>", lambda e: confirm())
        
        # 设置对话框为模态
        dialog.grab_set()
        dialog.wait_window()

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("FF14角色配置管理工具")
        
        # 获取 DPI 缩放因子
        try:
            self.scaling_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        except:
            self.scaling_factor = 1
        
        # 初始化最小
        self.min_width = int(500 * self.scaling_factor)
        self.min_height = int(600 * self.scaling_factor)
        self.root.minsize(self.min_width, self.min_height)
        
        # 设置窗口背景色
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabelframe", background="#f0f0f0")
        style.configure("TLabelframe.Label", background="#f0f0f0")
        style.configure("TButton", padding=5)
        style.configure("PathLabel.TLabel", background="#f0f0f0")
        style.configure("PathEntry.TEntry", fieldbackground="#f0f0f0")
        self.root.configure(background="#f0f0f0")

        # 创建主Frame
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        # 确保 data 目录存在
        self.data_dir = "data"
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        
        # 加载配置
        self.load_config()
        
        # 创建各个区域
        self.create_character_config_section()
        self.create_migration_section()
        self.create_backup_section()
        self.create_path_section()
        
        # 添加配置管理器实例变量
        self.config_manager = None

        # 设置窗口图标
        self.icon_path = get_resource_path("3.ico")  # 修改为 3.ico
        if os.path.exists(self.icon_path):
            self.root.iconbitmap(self.icon_path)
            self.root.tk.call('wm', 'iconbitmap', self.root._w, self.icon_path)

    def load_config(self):
        """加载配置"""
        self.config_file = os.path.join(self.data_dir, "config.json")
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            self.config = {
                "international_path": "",
                "china_path": "",
                "backup_path": ""
            }

    def save_config(self):
        """保存配置"""
        self.config = {
            "international_path": self.international_path.get(),
            "china_path": self.china_path.get(),
            "backup_path": self.backup_path.get()
        }
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def create_character_config_section(self):
        """创建角色配置管理区域"""
        frame = ttk.LabelFrame(self.main_frame, text="角色配置管理", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        
        # 按钮容器
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x")
        
        # 角色配置管理按钮
        ttk.Button(
            btn_frame,
            text="角色配置管理",
            style="primary.TButton",
            width=15,
            command=self.open_config_manager
        ).pack(side="left", padx=5)

    def open_config_manager(self):
        """打开配置管理器窗口"""
        if self.config_manager is None or not self.config_manager.window.winfo_exists():
            self.config_manager = ConfigManagerWindow(self.root, self.international_path, self.china_path, self.backup_path)
        else:
            self.config_manager.window.lift()

    def create_migration_section(self):
        """创建配置迁移区域"""
        frame = ttk.LabelFrame(self.main_frame, text="配置迁移", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        
        # 按钮容器
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x")
        
        # 迁移按钮
        ttk.Button(
            btn_frame,
            text="迁移",
            style="success.TButton",
            width=15,
            command=self.open_migration_window
        ).pack(side="left", padx=5)

    def create_backup_section(self):
        """创建备份与恢复区域"""
        frame = ttk.LabelFrame(self.main_frame, text="备份与恢复", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        
        # 按钮容器
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x")
        
        # 角色配置按钮
        ttk.Button(
            btn_frame,
            text="角色配置",
            style="primary.TButton",
            width=15,
            command=self.open_character_backup_window
        ).pack(side="left", padx=5)

    def create_path_section(self):
        """创建路径设置区域"""
        frame = ttk.LabelFrame(self.main_frame, text="路径设置", padding=10)
        frame.pack(fill="x")
        
        # 设置标签和输入框的样式
        style = ttk.Style()
        style.configure("PathLabel.TLabel", background="#f0f0f0")
        
        # 国际服路径
        self.create_path_row(frame, "国际服路径：", "international_path")
        
        # 国服路径
        self.create_path_row(frame, "国服路径：", "china_path")
        
        # 备份路径
        self.create_path_row(frame, "备份路径：", "backup_path")

    def create_path_row(self, parent, label_text, path_var_name):
        """创建路径设置行"""
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=2)
        
        # 标签
        ttk.Label(
            frame, 
            text=label_text, 
            style="PathLabel.TLabel"
        ).pack(side="left")
        
        # 路径输入框
        path_var = ttk.StringVar(value=self.config.get(path_var_name, ""))  # 从配置中加载初始值
        path_var.trace_add("write", lambda *args: self.save_config())  # 添加变更监听
        setattr(self, path_var_name, path_var)
        entry = ttk.Entry(
            frame, 
            textvariable=path_var, 
            state="readonly",
            style="PathEntry.TEntry"
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # 浏览按钮
        ttk.Button(
            frame,
            text="浏览",
            style="secondary.TButton",
            command=lambda: self.browse_folder(path_var)
        ).pack(side="right")

    def browse_folder(self, path_var):
        """浏览文件夹"""
        folder = filedialog.askdirectory()
        if folder:
            path_var.set(folder)

    def open_migration_window(self):
        """打开迁移口"""
        MigrationWindow(self.root, self.international_path, self.china_path)

    def show_custom_messagebox(self, type_, title, message, **kwargs):
        """显示自定义消息框"""
        # 播放提示音
        self.root.bell()
        
        dialog = ttk.Toplevel(self.root)
        dialog.withdraw()  # 隐对话框
        dialog.transient(self.root)  # 设置为主窗口的临时窗口
        
        if hasattr(self, 'icon_path') and self.icon_path:
            dialog.iconbitmap(self.icon_path)
        
        if type_ == "showinfo":
            result = messagebox.showinfo(title, message, parent=dialog, **kwargs)
        elif type_ == "showwarning":
            result = messagebox.showwarning(title, message, parent=dialog, **kwargs)
        elif type_ == "showerror":
            result = messagebox.showerror(title, message, parent=dialog, **kwargs)
        elif type_ == "askyesno":
            result = messagebox.askyesno(title, message, parent=dialog, **kwargs)
        
        dialog.destroy()
        return result

    def open_character_backup_window(self):
        """打开角色配置备份窗口"""
        CharacterBackupWindow(self.root, self.international_path, self.china_path, self.backup_path)

    def open_software_backup_window(self):
        """打开软件配置备份窗口"""
        SoftwareBackupWindow(self.root, self.international_path, self.china_path, self.backup_path)

    def format_path(self, path):
        """格式化路径用于显示"""
        return path.replace(os.sep, '/')

class MigrationWindow:
    def __init__(self, parent, international_path, china_path):
        # 创建新窗口
        self.window = ttk.Toplevel(parent)
        self.window.title("配置迁移")
        
        # 保存参数
        self.parent = parent
        self.international_path = international_path
        self.china_path = china_path
        
        # 设置窗口大小（增加宽度）
        window_width = 1000  # 从 800 改为 1000
        window_height = 500
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置窗口最小大小（增加最小宽度）
        self.window.minsize(900, 400)  # 从 700 改为 900
        
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # 设置样式
        style = ttk.Style()
        style.configure("Migration.TRadiobutton", background="#f0f0f0")
        style.configure("Migration.TCheckbutton", background="#f0f0f0")  # 添加复选框样式
        
        # 初始化变量（修改目标服务器的默认值）
        self.source_var = ttk.StringVar(value="international")
        self.target_var = ttk.StringVar(value="international")  # 改为 international
        self.source_var.trace_add("write", self.update_lists)
        self.target_var.trace_add("write", self.update_lists)
        
        # 创建左侧面板
        self.create_left_panel(main_frame)
        
        # 创建中间控制面板
        self.create_control_panel(main_frame)
        
        # 创建右侧面板
        self.create_right_panel(main_frame)
        
        # 加载配置数据
        self.load_configs()
        
        # 初始显示
        self.update_lists()
        
        # 加载选项配置
        self.load_options_config()
        
        # 加载选择状态
        self.load_selection_state()
        
        # 在窗口关闭时保存选择状态
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_left_panel(self, parent):
        """创左侧面板（源）"""
        panel = ttk.Frame(parent)
        panel.pack(side="left", fill="both", expand=True)
        
        # 创建单选按钮容器
        btn_frame = ttk.Frame(panel)
        btn_frame.pack(fill="x", pady=(0, 10))
        btn_frame.configure(style="TFrame")
        
        # 建单选按钮
        ttk.Radiobutton(
            btn_frame,
            text="国际服",
            value="international",
            variable=self.source_var,
            style="Migration.TRadiobutton"
        ).pack(side="left", padx=5)
        
        ttk.Radiobutton(
            btn_frame,
            text="国服",
            value="china",
            variable=self.source_var,
            style="Migration.TRadiobutton"
        ).pack(side="left", padx=5)
        
        # 创建列表框
        list_frame = ttk.LabelFrame(panel, text="角色列表", padding=5)
        list_frame.pack(fill="both", expand=True)
        
        self.left_listbox = ttk.Treeview(list_frame, show="tree", selectmode="browse")
        self.left_listbox.pack(fill="both", expand=True)

    def create_right_panel(self, parent):
        """创建右侧板（目标）"""
        panel = ttk.Frame(parent)
        panel.pack(side="right", fill="both", expand=True)
        
        # 创建单选按钮容器
        btn_frame = ttk.Frame(panel)
        btn_frame.pack(fill="x", pady=(0, 10))
        btn_frame.configure(style="TFrame")
        
        # 创建单选按钮
        ttk.Radiobutton(
            btn_frame,
            text="国际服",
            value="international",
            variable=self.target_var,
            style="Migration.TRadiobutton"
        ).pack(side="left", padx=5)
        
        ttk.Radiobutton(
            btn_frame,
            text="国服",
            value="china",
            variable=self.target_var,
            style="Migration.TRadiobutton"
        ).pack(side="left", padx=5)
        
        # 创建列表框
        list_frame = ttk.LabelFrame(panel, text="角色列表", padding=5)
        list_frame.pack(fill="both", expand=True)
        
        self.right_listbox = ttk.Treeview(list_frame, show="tree", selectmode="browse")
        self.right_listbox.pack(fill="both", expand=True)

    def create_control_panel(self, parent):
        """创建中间控制面板"""
        control_panel = ttk.Frame(parent)
        control_panel.pack(side="left", fill="y", padx=30)
        
        # 创建迁移按钮
        self.migrate_button = ttk.Button(
            control_panel,
            text="迁移 →",
            style="success.TButton",
            width=10,
            command=self.migrate_config
        )
        self.migrate_button.pack(pady=(0, 20))
        
        # 创建配置选项框架
        options_frame = ttk.LabelFrame(control_panel, text="配置选项", padding=5)
        options_frame.pack(fill="both", expand=True)
        
        # 定义配置选项
        self.config_options = {
            "ACQ.DAT": "近期悄悄话人员列表",
            "ADDON.DAT": "界面设置",
            "COMMON.DAT": "角色设置",
            "CONTROL0.DAT": "角色设置(鼠标模式)",
            "CONTROL1.DAT": "角色设置(手柄模式)",
            "GEARSET.DAT": "套装列表",
            "GS.DAT": "九宫幻卡卡组",
            "HOTBAR.DAT": "热键栏设置",
            "ITEMFDR.DAT": "雇员物品顺序",
            "ITEMODR.DAT": "物品栏、兵装库物品顺序",
            "KEYBIND.DAT": "键位设置",
            "LOGFLTR.DAT": "消息窗口设置",
            "MACRO.DAT": "用户宏(该角色专用)",
            "UISAVE.DAT": "UI使用记录"
        }
        
        # 创建复选框变量
        self.option_vars = {}
        
        # 创建复选框
        for filename, description in self.config_options.items():
            var = ttk.BooleanVar(value=True)  # 默认全选
            self.option_vars[filename] = var
            
            # 创建复选框（移除自动保存）
            ttk.Checkbutton(
                options_frame,
                text=f"{description} – {filename}",
                variable=var,
                style="Migration.TCheckbutton"
            ).pack(anchor="w", pady=2)
        
        # 添加全选/取消全选按钮
        select_frame = ttk.Frame(options_frame)
        select_frame.pack(fill="x", pady=(10, 0))
        
        # 建一个容器来居中放置按钮
        button_container = ttk.Frame(select_frame)
        button_container.pack(expand=True)
        
        ttk.Button(
            button_container,
            text="全选",
            command=lambda: self.toggle_all_options(True),
            width=8
        ).pack(side="left", padx=2)
        
        ttk.Button(
            button_container,
            text="取消全选",
            command=lambda: self.toggle_all_options(False),
            width=8
        ).pack(side="left", padx=2)

    def toggle_all_options(self, state: bool):
        """切换所有选项的状态"""
        for var in self.option_vars.values():
            var.set(state)

    def load_configs(self):
        """加载配置数据"""
        # 加载国际服和国服的标记数据
        self.international_marks = self.load_marks("international")
        self.china_marks = self.load_marks("china")

    def load_marks(self, server_type):
        """加载指定服务器的标记数"""
        config_file = os.path.join("data", f"{server_type}_marks.json")
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def update_lists(self, *args):
        """更新列表显示"""
        # 获取当前选择
        left_selection = self.left_listbox.selection()
        right_selection = self.right_listbox.selection()
        
        # 保存当前选择的原始文件夹名
        left_selected_folder = None
        right_selected_folder = None
        
        if left_selection:
            try:
                left_selected_folder = self.left_listbox.item(left_selection[0])["values"][0]
            except:
                pass
        
        if right_selection:
            try:
                right_selected_folder = self.right_listbox.item(right_selection[0])["values"][0]
            except:
                pass
        
        # 清空两个列表
        for listbox in [self.left_listbox, self.right_listbox]:
            for item in listbox.get_children():
                listbox.delete(item)
        
        # 获取源和目标的路径和标记
        source_type = self.source_var.get()
        target_type = self.target_var.get()
        
        # 获取对应的路径和标记
        source_path = self.international_path if source_type == "international" else self.china_path
        target_path = self.international_path if target_type == "international" else self.china_path
        
        source_marks = self.international_marks if source_type == "international" else self.china_marks
        target_marks = self.international_marks if target_type == "international" else self.china_marks
        
        # 加载源列表
        first_source_item = self.load_folder_list(self.left_listbox, source_path.get(), source_marks)
        # 加载目标列表
        first_target_item = self.load_folder_list(self.right_listbox, target_path.get(), target_marks)
        
        # 恢复左侧选择
        if left_selected_folder:
            # 查找匹配的项
            for item in self.left_listbox.get_children():
                try:
                    if self.left_listbox.item(item)["values"][0] == left_selected_folder:
                        self.left_listbox.selection_set(item)
                        break
                except:
                    continue
            else:
                # 如果没找到匹配项，选择第一项
                if first_source_item:
                    self.left_listbox.selection_set(first_source_item)
        elif first_source_item:
            self.left_listbox.selection_set(first_source_item)
        
        # 恢复右侧选择
        if right_selected_folder:
            # 查找匹配的项
            for item in self.right_listbox.get_children():
                try:
                    if self.right_listbox.item(item)["values"][0] == right_selected_folder:
                        self.right_listbox.selection_set(item)
                        break
                except:
                    continue
            else:
                # 如果没找到匹配项，选择第一项
                if first_target_item:
                    self.right_listbox.selection_set(first_target_item)
        elif first_target_item:
            self.right_listbox.selection_set(first_target_item)

    def load_folder_list(self, listbox, path, marks):
        """加载文件夹列表"""
        if not path:
            return None
            
        try:
            first_item = None
            for item in os.listdir(path):
                if "FFXIV_" in item and os.path.isdir(os.path.join(path, item)):
                    # 使用与角色配置管理相同的显示格式
                    display_name = f"{marks.get(item, item)} ({item})" if item in marks else item
                    # 为每个项目添加唯一标识符
                    unique_id = f"{listbox}_{item}"
                    listbox.insert("", "end", unique_id, text=display_name, values=(item,))
                    if first_item is None:
                        first_item = unique_id
            return first_item
        except Exception as e:
            self.show_message("error", "错误", f"扫描文件夹时出错：{str(e)}")
            return None

    def load_options_config(self):
        """加载选项配置"""
        config_file = os.path.join("data", "migration_options.json")
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                saved_options = json.load(f)
                # 更新选项状态
                for filename, state in saved_options.items():
                    if filename in self.option_vars:
                        self.option_vars[filename].set(state)
        except FileNotFoundError:
            # 如果配置文件不在，使用默认值（选）
            pass

    def save_options_config(self):
        """保存选项配置"""
        config_file = os.path.join("data", "migration_options.json")
        options_state = {
            filename: var.get()
            for filename, var in self.option_vars.items()
        }
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(options_state, f, ensure_ascii=False, indent=2)

    def migrate_config(self):
        """执行配置迁移"""
        # 获取选中的源和目标文件夹
        source_selection = self.left_listbox.selection()
        target_selection = self.right_listbox.selection()
        
        if not source_selection or not target_selection:
            self.show_message("warning", "警告", "请先选择源文件夹和目标文件夹！")
            return
        
        # 获取源和目标路径
        source_type = self.source_var.get()
        target_type = self.target_var.get()
        source_path = self.international_path if source_type == "international" else self.china_path
        target_path = self.international_path if target_type == "international" else self.china_path
        
        # 获取实际的文件夹名（从values中获取）
        source_folder = self.left_listbox.item(source_selection[0])["values"][0]
        target_folder = self.right_listbox.item(target_selection[0])["values"][0]
        
        # 构建完整路径
        source_folder_path = os.path.join(source_path.get(), source_folder)
        target_folder_path = os.path.join(target_path.get(), target_folder)
        
        # 获取选中的配置文件
        selected_files = [
            filename 
            for filename, var in self.option_vars.items() 
            if var.get()
        ]
        
        if not selected_files:
            self.show_message("warning", "警告", "请至少选择一个配置文件！")
            return
        
        # 确认对话框
        if not self.show_message(
            "askyesno", 
            "确认", 
            f"确定要将以下配置从\n{self.format_path(source_folder_path)}\n迁移到\n{self.format_path(target_folder_path)}？\n\n" +
            "\n".join(f"• {self.config_options[file]} – {file}" for file in selected_files)
        ):
            return
        
        # 执行迁移
        try:
            success_count = 0
            for filename in selected_files:
                source_file = os.path.join(source_folder_path, filename)
                target_file = os.path.join(target_folder_path, filename)
                
                if os.path.exists(source_file):
                    try:
                        import shutil
                        shutil.copy2(source_file, target_file)
                        success_count += 1
                    except Exception as e:
                        self.show_message("error", "错", f"复制文件失败{str(e)}")
            
            # 只在成功迁移后保存选项配置
            self.save_options_config()
            
            # 确保窗口在最前面
            self.window.lift()
            
            # 显示成功消息
            messagebox.showinfo(
                "迁移完成",
                f"迁移完成！成功迁移 {success_count} 个配置文件。\n\n" +
                f"从：{self.format_path(source_folder_path)}\n" +
                f"到：{self.format_path(target_folder_path)}",
                parent=self.window
            )
            
        except Exception as e:
            self.window.lift()
            messagebox.showerror("错误", f"迁移过程出错：{str(e)}", parent=self.window)

    def show_message(self, type_, title, message, **kwargs):
        """显示息框"""
        # 放提示音
        self.window.bell()
        
        if type_ == "showinfo":
            return messagebox.showinfo(title, message, parent=self.window, **kwargs)
        elif type_ == "showwarning":
            return messagebox.showwarning(title, message, parent=self.window, **kwargs)
        elif type_ == "showerror":
            return messagebox.showerror(title, message, parent=self.window, **kwargs)
        elif type_ == "askyesno":
            return messagebox.askyesno(title, message, parent=self.window, **kwargs)

    def load_selection_state(self):
        """加载选择状态"""
        try:
            with open(os.path.join("data", "migration_state.json"), 'r', encoding='utf-8') as f:
                state = json.load(f)
                self.source_var.set(state.get("source_server", "international"))
                self.target_var.set(state.get("target_server", "international"))
                
                # 加载选中的配置
                if "source_config" in state:
                    for item in self.left_listbox.get_children():
                        if self.left_listbox.item(item)["values"][0] == state["source_config"]:
                            self.left_listbox.selection_set(item)
                            break
                
                if "target_config" in state:
                    for item in self.right_listbox.get_children():
                        if self.right_listbox.item(item)["values"][0] == state["target_config"]:
                            self.right_listbox.selection_set(item)
                            break
        except FileNotFoundError:
            pass

    def save_selection_state(self):
        """保存选择状态"""
        state = {
            "source_server": self.source_var.get(),
            "target_server": self.target_var.get()
        }
        
        # 保存选中的配置
        source_selection = self.left_listbox.selection()
        if source_selection:
            state["source_config"] = self.left_listbox.item(source_selection[0])["values"][0]
        
        target_selection = self.right_listbox.selection()
        if target_selection:
            state["target_config"] = self.right_listbox.item(target_selection[0])["values"][0]
        
        with open(os.path.join("data", "migration_state.json"), 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False, indent=2)

    def on_closing(self):
        """窗口关闭时的处理"""
        self.save_selection_state()
        self.window.destroy()

    def format_path(self, path):
        """格式化路径用于显示"""
        return path.replace(os.sep, '/')

class CharacterBackupWindow:
    def __init__(self, parent, international_path, china_path, backup_path):
        # 创建新窗口
        self.window = ttk.Toplevel(parent)
        self.window.title("角色配置备份")
        
        # 设置窗口为模态
        self.window.transient(parent)
        
        # 保存参数
        self.parent = parent
        self.international_path = international_path
        self.china_path = china_path
        self.backup_path = backup_path
        
        # 初始化选择的文件夹和当前服务器
        self.selected_folder = None
        self.current_server = "international"  # 设置默认值
        
        # 加载配置数据
        self.load_configs()
        
        # 加载选择状态（在创建界面元素之前）
        self.load_selection_state()
        
        # 定义配置文件列表
        self.config_files = [
            ("ACQ.DAT", "近期悄悄话人员列表"),
            ("ADDON.DAT", "界面设置"),
            ("COMMON.DAT", "角色设置"),
            ("CONTROL0.DAT", "角色设置(鼠标模式)"),
            ("CONTROL1.DAT", "角色设置(手柄模式)"),
            ("GEARSET.DAT", "套装列表"),
            ("GS.DAT", "九宫幻卡卡组"),
            ("HOTBAR.DAT", "热键栏设置"),
            ("ITEMFDR.DAT", "雇员物品顺序"),
            ("ITEMODR.DAT", "物品栏、兵装库物品顺序"),
            ("KEYBIND.DAT", "键位设置"),
            ("LOGFLTR.DAT", "消息窗口设置"),
            ("MACRO.DAT", "用户宏(该角色专用)"),
            ("UISAVE.DAT", "UI使用记录")
        ]
        
        # 设置窗口大小
        window_width = 800
        window_height = 500
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置窗口最小大小
        self.window.minsize(700, 400)
        
        # 设置样式
        style = ttk.Style()
        style.configure("Backup.TRadiobutton", background="#f0f0f0")
        
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # 创建服务器选择区域
        server_frame = ttk.Frame(main_frame)
        server_frame.pack(fill="x", pady=(0, 10))
        
        # 创建单选按钮变量（使用加载的状态）
        self.server_var = ttk.StringVar(value=self.current_server)
        self.server_var.trace_add("write", self.on_server_change)
        
        # 创建单选按钮
        ttk.Radiobutton(
            server_frame,
            text="国际服",
            value="international",
            variable=self.server_var,
            style="Backup.TRadiobutton"
        ).pack(side="left", padx=5)
        
        ttk.Radiobutton(
            server_frame,
            text="国服",
            value="china",
            variable=self.server_var,
            style="Backup.TRadiobutton"
        ).pack(side="left", padx=5)
        
        # 创建列表框
        list_frame = ttk.LabelFrame(main_frame, text="角色列表", padding=5)
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        self.listbox = ttk.Treeview(list_frame, show="tree", selectmode="browse")
        self.listbox.pack(fill="both", expand=True)
        
        # 创建右侧操作区域
        operation_frame = ttk.LabelFrame(main_frame, text="操作", padding=5)
        operation_frame.pack(side="left", fill="both", padx=(5, 0))
        
        # 添加备份按钮
        self.backup_button = ttk.Button(
            operation_frame,
            text="备份",
            command=self.backup_config,
            style="primary.TButton",
            width=15
        )
        self.backup_button.pack(pady=5)
        
        # 添加恢复按钮
        self.restore_button = ttk.Button(
            operation_frame,
            text="恢复",
            command=self.restore_config,
            style="warning.TButton",
            width=15
        )
        self.restore_button.pack(pady=5)
        
        # 最后再扫描文件夹
        self.scan_folders()
        
        # 在窗口关闭时保存选择状态
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_configs(self):
        """加载所有配置数据"""
        # 加载国际服和国服的标记数据
        self.international_marks = self.load_marks("international")
        self.china_marks = self.load_marks("china")

    def load_marks(self, server_type):
        """加载指定服务器的标记数据"""
        config_file = os.path.join("data", f"{server_type}_marks.json")
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def on_server_change(self, *args):
        """服务器选择改变时的处理"""
        self.scan_folders()

    def scan_folders(self):
        """扫描并显示文件夹"""
        # 保存当前选择
        current_selection = None
        selected = self.listbox.selection()
        if selected:
            current_selection = selected[0]
            self.selected_folder = current_selection  # 更新 selected_folder
        
        # 清空列表
        for item in self.listbox.get_children():
            self.listbox.delete(item)
        
        # 获取当前选择的服务器和对应的路径
        server_type = self.server_var.get()
        path_var = self.international_path if server_type == "international" else self.china_path
        marks = self.international_marks if server_type == "international" else self.china_marks
        
        # 获取路径
        base_path = path_var.get()
        backup_base = self.backup_path.get()
        
        if not base_path:
            messagebox.showwarning("警告", "请先在路径设置中设置对应的游戏路径！", parent=self.window)
            return
        
        if not backup_base:
            messagebox.showwarning("警告", "请先在路径设置中设置备份路径！", parent=self.window)
            return
        
        try:
            first_item = None
            for item in os.listdir(base_path):
                if "FFXIV_" in item and os.path.isdir(os.path.join(base_path, item)):
                    # 检查是否有标记，如果有标记则显示"标记名 (文件夹名)"，否则直接显示文件夹名
                    display_name = f"{marks.get(item)} ({item})" if item in marks else item
                    
                    # 检查是否有备份（使用汉字识服务器类型）
                    server_folder = "国际服" if server_type == "international" else "国服"
                    backup_path = os.path.join(backup_base, server_folder, item)
                    has_backup = os.path.exists(backup_path)
                    
                    # 如果有备份，获取最新的修改时间
                    backup_time = ""
                    if has_backup:
                        try:
                            # 获取所有配置文件的最后修改时间
                            backup_files = [f for f in os.listdir(backup_path) if f.endswith('.DAT')]
                            if backup_files:
                                file_times = [
                                    os.path.getmtime(os.path.join(backup_path, f))
                                    for f in backup_files
                                ]
                                latest_time = max(file_times)
                                from datetime import datetime
                                time_str = datetime.fromtimestamp(latest_time).strftime('%Y-%m-%d %H:%M:%S')
                                backup_time = f" [{time_str}]"
                        except Exception:
                            backup_time = " [已备份]"
                    else:
                        backup_time = " [未备份]"
                    
                    # 在显示名称后添加备份状态
                    display_name = f"{display_name}{backup_time}"
                    
                    self.listbox.insert("", "end", item, text=display_name)
                    if first_item is None:
                        first_item = item
            
            # 优先使用保存的选择
            if self.selected_folder and self.selected_folder in os.listdir(base_path):
                self.listbox.selection_set(self.selected_folder)
            # 其次使用当前选择
            elif current_selection and current_selection in os.listdir(base_path):
                self.listbox.selection_set(current_selection)
            # 最后才使用第一项
            elif first_item:
                self.listbox.selection_set(first_item)
                
        except Exception as e:
            messagebox.showerror("错误", f"扫描文件夹时出错：{str(e)}", parent=self.window)

    def backup_config(self):
        """备份配置"""
        # 获取选中的配置
        selected = self.listbox.selection()
        if not selected:
            self.show_message("warning", "警告", "请先选择要备份的角色配置！")
            return
        
        # 获取当前服务器类型和路径
        server_type = self.server_var.get()
        source_path = self.international_path if server_type == "international" else self.china_path
        backup_base = self.backup_path.get()
        
        if not backup_base:
            self.show_message("warning", "警告", "请先设置备份路径！")
            return
        
        # 获取选中的文件夹
        folder_id = selected[0]
        folder_name = folder_id  # 使用原始文件夹名
        source_folder = os.path.join(source_path.get(), folder_name)
        
        # 建备份目录（使用汉字标识服务器类型）
        server_folder = "国际服" if server_type == "international" else "国服"
        backup_folder = os.path.normpath(os.path.join(backup_base, server_folder, folder_name))
        os.makedirs(backup_folder, exist_ok=True)
        
        # 确认备份操作
        if not self.show_message(
            "askyesno",
            "确认备份",
            f"确定要将以下位置的配置备份？\n\n" +
            f"从：{self.format_path(source_folder)}\n" +
            f"到：{self.format_path(backup_folder)}"
        ):
            return
        
        try:
            # 执行备份
            success_count = 0
            for config_file, description in self.config_files:
                source_file = os.path.normpath(os.path.join(source_folder, config_file))
                target_file = os.path.normpath(os.path.join(backup_folder, config_file))
                
                if os.path.exists(source_file):
                    try:
                        import shutil
                        shutil.copy2(source_file, target_file)
                        success_count += 1
                    except Exception as e:
                        self.show_message("error", "错误", f"备份文件失败：{config_file}\n{str(e)}")
            
            # 显示备份结果
            if success_count > 0:
                # 确保窗口在最前面
                self.window.lift()
                
                messagebox.showinfo(
                    "备份完成",
                    f"成功备份 {success_count} 个配置文件到：\n{self.format_path(backup_folder)}",
                    parent=self.window
                )
                
                # 更新选择状态并保存
                self.selected_folder = folder_id
                self.save_selection_state()
                
                # 刷新列表以更新备份状态显示
                self.scan_folders()
            else:
                self.window.lift()
                messagebox.showwarning(
                    "备份结果",
                    f"未能备份任何配置文件！\n请确认源文件夹中包含需要备份的配置文件。",
                    parent=self.window
                )
                
        except Exception as e:
            self.window.lift()
            messagebox.showerror("错误", f"备份过程出错：{str(e)}", parent=self.window)

    def restore_config(self):
        """恢复配置"""
        # 获取选中的配置
        selected = self.listbox.selection()
        if not selected:
            self.show_message("warning", "警告", "请先选择要恢复的角色配置！")
            return
        
        # 获取当前服务器类型和路径
        server_type = self.server_var.get()
        target_path = self.international_path if server_type == "international" else self.china_path
        backup_base = self.backup_path.get()
        
        if not backup_base:
            self.show_message("warning", "警告", "请先设置备份路径！")
            return
        
        # 获取选中的文件夹
        folder_id = selected[0]
        folder_name = folder_id  # 使用原始文件夹名
        target_folder = os.path.join(target_path.get(), folder_name)
        
        # 检查备份是否存在
        server_folder = "国际服" if server_type == "international" else "国服"
        backup_folder = os.path.normpath(os.path.join(backup_base, server_folder, folder_name))
        
        if not os.path.exists(backup_folder):
            self.show_message("warning", "警告", f"未找到该角色的备份：\n{backup_folder}")
            return
        
        # 确认恢复操作
        if not self.show_message(
            "askyesno",
            "确认恢复",
            f"确定要将备份恢复到以下位置？\n\n" +
            f"从：{self.format_path(backup_folder)}\n" +
            f"到：{self.format_path(target_folder)}\n\n" +
            "此操作将覆盖目标文件夹的同名文件！"
        ):
            return
        
        try:
            # 执行恢复
            success_count = 0
            for config_file, description in self.config_files:
                backup_file = os.path.normpath(os.path.join(backup_folder, config_file))
                target_file = os.path.normpath(os.path.join(target_folder, config_file))
                
                if os.path.exists(backup_file):
                    try:
                        import shutil
                        shutil.copy2(backup_file, target_file)
                        success_count += 1
                    except Exception as e:
                        self.show_message("error", "错误", f"恢复文件失败：{config_file}\n{str(e)}")
            
            # 显示恢复结果
            if success_count > 0:
                self.window.lift()
                messagebox.showinfo(
                    "恢复完成",
                    f"成功恢复 {success_count} 个配置文件到：\n{self.format_path(target_folder)}",
                    parent=self.window
                )
            else:
                self.window.lift()
                messagebox.showwarning(
                    "恢复结果",
                    f"未能恢复任何配置文件！\n请确认备份文件夹中包含需要恢复的配置文件。",
                    parent=self.window
                )
                
        except Exception as e:
            self.window.lift()
            messagebox.showerror("错误", f"恢复过程出错：{str(e)}", parent=self.window)

    def show_message(self, type_, title, message, **kwargs):
        """显示消息框"""
        # 播放提示音
        self.window.bell()
        
        if type_ == "showinfo":
            return messagebox.showinfo(title, message, parent=self.window, **kwargs)
        elif type_ == "showwarning":
            return messagebox.showwarning(title, message, parent=self.window, **kwargs)
        elif type_ == "showerror":
            return messagebox.showerror(title, message, parent=self.window, **kwargs)
        elif type_ == "askyesno":
            return messagebox.askyesno(title, message, parent=self.window, **kwargs)

    def load_selection_state(self):
        """加载选择状态"""
        try:
            with open(os.path.normpath(os.path.join("data", "backup_state.json")), 'r', encoding='utf-8') as f:
                state = json.load(f)
                self.current_server = state.get("server", "international")
                self.selected_folder = state.get("folder", None)
        except FileNotFoundError:
            self.current_server = "international"
            self.selected_folder = None

    def save_selection_state(self):
        """保存选择状态"""
        state = {
            "server": self.server_var.get(),
            "folder": self.selected_folder
        }
        with open(os.path.normpath(os.path.join("data", "backup_state.json")), 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False, indent=2)

    def on_closing(self):
        """窗口关闭时的处理"""
        selected = self.listbox.selection()
        if selected:
            self.selected_folder = selected[0]
        self.save_selection_state()
        self.window.destroy()

    def format_path(self, path):
        """格式化路径用于显示"""
        return path.replace(os.sep, '/')

class SoftwareBackupWindow:
    def __init__(self, parent, international_path, china_path, backup_path):
        # 创建新窗口
        self.window = ttk.Toplevel(parent)
        self.window.title("软件配置备份")
        
        # 保存参数
        self.parent = parent
        self.international_path = international_path
        self.china_path = china_path
        self.backup_path = backup_path
        
        # 设置窗口大小
        window_width = 800
        window_height = 500
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置窗口最小大小
        self.window.minsize(700, 400)
        
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # TODO: 添加软件配置备份界面的具体实现

if __name__ == "__main__":
    # 设置 DPI 感知
    set_dpi_awareness()
    
    # 创建主窗口
    root = ttk.Window(themename="cosmo")
    
    # 获取 DPI 缩放因子
    try:
        scaling_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
    except:
        scaling_factor = 1
    
    # 设置初始窗大小和位置
    window_width = int(500 * scaling_factor)
    window_height = int(600 * scaling_factor)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # 设置窗口图标
    icon_path = get_resource_path("3.ico")  # 修改为 3.ico
    if os.path.exists(icon_path):
        root.iconbitmap(icon_path)
    
    # 创建应用实例
    app = App(root)
    
    # 开始主循环
    root.mainloop() 
