import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import sqlite3
import datetime
import json
import os
import sys
import tkinter.font as tkFont
from tkcalendar import Calendar  # 保留但不使用日期选择器

try:
    import matplotlib
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.font_manager import FontProperties
    try:
        import numpy as np
    except ImportError:
        class NumpyStub:
            def __getattr__(self, name):
                return None
        np = NumpyStub()
    from matplotlib.dates import WeekdayLocator, DateFormatter
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    plt = None
    FigureCanvasTkAgg = None
    FontProperties = None
    np = None
    mdates = None
    MATPLOTLIB_AVAILABLE = False

def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def check_dependencies():
    """检查必要的依赖库是否安装"""
    missing = []
    try:
        import matplotlib
    except ImportError:
        missing.append("matplotlib")
    try:
        import numpy
    except ImportError:
        missing.append("numpy")
    try:
        import tkcalendar
    except ImportError:
        missing.append("tkcalendar")
    return missing

class JX3DungeonTracker:
    def __init__(self, root):
        """初始化应用程序"""
        self.root = root
        self.root.title("JX3DungeonTracker - 剑网3副本记录工具")
        
        # 数据库连接
        db_path = resource_path('jx3_dungeon.db')
        self.conn = sqlite3.connect(db_path)
        self.cursor = self.conn.cursor()
        
        # 窗口设置
        width = 1600
        height = 900
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.configure(bg="#f5f5f7")
        self.root.minsize(1024, 600)
        
        # 界面变量初始化
        self.trash_gold_var = tk.StringVar(value="")
        self.iron_gold_var = tk.StringVar(value="")
        self.other_gold_var = tk.StringVar(value="")
        self.fine_gold_var = tk.StringVar(value="")
        self.subsidy_gold_var = tk.StringVar(value="")
        self.lie_down_var = tk.StringVar(value="")
        self.team_type_var = tk.StringVar(value="十人本")
        self.total_gold_var = tk.StringVar(value="0")
        self.personal_gold_var = tk.StringVar(value="0")
        self.special_total_var = tk.StringVar(value="0")
        self.note_var = tk.StringVar(value="")
        self.start_date_var = tk.StringVar(value="")
        self.end_date_var = tk.StringVar(value="")
        
        # 图表字体设置
        if MATPLOTLIB_AVAILABLE:
            font_paths = [
                r'C:\Windows\Fonts\simhei.ttf',
                r'C:\Windows\Fonts\msyh.ttc',
                r'C:\Windows\Fonts\simsun.ttc',
                r'C:\Windows\Fonts\STHUPO.TTF'
            ]
            for path in font_paths:
                if os.path.exists(path):
                    self.chinese_font = FontProperties(fname=path)
                    break
            else:
                plt.rcParams['font.family'] = 'sans-serif'
                plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'KaiTi']
                self.chinese_font = None
        
        # 界面样式设置
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure(".", background="#f5f5f7", foreground="#333")
        self.style.configure("TFrame", background="#f5f5f7")
        self.style.configure("TLabel", background="#f5f5f7", font=("PingFang SC", 10))
        self.style.configure("TButton", font=("PingFang SC", 10), padding=6, background="#e6e6e6")
        self.style.map("TButton", background=[("active", "#d6d6d6")])
        self.style.configure("Treeview", font=("PingFang SC", 9), rowheight=24)
        self.style.configure("Treeview.Heading", font=("PingFang SC", 10), anchor="center")
        self.style.configure("TNotebook", background="#f5f5f7", borderwidth=0)
        self.style.configure("TNotebook.Tab", font=("PingFang SC", 10), padding=[10, 4])
        self.style.configure("TCombobox", padding=4)
        self.style.configure("TEntry", padding=4)
        self.style.configure("TLabelFrame", font=("PingFang SC", 10), padding=8, labelanchor="n")
        
        # 初始化数据库和界面
        self.create_tables()
        self.add_preset_data()
        self.create_ui()
        
        # 事件绑定
        self.trash_gold_var.trace_add("write", self.update_total_gold)
        self.iron_gold_var.trace_add("write", self.update_total_gold)
        self.other_gold_var.trace_add("write", self.update_total_gold)
        self.special_total_var.trace_add("write", self.update_total_gold)
        self.total_gold_var.trace_add("write", self.validate_difference)
        
        # 输入验证
        reg = self.root.register(self.validate_numeric_input)
        vcmd = (reg, '%P')
        self.trash_gold_entry.config(validate="key", validatecommand=vcmd)
        self.iron_gold_entry.config(validate="key", validatecommand=vcmd)
        self.other_gold_entry.config(validate="key", validatecommand=vcmd)
        self.fine_gold_entry.config(validate="key", validatecommand=vcmd)
        self.subsidy_gold_entry.config(validate="key", validatecommand=vcmd)
        self.lie_down_entry.config(validate="key", validatecommand=vcmd)
        self.personal_gold_entry.config(validate="key", validatecommand=vcmd)
        
        # 加载数据
        self.load_dungeon_records()
        self.update_stats()
        self.update_time()
        
        # 关闭事件处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def validate_numeric_input(self, new_value):
        """验证输入是否为有效数字"""
        if new_value == "":
            return True
        try:
            int(new_value)
            return True
        except ValueError:
            return False

    def create_tables(self):
        """创建数据库表结构"""
        # 副本表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS dungeons (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                special_drops TEXT
            )
        ''')
        
        # 记录表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dungeon_id INTEGER NOT NULL,
                trash_gold INTEGER DEFAULT 0,
                iron_gold INTEGER DEFAULT 0,
                other_gold INTEGER DEFAULT 0,
                special_auctions TEXT,
                total_gold INTEGER DEFAULT 0,
                black_owner TEXT,
                worker TEXT,
                time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                team_type TEXT,
                lie_down_count INTEGER DEFAULT 0,
                fine_gold INTEGER DEFAULT 0,
                subsidy_gold INTEGER DEFAULT 0,
                personal_gold INTEGER DEFAULT 0,
                note TEXT DEFAULT ''
            )
        ''')
        
        # 列宽保存表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS column_widths (
                tree_name TEXT PRIMARY KEY,
                widths TEXT
            )
        ''')
        
        self.conn.commit()

    def add_preset_data(self):
        """添加预设副本数据"""
        dungeons = [
            ("狼牙堡·狼神殿", "阿豪（宠物）,遗忘的书函（外观）,醉月玄晶（95级）"),
            ("敖龙岛", "赤纹野正宗（腰部挂件）,隐狐匿踪（特殊面部）,木木（宠物）,星云踏月骓（普通坐骑）,归墟玄晶（100级）"),
            ("范阳夜变", "簪花空竹（腰部挂件）,弃身·肆（特殊腰部）,幽明录（宠物）,润州绣舞筵（家具）,聆音（特殊腰部）,夜泊蝶影（披风）,归墟玄晶（100级）"),
            ("达摩洞", "活色生香（腰部挂件）,冰蚕龙渡（腰部挂件）,猿神发带（头饰）,漫漫香罗（奇趣坐骑）,阿修罗像（家具）,天乙玄晶（110级）"),
            ("白帝江关", "鲤跃龙门（背部挂件）,血佑铃（腰部挂件）,御马踏金·头饰（马具）,御马踏金·鞍饰（马具）,御马踏金·足饰（马具）,御马踏金（马具）,飞毛将军（普通坐骑）,阔豪（脚印）,天乙玄晶（110级）"),
            ("雷域大泽", "大眼崽（宠物）,灵虫石像（家具）,脊骨王座（家具）,掠影无迹（背部挂件）,荒原切（腰部挂件）,游空竹翼（背部挂件）"),
            ("河阳之战", "云鹤报捷（玩具）,玄域辟甲·头饰（马具）,玄域辟甲·鞍饰（马具）,玄域辟甲·足饰（马具）,玄域辟甲（马具）,扇风耳（宠物）,墨言（特殊背部）,天乙玄晶（110级）"),
            ("西津渡", "卯金修德（背部挂件）,相思尽（腰部挂件）,比翼剪（背部挂件）,静子（宠物）,泽心龙头像（家具）,焚金阙（外观）,赤发狻猊（头饰）,太一玄晶（120级）"),
            ("武狱黑牢", "驭己刃（腰部挂件）,象心灵犀（玩具）,心定（头饰）,幽兰引芳（脚印）,武氏挂旗（家具）,白鬼血泣（披风）,太一玄晶（120级）"),
            ("九老洞", "武圣（背部挂件）,不渡（特殊腰部）,灵龟·卜逆（奇趣坐骑）,朱雀·灼（家具）,青龙·木（家具）,麒麟·祝瑞（宠物）,幻月（特殊腰部）,太一玄晶（120级）"),
            ("冷龙峰", "涉海翎（帽子）,透骨香（腰部挂件）,转珠天轮（玩具）,鸷（宠物）,炽芒·邪锋（特殊腰部）,祆教神鸟像（家具）,太一玄晶（120级）")
        ]
        
        try:
            self.cursor.executemany('''
                INSERT OR IGNORE INTO dungeons (name, special_drops) 
                VALUES (?, ?)
            ''', dungeons)
            self.conn.commit()
        except:
            pass

    def create_ui(self):
        """创建主界面"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 标题栏
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 8))
        title_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, weight=0)
        
        title_label = ttk.Label(
            title_frame, 
            text="JX3DungeonTracker - 剑网3副本记录工具", 
            font=("PingFang SC", 16, "bold"),
            anchor="w"
        )
        title_label.grid(row=0, column=0, sticky="w", padx=10)
        
        self.time_var = tk.StringVar(value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        time_label = ttk.Label(
            title_frame, 
            textvariable=self.time_var, 
            font=("PingFang SC", 12),
            anchor="e"
        )
        time_label.grid(row=0, column=1, sticky="e")
        
        # 选项卡
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        record_frame = ttk.Frame(notebook)
        notebook.add(record_frame, text="副本记录")
        
        stats_frame = ttk.Frame(notebook)
        notebook.add(stats_frame, text="数据总览")
        
        preset_frame = ttk.Frame(notebook)
        notebook.add(preset_frame, text="副本预设")
        
        self.create_record_tab(record_frame)
        self.create_stats_tab(stats_frame)
        self.create_preset_tab(preset_frame)
        
        # 加载列宽设置
        self.load_column_widths()

    def on_close(self):
        """关闭窗口时的处理"""
        self.save_column_widths()
        self.root.destroy()

    def save_column_widths(self):
        """保存列宽设置到数据库"""
        trees = [
            (self.record_tree, "record_tree"),
            (self.worker_stats_tree, "worker_stats_tree"),
            (self.dungeon_tree, "dungeon_tree")
        ]
        
        for tree, tree_name in trees:
            if tree:
                widths = {}
                for col in tree["columns"]:
                    widths[col] = tree.column(col, "width")
                widths_json = json.dumps(widths)
                self.cursor.execute('''
                    INSERT OR REPLACE INTO column_widths (tree_name, widths)
                    VALUES (?, ?)
                ''', (tree_name, widths_json))
        self.conn.commit()

    def load_column_widths(self):
        """从数据库加载列宽设置"""
        trees = [
            (self.record_tree, "record_tree"),
            (self.worker_stats_tree, "worker_stats_tree"),
            (self.dungeon_tree, "dungeon_tree")
        ]
        
        for tree, tree_name in trees:
            if tree:
                self.cursor.execute('''
                    SELECT widths FROM column_widths WHERE tree_name = ?
                ''', (tree_name,))
                result = self.cursor.fetchone()
                if result and result[0]:
                    try:
                        widths = json.loads(result[0])
                        for col, width in widths.items():
                            if col in tree["columns"]:
                                tree.column(col, width=width)
                    except json.JSONDecodeError:
                        pass

    def setup_treeview_column_resizing(self, tree):
        """设置表格列宽调整功能"""
        tree.bind("<Double-1>", lambda event: self.on_treeview_double_click(event, tree))
        for col in tree["columns"]:
            tree.heading(col, command=lambda c=col: self.adjust_column_width(tree, c))

    def on_treeview_double_click(self, event, tree):
        """双击表头调整列宽"""
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            col = tree.identify_column(event.x)
            col_id = col.lstrip("#")
            try:
                col_index = int(col_id) - 1
                columns = tree["columns"]
                if 0 <= col_index < len(columns):
                    col_name = columns[col_index]
                    self.adjust_column_width(tree, col_name)
            except ValueError:
                pass

    def adjust_column_width(self, tree, column_id):
        """自动调整列宽"""
        heading_text = tree.heading(column_id, "text")
        font = tkFont.nametofont("TkDefaultFont")
        max_width = font.measure(heading_text) + 20
        
        for child in tree.get_children():
            item_text = tree.set(child, column_id)
            item_width = font.measure(item_text) + 20
            if item_width > max_width:
                max_width = item_width
        
        min_width = 50
        max_width = max(min_width, max_width)
        max_width = min(max_width, 500)
        tree.column(column_id, width=max_width)

    def update_time(self):
        """更新时间显示"""
        now = datetime.datetime.now()
        self.time_var.set(now.strftime("%Y-%m-%d %H:%M:%S"))
        self.root.after(1000, self.update_time)

    def create_record_tab(self, parent):
        """创建副本记录选项卡"""
        pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左侧表单区域
        form_frame = ttk.LabelFrame(pane, text="副本记录管理", padding=8)
        form_frame.config(width=350)
        
        # 副本选择行
        dungeon_row = ttk.Frame(form_frame)
        dungeon_row.pack(fill=tk.X, pady=3)
        ttk.Label(dungeon_row, text="副本名称:").pack(side=tk.LEFT, padx=(0, 5))
        self.dungeon_var = tk.StringVar()
        self.dungeon_combo = ttk.Combobox(dungeon_row, textvariable=self.dungeon_var, width=25)
        self.dungeon_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.dungeon_combo.bind("<<ComboboxSelected>>", self.on_dungeon_select)
        
        # 特殊掉落区域
        special_frame = ttk.LabelFrame(form_frame, text="特殊掉落", padding=6)
        special_frame.pack(fill=tk.BOTH, pady=(0, 5), expand=True)
        
        tree_frame = ttk.Frame(special_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=3)
        
        columns = ("item", "price")
        self.special_tree = ttk.Treeview(
            tree_frame, 
            columns=columns, 
            show="headings",
            height=3,
            selectmode="browse"
        )
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.special_tree.yview)
        self.special_tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.special_tree.xview)
        self.special_tree.configure(xscrollcommand=hsb.set)
        
        self.special_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        self.special_tree.heading("item", text="物品", anchor="center")
        self.special_tree.heading("price", text="金额", anchor="center")
        self.special_tree.column("item", width=120, anchor=tk.CENTER)
        self.special_tree.column("price", width=60, anchor=tk.CENTER)
        
        # 添加特殊掉落控件
        add_special_frame = ttk.Frame(special_frame)
        add_special_frame.pack(fill=tk.X, pady=(8, 0))
        
        ttk.Label(add_special_frame, text="物品:").grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.special_item_var = tk.StringVar()
        self.special_item_combo = ttk.Combobox(add_special_frame, textvariable=self.special_item_var, width=15)
        self.special_item_combo.grid(row=0, column=1, padx=(0, 5), sticky="ew")
        
        ttk.Label(add_special_frame, text="金额:").grid(row=0, column=2, padx=(5, 5), sticky="w")
        self.special_price_var = tk.StringVar()
        self.special_price_entry = ttk.Entry(add_special_frame, textvariable=self.special_price_var, width=10)
        self.special_price_entry.grid(row=0, column=3, padx=(0, 5), sticky="w")
        
        add_btn = ttk.Button(add_special_frame, text="添加", width=6, command=self.add_special_item)
        add_btn.grid(row=0, column=4, padx=(5, 0))
        add_special_frame.columnconfigure(1, weight=1)
        
        # 团队项目区域
        team_frame = ttk.LabelFrame(form_frame, text="团队项目", padding=6)
        team_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(team_frame, text="散件金额:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.trash_gold_entry = ttk.Entry(team_frame, textvariable=self.trash_gold_var, width=10)
        self.trash_gold_entry.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(team_frame, text="小铁金额:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.iron_gold_entry = ttk.Entry(team_frame, textvariable=self.iron_gold_var, width=10)
        self.iron_gold_entry.grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(team_frame, text="特殊金额:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.special_total_var = tk.StringVar(value="0")
        ttk.Label(team_frame, textvariable=self.special_total_var, width=10).grid(row=1, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(team_frame, text="其他金额:").grid(row=1, column=2, padx=3, pady=2, sticky=tk.W)
        self.other_gold_entry = ttk.Entry(team_frame, textvariable=self.other_gold_var, width=10)
        self.other_gold_entry.grid(row=1, column=3, padx=3, pady=2, sticky=tk.W)
        
        # 个人项目区域
        personal_frame = ttk.LabelFrame(form_frame, text="个人项目", padding=6)
        personal_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(personal_frame, text="补贴金额:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.subsidy_gold_entry = ttk.Entry(personal_frame, textvariable=self.subsidy_gold_var, width=10)
        self.subsidy_gold_entry.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(personal_frame, text="罚款金额:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.fine_gold_entry = ttk.Entry(personal_frame, textvariable=self.fine_gold_var, width=10)
        self.fine_gold_entry.grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)
        
        # 团队信息区域
        info_frame = ttk.LabelFrame(form_frame, text="团队信息", padding=6)
        info_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(info_frame, text="团队类型:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.team_type_combo = ttk.Combobox(info_frame, textvariable=self.team_type_var, 
                                           values=["十人本", "二十五人本"], width=10, state="readonly")
        self.team_type_combo.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        self.team_type_combo.current(0)
        
        ttk.Label(info_frame, text="躺拍人数:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.lie_down_entry = ttk.Entry(info_frame, textvariable=self.lie_down_var, width=6)
        self.lie_down_entry.grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(info_frame, text="黑本:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.black_owner_var = tk.StringVar()
        self.black_owner_combo = ttk.Combobox(info_frame, textvariable=self.black_owner_var, width=20)
        self.black_owner_combo.grid(row=1, column=1, padx=3, pady=2, sticky=tk.W, columnspan=3)
        
        ttk.Label(info_frame, text="打工仔:").grid(row=2, column=0, padx=3, pady=2, sticky=tk.W)
        self.worker_var = tk.StringVar()
        self.worker_combo = ttk.Combobox(info_frame, textvariable=self.worker_var, width=20)
        self.worker_combo.grid(row=2, column=1, padx=3, pady=2, sticky=tk.W, columnspan=3)
        
        ttk.Label(info_frame, text="备注:").grid(row=3, column=0, padx=3, pady=2, sticky=tk.W)
        self.note_entry = ttk.Entry(info_frame, textvariable=self.note_var, width=20)
        self.note_entry.grid(row=3, column=1, padx=3, pady=2, sticky=tk.W, columnspan=3)
        
        # 工资信息区域
        gold_frame = ttk.LabelFrame(form_frame, text="工资信息", padding=6)
        gold_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(gold_frame, text="团队总工资:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.total_gold_entry = ttk.Entry(gold_frame, textvariable=self.total_gold_var, width=10)
        self.total_gold_entry.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        self.difference_var = tk.StringVar(value="差额: 0")
        self.difference_label = ttk.Label(
            gold_frame, 
            textvariable=self.difference_var, 
            font=("PingFang SC", 9),
            foreground="#e74c3c"
        )
        self.difference_label.grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(gold_frame, text="个人工资:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.personal_gold_entry = ttk.Entry(gold_frame, textvariable=self.personal_gold_var, width=10)
        self.personal_gold_entry.grid(row=1, column=1, padx=3, pady=2, sticky=tk.W)
        
        # 按钮区域
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.add_btn = ttk.Button(btn_frame, text="保存记录", command=self.add_record, width=10)
        self.add_btn.grid(row=0, column=0, padx=2, sticky="ew")
        
        self.edit_btn = ttk.Button(btn_frame, text="编辑记录", command=self.edit_record, width=10)
        self.edit_btn.grid(row=0, column=1, padx=2, sticky="ew")
        
        self.update_btn = ttk.Button(btn_frame, text="更新记录", command=self.update_record, 
                                   state=tk.DISABLED, width=10)
        self.update_btn.grid(row=0, column=2, padx=2, sticky="ew")
        
        clear_btn = ttk.Button(btn_frame, text="清空表单", command=self.clear_form, width=10)
        clear_btn.grid(row=0, column=3, padx=2, sticky="ew")
        
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        btn_frame.columnconfigure(3, weight=1)
        
        # 右侧列表区域
        list_frame = ttk.LabelFrame(pane, text="副本记录列表", padding=8)
        pane.add(form_frame, weight=1)
        pane.add(list_frame, weight=2)
        
        # 搜索区域 - 修改为三行布局
        search_frame = ttk.Frame(list_frame)
        search_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        
        # 第一行：黑本、打工仔
        search_row1 = ttk.Frame(search_frame)
        search_row1.pack(fill=tk.X, pady=2)
        
        ttk.Label(search_row1, text="黑本:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_owner_var = tk.StringVar()
        self.search_owner_combo = ttk.Combobox(search_row1, textvariable=self.search_owner_var, width=15)
        self.search_owner_combo.pack(side=tk.LEFT, padx=(0, 8))
        
        ttk.Label(search_row1, text="打工仔:").pack(side=tk.LEFT, padx=(5, 5))
        self.search_worker_var = tk.StringVar()
        self.search_worker_combo = ttk.Combobox(search_row1, textvariable=self.search_worker_var, width=15)
        self.search_worker_combo.pack(side=tk.LEFT, padx=(0, 8))
        
        # 第二行：副本、特殊掉落
        search_row2 = ttk.Frame(search_frame)
        search_row2.pack(fill=tk.X, pady=2)
        
        ttk.Label(search_row2, text="副本:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_dungeon_var = tk.StringVar()
        self.search_dungeon_combo = ttk.Combobox(search_row2, textvariable=self.search_dungeon_var, width=15)
        self.search_dungeon_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.search_dungeon_combo.bind("<<ComboboxSelected>>", self.on_search_dungeon_select)
        
        ttk.Label(search_row2, text="特殊掉落:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_item_var = tk.StringVar()
        self.search_item_combo = ttk.Combobox(search_row2, textvariable=self.search_item_var, width=15)
        self.search_item_combo.pack(side=tk.LEFT, padx=(0, 8))
        
        # 第三行：开始时间、结束时间、团队类型
        search_row3 = ttk.Frame(search_frame)
        search_row3.pack(fill=tk.X, pady=2)
        
        ttk.Label(search_row3, text="开始时间:").pack(side=tk.LEFT, padx=(0, 3))
        self.start_date_entry = ttk.Entry(search_row3, textvariable=self.start_date_var, width=12)
        self.start_date_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(search_row3, text="结束时间:").pack(side=tk.LEFT, padx=(0, 3))
        self.end_date_entry = ttk.Entry(search_row3, textvariable=self.end_date_var, width=12)
        self.end_date_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(search_row3, text="团队类型:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_team_type_var = tk.StringVar()
        self.search_team_type_combo = ttk.Combobox(search_row3, textvariable=self.search_team_type_var, 
                                                  values=["", "十人本", "二十五人本"], width=10)
        self.search_team_type_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        # 第四行：按钮区域
        search_row4 = ttk.Frame(search_frame)
        search_row4.pack(fill=tk.X, pady=2)
        
        left_btn_frame = ttk.Frame(search_row4)
        left_btn_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        search_btn = ttk.Button(left_btn_frame, text="查询", command=self.search_records, width=10)
        search_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        reset_btn = ttk.Button(left_btn_frame, text="重置", command=self.reset_search, width=10)
        reset_btn.pack(side=tk.LEFT, padx=5)
        
        right_btn_frame = ttk.Frame(search_row4)
        right_btn_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        import_btn = ttk.Button(right_btn_frame, text="导入数据", command=self.import_data, width=10)
        import_btn.pack(side=tk.RIGHT, padx=(0, 5))
        
        export_btn = ttk.Button(right_btn_frame, text="导出数据", command=self.export_data, width=10)
        export_btn.pack(side=tk.RIGHT, padx=(0, 5))
        
        # 记录表格
        columns = ("id", "dungeon", "time", "team_type", "lie_down", "total", "personal", "black_owner", "worker", "note")
        self.record_tree = ttk.Treeview(
            list_frame, 
            columns=columns, 
            show="headings",
            selectmode="extended",
            height=10
        )
        
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.record_tree.yview)
        self.record_tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.record_tree.xview)
        self.record_tree.configure(xscrollcommand=hsb.set)
        
        self.record_tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")
        
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)
        
        # 表格列设置
        self.record_tree.heading("id", text="ID", anchor="center")
        self.record_tree.heading("dungeon", text="副本名称", anchor="center")
        self.record_tree.heading("time", text="时间", anchor="center")
        self.record_tree.heading("team_type", text="团队类型", anchor="center")
        self.record_tree.heading("lie_down", text="躺拍人数", anchor="center")
        self.record_tree.heading("total", text="团队总工资", anchor="center")
        self.record_tree.heading("personal", text="个人工资", anchor="center")
        self.record_tree.heading("black_owner", text="黑本", anchor="center")
        self.record_tree.heading("worker", text="打工仔", anchor="center")
        self.record_tree.heading("note", text="备注", anchor="center")
        
        self.record_tree.column("id", width=35, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("dungeon", width=100, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("time", width=100, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("team_type", width=70, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("lie_down", width=70, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("total", width=100, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("personal", width=100, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("black_owner", width=70, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("worker", width=70, anchor=tk.CENTER, stretch=False)
        self.record_tree.column("note", width=100, anchor=tk.CENTER, stretch=True)
        
        self.setup_treeview_column_resizing(self.record_tree)
        
        # 提示框设置
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.withdraw()
        self.tooltip.overrideredirect(True)
        self.tooltip.configure(bg="#ffffe0", relief="solid", borderwidth=1)
        self.tooltip_label = tk.Label(
            self.tooltip, 
            text="", 
            justify=tk.LEFT, 
            bg="#ffffe0",
            font=("PingFang SC", 9)
        )
        self.tooltip_label.pack(padx=5, pady=5)
        
        self.record_tree.bind("<Motion>", self.on_tree_motion)
        self.record_tree.bind("<Leave>", self.hide_tooltip)
        
        # 加载数据
        self.load_dungeon_options()
        self.load_black_owner_options()
        self.search_item_combo['values'] = self.get_all_special_items()

    def validate_date(self, date_str):
        """验证日期格式是否正确"""
        if not date_str:
            return True
        try:
            datetime.datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def on_tree_motion(self, event):
        """鼠标在表格上移动时显示详细信息"""
        self.tooltip.withdraw()
        row_id = self.record_tree.identify_row(event.y)
        col_id = self.record_tree.identify_column(event.x)
        
        if not row_id:
            return
            
        item = self.record_tree.item(row_id)
        values = item['values']
        col_name = self.record_tree.heading(col_id)["text"]
        
        if col_name == "团队总工资":
            record_id = values[0]
            self.cursor.execute('''
                SELECT trash_gold, iron_gold, other_gold, special_auctions
                FROM records
                WHERE id = ?
            ''', (record_id,))
            details = self.cursor.fetchone()
            
            if details:
                trash, iron, other, specials_json = details
                specials = json.loads(specials_json) if specials_json else []
                special_total = 0
                
                if specials:
                    for item in specials:
                        try:
                            special_total += int(item['price'])
                        except (ValueError, TypeError):
                            pass
                
                tooltip_text = f"散件金额: {trash}\n"
                tooltip_text += f"小铁金额: {iron}\n"
                tooltip_text += f"其他金额: {other}\n"
                tooltip_text += f"特殊拍卖总金额: {special_total}\n\n"
                
                if specials:
                    tooltip_text += "特殊拍卖明细:\n"
                    for item in specials:
                        tooltip_text += f"  {item['item']}: {item['price']}\n"
                else:
                    tooltip_text += "特殊拍卖明细: 无"
                    
                self.show_tooltip(event, tooltip_text)
                
        elif col_name == "个人工资":
            record_id = values[0]
            self.cursor.execute('''
                SELECT fine_gold, subsidy_gold
                FROM records
                WHERE id = ?
            ''', (record_id,))
            details = self.cursor.fetchone()
            
            if details:
                fine, subsidy = details
                tooltip_text = f"罚款金额: {fine}\n"
                tooltip_text += f"补贴金额: {subsidy}"
                self.show_tooltip(event, tooltip_text)

    def show_tooltip(self, event, text):
        """显示提示框"""
        self.tooltip_label.config(text=text)
        self.tooltip.update_idletasks()
        
        x = event.x_root + 15
        y = event.y_root + 15
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        if x + self.tooltip.winfo_width() > screen_width:
            x = event.x_root - self.tooltip.winfo_width() - 5
        if y + self.tooltip.winfo_height() > screen_height:
            y = event.y_root - self.tooltip.winfo_height() - 5
            
        self.tooltip.geometry(f"+{x}+{y}")
        self.tooltip.deiconify()

    def hide_tooltip(self, event=None):
        """隐藏提示框"""
        self.tooltip.withdraw()

    def on_search_dungeon_select(self, event=None):
        """选择副本时更新特殊掉落选项"""
        selected_dungeon = self.search_dungeon_var.get()
        if selected_dungeon:
            self.cursor.execute("SELECT special_drops FROM dungeons WHERE name=?", (selected_dungeon,))
            result = self.cursor.fetchone()
            if result and result[0]:
                drops = [item.strip() for item in result[0].split(',')]
            else:
                drops = []
            self.search_item_combo['values'] = drops
        else:
            self.search_item_combo['values'] = self.get_all_special_items()

    def reset_search(self):
        """重置搜索条件"""
        self.search_dungeon_var.set("")
        self.search_item_var.set("")
        self.search_owner_var.set("")
        self.search_worker_var.set("")
        self.search_team_type_var.set("")
        self.start_date_var.set("")
        self.end_date_var.set("")
        self.on_search_dungeon_select()
        self.load_dungeon_records()

    def on_dungeon_select(self, event):
        """选择副本时更新特殊掉落选项"""
        selected_dungeon = self.dungeon_var.get()
        if not selected_dungeon:
            return
            
        self.cursor.execute("SELECT special_drops FROM dungeons WHERE name=?", (selected_dungeon,))
        result = self.cursor.fetchone()
        if result and result[0]:
            drops = result[0].split(',')
            self.special_item_combo['values'] = drops
        else:
            self.special_item_combo['values'] = []

    def load_dungeon_options(self):
        """加载副本选项"""
        self.cursor.execute("SELECT name FROM dungeons")
        dungeons = [row[0] for row in self.cursor.fetchall()]
        self.dungeon_combo['values'] = dungeons
        self.search_dungeon_combo['values'] = dungeons
        self.search_item_combo['values'] = self.get_all_special_items()

    def load_black_owner_options(self):
        """加载黑本和打工仔选项"""
        self.cursor.execute("SELECT DISTINCT black_owner FROM records WHERE black_owner IS NOT NULL AND black_owner != ''")
        owners = [row[0] for row in self.cursor.fetchall()]
        self.black_owner_combo['values'] = owners
        self.search_owner_combo['values'] = owners
        
        self.cursor.execute("SELECT DISTINCT worker FROM records WHERE worker IS NOT NULL AND worker != ''")
        workers = [row[0] for row in self.cursor.fetchall()]
        self.worker_combo['values'] = workers
        self.search_worker_combo['values'] = workers

    def get_all_special_items(self):
        """获取所有特殊掉落物品"""
        items = set()
        self.cursor.execute("SELECT special_drops FROM dungeons")
        for row in self.cursor.fetchall():
            if row[0]:
                for item in row[0].split(','):
                    items.add(item.strip())
        return list(items)

    def get_special_items_by_dungeon(self, dungeon_name):
        """获取指定副本的特殊掉落物品"""
        self.cursor.execute("SELECT special_drops FROM dungeons WHERE name=?", (dungeon_name,))
        result = self.cursor.fetchone()
        if result and result[0]:
            return [item.strip() for item in result[0].split(',')]
        return []

    def create_stats_tab(self, parent):
        """创建数据统计选项卡"""
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 如果没有安装matplotlib，显示错误信息
        if not MATPLOTLIB_AVAILABLE:
            error_frame = ttk.Frame(main_frame)
            error_frame.pack(fill=tk.BOTH, expand=True, pady=20)
            error_label = ttk.Label(
                error_frame, 
                text="需要安装matplotlib和numpy库才能显示统计图表\n请运行: pip install matplotlib numpy", 
                foreground="red", 
                font=("PingFang SC", 10),
                anchor="center",
                justify="center"
            )
            error_label.pack(fill=tk.BOTH, expand=True)
            return
        
        # 左侧统计卡片区域
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # 总记录数卡片
        card1 = ttk.LabelFrame(left_frame, text="总记录数", padding=(10, 8))
        card1.pack(fill=tk.X, pady=5)
        self.total_records_var = tk.StringVar(value="0")
        ttk.Label(
            card1, 
            textvariable=self.total_records_var, 
            font=("PingFang SC", 14, "bold"),
            anchor="center"
        ).pack(fill=tk.BOTH, expand=True)
        
        # 团队总工资卡片
        card2 = ttk.LabelFrame(left_frame, text="团队总工资", padding=(10, 8))
        card2.pack(fill=tk.X, pady=5)
        self.team_total_gold_var = tk.StringVar(value="0")
        ttk.Label(
            card2, 
            textvariable=self.team_total_gold_var,
            font=("PingFang SC", 14, "bold"),
            anchor="center"
        ).pack(fill=tk.BOTH, expand=True)
        
        # 团队最高工资卡片
        card3 = ttk.LabelFrame(left_frame, text="团队最高工资", padding=(10, 8))
        card3.pack(fill=tk.X, pady=5)
        self.team_max_gold_var = tk.StringVar(value="0")
        ttk.Label(
            card3, 
            textvariable=self.team_max_gold_var,
            font=("PingFang SC", 14, "bold"),
            anchor="center"
        ).pack(fill=tk.BOTH, expand=True)
        
        # 个人总工资卡片
        card4 = ttk.LabelFrame(left_frame, text="个人总工资", padding=(10, 8))
        card4.pack(fill=tk.X, pady=5)
        self.personal_total_gold_var = tk.StringVar(value="0")
        ttk.Label(
            card4, 
            textvariable=self.personal_total_gold_var,
            font=("PingFang SC", 14, "bold"),
            anchor="center"
        ).pack(fill=tk.BOTH, expand=True)
        
        # 个人最高工资卡片
        card5 = ttk.LabelFrame(left_frame, text="个人最高工资", padding=(10, 8))
        card5.pack(fill=tk.X, pady=5)
        self.personal_max_gold_var = tk.StringVar(value="0")
        ttk.Label(
            card5, 
            textvariable=self.personal_max_gold_var,
            font=("PingFang SC", 14, "bold"),
            anchor="center"
        ).pack(fill=tk.BOTH, expand=True)
        
        # 右侧图表区域
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 图表区域
        chart_frame = ttk.LabelFrame(right_frame, text="每周数据统计", padding=(8, 6))
        chart_frame.pack(fill=tk.BOTH, expand=True)
        
        if MATPLOTLIB_AVAILABLE:
            # 设置中文字体
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'KaiTi', 'Arial Unicode MS']
            plt.rcParams['axes.unicode_minus'] = False
            
            # 创建图表
            self.fig, self.ax = plt.subplots(figsize=(6, 2.5), dpi=100)
            self.fig.patch.set_facecolor('#f5f5f7')
            self.ax.set_facecolor('#f5f5f7')
            
            # 将图表嵌入到Tkinter
            self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 打工仔统计区域
        worker_frame = ttk.LabelFrame(right_frame, text="打工仔统计", padding=(8, 6))
        worker_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        
        tree_frame = ttk.Frame(worker_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # 打工仔统计表格
        columns = ("worker", "count", "total_personal", "avg_personal")
        self.worker_stats_tree = ttk.Treeview(
            tree_frame, 
            columns=columns, 
            show="headings",
            height=5
        )
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.worker_stats_tree.yview)
        self.worker_stats_tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.worker_stats_tree.xview)
        self.worker_stats_tree.configure(xscrollcommand=hsb.set)
        
        self.worker_stats_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # 表格列设置
        self.worker_stats_tree.heading("worker", text="打工仔", anchor="center")
        self.worker_stats_tree.heading("count", text="参与次数", anchor="center")
        self.worker_stats_tree.heading("total_personal", text="总工资", anchor="center")
        self.worker_stats_tree.heading("avg_personal", text="平均工资", anchor="center")
        
        self.worker_stats_tree.column("worker", width=100, anchor=tk.CENTER)
        self.worker_stats_tree.column("count", width=70, anchor=tk.CENTER)
        self.worker_stats_tree.column("total_personal", width=100, anchor=tk.CENTER)
        self.worker_stats_tree.column("avg_personal", width=100, anchor=tk.CENTER)
        
        self.setup_treeview_column_resizing(self.worker_stats_tree)

    def create_preset_tab(self, parent):
        """创建副本预设选项卡"""
        pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True)
        
        # 左侧列表区域
        list_frame = ttk.LabelFrame(pane, text="副本列表", padding=8)
        list_frame.config(width=280)
        
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # 副本表格
        columns = ("id", "name", "drops")
        self.dungeon_tree = ttk.Treeview(
            tree_frame, 
            columns=columns, 
            show="headings",
            height=10
        )
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.dungeon_tree.yview)
        self.dungeon_tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.dungeon_tree.xview)
        self.dungeon_tree.configure(xscrollcommand=hsb.set)
        
        self.dungeon_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # 表格列设置
        self.dungeon_tree.heading("id", text="ID", anchor="center")
        self.dungeon_tree.heading("name", text="副本名称", anchor="center")
        self.dungeon_tree.heading("drops", text="特殊掉落", anchor="center")
        
        self.dungeon_tree.column("id", width=40, anchor=tk.CENTER)
        self.dungeon_tree.column("name", width=100, anchor=tk.CENTER)
        self.dungeon_tree.column("drops", width=100, anchor=tk.CENTER)
        
        self.setup_treeview_column_resizing(self.dungeon_tree)
        
        # 按钮区域
        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(fill=tk.X, pady=(8, 0))
        
        add_btn = ttk.Button(btn_frame, text="新增副本", command=self.add_dungeon, width=10)
        add_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        edit_btn = ttk.Button(btn_frame, text="编辑副本", command=self.edit_dungeon, width=10)
        edit_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        delete_btn = ttk.Button(btn_frame, text="删除副本", command=self.delete_dungeon, width=10)
        delete_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        # 右侧表单区域
        form_frame = ttk.LabelFrame(pane, text="副本详情", padding=8)
        pane.add(list_frame, weight=1)
        pane.add(form_frame, weight=2)
        
        # 副本名称
        ttk.Label(form_frame, text="副本名称:").grid(row=0, column=0, padx=4, pady=5, sticky=tk.W)
        self.preset_name_var = tk.StringVar()
        self.preset_name_entry = ttk.Entry(form_frame, textvariable=self.preset_name_var, width=30)
        self.preset_name_entry.grid(row=0, column=1, padx=4, pady=5, sticky=tk.W+tk.E, columnspan=2)
        
        # 特殊掉落
        ttk.Label(form_frame, text="特殊掉落:").grid(row=1, column=0, padx=4, pady=5, sticky=tk.NW)
        self.preset_drops_text = scrolledtext.ScrolledText(form_frame, width=40, height=6, font=("PingFang SC", 9))
        self.preset_drops_text.grid(row=1, column=1, padx=4, pady=5, sticky=tk.NSEW, columnspan=2)
        
        # 批量添加区域
        batch_frame = ttk.LabelFrame(form_frame, text="批量添加特殊掉落", padding=6)
        batch_frame.grid(row=2, column=0, columnspan=3, padx=4, pady=5, sticky=tk.W+tk.E)
        
        ttk.Label(batch_frame, text="输入多个物品，用逗号分隔:").pack(side=tk.TOP, anchor=tk.W, pady=(0, 4))
        self.batch_items_var = tk.StringVar()
        batch_entry = ttk.Entry(batch_frame, textvariable=self.batch_items_var, width=35)
        batch_entry.pack(side=tk.LEFT, padx=(0, 8), fill=tk.X, expand=True)
        
        batch_btn = ttk.Button(batch_frame, text="添加", command=self.batch_add_items, width=10)
        batch_btn.pack(side=tk.LEFT)
        
        # 按钮区域
        btn_frame2 = ttk.Frame(form_frame)
        btn_frame2.grid(row=3, column=0, columnspan=3, pady=8, sticky=tk.E)
        
        save_btn = ttk.Button(btn_frame2, text="保存", width=10, command=self.save_dungeon)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        clear_btn = ttk.Button(btn_frame2, text="清空", width=10, command=self.clear_preset_form)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        form_frame.columnconfigure(1, weight=1)
        form_frame.rowconfigure(1, weight=1)
        
        # 加载数据
        self.load_dungeon_presets()

    def load_dungeon_presets(self):
        """加载副本预设数据"""
        self.dungeon_tree.delete(*self.dungeon_tree.get_children())
        self.cursor.execute("SELECT id, name, special_drops FROM dungeons")
        for row in self.cursor.fetchall():
            self.dungeon_tree.insert("", "end", values=row)

    def add_dungeon(self):
        """新增副本"""
        self.clear_preset_form()
        self.preset_name_entry.focus()

    def edit_dungeon(self):
        """编辑副本"""
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        item = self.dungeon_tree.item(selected[0])
        values = item['values']
        self.preset_name_var.set(values[1])
        self.preset_drops_text.delete(1.0, tk.END)
        self.preset_drops_text.insert(tk.END, values[2])

    def delete_dungeon(self):
        """删除副本"""
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        item = self.dungeon_tree.item(selected[0])
        dungeon_id = item['values'][0]
        dungeon_name = item['values'][1]
        
        self.cursor.execute("SELECT COUNT(*) FROM records WHERE dungeon_id=?", (dungeon_id,))
        count = self.cursor.fetchone()[0]
        
        if count > 0:
            messagebox.showerror("错误", f"无法删除副本 '{dungeon_name}'，因为存在 {count} 条相关记录")
            return
            
        if messagebox.askyesno("确认", f"确定要删除副本 '{dungeon_name}' 吗？"):
            self.cursor.execute("DELETE FROM dungeons WHERE id=?", (dungeon_id,))
            self.conn.commit()
            self.load_dungeon_presets()
            self.load_dungeon_options()

    def batch_add_items(self):
        """批量添加物品"""
        items_str = self.batch_items_var.get()
        if not items_str:
            return
            
        items = [item.strip() for item in items_str.split(',') if item.strip()]
        current_text = self.preset_drops_text.get(1.0, tk.END).strip()
        
        if current_text:
            new_text = current_text + ", " + ", ".join(items)
        else:
            new_text = ", ".join(items)
            
        self.preset_drops_text.delete(1.0, tk.END)
        self.preset_drops_text.insert(tk.END, new_text)
        self.batch_items_var.set("")

    def save_dungeon(self):
        """保存副本设置"""
        name = self.preset_name_var.get().strip()
        drops = self.preset_drops_text.get(1.0, tk.END).strip()
        
        if not name:
            messagebox.showwarning("提示", "请输入副本名称")
            return
            
        self.cursor.execute("SELECT id FROM dungeons WHERE name=?", (name,))
        existing = self.cursor.fetchone()
        
        if existing:
            dungeon_id = existing[0]
            self.cursor.execute("UPDATE dungeons SET special_drops=? WHERE id=?", (drops, dungeon_id))
        else:
            self.cursor.execute("INSERT INTO dungeons (name, special_drops) VALUES (?, ?)", (name, drops))
            
        self.conn.commit()
        self.load_dungeon_presets()
        self.load_dungeon_options()
        self.clear_preset_form()
        messagebox.showinfo("成功", "副本预设已保存")

    def clear_preset_form(self):
        """清空副本表单"""
        self.preset_name_var.set("")
        self.preset_drops_text.delete(1.0, tk.END)
        self.batch_items_var.set("")

    def add_special_item(self):
        """添加特殊掉落物品"""
        item = self.special_item_var.get().strip()
        price = self.special_price_var.get().strip()
        
        if not item or not price:
            messagebox.showwarning("提示", "请填写物品名称和金额")
            return
            
        try:
            price_int = int(price)
        except ValueError:
            messagebox.showerror("错误", "金额必须是整数")
            return
            
        self.special_tree.insert("", "end", values=(item, price_int))
        self.special_item_var.set("")
        self.special_price_var.set("")
        self.update_special_total()

    def update_special_total(self):
        """更新特殊掉落总金额"""
        total = 0
        for child in self.special_tree.get_children():
            values = self.special_tree.item(child, 'values')
            if values and len(values) >= 2:
                try:
                    total += int(values[1])
                except ValueError:
                    pass
        self.special_total_var.set(str(total))

    def update_total_gold(self, *args):
        """更新团队总工资"""
        def safe_int(value):
            try:
                return int(value) if value != "" else 0
            except ValueError:
                return 0
                
        trash = safe_int(self.trash_gold_var.get())
        iron = safe_int(self.iron_gold_var.get())
        other = safe_int(self.other_gold_var.get())
        special = safe_int(self.special_total_var.get())
        
        total = trash + iron + other + special
        self.total_gold_var.set(str(total))

    def validate_difference(self, *args):
        """验证总工资差异"""
        def safe_int(value):
            try:
                return int(value) if value != "" else 0
            except ValueError:
                return 0
                
        total = safe_int(self.total_gold_var.get())
        trash = safe_int(self.trash_gold_var.get())
        iron = safe_int(self.iron_gold_var.get())
        other = safe_int(self.other_gold_var.get())
        special = safe_int(self.special_total_var.get())
        
        calculated_total = trash + iron + other + special
        difference = total - calculated_total
        
        if difference == 0:
            self.difference_var.set("差额: 0")
            self.difference_label.configure(foreground="green")
        else:
            self.difference_var.set(f"差额: {difference}")
            self.difference_label.configure(foreground="red")

    def add_record(self):
        """添加副本记录"""
        dungeon = self.dungeon_var.get()
        if not dungeon:
            messagebox.showwarning("提示", "请选择副本")
            return
            
        self.cursor.execute("SELECT id FROM dungeons WHERE name=?", (dungeon,))
        result = self.cursor.fetchone()
        if not result:
            messagebox.showerror("错误", f"找不到副本 '{dungeon}'")
            return
            
        dungeon_id = result[0]
        special_auctions = []
        
        # 收集特殊掉落
        for child in self.special_tree.get_children():
            values = self.special_tree.item(child, 'values')
            if values and len(values) >= 2:
                special_auctions.append({
                    "item": values[0],
                    "price": values[1]
                })
        
        # 插入记录
        self.cursor.execute('''
            INSERT INTO records (
                dungeon_id, 
                trash_gold, 
                iron_gold, 
                other_gold, 
                special_auctions, 
                total_gold, 
                black_owner, 
                worker, 
                team_type, 
                lie_down_count, 
                fine_gold, 
                subsidy_gold, 
                personal_gold,
                note
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            dungeon_id,
            int(self.trash_gold_var.get() or 0),
            int(self.iron_gold_var.get() or 0),
            int(self.other_gold_var.get() or 0),
            json.dumps(special_auctions, ensure_ascii=False),
            int(self.total_gold_var.get() or 0),
            self.black_owner_var.get() or None,
            self.worker_var.get() or None,
            self.team_type_var.get(),
            int(self.lie_down_var.get() or 0),
            int(self.fine_gold_var.get() or 0),
            int(self.subsidy_gold_var.get() or 0),
            int(self.personal_gold_var.get() or 0),
            self.note_var.get() or ""
        ))
        
        self.conn.commit()
        self.load_dungeon_records()
        self.update_stats()
        messagebox.showinfo("成功", "记录已添加")
        self.clear_form()

    def load_dungeon_records(self):
        """加载副本记录"""
        self.record_tree.delete(*self.record_tree.get_children())
        self.cursor.execute('''
            SELECT 
                r.id, 
                d.name, 
                strftime('%Y-%m-%d %H:%M', r.time), 
                r.team_type, 
                r.lie_down_count,
                r.total_gold,
                r.personal_gold,
                r.black_owner,
                r.worker,
                r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            ORDER BY r.time DESC
        ''')
        
        for row in self.cursor.fetchall():
            note = row[9] or ""
            if len(note) > 15:
                note = note[:15] + "..."
                
            self.record_tree.insert("", "end", values=(
                row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], note
            ))

    def search_records(self):
        """搜索副本记录"""
        self.record_tree.delete(*self.record_tree.get_children())
        conditions = []
        params = []
        
        # 黑本条件
        owner = self.search_owner_var.get()
        if owner:
            conditions.append("r.black_owner = ?")
            params.append(owner)
        
        # 打工仔条件
        worker = self.search_worker_var.get()
        if worker:
            conditions.append("r.worker = ?")
            params.append(worker)
        
        # 副本条件
        dungeon = self.search_dungeon_var.get()
        if dungeon:
            conditions.append("d.name = ?")
            params.append(dungeon)
        
        # 特殊掉落条件
        item = self.search_item_var.get()
        if item:
            conditions.append("r.special_auctions LIKE ?")
            params.append(f'%{item}%')
        
        # 开始时间条件
        start_date = self.start_date_var.get().strip()
        if start_date:
            if not self.validate_date(start_date):
                messagebox.showwarning("警告", "开始日期格式不正确，请使用YYYY-MM-DD格式")
                return
            conditions.append("r.time >= ?")
            params.append(f"{start_date} 00:00:00")
        
        # 结束时间条件
        end_date = self.end_date_var.get().strip()
        if end_date:
            if not self.validate_date(end_date):
                messagebox.showwarning("警告", "结束日期格式不正确，请使用YYYY-MM-DD格式")
                return
            conditions.append("r.time <= ?")
            params.append(f"{end_date} 23:59:59")
        
        # 团队类型条件
        team_type = self.search_team_type_var.get()
        if team_type:
            conditions.append("r.team_type = ?")
            params.append(team_type)
        
        # 构建SQL查询
        sql = '''
            SELECT 
                r.id, 
                d.name, 
                strftime('%Y-%m-%d %H:%M', r.time), 
                r.team_type, 
                r.lie_down_count,
                r.total_gold,
                r.personal_gold,
                r.black_owner,
                r.worker,
                r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
        '''
        
        if conditions:
            sql += " WHERE " + " AND ".join(conditions)
            
        sql += " ORDER BY r.time DESC"
        
        # 执行查询
        self.cursor.execute(sql, params)
        
        for row in self.cursor.fetchall():
            note = row[9] or ""
            if len(note) > 15:
                note = note[:15] + "..."
                
            self.record_tree.insert("", "end", values=(
                row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], note
            ))

    def edit_record(self):
        """编辑选中的记录"""
        selected = self.record_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一条记录")
            return
            
        item = self.record_tree.item(selected[0])
        record_id = item['values'][0]
        
        self.cursor.execute('''
            SELECT 
                d.name,
                r.trash_gold,
                r.iron_gold,
                r.other_gold,
                r.special_auctions,
                r.total_gold,
                r.black_owner,
                r.worker,
                r.team_type,
                r.lie_down_count,
                r.fine_gold,
                r.subsidy_gold,
                r.personal_gold,
                r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            WHERE r.id = ?
        ''', (record_id,))
        
        record = self.cursor.fetchone()
        if not record:
            messagebox.showerror("错误", "找不到记录")
            return
            
        # 填充表单
        self.dungeon_var.set(record[0])
        self.trash_gold_var.set(str(record[1]) if record[1] != 0 else "")
        self.iron_gold_var.set(str(record[2]) if record[2] != 0 else "")
        self.other_gold_var.set(str(record[3]) if record[3] != 0 else "")
        self.total_gold_var.set(str(record[5]))
        self.black_owner_var.set(record[6] or "")
        self.worker_var.set(record[7] or "")
        self.team_type_var.set(record[8])
        self.lie_down_var.set(str(record[9]) if record[9] != 0 else "")
        self.fine_gold_var.set(str(record[10]) if record[10] != 0 else "")
        self.subsidy_gold_var.set(str(record[11]) if record[11] != 0 else "")
        self.personal_gold_var.set(str(record[12]))
        self.note_var.set(record[13] or "")
        
        # 填充特殊掉落
        self.special_tree.delete(*self.special_tree.get_children())
        special_auctions = json.loads(record[4]) if record[4] else []
        for item in special_auctions:
            self.special_tree.insert("", "end", values=(item['item'], item['price']))
        
        self.update_special_total()
        self.current_edit_id = record_id
        
        # 更新按钮状态
        self.add_btn.configure(state=tk.DISABLED)
        self.edit_btn.configure(state=tk.DISABLED)
        self.update_btn.configure(state=tk.NORMAL)

    def update_record(self):
        """更新记录"""
        if not hasattr(self, 'current_edit_id'):
            return
            
        dungeon = self.dungeon_var.get()
        if not dungeon:
            messagebox.showwarning("提示", "请选择副本")
            return
            
        self.cursor.execute("SELECT id FROM dungeons WHERE name=?", (dungeon,))
        result = self.cursor.fetchone()
        if not result:
            messagebox.showerror("错误", f"找不到副本 '{dungeon}'")
            return
            
        dungeon_id = result[0]
        special_auctions = []
        
        # 收集特殊掉落
        for child in self.special_tree.get_children():
            values = self.special_tree.item(child, 'values')
            if values and len(values) >= 2:
                special_auctions.append({
                    "item": values[0],
                    "price": values[1]
                })
        
        # 更新记录
        self.cursor.execute('''
            UPDATE records SET
                dungeon_id = ?, 
                trash_gold = ?, 
                iron_gold = ?, 
                other_gold = ?, 
                special_auctions = ?, 
                total_gold = ?, 
                black_owner = ?, 
                worker = ?, 
                team_type = ?, 
                lie_down_count = ?, 
                fine_gold = ?, 
                subsidy_gold = ?, 
                personal_gold = ?,
                note = ?
            WHERE id = ?
        ''', (
            dungeon_id,
            int(self.trash_gold_var.get() or 0),
            int(self.iron_gold_var.get() or 0),
            int(self.other_gold_var.get() or 0),
            json.dumps(special_auctions, ensure_ascii=False),
            int(self.total_gold_var.get() or 0),
            self.black_owner_var.get() or None,
            self.worker_var.get() or None,
            self.team_type_var.get(),
            int(self.lie_down_var.get() or 0),
            int(self.fine_gold_var.get() or 0),
            int(self.subsidy_gold_var.get() or 0),
            int(self.personal_gold_var.get() or 0),
            self.note_var.get() or "",
            self.current_edit_id
        ))
        
        self.conn.commit()
        self.load_dungeon_records()
        self.update_stats()
        messagebox.showinfo("成功", "记录已更新")
        self.clear_form()

    def clear_form(self):
        """清空表单"""
        self.dungeon_var.set("")
        self.trash_gold_var.set("")
        self.iron_gold_var.set("")
        self.other_gold_var.set("")
        self.fine_gold_var.set("")
        self.subsidy_gold_var.set("")
        self.lie_down_var.set("")
        self.team_type_var.set("十人本")
        self.total_gold_var.set("0")
        self.personal_gold_var.set("0")
        self.black_owner_var.set("")
        self.worker_var.set("")
        self.note_var.set("")
        self.special_tree.delete(*self.special_tree.get_children())
        self.special_total_var.set("0")
        self.special_item_var.set("")
        self.special_price_var.set("")
        
        # 重置按钮状态
        self.add_btn.configure(state=tk.NORMAL)
        self.edit_btn.configure(state=tk.NORMAL)
        self.update_btn.configure(state=tk.DISABLED)
        
        if hasattr(self, 'current_edit_id'):
            del self.current_edit_id

    def update_stats(self):
        """更新统计数据"""
        # 总记录数
        self.cursor.execute("SELECT COUNT(*) FROM records")
        total_records = self.cursor.fetchone()[0]
        self.total_records_var.set(str(total_records))
        
        # 团队总工资
        self.cursor.execute("SELECT SUM(total_gold) FROM records")
        team_total_gold = self.cursor.fetchone()[0] or 0
        self.team_total_gold_var.set(f"{team_total_gold:,}")
        
        # 团队最高工资
        self.cursor.execute("SELECT MAX(total_gold) FROM records")
        team_max_gold = self.cursor.fetchone()[0] or 0
        self.team_max_gold_var.set(f"{team_max_gold:,}")
        
        # 个人总工资
        self.cursor.execute("SELECT SUM(personal_gold) FROM records")
        personal_total_gold = self.cursor.fetchone()[0] or 0
        self.personal_total_gold_var.set(f"{personal_total_gold:,}")
        
        # 个人最高工资
        self.cursor.execute("SELECT MAX(personal_gold) FROM records")
        personal_max_gold = self.cursor.fetchone()[0] or 0
        self.personal_max_gold_var.set(f"{personal_max_gold:,}")
        
        # 更新其他统计
        self.update_worker_stats()
        self.update_chart()

    def update_worker_stats(self):
        """更新打工仔统计"""
        self.worker_stats_tree.delete(*self.worker_stats_tree.get_children())
        
        self.cursor.execute('''
            SELECT 
                worker,
                COUNT(*) AS count,
                SUM(personal_gold) AS total_personal,
                AVG(personal_gold) AS avg_personal
            FROM records
            WHERE worker IS NOT NULL AND worker != ''
            GROUP BY worker
            ORDER BY total_personal DESC
        ''')
        
        for row in self.cursor.fetchall():
            self.worker_stats_tree.insert("", "end", values=(
                row[0],
                row[1],
                f"{row[2]:,}" if row[2] else "0",
                f"{row[3]:,.0f}" if row[3] else "0"
            ))

    def update_chart(self):
        """更新图表"""
        if not MATPLOTLIB_AVAILABLE:
            return
            
        self.ax.clear()
        end_date = datetime.datetime.now()
        start_date = end_date - datetime.timedelta(weeks=4)
        
        # 查询每周数据
        self.cursor.execute('''
            SELECT 
                strftime('%Y-%W', time) AS week,
                SUM(total_gold) AS total_weekly_gold,
                SUM(personal_gold) AS total_weekly_personal
            FROM records
            WHERE time >= ?
            GROUP BY week
            ORDER BY week
        ''', (start_date.strftime("%Y-%m-%d"),))
        
        results = self.cursor.fetchall()
        
        if not results:
            self.ax.text(0.5, 0.5, "暂无数据", ha='center', va='center', fontsize=10)
            self.canvas.draw()
            return
            
        # 准备图表数据
        weeks = []
        team_gold = []
        personal_gold = []
        
        for row in results:
            year, week = row[0].split('-')
            week_start = datetime.datetime.strptime(f"{year}-{week}-1", "%Y-%W-%w")
            weeks.append(week_start)
            team_gold.append(row[1] or 0)
            personal_gold.append(row[2] or 0)
        
        # 绘制图表
        self.ax.plot(weeks, team_gold, 'o-', label='团队总工资')
        self.ax.plot(weeks, personal_gold, 's-', label='个人总工资')
        
        self.ax.set_title('每周工资统计', fontsize=12)
        self.ax.set_xlabel('日期', fontsize=10)
        self.ax.set_ylabel('金额', fontsize=10)
        self.ax.legend(fontsize=10)
        
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
        self.ax.xaxis.set_major_locator(mdates.WeekdayLocator())
        self.fig.autofmt_xdate()
        self.ax.grid(True, linestyle='--', alpha=0.7)
        
        self.canvas.draw()

    def import_data(self):
        """导入数据"""
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[("JSON文件", "*.json")]
        )
        if not file_path:
            return
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 导入副本数据
            if 'dungeons' in data:
                for dungeon in data['dungeons']:
                    self.cursor.execute('''
                        INSERT OR IGNORE INTO dungeons (name, special_drops)
                        VALUES (?, ?)
                    ''', (dungeon['name'], dungeon['special_drops']))
            
            # 导入记录数据
            if 'records' in data:
                for record in data['records']:
                    self.cursor.execute("SELECT id FROM dungeons WHERE name=?", (record['dungeon'],))
                    result = self.cursor.fetchone()
                    if result:
                        dungeon_id = result[0]
                        self.cursor.execute('''
                            INSERT INTO records (
                                dungeon_id, 
                                trash_gold, 
                                iron_gold, 
                                other_gold, 
                                special_auctions, 
                                total_gold, 
                                black_owner, 
                                worker, 
                                time, 
                                team_type, 
                                lie_down_count, 
                                fine_gold, 
                                subsidy_gold, 
                                personal_gold,
                                note
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            dungeon_id,
                            record['trash_gold'],
                            record['iron_gold'],
                            record['other_gold'],
                            json.dumps(record['special_auctions'], ensure_ascii=False),
                            record['total_gold'],
                            record['black_owner'],
                            record['worker'],
                            record['time'],
                            record['team_type'],
                            record['lie_down_count'],
                            record['fine_gold'],
                            record['subsidy_gold'],
                            record['personal_gold'],
                            record.get('note', '')
                        ))
            
            self.conn.commit()
            self.load_dungeon_records()
            self.update_stats()
            self.load_dungeon_options()
            self.load_black_owner_options()
            messagebox.showinfo("成功", "数据导入完成")
        except Exception as e:
            messagebox.showerror("错误", f"导入数据失败: {str(e)}")

    def export_data(self):
        """导出数据"""
        file_path = filedialog.asksaveasfilename(
            title="保存数据文件",
            filetypes=[("JSON文件", "*.json")],
            defaultextension=".json"
        )
        if not file_path:
            return
            
        data = {
            "dungeons": [],
            "records": []
        }
        
        # 导出副本数据
        self.cursor.execute("SELECT name, special_drops FROM dungeons")
        for row in self.cursor.fetchall():
            data['dungeons'].append({
                "name": row[0],
                "special_drops": row[1]
            })
        
        # 导出记录数据
        self.cursor.execute('''
            SELECT 
                d.name,
                r.trash_gold,
                r.iron_gold,
                r.other_gold,
                r.special_auctions,
                r.total_gold,
                r.black_owner,
                r.worker,
                r.time,
                r.team_type,
                r.lie_down_count,
                r.fine_gold,
                r.subsidy_gold,
                r.personal_gold,
                r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
        ''')
        
        for row in self.cursor.fetchall():
            data['records'].append({
                "dungeon": row[0],
                "trash_gold": row[1],
                "iron_gold": row[2],
                "other_gold": row[3],
                "special_auctions": json.loads(row[4]) if row[4] else [],
                "total_gold": row[5],
                "black_owner": row[6],
                "worker": row[7],
                "time": row[8],
                "team_type": row[9],
                "lie_down_count": row[10],
                "fine_gold": row[11],
                "subsidy_gold": row[12],
                "personal_gold": row[13],
                "note": row[14]
            })
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "数据导出完成")
        except Exception as e:
            messagebox.showerror("错误", f"导出数据失败: {str(e)}")

if __name__ == "__main__":
    # 检查依赖
    missing_libs = check_dependencies()
    if missing_libs:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少依赖", f"请先安装以下库: {', '.join(missing_libs)}\n运行命令: pip install {' '.join(missing_libs)}")
        root.destroy()
    else:
        root = tk.Tk()
        app = JX3DungeonTracker(root)
        root.mainloop()