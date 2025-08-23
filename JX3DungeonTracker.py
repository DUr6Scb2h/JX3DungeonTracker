import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import datetime
import json
import os
import sys
import tkinter.font as tkFont
import platform
import shutil
import locale

try:
    import matplotlib
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import numpy as np
    from matplotlib.dates import WeekdayLocator, DateFormatter
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    plt = None
    FigureCanvasTkAgg = None
    np = None
    mdates = None
    MATPLOTLIB_AVAILABLE = False

def resource_path(relative_path):
    """获取资源路径，用于PyInstaller打包后的资源访问"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_app_data_path():
    """获取应用程序数据目录，确保可写入"""
    if platform.system() == "Windows":
        app_data = os.getenv('APPDATA')
        app_dir = os.path.join(app_data, "JX3DungeonTracker")
    elif platform.system() == "Darwin":
        app_dir = os.path.expanduser("~/Library/Application Support/JX3DungeonTracker")
    else:
        app_dir = os.path.expanduser("~/.jx3dungeontracker")
    
    os.makedirs(app_dir, exist_ok=True)
    return app_dir

def check_dependencies():
    missing = []
    try:
        import matplotlib
    except ImportError:
        missing.append("matplotlib")
    try:
        import numpy
    except ImportError:
        missing.append("numpy")
    return missing

def get_current_time():
    """获取当前时间，确保在打包环境中也能正确获取"""
    try:
        now = datetime.datetime.now()
        return now.strftime("%Y-%m-%d %H:%M:%S")
    except:
        now = datetime.datetime.utcnow()
        return now.strftime("%Y-%m-%d %H:%M:%S")

class DatabaseManager:
    def __init__(self, db_path):
        db_dir = os.path.dirname(db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
            
        self.conn = sqlite3.connect(db_path)
        self.cursor = self.conn.cursor()
        self.initialize_tables()
        self.load_preset_dungeons()

    def initialize_tables(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS dungeons (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                special_drops TEXT
            )
        ''')
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
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS column_widths (
                tree_name TEXT PRIMARY KEY,
                widths TEXT
            )
        ''')
        self.conn.commit()

    def load_preset_dungeons(self):
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
        self.cursor.executemany('''
            INSERT OR IGNORE INTO dungeons (name, special_drops) 
            VALUES (?, ?)
        ''', dungeons)
        self.conn.commit()

    def execute_query(self, query, params=()):
        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def execute_update(self, query, params=()):
        self.cursor.execute(query, params)
        self.conn.commit()

    def close(self):
        self.conn.commit()
        self.conn.close()

class SpecialItemsTree:
    def __init__(self, parent):
        self.tree = ttk.Treeview(parent, columns=("item", "price"), show="headings", height=3, selectmode="browse")
        self.tree.heading("item", text="物品", anchor="center")
        self.tree.heading("price", text="金额", anchor="center")
        self.tree.column("item", width=120, anchor=tk.CENTER)
        self.tree.column("price", width=60, anchor=tk.CENTER)
        
        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

    def clear(self):
        self.tree.delete(*self.tree.get_children())

    def add_item(self, item, price):
        self.tree.insert("", "end", values=(item, price))

    def get_items(self):
        items = []
        for child in self.tree.get_children():
            values = self.tree.item(child, 'values')
            items.append({"item": values[0], "price": values[1]})
        return items

    def calculate_total(self):
        total = 0
        for child in self.tree.get_children():
            values = self.tree.item(child, 'values')
            try:
                total += int(values[1])
            except ValueError:
                pass
        return total

class GoldCalculator:
    @staticmethod
    def safe_int(value):
        try:
            return int(value) if value != "" else 0
        except ValueError:
            return 0

    @classmethod
    def calculate_total(cls, trash, iron, other, special):
        return cls.safe_int(trash) + cls.safe_int(iron) + cls.safe_int(other) + cls.safe_int(special)

    @classmethod
    def calculate_difference(cls, total, trash, iron, other, special):
        calculated = cls.safe_int(trash) + cls.safe_int(iron) + cls.safe_int(other) + cls.safe_int(special)
        return cls.safe_int(total) - calculated

class JX3DungeonTracker:
    def auto_resize_column(self, tree, column_id):
        font = tkFont.Font(family="PingFang SC", size=10)
        heading_text = tree.heading(column_id)["text"]
        max_width = font.measure(heading_text) + 20
        
        for item in tree.get_children():
            cell_value = tree.set(item, column_id)
            if cell_value:
                cell_width = font.measure(cell_value) + 20
                if cell_width > max_width:
                    max_width = cell_width
        
        tree.column(column_id, width=max_width)

    def setup_column_resizing(self, tree):
        columns = tree["columns"]
        for col in columns:
            tree.heading(col, command=lambda c=col: self.auto_resize_column(tree, c))

    def __init__(self, root):
        self.root = root
        self.root.title("JX3DungeonTracker - 剑网3副本记录工具")
        
        try:
            locale.setlocale(locale.LC_TIME, '')
        except:
            pass 
        
        app_data_dir = get_app_data_path()
        db_path = os.path.join(app_data_dir, 'jx3_dungeon.db')
        
        if not os.path.exists(db_path):
            try:
                default_db_path = resource_path('jx3_dungeon.db')
                if os.path.exists(default_db_path):
                    shutil.copy2(default_db_path, db_path)
            except Exception as e:
                print(f"无法复制默认数据库: {e}")
        
        self.db = DatabaseManager(db_path)
        self.setup_ui()
        self.setup_events()
        self.load_data()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        self.setup_window()
        self.setup_variables()
        self.setup_styles()
        self.create_main_ui()

    def setup_window(self):
        width, height = 1600, 900
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.configure(bg="#f5f5f7")
        self.root.minsize(1024, 600)

    def setup_variables(self):
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
        self.time_var = tk.StringVar(value=get_current_time())
        self.difference_var = tk.StringVar(value="差额: 0")
        self.dungeon_var = tk.StringVar()
        self.special_item_var = tk.StringVar()
        self.special_price_var = tk.StringVar()
        self.black_owner_var = tk.StringVar()
        self.worker_var = tk.StringVar()
        self.search_dungeon_var = tk.StringVar()
        self.search_item_var = tk.StringVar()
        self.search_owner_var = tk.StringVar()
        self.search_worker_var = tk.StringVar()
        self.search_team_type_var = tk.StringVar()
        self.preset_drops_var = tk.StringVar()
        self.preset_name_var = tk.StringVar()

    def setup_styles(self):
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

    def create_main_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 8))
        title_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, weight=0)
        
        ttk.Label(title_frame, text="JX3DungeonTracker - 剑网3副本记录工具", 
                 font=("PingFang SC", 16, "bold"), anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=10)
        
        ttk.Label(title_frame, textvariable=self.time_var, 
                 font=("PingFang SC", 12), anchor="e"
        ).grid(row=0, column=1, sticky="e")
        
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        record_frame = ttk.Frame(notebook)
        stats_frame = ttk.Frame(notebook)
        preset_frame = ttk.Frame(notebook)
        
        notebook.add(record_frame, text="副本记录")
        notebook.add(stats_frame, text="数据总览")
        notebook.add(preset_frame, text="副本预设")
        
        self.create_record_tab(record_frame)
        self.create_stats_tab(stats_frame)
        self.create_preset_tab(preset_frame)

    def create_record_tab(self, parent):
        pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        form_frame = ttk.LabelFrame(pane, text="副本记录管理", padding=8, width=350)
        self.build_record_form(form_frame)
        
        list_frame = ttk.LabelFrame(pane, text="副本记录列表", padding=8)
        self.build_record_list(list_frame)
        
        pane.add(form_frame, weight=1)
        pane.add(list_frame, weight=2)

    def build_record_form(self, parent):
        dungeon_row = ttk.Frame(parent)
        dungeon_row.pack(fill=tk.X, pady=3)
        ttk.Label(dungeon_row, text="副本名称:").pack(side=tk.LEFT, padx=(0, 5))
        self.dungeon_combo = ttk.Combobox(dungeon_row, textvariable=self.dungeon_var, width=25)
        self.dungeon_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        special_frame = ttk.LabelFrame(parent, text="特殊掉落", padding=6)
        special_frame.pack(fill=tk.BOTH, pady=(0, 5), expand=True)
        
        tree_frame = ttk.Frame(special_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=3)
        self.special_tree = SpecialItemsTree(tree_frame)
        
        add_special_frame = ttk.Frame(special_frame)
        add_special_frame.pack(fill=tk.X, pady=(8, 0))
        
        ttk.Label(add_special_frame, text="物品:").grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.special_item_combo = ttk.Combobox(add_special_frame, textvariable=self.special_item_var, width=15)
        self.special_item_combo.grid(row=0, column=1, padx=(0, 5), sticky="ew")
        
        ttk.Label(add_special_frame, text="金额:").grid(row=0, column=2, padx=(5, 5), sticky="w")
        self.special_price_entry = ttk.Entry(add_special_frame, textvariable=self.special_price_var, width=10)
        self.special_price_entry.grid(row=0, column=3, padx=(0, 5), sticky="w")
        
        ttk.Button(add_special_frame, text="添加", width=6, command=self.add_special_item
        ).grid(row=0, column=4, padx=(5, 0))
        add_special_frame.columnconfigure(1, weight=1)
        
        team_frame = ttk.LabelFrame(parent, text="团队项目", padding=6)
        team_frame.pack(fill=tk.X, pady=(0, 5))
        self.build_team_fields(team_frame)
        
        personal_frame = ttk.LabelFrame(parent, text="个人项目", padding=6)
        personal_frame.pack(fill=tk.X, pady=(0, 5))
        self.build_personal_fields(personal_frame)
        
        info_frame = ttk.LabelFrame(parent, text="团队信息", padding=6)
        info_frame.pack(fill=tk.X, pady=(0, 5))
        self.build_info_fields(info_frame)
        
        gold_frame = ttk.LabelFrame(parent, text="工资信息", padding=6)
        gold_frame.pack(fill=tk.X, pady=(0, 5))
        self.build_gold_fields(gold_frame)
        
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=5)
        self.build_form_buttons(btn_frame)

    def build_team_fields(self, parent):
        ttk.Label(parent, text="散件金额:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.trash_gold_entry = ttk.Entry(parent, textvariable=self.trash_gold_var, width=10)
        self.trash_gold_entry.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(parent, text="小铁金额:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.iron_gold_entry = ttk.Entry(parent, textvariable=self.iron_gold_var, width=10)
        self.iron_gold_entry.grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(parent, text="特殊金额:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        ttk.Label(parent, textvariable=self.special_total_var, width=10
        ).grid(row=1, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(parent, text="其他金额:").grid(row=1, column=2, padx=3, pady=2, sticky=tk.W)
        self.other_gold_entry = ttk.Entry(parent, textvariable=self.other_gold_var, width=10)
        self.other_gold_entry.grid(row=1, column=3, padx=3, pady=2, sticky=tk.W)

    def build_personal_fields(self, parent):
        ttk.Label(parent, text="补贴金额:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.subsidy_gold_entry = ttk.Entry(parent, textvariable=self.subsidy_gold_var, width=10)
        self.subsidy_gold_entry.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(parent, text="罚款金额:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.fine_gold_entry = ttk.Entry(parent, textvariable=self.fine_gold_var, width=10)
        self.fine_gold_entry.grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)

    def build_info_fields(self, parent):
        ttk.Label(parent, text="团队类型:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.team_type_combo = ttk.Combobox(parent, textvariable=self.team_type_var, 
                                           values=["十人本", "二十五人本"], width=10, state="readonly")
        self.team_type_combo.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        self.team_type_combo.current(0)
        
        ttk.Label(parent, text="躺拍人数:").grid(row=0, column=2, padx=3, pady=2, sticky=tk.W)
        self.lie_down_entry = ttk.Entry(parent, textvariable=self.lie_down_var, width=6)
        self.lie_down_entry.grid(row=0, column=3, padx=3, pady=2, sticky=tk.W)
        
        ttk.Label(parent, text="黑本:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.black_owner_combo = ttk.Combobox(parent, textvariable=self.black_owner_var, width=20)
        self.black_owner_combo.grid(row=1, column=1, padx=3, pady=2, sticky=tk.W, columnspan=3)
        
        ttk.Label(parent, text="打工仔:").grid(row=2, column=0, padx=3, pady=2, sticky=tk.W)
        self.worker_combo = ttk.Combobox(parent, textvariable=self.worker_var, width=20)
        self.worker_combo.grid(row=2, column=1, padx=3, pady=2, sticky=tk.W, columnspan=3)
        
        ttk.Label(parent, text="备注:").grid(row=3, column=0, padx=3, pady=2, sticky=tk.W)
        self.note_entry = ttk.Entry(parent, textvariable=self.note_var, width=20)
        self.note_entry.grid(row=3, column=1, padx=3, pady=2, sticky=tk.W, columnspan=3)

    def build_gold_fields(self, parent):
        ttk.Label(parent, text="团队总工资:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.total_gold_entry = ttk.Entry(parent, textvariable=self.total_gold_var, width=10)
        self.total_gold_entry.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)
        
        self.difference_label = ttk.Label(parent, textvariable=self.difference_var, 
                                         font=("PingFang SC", 9), foreground="#e74c3c")
        self.difference_label.grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(parent, text="个人工资:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.personal_gold_entry = ttk.Entry(parent, textvariable=self.personal_gold_var, width=10)
        self.personal_gold_entry.grid(row=1, column=1, padx=3, pady=2, sticky=tk.W)

    def build_form_buttons(self, parent):
        self.add_btn = ttk.Button(parent, text="保存记录", command=self.add_record, width=10)
        self.edit_btn = ttk.Button(parent, text="编辑记录", command=self.edit_record, width=10)
        self.update_btn = ttk.Button(parent, text="更新记录", command=self.update_record, 
                                   state=tk.DISABLED, width=10)
        clear_btn = ttk.Button(parent, text="清空表单", command=self.clear_form, width=10)
        
        self.add_btn.grid(row=0, column=0, padx=2, sticky="ew")
        self.edit_btn.grid(row=0, column=1, padx=2, sticky="ew")
        self.update_btn.grid(row=0, column=2, padx=2, sticky="ew")
        clear_btn.grid(row=0, column=3, padx=2, sticky="ew")
        
        for i in range(4):
            parent.columnconfigure(i, weight=1)

    def build_record_list(self, parent):
        search_frame = ttk.Frame(parent)
        search_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        self.build_search_controls(search_frame)
        
        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        
        columns = ("id", "dungeon", "time", "team_type", "lie_down", "total", "personal", "black_owner", "worker", "note")
        self.record_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="extended", height=10)
        
        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.record_tree.yview)
        hsb = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.record_tree.xview)
        self.record_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.record_tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        
        self.setup_tree_columns()
        self.setup_column_resizing(self.record_tree)
        self.setup_tooltip()
        self.setup_context_menu()

    def build_search_controls(self, parent):
        search_row1 = ttk.Frame(parent)
        search_row1.pack(fill=tk.X, pady=2)
        ttk.Label(search_row1, text="黑本:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_owner_combo = ttk.Combobox(search_row1, textvariable=self.search_owner_var, width=15)
        self.search_owner_combo.pack(side=tk.LEFT, padx=(0, 8))
        
        ttk.Label(search_row1, text="打工仔:").pack(side=tk.LEFT, padx=(5, 5))
        self.search_worker_combo = ttk.Combobox(search_row1, textvariable=self.search_worker_var, width=15)
        self.search_worker_combo.pack(side=tk.LEFT, padx=(0, 8))
        
        search_row2 = ttk.Frame(parent)
        search_row2.pack(fill=tk.X, pady=2)
        ttk.Label(search_row2, text="副本:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_dungeon_combo = ttk.Combobox(search_row2, textvariable=self.search_dungeon_var, width=15)
        self.search_dungeon_combo.pack(side=tk.LEFT, padx=(0, 8))
        
        ttk.Label(search_row2, text="特殊掉落:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_item_combo = ttk.Combobox(search_row2, textvariable=self.search_item_var, width=15)
        self.search_item_combo.pack(side=tk.LEFT, padx=(0, 8))
        
        search_row3 = ttk.Frame(parent)
        search_row3.pack(fill=tk.X, pady=2)
        ttk.Label(search_row3, text="开始时间:").pack(side=tk.LEFT, padx=(0, 3))
        self.start_date_entry = ttk.Entry(search_row3, textvariable=self.start_date_var, width=12)
        self.start_date_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(search_row3, text="结束时间:").pack(side=tk.LEFT, padx=(0, 3))
        self.end_date_entry = ttk.Entry(search_row3, textvariable=self.end_date_var, width=12)
        self.end_date_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(search_row3, text="团队类型:").pack(side=tk.LEFT, padx=(0, 3))
        self.search_team_type_combo = ttk.Combobox(search_row3, textvariable=self.search_team_type_var, 
                                                  values=["", "十人本", "二十五人本"], width=10)
        self.search_team_type_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        search_row4 = ttk.Frame(parent)
        search_row4.pack(fill=tk.X, pady=2)
        
        left_btn_frame = ttk.Frame(search_row4)
        left_btn_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(left_btn_frame, text="查询", command=self.search_records, width=10
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(left_btn_frame, text="重置", command=self.reset_search, width=10
        ).pack(side=tk.LEFT, padx=5)
        
        right_btn_frame = ttk.Frame(search_row4)
        right_btn_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        ttk.Button(right_btn_frame, text="导入数据", command=self.import_data, width=10
        ).pack(side=tk.RIGHT, padx=(0, 5))
        ttk.Button(right_btn_frame, text="导出数据", command=self.export_data, width=10
        ).pack(side=tk.RIGHT, padx=(0, 5))

    def setup_tree_columns(self):
        columns = [
            ("id", "ID", 35),
            ("dungeon", "副本名称", 100),
            ("time", "时间", 100),
            ("team_type", "团队类型", 70),
            ("lie_down", "躺拍人数", 70),
            ("total", "团队总工资", 100),
            ("personal", "个人工资", 100),
            ("black_owner", "黑本", 70),
            ("worker", "打工仔", 70),
            ("note", "备注", 100)
        ]
        
        for col_id, heading, width in columns:
            self.record_tree.heading(col_id, text=heading, anchor="center")
            self.record_tree.column(col_id, width=width, anchor=tk.CENTER, stretch=(col_id == "note"))

    def setup_tooltip(self):
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.withdraw()
        self.tooltip.overrideredirect(True)
        self.tooltip.configure(bg="#ffffe0", relief="solid", borderwidth=1)
        self.tooltip_label = tk.Label(self.tooltip, text="", justify=tk.LEFT, 
                                     bg="#ffffe0", font=("PingFang SC", 9))
        self.tooltip_label.pack(padx=5, pady=5)
        
        self.record_tree.bind("<Motion>", self.on_tree_motion)
        self.record_tree.bind("<Leave>", self.hide_tooltip)

    def setup_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="删除记录", command=self.delete_selected_records)
        self.record_tree.bind("<Button-3>", self.show_record_context_menu)

    def create_stats_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        if not MATPLOTLIB_AVAILABLE:
            self.show_matplotlib_error(main_frame)
            return
        
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        self.build_stats_cards(left_frame)
        
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.build_chart_area(right_frame)
        self.build_worker_stats(right_frame)

    def show_matplotlib_error(self, parent):
        error_frame = ttk.Frame(parent)
        error_frame.pack(fill=tk.BOTH, expand=True, pady=20)
        ttk.Label(error_frame, 
                 text="需要安装matplotlib和numpy库才能显示统计图表\n请运行: pip install matplotlib numpy", 
                 foreground="red", font=("PingFang SC", 10), anchor="center", justify="center"
        ).pack(fill=tk.BOTH, expand=True)

    def build_stats_cards(self, parent):
        self.total_records_var = tk.StringVar(value="0")
        self.team_total_gold_var = tk.StringVar(value="0")
        self.team_max_gold_var = tk.StringVar(value="0")
        self.personal_total_gold_var = tk.StringVar(value="0")
        self.personal_max_gold_var = tk.StringVar(value="0")

        cards = [
            ("总记录数", self.total_records_var),
            ("团队总工资", self.team_total_gold_var),
            ("团队最高工资", self.team_max_gold_var),
            ("个人总工资", self.personal_total_gold_var),
            ("个人最高工资", self.personal_max_gold_var)
        ]
        
        for title, var in cards:
            card = ttk.LabelFrame(parent, text=title, padding=(10, 8))
            card.pack(fill=tk.X, pady=5)
            ttk.Label(card, textvariable=var, font=("PingFang SC", 14, "bold"), anchor="center"
            ).pack(fill=tk.BOTH, expand=True)

    def build_chart_area(self, parent):
        chart_frame = ttk.LabelFrame(parent, text="每周数据统计", padding=(8, 6))
        chart_frame.pack(fill=tk.BOTH, expand=True)
        
        if MATPLOTLIB_AVAILABLE:
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'KaiTi', 'Arial Unicode MS']
            plt.rcParams['axes.unicode_minus'] = False
            
            self.fig, self.ax = plt.subplots(figsize=(6, 2.5), dpi=100)
            self.fig.patch.set_facecolor('#f5f5f7')
            self.ax.set_facecolor('#f5f5f7')
            
            self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def build_worker_stats(self, parent):
        worker_frame = ttk.LabelFrame(parent, text="打工仔统计", padding=(8, 6))
        worker_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        
        tree_frame = ttk.Frame(worker_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.worker_stats_tree = ttk.Treeview(tree_frame, columns=("worker", "count", "total_personal", "avg_personal"), 
                                            show="headings", height=5)
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.worker_stats_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.worker_stats_tree.xview)
        self.worker_stats_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.worker_stats_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        columns = [
            ("worker", "打工仔", 100),
            ("count", "参与次数", 70),
            ("total_personal", "总工资", 100),
            ("avg_personal", "平均工资", 100)
        ]
        
        for col_id, heading, width in columns:
            self.worker_stats_tree.heading(col_id, text=heading, anchor="center")
            self.worker_stats_tree.column(col_id, width=width, anchor=tk.CENTER)

        self.setup_column_resizing(self.worker_stats_tree)

    def create_preset_tab(self, parent):
        pane = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        pane.pack(fill=tk.BOTH, expand=True)
        
        list_frame = ttk.LabelFrame(pane, text="副本列表", padding=8)
        self.build_dungeon_list(list_frame)
        
        form_frame = ttk.LabelFrame(pane, text="副本详情", padding=8)
        self.build_dungeon_form(form_frame)
        
        pane.add(list_frame, weight=2)
        pane.add(form_frame, weight=1)

    def build_dungeon_list(self, parent):
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.dungeon_tree = ttk.Treeview(tree_frame, columns=("id", "name", "drops"), show="headings", height=19)
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.dungeon_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.dungeon_tree.xview)
        self.dungeon_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.dungeon_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        columns = [
            ("id", "ID", 40),
            ("name", "副本名称", 120),
            ("drops", "特殊掉落", 150)
        ]

        for col_id, heading, width in columns:
            self.dungeon_tree.heading(col_id, text=heading, anchor="center")
            self.dungeon_tree.column(col_id, width=width, anchor=tk.CENTER)

        self.setup_column_resizing(self.dungeon_tree)
        
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(0, 0))
        
        buttons = [
            ("新增副本", self.add_dungeon),
            ("编辑副本", self.edit_dungeon),
            ("删除副本", self.delete_dungeon)
        ]
        
        for text, command in buttons:
            ttk.Button(btn_frame, text=text, command=command, width=10
            ).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

    def build_dungeon_form(self, parent):
        parent.grid_columnconfigure(1, weight=1)
        
        name_row = ttk.Frame(parent)
        name_row.grid(row=0, column=0, columnspan=3, sticky="ew", pady=5)
        name_row.columnconfigure(1, weight=1)
        
        ttk.Label(name_row, text="副本名称:").grid(row=0, column=0, padx=4, sticky=tk.W)
        self.preset_name_entry = ttk.Entry(name_row, textvariable=self.preset_name_var, width=15)
        self.preset_name_entry.grid(row=0, column=1, padx=4, sticky=tk.W)
        
        btn_frame = ttk.Frame(name_row)
        btn_frame.grid(row=0, column=2, padx=4, sticky=tk.E)
        
        ttk.Button(btn_frame, text="保存", width=8, command=self.save_dungeon
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空", width=8, command=self.clear_preset_form
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(parent, text="特殊掉落:").grid(row=1, column=0, padx=4, pady=5, sticky=tk.W)
        self.preset_drops_entry = ttk.Entry(parent, textvariable=self.preset_drops_var, width=180)
        self.preset_drops_entry.grid(row=1, column=1, padx=4, pady=5, sticky=tk.W, columnspan=2)
        
        batch_frame = ttk.LabelFrame(parent, text="批量添加特殊掉落", padding=6)
        batch_frame.grid(row=2, column=0, columnspan=3, padx=4, pady=5, sticky=tk.W+tk.E)
        batch_frame.columnconfigure(0, weight=1)
        
        ttk.Label(batch_frame, text="输入多个物品，用逗号或方括号分隔:").pack(side=tk.TOP, anchor=tk.W, pady=(0, 4))
        
        input_frame = ttk.Frame(batch_frame)
        input_frame.pack(fill=tk.X, pady=2)
        
        self.batch_items_var = tk.StringVar()
        batch_entry = ttk.Entry(input_frame, textvariable=self.batch_items_var)
        batch_entry.pack(side=tk.LEFT, padx=(0, 8), fill=tk.X, expand=True)
        
        ttk.Button(input_frame, text="添加", command=self.batch_add_items, width=8
        ).pack(side=tk.LEFT)

    def setup_events(self):
        self.trash_gold_var.trace_add("write", self.update_total_gold)
        self.iron_gold_var.trace_add("write", self.update_total_gold)
        self.other_gold_var.trace_add("write", self.update_total_gold)
        self.special_total_var.trace_add("write", self.update_total_gold)
        self.total_gold_var.trace_add("write", self.validate_difference)
        
        reg = self.root.register(self.validate_numeric_input)
        vcmd = (reg, '%P')
        for entry in [self.trash_gold_entry, self.iron_gold_entry, self.other_gold_entry, 
                     self.fine_gold_entry, self.subsidy_gold_entry, self.lie_down_entry, 
                     self.personal_gold_entry]:
            entry.config(validate="key", validatecommand=vcmd)
        
        self.dungeon_combo.bind("<<ComboboxSelected>>", self.on_dungeon_select)
        self.search_dungeon_combo.bind("<<ComboboxSelected>>", self.on_search_dungeon_select)
        
        self.root.after(300, self.load_column_widths)
        self.root.after(1000, self.update_time)

    def load_data(self):
        self.load_dungeon_options()
        self.load_dungeon_records()
        self.load_dungeon_presets()
        self.update_stats()
        self.load_black_owner_options()

    def load_dungeon_options(self):
        dungeons = [row[0] for row in self.db.execute_query("SELECT name FROM dungeons")]
        self.dungeon_combo['values'] = dungeons
        self.search_dungeon_combo['values'] = dungeons
        self.search_item_combo['values'] = self.get_all_special_items()

    def get_all_special_items(self):
        items = set()
        for row in self.db.execute_query("SELECT special_drops FROM dungeons"):
            if row[0]:
                for item in row[0].split(','):
                    items.add(item.strip())
        return list(items)

    def load_dungeon_records(self):
        self.record_tree.delete(*self.record_tree.get_children())
        records = self.db.execute_query('''
            SELECT r.id, d.name, strftime('%Y-%m-%d %H:%M', r.time), 
                   r.team_type, r.lie_down_count, r.total_gold, 
                   r.personal_gold, r.black_owner, r.worker, r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            ORDER BY r.time DESC
        ''')
        
        for row in records:
            note = row[9] or ""
            if len(note) > 15:
                note = note[:15] + "..."
            self.record_tree.insert("", "end", values=row[:9] + (note,))

    def load_dungeon_presets(self):
        self.dungeon_tree.delete(*self.dungeon_tree.get_children())
        for row in self.db.execute_query("SELECT id, name, special_drops FROM dungeons"):
            self.dungeon_tree.insert("", "end", values=row)

    def load_column_widths(self):
        trees = [
            (self.record_tree, "record_tree"),
            (self.worker_stats_tree, "worker_stats_tree"),
            (self.dungeon_tree, "dungeon_tree")
        ]
        
        for tree, tree_name in trees:
            result = self.db.execute_query("SELECT widths FROM column_widths WHERE tree_name = ?", (tree_name,))
            if result and result[0]:
                try:
                    widths = json.loads(result[0][0])
                    for col, width in widths.items():
                        if col in tree["columns"]:
                            tree.column(col, width=width)
                except json.JSONDecodeError:
                    pass

    def save_column_widths(self):
        trees = [
            (self.record_tree, "record_tree"),
            (self.worker_stats_tree, "worker_stats_tree"),
            (self.dungeon_tree, "dungeon_tree")
        ]
        
        for tree, tree_name in trees:
            if tree and tree.winfo_exists():
                widths = {col: tree.column(col, "width") for col in tree["columns"]}
                self.db.execute_update('''
                    INSERT OR REPLACE INTO column_widths (tree_name, widths)
                    VALUES (?, ?)
                ''', (tree_name, json.dumps(widths)))

    def validate_numeric_input(self, new_value):
        return new_value == "" or new_value.isdigit()

    def on_dungeon_select(self, event):
        selected = self.dungeon_var.get()
        if not selected:
            return
            
        result = self.db.execute_query("SELECT special_drops FROM dungeons WHERE name=?", (selected,))
        if result and result[0][0]:
            self.special_item_combo['values'] = [item.strip() for item in result[0][0].split(',')]
        else:
            self.special_item_combo['values'] = []

    def on_search_dungeon_select(self, event=None):
        selected = self.search_dungeon_var.get()
        if selected:
            result = self.db.execute_query("SELECT special_drops FROM dungeons WHERE name=?", (selected,))
            if result and result[0][0]:
                drops = [item.strip() for item in result[0][0].split(',')]
            else:
                drops = []
            self.search_item_combo['values'] = drops
        else:
            self.search_item_combo['values'] = self.get_all_special_items()

    def update_time(self):
        self.time_var.set(get_current_time())
        self.root.after(1000, self.update_time)

    def add_special_item(self):
        item = self.special_item_var.get().strip()
        price = self.special_price_var.get().strip()
        
        if not item or not price:
            messagebox.showwarning("提示", "请填写物品名称和金额")
            return
            
        if not price.isdigit():
            messagebox.showerror("错误", "金额必须是整数")
            return
            
        self.special_tree.add_item(item, price)
        self.special_item_var.set("")
        self.special_price_var.set("")
        self.special_total_var.set(str(self.special_tree.calculate_total()))

    def update_total_gold(self, *args):
        trash = self.trash_gold_var.get()
        iron = self.iron_gold_var.get()
        other = self.other_gold_var.get()
        special = self.special_total_var.get()
        total = GoldCalculator.calculate_total(trash, iron, other, special)
        self.total_gold_var.set(str(total))

    def validate_difference(self, *args):
        trash = self.trash_gold_var.get()
        iron = self.iron_gold_var.get()
        other = self.other_gold_var.get()
        special = self.special_total_var.get()
        total = self.total_gold_var.get()
        
        difference = GoldCalculator.calculate_difference(total, trash, iron, other, special)
        self.difference_var.set(f"差额: {difference}")
        self.difference_label.configure(foreground="green" if difference == 0 else "red")

    def add_record(self):
        dungeon = self.dungeon_var.get()
        if not dungeon:
            messagebox.showwarning("提示", "请选择副本")
            return
            
        result = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (dungeon,))
        if not result:
            messagebox.showerror("错误", f"找不到副本 '{dungeon}'")
            return
            
        dungeon_id = result[0][0]
        special_auctions = self.special_tree.get_items()
        personal_gold = GoldCalculator.safe_int(self.personal_gold_var.get())
        
        current_time = get_current_time()
        
        self.db.execute_update('''
            INSERT INTO records (
                dungeon_id, trash_gold, iron_gold, other_gold, special_auctions, 
                total_gold, black_owner, worker, time, team_type, lie_down_count, 
                fine_gold, subsidy_gold, personal_gold, note
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            dungeon_id,
            GoldCalculator.safe_int(self.trash_gold_var.get()),
            GoldCalculator.safe_int(self.iron_gold_var.get()),
            GoldCalculator.safe_int(self.other_gold_var.get()),
            json.dumps(special_auctions, ensure_ascii=False),
            GoldCalculator.safe_int(self.total_gold_var.get()),
            self.black_owner_var.get() or None,
            self.worker_var.get() or None,
            current_time,
            self.team_type_var.get(),
            GoldCalculator.safe_int(self.lie_down_var.get()),
            GoldCalculator.safe_int(self.fine_gold_var.get()),
            GoldCalculator.safe_int(self.subsidy_gold_var.get()),
            personal_gold,
            self.note_var.get() or ""
        ))
        
        self.load_dungeon_records()
        self.update_stats()
        self.load_black_owner_options()
        messagebox.showinfo("成功", "记录已添加")
        self.clear_form()

    def edit_record(self):
        selected = self.record_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一条记录")
            return
            
        record_id = self.record_tree.item(selected[0], 'values')[0]
        record = self.db.execute_query('''
            SELECT d.name, r.trash_gold, r.iron_gold, r.other_gold, r.special_auctions, 
                   r.total_gold, r.black_owner, r.worker, r.time, r.team_type, r.lie_down_count, 
                   r.fine_gold, r.subsidy_gold, r.personal_gold, r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            WHERE r.id = ?
        ''', (record_id,))
        
        if not record:
            messagebox.showerror("错误", "找不到记录")
            return
            
        r = record[0]
        self.dungeon_var.set(r[0])
        self.trash_gold_var.set(str(r[1]) if r[1] != 0 else "")
        self.iron_gold_var.set(str(r[2]) if r[2] != 0 else "")
        self.other_gold_var.set(str(r[3]) if r[3] != 0 else "")
        self.total_gold_var.set(str(r[5]))
        self.black_owner_var.set(r[6] or "")
        self.worker_var.set(r[7] or "")
        self.team_type_var.set(r[9])
        self.lie_down_var.set(str(r[10]) if r[10] != 0 else "")
        self.fine_gold_var.set(str(r[11]) if r[11] != 0 else "")
        self.subsidy_gold_var.set(str(r[12]) if r[12] != 0 else "")
        self.personal_gold_var.set(str(r[13]))
        self.note_var.set(r[14] or "")
        
        self.special_tree.clear()
        special_auctions = json.loads(r[4]) if r[4] else []
        for item in special_auctions:
            self.special_tree.add_item(item['item'], item['price'])
        self.special_total_var.set(str(self.special_tree.calculate_total()))
        
        self.current_edit_id = record_id
        self.add_btn.configure(state=tk.DISABLED)
        self.edit_btn.configure(state=tk.DISABLED)
        self.update_btn.configure(state=tk.NORMAL)

    def update_record(self):
        if not hasattr(self, 'current_edit_id'):
            return
            
        dungeon = self.dungeon_var.get()
        if not dungeon:
            messagebox.showwarning("提示", "请选择副本")
            return
            
        result = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (dungeon,))
        if not result:
            messagebox.showerror("错误", f"找不到副本 '{dungeon}'")
            return
            
        dungeon_id = result[0][0]
        special_auctions = self.special_tree.get_items()
        
        self.db.execute_update('''
            UPDATE records SET
                dungeon_id = ?, trash_gold = ?, iron_gold = ?, other_gold = ?, 
                special_auctions = ?, total_gold = ?, black_owner = ?, worker = ?, 
                team_type = ?, lie_down_count = ?, fine_gold = ?, subsidy_gold = ?, 
                personal_gold = ?, note = ?
            WHERE id = ?
        ''', (
            dungeon_id,
            GoldCalculator.safe_int(self.trash_gold_var.get()),
            GoldCalculator.safe_int(self.iron_gold_var.get()),
            GoldCalculator.safe_int(self.other_gold_var.get()),
            json.dumps(special_auctions, ensure_ascii=False),
            GoldCalculator.safe_int(self.total_gold_var.get()),
            self.black_owner_var.get() or None,
            self.worker_var.get() or None,
            self.team_type_var.get(),
            GoldCalculator.safe_int(self.lie_down_var.get()),
            GoldCalculator.safe_int(self.fine_gold_var.get()),
            GoldCalculator.safe_int(self.subsidy_gold_var.get()),
            GoldCalculator.safe_int(self.personal_gold_var.get()),
            self.note_var.get() or "",
            self.current_edit_id
        ))
        
        self.load_dungeon_records()
        self.update_stats()
        self.load_black_owner_options()
        messagebox.showinfo("成功", "记录已更新")
        self.clear_form()

    def clear_form(self):
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
        self.special_tree.clear()
        self.special_total_var.set("0")
        self.special_item_var.set("")
        self.special_price_var.set("")
        
        self.add_btn.configure(state=tk.NORMAL)
        self.edit_btn.configure(state=tk.NORMAL)
        self.update_btn.configure(state=tk.DISABLED)
        if hasattr(self, 'current_edit_id'):
            del self.current_edit_id

    def search_records(self):
        self.record_tree.delete(*self.record_tree.get_children())
        conditions, params = [], []
        
        if owner := self.search_owner_var.get():
            conditions.append("r.black_owner = ?")
            params.append(owner)
        if worker := self.search_worker_var.get():
            conditions.append("r.worker = ?")
            params.append(worker)
        if dungeon := self.search_dungeon_var.get():
            conditions.append("d.name = ?")
            params.append(dungeon)
        if item := self.search_item_var.get():
            conditions.append("r.special_auctions LIKE ?")
            params.append(f'%{item}%')
        if start_date := self.start_date_var.get().strip():
            if not self.validate_date(start_date):
                messagebox.showwarning("警告", "开始日期格式不正确，请使用YYYY-MM-DD格式")
                return
            conditions.append("r.time >= ?")
            params.append(f"{start_date} 00:00:00")
        if end_date := self.end_date_var.get().strip():
            if not self.validate_date(end_date):
                messagebox.showwarning("警告", "结束日期格式不正确，请使用YYYY-MM-DD格式")
                return
            conditions.append("r.time <= ?")
            params.append(f"{end_date} 23:59:59")
        if team_type := self.search_team_type_var.get():
            conditions.append("r.team_type = ?")
            params.append(team_type)
        
        sql = '''
            SELECT r.id, d.name, strftime('%Y-%m-%d %H:%M', r.time), 
                   r.team_type, r.lie_down_count, r.total_gold, 
                   r.personal_gold, r.black_owner, r.worker, r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
        '''
        if conditions:
            sql += " WHERE " + " AND ".join(conditions)
        sql += " ORDER BY r.time DESC"
        
        for row in self.db.execute_query(sql, params):
            note = row[9] or ""
            if len(note) > 15:
                note = note[:15] + "..."
            self.record_tree.insert("", "end", values=row[:9] + (note,))

    def validate_date(self, date_str):
        if not date_str:
            return True
        try:
            datetime.datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def reset_search(self):
        self.search_dungeon_var.set("")
        self.search_item_var.set("")
        self.search_owner_var.set("")
        self.search_worker_var.set("")
        self.search_team_type_var.set("")
        self.start_date_var.set("")
        self.end_date_var.set("")
        self.on_search_dungeon_select()
        self.load_dungeon_records()

    def show_record_context_menu(self, event):
        row_id = self.record_tree.identify_row(event.y)
        if not row_id:
            return
        self.record_tree.selection_set(row_id)
        self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_records(self):
        selected = self.record_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择要删除的记录")
            return
            
        if not messagebox.askyesno("确认", f"确定要删除选中的 {len(selected)} 条记录吗？"):
            return
            
        for item in selected:
            record_id = self.record_tree.item(item, 'values')[0]
            self.db.execute_update("DELETE FROM records WHERE id=?", (record_id,))
            
        self.load_dungeon_records()
        self.update_stats()
        self.load_black_owner_options()
        messagebox.showinfo("成功", "记录已删除")

    def on_tree_motion(self, event):
        self.tooltip.withdraw()
        row_id = self.record_tree.identify_row(event.y)
        col_id = self.record_tree.identify_column(event.x)
        
        if not row_id:
            return
            
        item = self.record_tree.item(row_id)
        col_name = self.record_tree.heading(col_id)["text"]
        record_id = item['values'][0]
        
        if col_name == "团队总工资":
            result = self.db.execute_query('''
                SELECT trash_gold, iron_gold, other_gold, special_auctions
                FROM records WHERE id = ?
            ''', (record_id,))
            
            if result:
                trash, iron, other, specials_json = result[0]
                specials = json.loads(specials_json) if specials_json else []
                special_total = sum(int(item['price']) for item in specials if 'price' in item)
                
                tooltip_text = f"散件金额: {trash}\n小铁金额: {iron}\n其他金额: {other}\n特殊拍卖总金额: {special_total}\n\n"
                if specials:
                    tooltip_text += "特殊拍卖明细:\n" + "\n".join(f"  {item['item']}: {item['price']}" for item in specials)
                else:
                    tooltip_text += "特殊拍卖明细: 无"
                self.show_tooltip(event, tooltip_text)
                
        elif col_name == "个人工资":
            result = self.db.execute_query('''
                SELECT fine_gold, subsidy_gold FROM records WHERE id = ?
            ''', (record_id,))
            
            if result:
                fine, subsidy = result[0]
                self.show_tooltip(event, f"罚款金额: {fine}\n补贴金额: {subsidy}")

    def show_tooltip(self, event, text):
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
        self.tooltip.withdraw()

    def update_stats(self):
        total_records = self.db.execute_query("SELECT COUNT(*) FROM records")[0][0] or 0
        self.total_records_var.set(str(total_records))
        
        team_total_gold = self.db.execute_query("SELECT SUM(total_gold) FROM records")[0][0] or 0
        self.team_total_gold_var.set(f"{team_total_gold:,}")
        
        team_max_gold = self.db.execute_query("SELECT MAX(total_gold) FROM records")[0][0] or 0
        self.team_max_gold_var.set(f"{team_max_gold:,}")
        
        personal_total_gold = self.db.execute_query("SELECT SUM(personal_gold) FROM records")[0][0] or 0
        self.personal_total_gold_var.set(f"{personal_total_gold:,}")
        
        personal_max_gold = self.db.execute_query("SELECT MAX(personal_gold) FROM records")[0][0] or 0
        self.personal_max_gold_var.set(f"{personal_max_gold:,}")
        
        self.update_worker_stats()
        self.update_chart()

    def update_worker_stats(self):
        self.worker_stats_tree.delete(*self.worker_stats_tree.get_children())
        workers = self.db.execute_query('''
            SELECT worker, COUNT(*), SUM(personal_gold), AVG(personal_gold)
            FROM records
            WHERE worker IS NOT NULL AND worker != ''
            GROUP BY worker
            ORDER BY SUM(personal_gold) DESC
        ''')
        
        for row in workers:
            worker, count, total, avg = row
            self.worker_stats_tree.insert("", "end", values=(
                worker,
                count,
                f"{total:,}" if total else "0",
                f"{avg:,.0f}" if avg else "0"
            ))

    def update_chart(self):
        if not MATPLOTLIB_AVAILABLE:
            return
            
        self.ax.clear()
        end_date = datetime.datetime.now()
        start_date = end_date - datetime.timedelta(weeks=4)
        
        results = self.db.execute_query('''
            SELECT strftime('%Y-%W', time) AS week,
                   SUM(total_gold) AS total_weekly_gold,
                   SUM(personal_gold) AS total_weekly_personal
            FROM records
            WHERE time >= ?
            GROUP BY week
            ORDER BY week
        ''', (start_date.strftime("%Y-%m-%d"),))
        
        if not results:
            self.ax.text(0.5, 0.5, "暂无数据", ha='center', va='center', fontsize=10)
            self.canvas.draw()
            return
            
        weeks, team_gold, personal_gold = [], [], []
        for row in results:
            year, week = row[0].split('-')
            week_start = datetime.datetime.strptime(f"{year}-{week}-1", "%Y-%W-%w")
            weeks.append(week_start)
            team_gold.append(row[1] or 0)
            personal_gold.append(row[2] or 0)
        
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

    def add_dungeon(self):
        self.clear_preset_form()
        self.preset_name_entry.focus()

    def edit_dungeon(self):
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        values = self.dungeon_tree.item(selected[0], 'values')
        self.preset_name_var.set(values[1])
        self.preset_drops_var.set(values[2])

    def delete_dungeon(self):
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        values = self.dungeon_tree.item(selected[0], 'values')
        dungeon_id, dungeon_name = values[0], values[1]
        
        count = self.db.execute_query("SELECT COUNT(*) FROM records WHERE dungeon_id=?", (dungeon_id,))[0][0]
        if count > 0:
            messagebox.showerror("错误", f"无法删除副本 '{dungeon_name}'，因为存在 {count} 条相关记录")
            return
            
        if messagebox.askyesno("确认", f"确定要删除副本 '{dungeon_name}' 吗？"):
            self.db.execute_update("DELETE FROM dungeons WHERE id=?", (dungeon_id,))
            self.load_dungeon_presets()
            self.load_dungeon_options()

    def batch_add_items(self):
        items_str = self.batch_items_var.get()
        if not items_str:
            return
            
        items = []
        for item in items_str.split(','):
            item = item.strip()
            if item.startswith('[') and item.endswith(']'):
                item = item[1:-1].strip()
            if item:
                items.append(item)
        
        if not items:
            return
            
        current_drops = self.preset_drops_var.get().strip()
        
        if current_drops:
            new_drops = current_drops + ", " + ", ".join(items)
        else:
            new_drops = ", ".join(items)
            
        self.preset_drops_var.set(new_drops)
        self.batch_items_var.set("")

    def save_dungeon(self):
        name = self.preset_name_var.get().strip()
        drops = self.preset_drops_var.get().strip()
        
        if not name:
            messagebox.showwarning("提示", "请输入副本名称")
            return
            
        existing = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (name,))
        if existing:
            dungeon_id = existing[0][0]
            self.db.execute_update("UPDATE dungeons SET special_drops=? WHERE id=?", (drops, dungeon_id))
        else:
            self.db.execute_update("INSERT INTO dungeons (name, special_drops) VALUES (?, ?)", (name, drops))
            
        self.load_dungeon_presets()
        self.load_dungeon_options()
        self.clear_preset_form()
        messagebox.showinfo("成功", "副本预设已保存")

    def clear_preset_form(self):
        self.preset_name_var.set("")
        self.preset_drops_var.set("")
        self.batch_items_var.set("")

    def import_data(self):
        file_path = filedialog.askopenfilename(title="选择数据文件", filetypes=[("JSON文件", "*.json")])
        if not file_path:
            return
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if 'dungeons' in data:
                for dungeon in data['dungeons']:
                    self.db.execute_update('''
                        INSERT OR IGNORE INTO dungeons (name, special_drops) VALUES (?, ?)
                    ''', (dungeon['name'], dungeon['special_drops']))
            
            if 'records' in data:
                for record in data['records']:
                    result = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (record['dungeon'],))
                    if result:
                        dungeon_id = result[0][0]
                        self.db.execute_update('''
                            INSERT INTO records (dungeon_id, trash_gold, iron_gold, other_gold, 
                                special_auctions, total_gold, black_owner, worker, time, 
                                team_type, lie_down_count, fine_gold, subsidy_gold, personal_gold, note)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            dungeon_id, record['trash_gold'], record['iron_gold'], record['other_gold'],
                            json.dumps(record['special_auctions'], ensure_ascii=False),
                            record['total_gold'], record['black_owner'], record['worker'], record['time'],
                            record['team_type'], record['lie_down_count'], record['fine_gold'],
                            record['subsidy_gold'], record['personal_gold'], record.get('note', '')
                        ))
            
            self.load_dungeon_records()
            self.update_stats()
            self.load_dungeon_options()
            self.load_black_owner_options()
            messagebox.showinfo("成功", "数据导入完成")
        except Exception as e:
            messagebox.showerror("错误", f"导入数据失败: {str(e)}")

    def export_data(self):
        file_path = filedialog.asksaveasfilename(
            title="保存数据文件", filetypes=[("JSON文件", "*.json")], defaultextension=".json")
        if not file_path:
            return
            
        data = {"dungeons": [], "records": []}
        
        for row in self.db.execute_query("SELECT name, special_drops FROM dungeons"):
            data['dungeons'].append({"name": row[0], "special_drops": row[1]})
        
        for row in self.db.execute_query('''
            SELECT d.name, r.trash_gold, r.iron_gold, r.other_gold, r.special_auctions, 
                   r.total_gold, r.black_owner, r.worker, r.time, r.team_type, 
                   r.lie_down_count, r.fine_gold, r.subsidy_gold, r.personal_gold, r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
        '''):
            data['records'].append({
                "dungeon": row[0], "trash_gold": row[1], "iron_gold": row[2], "other_gold": row[3],
                "special_auctions": json.loads(row[4]) if row[4] else [],
                "total_gold": row[5], "black_owner": row[6], "worker": row[7], "time": row[8],
                "team_type": row[9], "lie_down_count": row[10], "fine_gold": row[11],
                "subsidy_gold": row[12], "personal_gold": row[13], "note": row[14]
            })
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "数据导出完成")
        except Exception as e:
            messagebox.showerror("错误", f"导出数据失败: {str(e)}")

    def load_black_owner_options(self):
        owners = [row[0] for row in self.db.execute_query("SELECT DISTINCT black_owner FROM records WHERE black_owner IS NOT NULL AND black_owner != ''")]
        workers = [row[0] for row in self.db.execute_query("SELECT DISTINCT worker FROM records WHERE worker IS NOT NULL AND worker != ''")]
        
        self.black_owner_combo['values'] = owners
        self.search_owner_combo['values'] = owners
         
        self.worker_combo['values'] = workers
        self.search_worker_combo['values'] = workers

    def on_close(self):
        self.save_column_widths()
        self.db.close()
        self.root.destroy()

if __name__ == "__main__":
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
