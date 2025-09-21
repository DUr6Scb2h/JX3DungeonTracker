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

# 设置缩放因子
SCALE_FACTOR = 1

# 延迟导入matplotlib，减少启动时间
MATPLOTLIB_AVAILABLE = False
plt = None
FigureCanvasTkAgg = None
np = None
mdates = None

try:
    import matplotlib
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import numpy as np
    from matplotlib.dates import WeekdayLocator, DateFormatter
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    pass

def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_app_data_path():
    """获取应用数据目录"""
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
    """检查依赖库"""
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
    """获取当前时间字符串"""
    try:
        now = datetime.datetime.now()
        return now.strftime("%Y-%m-%d %H:%M:%S")
    except:
        now = datetime.datetime.utcnow()
        return now.strftime("%Y-%m-%d %H:%M:%S")

class DatabaseManager:
    """数据库管理类"""
    
    def __init__(self, db_path):
        # 确保数据库目录存在
        db_dir = os.path.dirname(db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
            
        # 连接数据库
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.cursor = self.conn.cursor()
        
        # 启用外键约束
        self.cursor.execute("PRAGMA foreign_keys = ON")
        
        # 初始化表和升级数据库
        self.initialize_tables()
        self.load_preset_dungeons()
        self.upgrade_database()
        
        print(f"数据库管理器已初始化 - 路径: {db_path}")
        
        # 检查数据库完整性
        self.check_integrity()
    
    def check_integrity(self):
        """检查数据库完整性"""
        tables = self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        print(f"数据库中的表: {[table[0] for table in tables]}")
        
        # 检查各表的记录数量
        for table in ['dungeons', 'records']:
            count = self.cursor.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            print(f"{table}表中的记录数量: {count}")

    def initialize_tables(self):
        """初始化数据库表结构"""
        # 创建副本表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS dungeons (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                special_drops TEXT
            )
        ''')
        
        # 创建记录表
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
                note TEXT DEFAULT '',
                is_new INTEGER DEFAULT 0,
                FOREIGN KEY (dungeon_id) REFERENCES dungeons (id)
            )
        ''')
        
        # 创建列宽存储表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS column_widths (
                tree_name TEXT PRIMARY KEY,
                widths TEXT
            )
        ''')
        
        # 创建窗口状态存储表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS window_state (
                width INTEGER,
                height INTEGER,
                maximized INTEGER DEFAULT 0,
                x INTEGER,
                y INTEGER
            )
        ''')
        
        self.conn.commit()
        print("数据库表结构已初始化")

    def upgrade_database(self):
        """升级数据库结构"""
        try:
            # 检查是否已存在is_new列
            self.cursor.execute("PRAGMA table_info(records)")
            columns = [column[1] for column in self.cursor.fetchall()]
            
            if 'is_new' not in columns:
                self.cursor.execute("ALTER TABLE records ADD COLUMN is_new INTEGER DEFAULT 0")
                
            # 检查window_state表是否存在
            self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='window_state'")
            if not self.cursor.fetchone():
                self.cursor.execute('''
                    CREATE TABLE window_state (
                        width INTEGER,
                        height INTEGER,
                        maximized INTEGER DEFAULT 0,
                        x INTEGER,
                        y INTEGER
                    )
                ''')
                
            self.conn.commit()
        except Exception as e:
            print(f"Database upgrade failed: {e}")

    def load_preset_dungeons(self):
        """加载预设副本"""
        dungeons = [
            ("狼牙堡·狼神殿", "阿豪（宠物）,遗忘的书函（外观）,醉月玄晶（95级）"),
            ("敖龙岛", "赤纹野正宗（腰部挂件）,隐狐匿踪（特殊面部）,木木（宠物）,星云踏月骓（普通坐骑）,归墟玄晶（100级）"),
            ("范阳夜变", "簪花空竹（腰部挂件）,弃身·肆（特殊腰部）,幽明录（宠物）,润州绣舞筵（家具）,聆音（特殊腰部）,夜泊蝶影（披风）,归墟玄晶（100级）"),
            ("达摩洞", "活色生香（腰部挂件）,冰蚕龙渡（腰部挂件）,猿神发带（头饰）,漫漫香罗（奇趣坐骑）,阿修罗像（家具）,天乙玄晶（110级）"),
            ("白帝江关", "鲤跃龙门（背部挂件）,血佑铃（腰部挂件）,御马踏金·头饰（马具）,御马踏金·鞍饰（马具）,御马踏金·足饰（马具）,御马踏金（马具）,飞毛将军（普通坐骑）,阔豪（脚印）,天乙玄晶（110级）"),
            ("雷域大泽", "大眼崽（宠物）,灵虫石像（家具）,脊骨王座（家具）,掠影无迹（背部挂件）,荒原切（腰部挂件）,游空竹翼（背部挂件）,天乙玄晶（110级）"),
            ("河阳之战", "爆炸（头顶表情）,北拒风狼（家具）,百战同心（家具）,云鹤报捷（玩具）,玄域辟甲·头饰（马具）,玄域辟甲·鞍饰（马具）,玄域辟甲·足饰（马具）,玄域辟甲（马具）,扇风耳（宠物）,墨言（特殊背部）,天乙玄晶（110级）"),
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
        print("预设副本已加载")

    def execute_query(self, query, params=()):
        """执行查询语句"""
        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def execute_update(self, query, params=()):
        """执行更新语句"""
        self.cursor.execute(query, params)
        self.conn.commit()

    def close(self):
        """关闭数据库连接"""
        self.conn.commit()
        self.conn.close()

class SpecialItemsTree:
    """特殊物品树形视图"""
    def __init__(self, parent):
        self.tree = ttk.Treeview(parent, columns=("item", "price"), show="headings", 
                                height=int(3*SCALE_FACTOR), selectmode="browse")
        self.tree.heading("item", text="物品", anchor="center")
        self.tree.heading("price", text="金额", anchor="center")
        self.tree.column("item", width=int(120*SCALE_FACTOR), anchor=tk.CENTER)
        self.tree.column("price", width=int(60*SCALE_FACTOR), anchor=tk.CENTER)
        
        # 设置Treeview样式
        style = ttk.Style()
        style.configure("Special.Treeview", font=("PingFang SC", int(9*SCALE_FACTOR)), 
                       rowheight=int(24*SCALE_FACTOR))
        self.tree.configure(style="Special.Treeview")
        
        # 添加滚动条
        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 布局
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        
        # 添加上下文菜单
        self.setup_context_menu()

    def setup_context_menu(self):
        """设置上下文菜单"""
        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="删除选中项", command=self.delete_selected_items)
        
        # 绑定右键事件
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        """显示上下文菜单"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_items(self):
        """删除选中的物品"""
        selected_items = self.tree.selection()
        for item in selected_items:
            self.tree.delete(item)

    def clear(self):
        """清空所有项"""
        self.tree.delete(*self.tree.get_children())

    def add_item(self, item, price):
        """添加物品"""
        self.tree.insert("", "end", values=(item, price))

    def get_items(self):
        """获取所有物品"""
        items = []
        for child in self.tree.get_children():
            values = self.tree.item(child, 'values')
            items.append({"item": values[0], "price": values[1]})
        return items

    def calculate_total(self):
        """计算总金额"""
        total = 0
        for child in self.tree.get_children():
            values = self.tree.item(child, 'values')
            try:
                total += int(values[1])
            except ValueError:
                pass
        return total

class GoldCalculator:
    """金币计算工具类"""
    @staticmethod
    def safe_int(value):
        """安全转换为整数"""
        try:
            return int(value) if value != "" else 0
        except ValueError:
            return 0

    @classmethod
    def calculate_total(cls, trash, iron, other, special):
        """计算总金额"""
        return cls.safe_int(trash) + cls.safe_int(iron) + cls.safe_int(other) + cls.safe_int(special)

    @classmethod
    def calculate_difference(cls, total, trash, iron, other, special):
        """计算差额"""
        calculated = cls.safe_int(trash) + cls.safe_int(iron) + cls.safe_int(other) + cls.safe_int(special)
        return cls.safe_int(total) - calculated

class JX3DungeonTracker:
    """主应用程序类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("JX3DungeonTracker - 剑网3副本记录工具")
        
        # 设置本地化
        try:
            locale.setlocale(locale.LC_TIME, '')
        except:
            pass 
        
        # 初始化数据库
        app_data_dir = get_app_data_path()
        db_path = os.path.join(app_data_dir, 'jx3_dungeon.db')
        
        print(f"应用数据目录: {app_data_dir}")
        print(f"数据库路径: {db_path}")
        print(f"数据库文件是否存在: {os.path.exists(db_path)}")
        
        os.makedirs(app_data_dir, exist_ok=True)
        
        # 数据库初始化
        try:
            self.db = DatabaseManager(db_path)
            print(f"数据库已初始化: {db_path}")
            
            # 测试数据库连接
            test_result = self.db.execute_query("SELECT COUNT(*) FROM records")
            print(f"数据库中的记录数量: {test_result[0][0]}")
            
        except Exception as e:
            messagebox.showerror("数据库错误", f"无法初始化数据库: {str(e)}")
            self.root.destroy()
            return
        
        # 初始化变量
        self.new_record_ids = set()
        self.cached_dungeons = None
        self.cached_owners = None
        self.cached_workers = None
        
        # 设置UI和事件
        self.setup_ui()
        self.setup_events()
        self.load_data()
        
        # 设置关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        """设置用户界面"""
        self.setup_window()
        self.setup_variables()
        self.setup_styles()
        self.create_main_ui()

    def setup_window(self):
        """设置窗口属性"""
        # 从数据库加载保存的窗口状态
        result = self.db.execute_query("SELECT width, height, maximized, x, y FROM window_state")
        if result:
            width, height, maximized, x, y = result[0]
            # 确保窗口大小不会超出屏幕
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            width = min(width, screen_width)
            height = min(height, screen_height)
            
            # 设置窗口位置
            if x is not None and y is not None:
                x = max(0, min(x, screen_width - width))
                y = max(0, min(y, screen_height - height))
                self.root.geometry(f"{width}x{height}+{x}+{y}")
            else:
                self.root.geometry(f"{width}x{height}")
            
            # 恢复最大化状态
            if maximized:
                self.root.state('zoomed')
        else:
            # 默认窗口大小和位置
            width, height = int(1600*SCALE_FACTOR), int(900*SCALE_FACTOR)
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        self.root.configure(bg="#f5f5f7")
        self.root.minsize(int(1024*SCALE_FACTOR), int(600*SCALE_FACTOR))

    def setup_variables(self):
        """设置界面变量"""
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
        """设置界面样式"""
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure(".", background="#f5f5f7", foreground="#333")
        self.style.configure("TFrame", background="#f5f5f7")
        self.style.configure("TLabel", background="#f5f5f7", font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.style.configure("TButton",
                            font=("PingFang SC", int(10*SCALE_FACTOR)),
                            padding=int(6*SCALE_FACTOR),
                            background="#e6e6e6")
        self.style.map("TButton", background=[("active", "#d6d6d6")])
        self.style.configure("Treeview", font=("PingFang SC", int(9*SCALE_FACTOR)), rowheight=int(24*SCALE_FACTOR))
        self.style.configure("Treeview.Heading", font=("PingFang SC", int(10*SCALE_FACTOR)), anchor="center")
        self.style.configure("TNotebook", background="#f5f5f7", borderwidth=0)
        self.style.configure("TNotebook.Tab", font=("PingFang SC", int(10*SCALE_FACTOR)), padding=[int(10*SCALE_FACTOR), int(4*SCALE_FACTOR)])
        self.style.configure("TCombobox", font=("PingFang SC", int(10*SCALE_FACTOR)), padding=int(4*SCALE_FACTOR))
        self.style.configure("TEntry", font=("PingFang SC", int(10*SCALE_FACTOR)), padding=int(4*SCALE_FACTOR))
        self.style.configure("TLabelFrame", font=("PingFang SC", int(10*SCALE_FACTOR)), padding=int(8*SCALE_FACTOR), labelanchor="n")
        
        self.style.configure("NewRecord.Treeview", background="#e6f7ff")
        
        # 设置下拉列表字体
        self.root.option_add("*TCombobox*Listbox*Font", ("PingFang SC", int(10*SCALE_FACTOR)))

    def create_main_ui(self):
        """创建主界面"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=int(10*SCALE_FACTOR), pady=int(10*SCALE_FACTOR))
        
        # 标题栏
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, int(8*SCALE_FACTOR)))
        title_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, weight=0)
        
        ttk.Label(title_frame, text="JX3DungeonTracker - 剑网3副本记录工具", 
                 font=("PingFang SC", int(16*SCALE_FACTOR), "bold"), anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=int(10*SCALE_FACTOR))
        
        ttk.Label(title_frame, textvariable=self.time_var, 
                 font=("PingFang SC", int(12*SCALE_FACTOR)), anchor="e"
        ).grid(row=0, column=1, sticky="e")
        
        # 创建选项卡
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=int(5*SCALE_FACTOR), pady=int(5*SCALE_FACTOR))
        
        record_frame = ttk.Frame(notebook)
        stats_frame = ttk.Frame(notebook)
        preset_frame = ttk.Frame(notebook)
        
        notebook.add(record_frame, text="副本记录")
        notebook.add(stats_frame, text="数据总览")
        notebook.add(preset_frame, text="副本预设")
        
        # 创建各个选项卡内容
        self.create_record_tab(record_frame)
        self.create_stats_tab(stats_frame)
        self.create_preset_tab(preset_frame)

    def create_record_tab(self, parent):
        """创建记录选项卡"""
        pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=int(5*SCALE_FACTOR), pady=int(5*SCALE_FACTOR))
        
        form_frame = ttk.LabelFrame(pane, text="副本记录管理", padding=int(8*SCALE_FACTOR), width=int(350*SCALE_FACTOR))
        self.build_record_form(form_frame)
        
        list_frame = ttk.LabelFrame(pane, text="副本记录列表", padding=int(8*SCALE_FACTOR))
        self.build_record_list(list_frame)
        
        pane.add(form_frame, weight=1)
        pane.add(list_frame, weight=2)

    def build_record_form(self, parent):
        """构建记录表单"""
        # 副本选择
        dungeon_row = ttk.Frame(parent)
        dungeon_row.pack(fill=tk.X, pady=int(3*SCALE_FACTOR))
        ttk.Label(dungeon_row, text="副本名称:").pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        self.dungeon_combo = ttk.Combobox(dungeon_row, textvariable=self.dungeon_var, 
                                         width=int(25*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.dungeon_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 特殊掉落区域
        special_frame = ttk.LabelFrame(parent, text="特殊掉落", padding=int(6*SCALE_FACTOR))
        special_frame.pack(fill=tk.BOTH, pady=(0, int(5*SCALE_FACTOR)), expand=True)
        
        tree_frame = ttk.Frame(special_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=int(3*SCALE_FACTOR))
        self.special_tree = SpecialItemsTree(tree_frame)
        
        # 添加特殊掉落控件
        add_special_frame = ttk.Frame(special_frame)
        add_special_frame.pack(fill=tk.X, pady=(int(8*SCALE_FACTOR), 0))
        
        ttk.Label(add_special_frame, text="物品:").grid(row=0, column=0, padx=(0, int(5*SCALE_FACTOR)), sticky="w")
        self.special_item_combo = ttk.Combobox(add_special_frame, textvariable=self.special_item_var, 
                                              width=int(15*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.special_item_combo.grid(row=0, column=1, padx=(0, int(5*SCALE_FACTOR)), sticky="ew")
        
        ttk.Label(add_special_frame, text="金额:").grid(row=0, column=2, padx=(int(5*SCALE_FACTOR), int(5*SCALE_FACTOR)), sticky="w")
        self.special_price_entry = ttk.Entry(add_special_frame, textvariable=self.special_price_var, 
                                            width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.special_price_entry.grid(row=0, column=3, padx=(0, int(5*SCALE_FACTOR)), sticky="w")
        
        ttk.Button(add_special_frame, text="添加", width=int(6*SCALE_FACTOR), command=self.add_special_item
        ).grid(row=0, column=4, padx=(int(5*SCALE_FACTOR), 0))
        add_special_frame.columnconfigure(1, weight=1)
        
        # 团队项目
        team_frame = ttk.LabelFrame(parent, text="团队项目", padding=int(6*SCALE_FACTOR))
        team_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_team_fields(team_frame)
        
        # 个人项目
        personal_frame = ttk.LabelFrame(parent, text="个人项目", padding=int(6*SCALE_FACTOR))
        personal_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_personal_fields(personal_frame)
        
        # 团队信息
        info_frame = ttk.LabelFrame(parent, text="团队信息", padding=int(6*SCALE_FACTOR))
        info_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_info_fields(info_frame)
        
        # 工资信息
        gold_frame = ttk.LabelFrame(parent, text="工资信息", padding=int(6*SCALE_FACTOR))
        gold_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_gold_fields(gold_frame)
        
        # 按钮区域
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=int(5*SCALE_FACTOR))
        self.build_form_buttons(btn_frame)

    def build_team_fields(self, parent):
        """构建团队项目字段"""
        ttk.Label(parent, text="散件金额:").grid(row=0, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.trash_gold_entry = ttk.Entry(parent, textvariable=self.trash_gold_var, 
                                         width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.trash_gold_entry.grid(row=0, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="小铁金额:").grid(row=0, column=2, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.iron_gold_entry = ttk.Entry(parent, textvariable=self.iron_gold_var, 
                                        width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.iron_gold_entry.grid(row=0, column=3, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="特殊金额:").grid(row=1, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        ttk.Label(parent, textvariable=self.special_total_var, width=int(10*SCALE_FACTOR), 
                 font=("PingFang SC", int(10*SCALE_FACTOR))
        ).grid(row=1, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="其他金额:").grid(row=1, column=2, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.other_gold_entry = ttk.Entry(parent, textvariable=self.other_gold_var, 
                                         width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.other_gold_entry.grid(row=1, column=3, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)

    def build_personal_fields(self, parent):
        """构建个人项目字段"""
        ttk.Label(parent, text="补贴金额:").grid(row=0, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.subsidy_gold_entry = ttk.Entry(parent, textvariable=self.subsidy_gold_var, 
                                           width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.subsidy_gold_entry.grid(row=0, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="罚款金额:").grid(row=0, column=2, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.fine_gold_entry = ttk.Entry(parent, textvariable=self.fine_gold_var, 
                                        width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.fine_gold_entry.grid(row=0, column=3, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)

    def build_info_fields(self, parent):
        """构建信息字段"""
        ttk.Label(parent, text="团队类型:").grid(row=0, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.team_type_combo = ttk.Combobox(parent, textvariable=self.team_type_var, 
                                           values=["十人本", "二十五人本"], width=int(10*SCALE_FACTOR), 
                                           state="readonly", font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.team_type_combo.grid(row=0, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.team_type_combo.current(0)
        
        ttk.Label(parent, text="躺拍人数:").grid(row=0, column=2, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.lie_down_entry = ttk.Entry(parent, textvariable=self.lie_down_var, 
                                       width=int(6*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.lie_down_entry.grid(row=0, column=3, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="黑本:").grid(row=1, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.black_owner_combo = ttk.Combobox(parent, textvariable=self.black_owner_var, 
                                             width=int(20*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.black_owner_combo.grid(row=1, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W, columnspan=3)
        
        ttk.Label(parent, text="打工仔:").grid(row=2, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.worker_combo = ttk.Combobox(parent, textvariable=self.worker_var, 
                                        width=int(20*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.worker_combo.grid(row=2, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W, columnspan=3)
        
        ttk.Label(parent, text="备注:").grid(row=3, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.note_entry = ttk.Entry(parent, textvariable=self.note_var, 
                                   width=int(20*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.note_entry.grid(row=3, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W, columnspan=3)

    def build_gold_fields(self, parent):
        """构建工资字段"""
        ttk.Label(parent, text="团队总工资:").grid(row=0, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.total_gold_entry = ttk.Entry(parent, textvariable=self.total_gold_var, 
                                         width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.total_gold_entry.grid(row=0, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        self.difference_label = ttk.Label(parent, textvariable=self.difference_var, 
                                         font=("PingFang SC", int(9*SCALE_FACTOR)), foreground="#e74c3c")
        self.difference_label.grid(row=0, column=2, padx=int(5*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="个人工资:").grid(row=1, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.personal_gold_entry = ttk.Entry(parent, textvariable=self.personal_gold_var, 
                                            width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.personal_gold_entry.grid(row=1, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)

    def build_form_buttons(self, parent):
        """构建表单按钮"""
        self.add_btn = ttk.Button(parent, text="保存记录", command=self.validate_and_save, width=int(10*SCALE_FACTOR))
        self.edit_btn = ttk.Button(parent, text="编辑记录", command=self.edit_record, width=int(10*SCALE_FACTOR))
        self.update_btn = ttk.Button(parent, text="更新记录", command=self.update_record, 
                                   state=tk.DISABLED, width=int(10*SCALE_FACTOR))
        clear_btn = ttk.Button(parent, text="清空表单", command=self.clear_form, width=int(10*SCALE_FACTOR))
        
        self.add_btn.grid(row=0, column=0, padx=int(2*SCALE_FACTOR), sticky="ew")
        self.edit_btn.grid(row=0, column=1, padx=int(2*SCALE_FACTOR), sticky="ew")
        self.update_btn.grid(row=0, column=2, padx=int(2*SCALE_FACTOR), sticky="ew")
        clear_btn.grid(row=0, column=3, padx=int(2*SCALE_FACTOR), sticky="ew")
        
        for i in range(4):
            parent.columnconfigure(i, weight=1)

    def build_record_list(self, parent):
        """构建记录列表"""
        search_frame = ttk.Frame(parent)
        search_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, int(5*SCALE_FACTOR)))
        self.build_search_controls(search_frame)
        
        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        
        # 创建记录树
        columns = ("row_num", "dungeon", "time", "team_type", "lie_down", "total", "personal", "black_owner", "worker", "note")
        self.record_tree = ttk.Treeview(parent, columns=columns, show="headings", 
                                       selectmode="extended", height=int(10*SCALE_FACTOR))
        
        # 添加滚动条
        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.record_tree.yview)
        hsb = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.record_tree.xview)
        self.record_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.record_tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        
        # 设置列
        self.setup_tree_columns()
        self.setup_column_resizing(self.record_tree)
        self.setup_tooltip()
        self.setup_context_menu()
        
        # 设置新记录高亮样式
        self.record_tree.tag_configure('new_record', background='#e6f7ff')

    def build_search_controls(self, parent):
        """构建搜索控件"""
        search_row1 = ttk.Frame(parent)
        search_row1.pack(fill=tk.X, pady=int(2*SCALE_FACTOR))
        ttk.Label(search_row1, text="黑本:").pack(side=tk.LEFT, padx=(0, int(3*SCALE_FACTOR)))
        self.search_owner_combo = ttk.Combobox(search_row1, textvariable=self.search_owner_var, 
                                              width=int(15*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.search_owner_combo.pack(side=tk.LEFT, padx=(0, int(8*SCALE_FACTOR)))
        
        ttk.Label(search_row1, text="打工仔:").pack(side=tk.LEFT, padx=(int(5*SCALE_FACTOR), int(5*SCALE_FACTOR)))
        self.search_worker_combo = ttk.Combobox(search_row1, textvariable=self.search_worker_var, 
                                               width=int(15*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.search_worker_combo.pack(side=tk.LEFT, padx=(0, int(8*SCALE_FACTOR)))
        
        search_row2 = ttk.Frame(parent)
        search_row2.pack(fill=tk.X, pady=int(2*SCALE_FACTOR))
        ttk.Label(search_row2, text="副本:").pack(side=tk.LEFT, padx=(0, int(3*SCALE_FACTOR)))
        self.search_dungeon_combo = ttk.Combobox(search_row2, textvariable=self.search_dungeon_var, 
                                                width=int(15*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.search_dungeon_combo.pack(side=tk.LEFT, padx=(0, int(8*SCALE_FACTOR)))
        
        ttk.Label(search_row2, text="特殊掉落:").pack(side=tk.LEFT, padx=(0, int(3*SCALE_FACTOR)))
        self.search_item_combo = ttk.Combobox(search_row2, textvariable=self.search_item_var, 
                                             width=int(15*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.search_item_combo.pack(side=tk.LEFT, padx=(0, int(8*SCALE_FACTOR)))
        
        search_row3 = ttk.Frame(parent)
        search_row3.pack(fill=tk.X, pady=int(2*SCALE_FACTOR))
        ttk.Label(search_row3, text="开始时间:").pack(side=tk.LEFT, padx=(0, int(3*SCALE_FACTOR)))
        self.start_date_entry = ttk.Entry(search_row3, textvariable=self.start_date_var, 
                                         width=int(12*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.start_date_entry.pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        
        ttk.Label(search_row3, text="结束时间:").pack(side=tk.LEFT, padx=(0, int(3*SCALE_FACTOR)))
        self.end_date_entry = ttk.Entry(search_row3, textvariable=self.end_date_var, 
                                       width=int(12*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.end_date_entry.pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        
        ttk.Label(search_row3, text="团队类型:").pack(side=tk.LEFT, padx=(0, int(3*SCALE_FACTOR)))
        self.search_team_type_combo = ttk.Combobox(search_row3, textvariable=self.search_team_type_var, 
                                                  values=["", "十人本", "二十五人本"], 
                                                  width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.search_team_type_combo.pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        
        search_row4 = ttk.Frame(parent)
        search_row4.pack(fill=tk.X, pady=int(2*SCALE_FACTOR))
        
        left_btn_frame = ttk.Frame(search_row4)
        left_btn_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(left_btn_frame, text="查询", command=self.search_records, width=int(10*SCALE_FACTOR)
        ).pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        ttk.Button(left_btn_frame, text="重置", command=self.reset_search, width=int(10*SCALE_FACTOR)
        ).pack(side=tk.LEFT, padx=int(5*SCALE_FACTOR))
        ttk.Button(left_btn_frame, text="修复数据库", command=self.repair_database, width=int(10*SCALE_FACTOR)
        ).pack(side=tk.LEFT, padx=int(5*SCALE_FACTOR))
        
        right_btn_frame = ttk.Frame(search_row4)
        right_btn_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        ttk.Button(right_btn_frame, text="导入数据", command=self.import_data, width=int(10*SCALE_FACTOR)
        ).pack(side=tk.RIGHT, padx=(0, int(5*SCALE_FACTOR)))
        ttk.Button(right_btn_frame, text="导出数据", command=self.export_data, width=int(10*SCALE_FACTOR)
        ).pack(side=tk.RIGHT, padx=(0, int(5*SCALE_FACTOR)))

    def setup_tree_columns(self):
        """设置树形视图列"""
        columns = [
            ("row_num", "序号", int(35*SCALE_FACTOR)),
            ("dungeon", "副本名称", int(100*SCALE_FACTOR)),
            ("time", "时间", int(100*SCALE_FACTOR)),
            ("team_type", "团队类型", int(70*SCALE_FACTOR)),
            ("lie_down", "躺拍人数", int(70*SCALE_FACTOR)),
            ("total", "团队总工资", int(100*SCALE_FACTOR)),
            ("personal", "个人工资", int(100*SCALE_FACTOR)),
            ("black_owner", "黑本", int(70*SCALE_FACTOR)),
            ("worker", "打工仔", int(70*SCALE_FACTOR)),
            ("note", "备注", int(100*SCALE_FACTOR))
        ]
        
        for col_id, heading, width in columns:
            self.record_tree.heading(col_id, text=heading, anchor="center")
            self.record_tree.column(col_id, width=width, anchor=tk.CENTER, stretch=(col_id == "note"))

    def setup_tooltip(self):
        """设置工具提示"""
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.withdraw()
        self.tooltip.overrideredirect(True)
        self.tooltip.configure(bg="#ffffe0", relief="solid", borderwidth=1)
        self.tooltip_label = tk.Label(self.tooltip, text="", justify=tk.LEFT, 
                                     bg="#ffffe0", font=("PingFang SC", int(9*SCALE_FACTOR)))
        self.tooltip_label.pack(padx=int(5*SCALE_FACTOR), pady=int(5*SCALE_FACTOR))
        
        self.record_tree.bind("<Motion>", self.on_tree_motion)
        self.record_tree.bind("<Leave>", self.hide_tooltip)

    def setup_context_menu(self):
        """设置上下文菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="删除记录", command=self.delete_selected_records)
        self.record_tree.bind("<Button-3>", self.show_record_context_menu)

    def create_stats_tab(self, parent):
        """创建统计选项卡"""
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=int(10*SCALE_FACTOR), pady=int(10*SCALE_FACTOR))
        
        if not MATPLOTLIB_AVAILABLE:
            self.show_matplotlib_error(main_frame)
            return
        
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, int(10*SCALE_FACTOR)))
        self.build_stats_cards(left_frame)
        
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.build_chart_area(right_frame)
        self.build_worker_stats(right_frame)

    def show_matplotlib_error(self, parent):
        """显示matplotlib错误信息"""
        error_frame = ttk.Frame(parent)
        error_frame.pack(fill=tk.BOTH, expand=True, pady=int(20*SCALE_FACTOR))
        ttk.Label(error_frame, 
                 text="需要安装matplotlib和numpy库才能显示统计图表\n请运行: pip install matplotlib numpy", 
                 foreground="red", font=("PingFang SC", int(10*SCALE_FACTOR)), anchor="center", justify="center"
        ).pack(fill=tk.BOTH, expand=True)

    def build_stats_cards(self, parent):
        """构建统计卡片"""
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
            card = ttk.LabelFrame(parent, text=title, padding=(int(10*SCALE_FACTOR), int(8*SCALE_FACTOR)))
            card.pack(fill=tk.X, pady=int(5*SCALE_FACTOR))
            ttk.Label(card, textvariable=var, font=("PingFang SC", int(14*SCALE_FACTOR), "bold"), anchor="center"
            ).pack(fill=tk.BOTH, expand=True)

    def build_chart_area(self, parent):
        """构建图表区域"""
        chart_frame = ttk.LabelFrame(parent, text="每周数据统计", padding=(int(8*SCALE_FACTOR), int(6*SCALE_FACTOR)))
        chart_frame.pack(fill=tk.BOTH, expand=True)
        
        if MATPLOTLIB_AVAILABLE:
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'KaiTi', 'Arial Unicode MS']
            plt.rcParams['axes.unicode_minus'] = False
        
            self.fig, self.ax = plt.subplots(figsize=(int(6*SCALE_FACTOR), int(2.5*SCALE_FACTOR)), dpi=100)
            self.fig.patch.set_facecolor('#f5f5f7')
            self.ax.set_facecolor('#f5f5f7')
        
            # 设置图表标题和标签的字体大小
            plt.rcParams['axes.titlesize'] = int(12*SCALE_FACTOR)
            plt.rcParams['axes.labelsize'] = int(10*SCALE_FACTOR)
            plt.rcParams['xtick.labelsize'] = int(8*SCALE_FACTOR)
            plt.rcParams['ytick.labelsize'] = int(8*SCALE_FACTOR)
            plt.rcParams['legend.fontsize'] = int(10*SCALE_FACTOR)
        
            self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def build_worker_stats(self, parent):
        """构建打工仔统计"""
        worker_frame = ttk.LabelFrame(parent, text="打工仔统计", padding=(int(8*SCALE_FACTOR), int(6*SCALE_FACTOR)))
        worker_frame.pack(fill=tk.BOTH, expand=True, pady=(int(8*SCALE_FACTOR), 0))
        
        tree_frame = ttk.Frame(worker_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.worker_stats_tree = ttk.Treeview(tree_frame, columns=("worker", "count", "total_personal", "avg_personal"), 
                                            show="headings", height=int(5*SCALE_FACTOR))
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.worker_stats_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.worker_stats_tree.xview)
        self.worker_stats_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.worker_stats_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        columns = [
            ("worker", "打工仔", int(100*SCALE_FACTOR)),
            ("count", "参与次数", int(70*SCALE_FACTOR)),
            ("total_personal", "总工资", int(100*SCALE_FACTOR)),
            ("avg_personal", "平均工资", int(100*SCALE_FACTOR))
        ]
        
        for col_id, heading, width in columns:
            self.worker_stats_tree.heading(col_id, text=heading, anchor="center")
            self.worker_stats_tree.column(col_id, width=width, anchor=tk.CENTER)

        self.setup_column_resizing(self.worker_stats_tree)

    def create_preset_tab(self, parent):
        """创建预设选项卡"""
        pane = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        pane.pack(fill=tk.BOTH, expand=True)
        
        list_frame = ttk.LabelFrame(pane, text="副本列表", padding=int(8*SCALE_FACTOR))
        self.build_dungeon_list(list_frame)
        
        form_frame = ttk.LabelFrame(pane, text="副本详情", padding=int(8*SCALE_FACTOR))
        self.build_dungeon_form(form_frame)
        
        pane.add(list_frame, weight=2)
        pane.add(form_frame, weight=1)

    def build_dungeon_list(self, parent):
        """构建副本列表"""
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, int(8*SCALE_FACTOR)))
        
        self.dungeon_tree = ttk.Treeview(tree_frame, columns=("name", "drops"), show="headings", height=int(22*SCALE_FACTOR))
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.dungeon_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.dungeon_tree.xview)
        self.dungeon_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.dungeon_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        columns = [
            ("name", "副本名称", int(120*SCALE_FACTOR)),
            ("drops", "特殊掉落", int(150*SCALE_FACTOR))
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
            ttk.Button(btn_frame, text=text, command=command, width=int(10*SCALE_FACTOR)
            ).pack(side=tk.LEFT, padx=int(2*SCALE_FACTOR), fill=tk.X, expand=True)

    def build_dungeon_form(self, parent):
        """构建副本表单"""
        parent.grid_columnconfigure(1, weight=1)
        
        name_row = ttk.Frame(parent)
        name_row.grid(row=0, column=0, columnspan=3, sticky="ew", pady=int(5*SCALE_FACTOR))
        name_row.columnconfigure(1, weight=1)
        
        ttk.Label(name_row, text="副本名称:").grid(row=0, column=0, padx=int(4*SCALE_FACTOR), sticky=tk.W)
        self.preset_name_entry = ttk.Entry(name_row, textvariable=self.preset_name_var, 
                                          width=int(15*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.preset_name_entry.grid(row=0, column=1, padx=int(4*SCALE_FACTOR), sticky=tk.W)
        
        btn_frame = ttk.Frame(name_row)
        btn_frame.grid(row=0, column=2, padx=int(4*SCALE_FACTOR), sticky=tk.E)
        
        ttk.Button(btn_frame, text="保存", width=int(8*SCALE_FACTOR), command=self.save_dungeon
        ).pack(side=tk.LEFT, padx=int(2*SCALE_FACTOR))
        ttk.Button(btn_frame, text="清空", width=int(8*SCALE_FACTOR), command=self.clear_preset_form
        ).pack(side=tk.LEFT, padx=int(2*SCALE_FACTOR))
        
        ttk.Label(parent, text="特殊掉落:").grid(row=1, column=0, padx=int(4*SCALE_FACTOR), pady=int(5*SCALE_FACTOR), sticky=tk.W)
        self.preset_drops_entry = ttk.Entry(parent, textvariable=self.preset_drops_var, 
                                           width=int(180*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.preset_drops_entry.grid(row=1, column=1, padx=int(4*SCALE_FACTOR), pady=int(5*SCALE_FACTOR), sticky=tk.W, columnspan=2)
        
        batch_frame = ttk.LabelFrame(parent, text="批量添加特殊掉落", padding=int(6*SCALE_FACTOR))
        batch_frame.grid(row=2, column=0, columnspan=3, padx=int(4*SCALE_FACTOR), pady=int(5*SCALE_FACTOR), sticky=tk.W+tk.E)
        batch_frame.columnconfigure(0, weight=1)
        
        ttk.Label(batch_frame, text="输入多个物品，用逗号或方括号分隔:").pack(side=tk.TOP, anchor=tk.W, pady=(0, int(4*SCALE_FACTOR)))
        
        input_frame = ttk.Frame(batch_frame)
        input_frame.pack(fill=tk.X, pady=int(2*SCALE_FACTOR))
        
        self.batch_items_var = tk.StringVar()
        batch_entry = ttk.Entry(input_frame, textvariable=self.batch_items_var, font=("PingFang SC", int(10*SCALE_FACTOR)))
        batch_entry.pack(side=tk.LEFT, padx=(0, int(8*SCALE_FACTOR)), fill=tk.X, expand=True)
        
        ttk.Button(input_frame, text="添加", command=self.batch_add_items, width=int(8*SCALE_FACTOR)
        ).pack(side=tk.LEFT)

    def setup_events(self):
        """设置事件绑定"""
        # 绑定变量变化事件
        self.trash_gold_var.trace_add("write", self.update_total_gold)
        self.iron_gold_var.trace_add("write", self.update_total_gold)
        self.other_gold_var.trace_add("write", self.update_total_gold)
        self.special_total_var.trace_add("write", self.update_total_gold)
        self.total_gold_var.trace_add("write", self.validate_difference)
        
        # 设置数字输入验证
        reg = self.root.register(self.validate_numeric_input)
        vcmd = (reg, '%P')
        for entry in [self.trash_gold_entry, self.iron_gold_entry, self.other_gold_entry, 
                     self.fine_gold_entry, self.subsidy_gold_entry, self.lie_down_entry, 
                     self.personal_gold_entry]:
            entry.config(validate="key", validatecommand=vcmd)
        
        # 绑定组合框事件
        self.dungeon_combo.bind("<<ComboboxSelected>>", self.on_dungeon_select)
        self.search_dungeon_combo.bind("<<ComboboxSelected>>", self.on_search_dungeon_select)
        
        # 绑定记录树点击事件
        self.record_tree.bind('<ButtonRelease-1>', self.on_record_click)
        
        # 延迟加载列宽和时间更新
        self.root.after(300, self.load_column_widths)
        self.root.after(1000, self.update_time)

    def auto_resize_column(self, tree, column_id):
        """自动调整列宽"""
        font = tkFont.Font(family="PingFang SC", size=int(10*SCALE_FACTOR))
        heading_text = tree.heading(column_id)["text"]
        max_width = font.measure(heading_text) + int(20*SCALE_FACTOR)
        
        for item in tree.get_children():
            cell_value = tree.set(item, column_id)
            if cell_value:
                cell_width = font.measure(cell_value) + int(20*SCALE_FACTOR)
                if cell_width > max_width:
                    max_width = cell_width
        
        tree.column(column_id, width=max_width)

    def setup_column_resizing(self, tree):
        """设置列调整"""
        columns = tree["columns"]
        for col in columns:
            tree.heading(col, command=lambda c=col: self.auto_resize_column(tree, c))

    def on_record_click(self, event):
        """处理记录树点击事件"""
        item = self.record_tree.identify('item', event.x, event.y)
        if not item:
            return
            
        selected = self.record_tree.selection()
        if len(selected) == 1 and item in selected:
            self.fill_form_from_record(item)

    def fill_form_from_record(self, item):
        """从记录填充表单"""
        values = self.record_tree.item(item, 'values')
        if not values:
            return
            
        dungeon_name = values[1]
        time_str = values[2]
        
        # 查询数据库获取完整记录
        record = self.db.execute_query('''
            SELECT r.id, d.name, r.trash_gold, r.iron_gold, r.other_gold, r.special_auctions, 
                   r.total_gold, r.black_owner, r.worker, r.time, r.team_type, r.lie_down_count, 
                   r.fine_gold, r.subsidy_gold, r.personal_gold, r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            WHERE d.name = ? AND r.time LIKE ?
        ''', (dungeon_name, f"{time_str}%"))
        
        if not record:
            return
            
        r = record[0]
        self.dungeon_var.set(r[1])
        self.trash_gold_var.set(str(r[2]) if r[2] != 0 else "")
        self.iron_gold_var.set(str(r[3]) if r[3] != 0 else "")
        self.other_gold_var.set(str(r[4]) if r[4] != 0 else "")
        self.total_gold_var.set(str(r[6]))
        self.black_owner_var.set(r[7] or "")
        self.worker_var.set(r[8] or "")
        self.team_type_var.set(r[10])
        self.lie_down_var.set(str(r[11]) if r[11] != 0 else "")
        self.fine_gold_var.set(str(r[12]) if r[12] != 0 else "")
        self.subsidy_gold_var.set(str(r[13]) if r[13] != 0 else "")
        self.personal_gold_var.set(str(r[14]))
        self.note_var.set(r[15] or "")
        
        # 填充特殊掉落
        self.special_tree.clear()
        special_auctions = json.loads(r[5]) if r[5] else []
        for item_data in special_auctions:
            self.special_tree.add_item(item_data['item'], item_data['price'])
        self.special_total_var.set(str(self.special_tree.calculate_total()))
        
        # 更新按钮状态
        self.add_btn.configure(state=tk.NORMAL)
        self.edit_btn.configure(state=tk.NORMAL)
        self.update_btn.configure(state=tk.DISABLED)
        
        # 清除当前编辑ID
        if hasattr(self, 'current_edit_id'):
            del self.current_edit_id

    def load_data(self):
        """加载数据"""
        self.load_dungeon_options()
        
        # 检查数据库完整性
        self.check_database_integrity()
        
        self.load_dungeon_records()
        self.load_dungeon_presets()
        self.update_stats()
        self.load_black_owner_options()
        self.clear_new_record_highlights()
        
        print(f"数据加载完成 - 记录树中的项目数量: {len(self.record_tree.get_children())}")

    def load_dungeon_options(self):
        """加载副本选项"""
        if self.cached_dungeons is None:
            self.cached_dungeons = [row[0] for row in self.db.execute_query("SELECT name FROM dungeons")]
        
        self.dungeon_combo['values'] = self.cached_dungeons
        self.search_dungeon_combo['values'] = self.cached_dungeons
        self.search_item_combo['values'] = self.get_all_special_items()

    def get_all_special_items(self):
        """获取所有特殊物品"""
        items = set()
        for row in self.db.execute_query("SELECT special_drops FROM dungeons"):
            if row[0]:
                for item in row[0].split(','):
                    items.add(item.strip())
        return list(items)

    def load_dungeon_records(self):
        """加载副本记录"""
        # 清除现有记录
        for item in self.record_tree.get_children():
            self.record_tree.delete(item)
        
        # 使用LEFT JOIN确保即使dungeon_id无效也能显示记录
        records = self.db.execute_query('''
            SELECT r.id, 
                   COALESCE(d.name, '未知副本') as dungeon_name, 
                   strftime('%Y-%m-%d %H:%M', r.time), 
                   r.team_type, r.lie_down_count, r.total_gold, 
                   r.personal_gold, r.black_owner, r.worker, r.note, r.is_new
            FROM records r
            LEFT JOIN dungeons d ON r.dungeon_id = d.id
            ORDER BY r.time DESC
        ''')
        
        print(f"从数据库查询到的记录数量: {len(records)}")
        
        total_records = len(records)
        row_num = total_records
        
        # 插入记录到树形视图
        for row in records:
            note = row[9] or ""
            if len(note) > 15:
                note = note[:15] + "..."
            
            tags = ()
            if row[10] == 1:
                tags = ("new_record",)
                self.new_record_ids.add(row[0])
            
            self.record_tree.insert("", "end", values=(
                row_num, row[1], row[2], row[3], row[4], 
                f"{row[5]:,}", f"{row[6]:,}", row[7], row[8], note
            ), tags=tags)
            row_num -= 1
        
        # 刷新UI
        self.record_tree.update_idletasks()
        print(f"记录树已刷新，包含 {len(self.record_tree.get_children())} 条记录")

    def clear_new_record_highlights(self):
        """清除新记录高亮"""
        self.db.execute_update("UPDATE records SET is_new = 0 WHERE is_new = 1")
        for item in self.record_tree.get_children():
            self.record_tree.item(item, tags=())

    def load_dungeon_presets(self):
        """加载副本预设"""
        self.dungeon_tree.delete(*self.dungeon_tree.get_children())
        for row in self.db.execute_query("SELECT name, special_drops FROM dungeons"):
            self.dungeon_tree.insert("", "end", values=(row[0], row[1]))

    def load_column_widths(self):
        """加载列宽设置"""
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
        """保存列宽设置"""
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
        """验证数字输入"""
        return new_value == "" or new_value.isdigit()

    def on_dungeon_select(self, event):
        """处理副本选择事件"""
        selected = self.dungeon_var.get()
        if not selected:
            return
            
        result = self.db.execute_query("SELECT special_drops FROM dungeons WHERE name=?", (selected,))
        if result and result[0][0]:
            self.special_item_combo['values'] = [item.strip() for item in result[0][0].split(',')]
        else:
            self.special_item_combo['values'] = []

    def on_search_dungeon_select(self, event=None):
        """处理搜索副本选择事件"""
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
        """更新时间显示"""
        self.time_var.set(get_current_time())
        self.root.after(1000, self.update_time)

    def add_special_item(self):
        """添加特殊物品"""
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
        """更新总金额"""
        trash = self.trash_gold_var.get()
        iron = self.iron_gold_var.get()
        other = self.other_gold_var.get()
        special = self.special_total_var.get()
        total = GoldCalculator.calculate_total(trash, iron, other, special)
        self.total_gold_var.set(str(total))

    def validate_difference(self, *args):
        """验证差额"""
        trash = self.trash_gold_var.get()
        iron = self.iron_gold_var.get()
        other = self.other_gold_var.get()
        special = self.special_total_var.get()
        total = self.total_gold_var.get()
        
        difference = GoldCalculator.calculate_difference(total, trash, iron, other, special)
        self.difference_var.set(f"差额: {difference}")
        self.difference_label.configure(foreground="green" if difference == 0 else "red")

    def validate_and_save(self):
        """验证并保存数据"""
        if not self.dungeon_var.get():
            messagebox.showwarning("提示", "请选择副本")
            return False
            
        if not self.total_gold_var.get().isdigit() or int(self.total_gold_var.get()) <= 0:
            messagebox.showwarning("提示", "请输入有效的团队总工资")
            return False
            
        # 验证差额
        trash = self.trash_gold_var.get()
        iron = self.iron_gold_var.get()
        other = self.other_gold_var.get()
        special = self.special_total_var.get()
        total = self.total_gold_var.get()
        
        difference = GoldCalculator.calculate_difference(total, trash, iron, other, special)
        if difference != 0:
            if not messagebox.askyesno("确认", f"工资计算存在差额: {difference}，是否继续保存？"):
                return False
        
        # 保存数据
        self.add_record()
        return True

    def add_record(self):
        """添加记录"""
        dungeon = self.dungeon_var.get()
        if not dungeon:
            messagebox.showwarning("提示", "请选择副本")
            return
            
        # 获取副本ID
        result = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (dungeon,))
        if not result:
            messagebox.showerror("错误", f"找不到副本 '{dungeon}'")
            return
            
        dungeon_id = result[0][0]
        special_auctions = self.special_tree.get_items()
        
        # 计算个人工资
        personal_gold = GoldCalculator.safe_int(self.personal_gold_var.get())
        
        # 获取当前时间
        current_time = get_current_time()
        
        # 构建SQL插入语句
        try:
            self.db.execute_update('''
                INSERT INTO records (
                    dungeon_id, trash_gold, iron_gold, other_gold, special_auctions, 
                    total_gold, black_owner, worker, time, team_type, lie_down_count, 
                    fine_gold, subsidy_gold, personal_gold, note, is_new
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
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
            
            print("记录已成功添加到数据库")
            
            # 刷新数据
            self.load_dungeon_records()
            self.update_stats()
            self.load_black_owner_options()
            
            messagebox.showinfo("成功", "记录已添加")
            self.clear_form()
            
        except Exception as e:
            messagebox.showerror("错误", f"保存记录时出错: {str(e)}")
            print(f"保存记录错误: {e}")

    def edit_record(self):
        """编辑记录"""
        selected = self.record_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一条记录")
            return
            
        values = self.record_tree.item(selected[0], 'values')
        dungeon_name = values[1]
        time_str = values[2]
        
        record = self.db.execute_query('''
            SELECT r.id, d.name, r.trash_gold, r.iron_gold, r.other_gold, r.special_auctions, 
                   r.total_gold, r.black_owner, r.worker, r.time, r.team_type, r.lie_down_count, 
                   r.fine_gold, r.subsidy_gold, r.personal_gold, r.note
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            WHERE d.name = ? AND r.time LIKE ?
        ''', (dungeon_name, f"{time_str}%"))
        
        if not record:
            messagebox.showerror("错误", "找不到记录")
            return
            
        r = record[0]
        self.dungeon_var.set(r[1])
        self.trash_gold_var.set(str(r[2]) if r[2] != 0 else "")
        self.iron_gold_var.set(str(r[3]) if r[3] != 0 else "")
        self.other_gold_var.set(str(r[4]) if r[4] != 0 else "")
        self.total_gold_var.set(str(r[6]))
        self.black_owner_var.set(r[7] or "")
        self.worker_var.set(r[8] or "")
        self.team_type_var.set(r[10])
        self.lie_down_var.set(str(r[11]) if r[11] != 0 else "")
        self.fine_gold_var.set(str(r[12]) if r[12] != 0 else "")
        self.subsidy_gold_var.set(str(r[13]) if r[13] != 0 else "")
        self.personal_gold_var.set(str(r[14]))
        self.note_var.set(r[15] or "")
        
        self.special_tree.clear()
        special_auctions = json.loads(r[5]) if r[5] else []
        for item in special_auctions:
            self.special_tree.add_item(item['item'], item['price'])
        self.special_total_var.set(str(self.special_tree.calculate_total()))
        
        self.current_edit_id = r[0]
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
        """搜索记录"""
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
                   r.personal_gold, r.black_owner, r.worker, r.note, r.is_new
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
        '''
        if conditions:
            sql += " WHERE " + " AND ".join(conditions)
        sql += " ORDER BY r.time DESC"
        
        records = self.db.execute_query(sql, params)
        total_records = len(records)
        row_num = total_records
        
        for row in records:
            note = row[9] or ""
            if len(note) > 15:
                note = note[:15] + "..."
            
            tags = ()
            if row[10] == 1:
                tags = ("new_record",)
            
            self.record_tree.insert("", "end", values=(row_num, row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], note), tags=tags)
            row_num -= 1

    def validate_date(self, date_str):
        """验证日期格式"""
        if not date_str:
            return True
        try:
            datetime.datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def reset_search(self):
        """重置搜索"""
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
        """显示记录上下文菜单"""
        row_id = self.record_tree.identify_row(event.y)
        if not row_id:
            return
        # 如果右键点击的项目不在当前选择中，则清除选择并选择该项目
        if row_id not in self.record_tree.selection():
            self.record_tree.selection_set(row_id)
        self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_records(self):
        """删除选中记录"""
        selected = self.record_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择要删除的记录")
            return
            
        if not messagebox.askyesno("确认", f"确定要删除选中的 {len(selected)} 条记录吗？"):
            return
            
        for item in selected:
            values = self.record_tree.item(item, 'values')
            dungeon_name = values[1]
            time_str = values[2]
            
            result = self.db.execute_query('''
                SELECT r.id FROM records r
                JOIN dungeons d ON r.dungeon_id = d.id
                WHERE d.name = ? AND r.time LIKE ?
            ''', (dungeon_name, f"{time_str}%"))
            
            if result:
                record_id = result[0][0]
                self.db.execute_update("DELETE FROM records WHERE id=?", (record_id,))
            
        self.load_dungeon_records()
        self.update_stats()
        self.load_black_owner_options()
        messagebox.showinfo("成功", "记录已删除")

    def on_tree_motion(self, event):
        """处理树形视图鼠标移动事件"""
        self.tooltip.withdraw()
        row_id = self.record_tree.identify_row(event.y)
        col_id = self.record_tree.identify_column(event.x)
        
        if not row_id:
            return
            
        item = self.record_tree.item(row_id)
        col_name = self.record_tree.heading(col_id)["text"]
        values = item['values']
        
        dungeon_name = values[1]
        time_str = values[2]
        
        result = self.db.execute_query('''
            SELECT r.id FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            WHERE d.name = ? AND r.time LIKE ?
        ''', (dungeon_name, f"{time_str}%"))
        
        if not result:
            return
            
        record_id = result[0][0]
        
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
        """显示工具提示"""
        self.tooltip_label.config(text=text)
        self.tooltip.update_idletasks()
        
        x = event.x_root + int(15*SCALE_FACTOR)
        y = event.y_root + int(15*SCALE_FACTOR)
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        if x + self.tooltip.winfo_width() > screen_width:
            x = event.x_root - self.tooltip.winfo_width() - int(5*SCALE_FACTOR)
        if y + self.tooltip.winfo_height() > screen_height:
            y = event.y_root - self.tooltip.winfo_height() - int(5*SCALE_FACTOR)
            
        self.tooltip.geometry(f"+{x}+{y}")
        self.tooltip.deiconify()

    def hide_tooltip(self, event=None):
        """隐藏工具提示"""
        self.tooltip.withdraw()

    def update_stats(self):
        """更新统计信息"""
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
        
        print(f"统计信息已更新 - 总记录数: {total_records}")

    def update_worker_stats(self):
        """更新打工仔统计"""
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
        """更新图表"""
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
            ORDER by week
        ''', (start_date.strftime("%Y-%m-%d"),))
        
        if not results:
            self.ax.text(0.5, 0.5, "暂无数据", ha='center', va='center', fontsize=int(10*SCALE_FACTOR))
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
        self.ax.set_title('每周工资统计', fontsize=int(12*SCALE_FACTOR))
        self.ax.set_xlabel('日期', fontsize=int(10*SCALE_FACTOR))
        self.ax.set_ylabel('金额', fontsize=int(10*SCALE_FACTOR))
        self.ax.legend(fontsize=int(10*SCALE_FACTOR))
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
        self.ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
        self.fig.autofmt_xdate()
        self.ax.grid(True, linestyle='--', alpha=0.7)
        self.canvas.draw()

    def add_dungeon(self):
        """添加副本"""
        self.clear_preset_form()
        self.preset_name_entry.focus()

    def edit_dungeon(self):
        """编辑副本"""
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        values = self.dungeon_tree.item(selected[0], 'values')
        self.preset_name_var.set(values[0])
        self.preset_drops_var.set(values[1])

    def delete_dungeon(self):
        """删除副本"""
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        values = self.dungeon_tree.item(selected[0], 'values')
        dungeon_name = values[0]
        
        result = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (dungeon_name,))
        if not result:
            return
            
        dungeon_id = result[0][0]
        
        count = self.db.execute_query("SELECT COUNT(*) FROM records WHERE dungeon_id=?", (dungeon_id,))[0][0]
        if count > 0:
            messagebox.showerror("错误", f"无法删除副本 '{dungeon_name}'，因为存在 {count} 条相关记录")
            return
            
        if messagebox.askyesno("确认", f"确定要删除副本 '{dungeon_name}' 吗？"):
            self.db.execute_update("DELETE FROM dungeons WHERE id=?", (dungeon_id,))
            self.load_dungeon_presets()
            self.load_dungeon_options()

    def batch_add_items(self):
        """批量添加物品"""
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
        """保存副本"""
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
        """清空预设表单"""
        self.preset_name_var.set("")
        self.preset_drops_var.set("")
        self.batch_items_var.set("")

    def import_data(self):
        """导入数据"""
        file_path = filedialog.askopenfilename(title="选择数据文件", filetypes=[("JSON文件", "*.json")])
        if not file_path:
            return
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            imported_count = 0
            skipped_count = 0
            
            if 'dungeons' in data:
                for dungeon in data['dungeons']:
                    existing = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (dungeon['name'],))
                    if not existing:
                        self.db.execute_update('''
                            INSERT INTO dungeons (name, special_drops) VALUES (?, ?)
                        ''', (dungeon['name'], dungeon['special_drops']))
            
            if 'records' in data:
                for record in data['records']:
                    result = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (record['dungeon'],))
                    if not result:
                        continue
                    
                    dungeon_id = result[0][0]
                    
                    existing_record = self.db.execute_query('''
                        SELECT id FROM records 
                        WHERE dungeon_id=? AND time=? AND total_gold=?
                        AND COALESCE(black_owner, '') = COALESCE(?, '')
                        AND COALESCE(worker, '') = COALESCE(?, '')
                    ''', (
                        dungeon_id, 
                        record['time'], 
                        record['total_gold'],
                        record.get('black_owner', '') or '',
                        record.get('worker', '') or ''
                    ))
                    
                    if existing_record:
                        skipped_count += 1
                        continue
                    
                    self.db.execute_update('''
                        INSERT INTO records (dungeon_id, trash_gold, iron_gold, other_gold, 
                            special_auctions, total_gold, black_owner, worker, time, 
                            team_type, lie_down_count, fine_gold, subsidy_gold, personal_gold, note, is_new)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                    ''', (
                        dungeon_id, 
                        record['trash_gold'], 
                        record['iron_gold'], 
                        record['other_gold'],
                        json.dumps(record['special_auctions'], ensure_ascii=False),
                        record['total_gold'], 
                        record.get('black_owner', '') or None, 
                        record.get('worker', '') or None, 
                        record['time'],
                        record['team_type'], 
                        record['lie_down_count'], 
                        record['fine_gold'],
                        record['subsidy_gold'], 
                        record['personal_gold'], 
                        record.get('note', '') or ''
                    ))
                    imported_count += 1
            
            self.load_dungeon_records()
            self.update_stats()
            self.load_dungeon_options()
            self.load_black_owner_options()
            
            message = f"数据导入完成\n成功导入: {imported_count} 条记录\n跳过重复: {skipped_count} 条记录"
            if imported_count > 0:
                message += "\n\n新增记录已高亮显示，下次打开应用后将取消高亮"
            messagebox.showinfo("成功", message)
        except Exception as e:
            messagebox.showerror("错误", f"导入数据失败: {str(e)}")

    def export_data(self):
        """导出数据"""
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
            ORDER BY r.time DESC
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
        """加载黑本和打工仔选项"""
        if self.cached_owners is None:
            self.cached_owners = [row[0] for row in self.db.execute_query("SELECT DISTINCT black_owner FROM records WHERE black_owner IS NOT NULL AND black_owner != ''")]
        
        if self.cached_workers is None:
            self.cached_workers = [row[0] for row in self.db.execute_query("SELECT DISTINCT worker FROM records WHERE worker IS NOT NULL AND worker != ''")]
        
        self.black_owner_combo['values'] = self.cached_owners
        self.search_owner_combo['values'] = self.cached_owners
         
        self.worker_combo['values'] = self.cached_workers
        self.search_worker_combo['values'] = self.cached_workers

    def check_database_integrity(self):
        """检查数据库完整性"""
        orphaned_records = self.db.execute_query('''
            SELECT r.id, r.dungeon_id 
            FROM records r 
            WHERE r.dungeon_id NOT IN (SELECT id FROM dungeons) AND r.dungeon_id IS NOT NULL
        ''')
        
        if orphaned_records:
            print(f"找到 {len(orphaned_records)} 条无效的dungeon_id记录")
            
            valid_dungeon = self.db.execute_query("SELECT id FROM dungeons LIMIT 1")
            if valid_dungeon:
                valid_dungeon_id = valid_dungeon[0][0]
                print(f"使用有效的dungeon_id进行修复: {valid_dungeon_id}")
                
                for record in orphaned_records:
                    record_id, old_dungeon_id = record
                    self.db.execute_update(
                        "UPDATE records SET dungeon_id = ? WHERE id = ?",
                        (valid_dungeon_id, record_id)
                    )
                    print(f"修复记录 {record_id}: dungeon_id {old_dungeon_id} -> {valid_dungeon_id}")
                
                print("数据库自动修复完成")
        
        return orphaned_records

    def repair_database(self):
        """修复数据库"""
        orphaned_records = self.check_database_integrity()
        if orphaned_records:
            messagebox.showinfo("数据库修复", f"已修复 {len(orphaned_records)} 条无效记录")
            self.load_dungeon_records()
        else:
            messagebox.showinfo("数据库修复", "没有发现需要修复的记录")

    def on_close(self):
        """处理关闭事件"""
        self.clear_new_record_highlights()
        self.save_column_widths()
        
        # 保存窗口状态
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        maximized = 1 if self.root.state() == 'zoomed' else 0
        
        self.db.execute_update("DELETE FROM window_state")
        self.db.execute_update("INSERT INTO window_state (width, height, maximized, x, y) VALUES (?, ?, ?, ?, ?)", 
                              (width, height, maximized, x, y))
        
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
