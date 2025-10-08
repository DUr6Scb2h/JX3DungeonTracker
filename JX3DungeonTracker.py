import calendar
from datetime import timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import datetime as dt
import json
import os
import sys
import tkinter.font as tkFont
import platform
import shutil
import locale
import atexit
import re
import threading
import time

SCALE_FACTOR = 1
MATPLOTLIB_AVAILABLE = False

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
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_app_data_path():
    if platform.system() == "Windows":
        app_data = os.getenv('APPDATA')
        app_dir = os.path.join(app_data, "JX3DungeonTracker")
    elif platform.system() == "Darwin":
        app_dir = os.path.expanduser("~/Library/Application Support/JX3DungeonTracker")
    else:
        app_dir = os.path.expanduser("~/.jx3dungeontracker")
    
    os.makedirs(app_dir, exist_ok=True)
    return app_dir

def get_current_time():
    try:
        now = dt.datetime.now()
        return now.strftime("%Y-%m-%d %H:%M:%S")
    except:
        now = dt.datetime.utcnow()
        return now.strftime("%Y-%m-%d %H:%M:%S")

class DatabaseManager:
    def __init__(self, db_path):
        db_dir = os.path.dirname(db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
            
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self.cursor.execute("PRAGMA foreign_keys = ON")
        
        self.initialize_tables()
        self.load_preset_dungeons()
        self.upgrade_database()

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
                note TEXT DEFAULT '',
                is_new INTEGER DEFAULT 0,
                scattered_consumption INTEGER DEFAULT 0,
                iron_consumption INTEGER DEFAULT 0,
                special_consumption INTEGER DEFAULT 0,
                other_consumption INTEGER DEFAULT 0,
                total_consumption INTEGER DEFAULT 0,
                FOREIGN KEY (dungeon_id) REFERENCES dungeons (id)
            )
        ''')

        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS column_widths (
                tree_name TEXT PRIMARY KEY,
                widths TEXT
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS window_state (
                width INTEGER,
                height INTEGER,
                maximized INTEGER DEFAULT 0,
                x INTEGER,
                y INTEGER
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS pane_positions (
                pane_name TEXT PRIMARY KEY,
                position INTEGER
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS analysis_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT NOT NULL,
                remark TEXT
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS filled_uids (
                uid TEXT PRIMARY KEY,
                fill_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.conn.commit()

    def upgrade_database(self):
        try:
            self.cursor.execute("PRAGMA table_info(records)")
            columns = [column[1] for column in self.cursor.fetchall()]
            
            if 'is_new' not in columns:
                self.cursor.execute("ALTER TABLE records ADD COLUMN is_new INTEGER DEFAULT 0")
            
            consumption_columns = [
                'scattered_consumption', 
                'iron_consumption', 
                'special_consumption', 
                'other_consumption', 
                'total_consumption'
            ]
            
            for col in consumption_columns:
                if col not in columns:
                    self.cursor.execute(f"ALTER TABLE records ADD COLUMN {col} INTEGER DEFAULT 0")
                    
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
            pass

    def load_preset_dungeons(self):
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

    def execute_query(self, query, params=()):
        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def execute_update(self, query, params=()):
        self.cursor.execute(query, params)
        self.conn.commit()

    def close(self):
        try:
            if hasattr(self, 'cursor'):
                self.cursor.close()
            if hasattr(self, 'conn'):
                self.conn.commit()
                self.conn.close()
        except Exception:
            pass

    def get_pane_position(self, pane_name):
        result = self.execute_query("SELECT position FROM pane_positions WHERE pane_name = ?", (pane_name,))
        return result[0][0] if result else None

    def save_pane_position(self, pane_name, position):
        self.execute_update('''
            INSERT OR REPLACE INTO pane_positions (pane_name, position) 
            VALUES (?, ?)
        ''', (pane_name, position))

class SpecialItemsTree:
    def __init__(self, parent):
        self.tree = ttk.Treeview(parent, columns=("item", "price"), show="headings", 
                                height=int(3*SCALE_FACTOR), selectmode="browse")
        self.tree.heading("item", text="物品", anchor="center")
        self.tree.heading("price", text="金额", anchor="center")
        self.tree.column("item", width=int(120*SCALE_FACTOR), anchor=tk.CENTER)
        self.tree.column("price", width=int(60*SCALE_FACTOR), anchor=tk.CENTER)
        
        style = ttk.Style()
        style.configure("Special.Treeview", font=("PingFang SC", int(9*SCALE_FACTOR)), 
                       rowheight=int(24*SCALE_FACTOR))
        self.tree.configure(style="Special.Treeview")
        
        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        
        self.setup_context_menu()

    def setup_context_menu(self):
        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="删除选中项", command=self.delete_selected_items)
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_items(self):
        selected_items = self.tree.selection()
        for item in selected_items:
            self.tree.delete(item)

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

class DBAnalyzer:
    def __init__(self, parent, main_app):
        self.parent = parent
        self.main_app = main_app
        self.db_folders = {}
        self.analysis_results = []
        self.filled_uids = set()
        
        self.optimize_patterns()
        self.batch_size = 5000
        self.max_file_size_mb = 100
        
        self.setup_ui()
        self.load_folder_list()
        self.load_filled_uids()

    def optimize_patterns(self):
        self.patterns = {
            'start': re.compile(r'^你悄悄地对\[[^\]]+\]说：开始自动记录\[(.*?)\]$'),
            'end': re.compile(r'^你悄悄地对\[[^\]]+\]说：结束自动记录\[(.*?)\]$'),
            'team_info': re.compile(
                r'\[房间\]\[([^\]]+)\]：拍团目前总收入为：(\d+)金，'
                r'补贴总费用：(\d+)金，\s*实际可用分配金额：(\d+)金，'
                r'\s*分配人数：(\d+)，\s*每人底薪：(\d+)金'
            ),
            'personal_salary_named': re.compile(r'text="(\d+)"[^>]*name="Text_(Gold|Silver|Copper)"'),
            'penalty': re.compile(r'\[房间\]\[([^\]]+)\]：.*?向团队里追加了\[(\d+)金\]'),
            'item_purchase': re.compile(r'\[房间\]\[([^\]]+)\]：\[([^\]]+)\]花费\[(.*?)\]购买了\[(.*?)\]'),
            'gold_amount': re.compile(r'(\d+)金砖|(\d+)金')
        }
        
        self.fixed_rules = {
            "scattered_keywords": ["五行石", "五彩石", "上品茶饼", "猫眼石", "玛瑙"],
            "iron_keywords": ["陨铁"]
        }

    def setup_ui(self):
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=int(10*SCALE_FACTOR), pady=int(10*SCALE_FACTOR))
        
        file_frame = ttk.LabelFrame(main_frame, text="数据库文件夹列表", padding=int(8*SCALE_FACTOR))
        file_frame.pack(fill=tk.X, pady=(0, int(10*SCALE_FACTOR)))
        
        tree_container = ttk.Frame(file_frame)
        tree_container.pack(fill=tk.BOTH, expand=True, pady=(0, int(5*SCALE_FACTOR)))
        
        columns = ("folder", "remark")
        self.file_treeview = ttk.Treeview(tree_container, columns=columns, show="headings", height=6)
        self.file_treeview.heading("folder", text="文件夹路径", anchor="center")
        self.file_treeview.heading("remark", text="打工仔", anchor="center")
        self.file_treeview.column("folder", width=int(400*SCALE_FACTOR), anchor=tk.CENTER)
        self.file_treeview.column("remark", width=int(150*SCALE_FACTOR), anchor=tk.CENTER)
        
        file_vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.file_treeview.yview)
        file_hsb = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.file_treeview.xview)
        self.file_treeview.configure(yscrollcommand=file_vsb.set, xscrollcommand=file_hsb.set)
        
        self.file_treeview.grid(row=0, column=0, sticky="nsew")
        file_vsb.grid(row=0, column=1, sticky="ns")
        file_hsb.grid(row=1, column=0, sticky="ew")
        
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)
        
        self.file_treeview.bind('<<TreeviewSelect>>', self.on_treeview_select)
        
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        
        ttk.Button(btn_frame, text="添加文件夹", command=self.add_folder).pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        ttk.Button(btn_frame, text="移除文件夹", command=self.remove_folder).pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        ttk.Button(btn_frame, text="清空列表", command=self.clear_folders).pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        ttk.Button(btn_frame, text="保存列表", command=self.save_folder_list).pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        
        remark_frame = ttk.Frame(file_frame)
        remark_frame.pack(fill=tk.X)
        
        ttk.Label(remark_frame, text="打工仔备注:").pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        self.remark_entry = ttk.Entry(remark_frame, width=int(30*SCALE_FACTOR))
        self.remark_entry.pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        ttk.Button(remark_frame, text="修改选中文件夹备注", command=self.edit_selected_remark).pack(side=tk.LEFT)
        
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, int(10*SCALE_FACTOR)))
        
        ttk.Button(control_frame, text="开始分析", command=self.start_analysis).pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        ttk.Button(control_frame, text="填充到表单", command=self.fill_form).pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        
        self.progress_frame = ttk.LabelFrame(main_frame, text="分析进度", padding=int(8*SCALE_FACTOR))
        self.progress_frame.pack(fill=tk.X, pady=(0, int(10*SCALE_FACTOR)))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        
        self.status_var = tk.StringVar(value="准备就绪")
        self.status_label = ttk.Label(self.progress_frame, textvariable=self.status_var)
        self.status_label.pack(fill=tk.X)
        
        result_frame = ttk.LabelFrame(main_frame, text="分析结果", padding=int(8*SCALE_FACTOR))
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("uid", "start_time", "end_time", "dungeon_name", "black_person", "worker", 
                "team_total", "personal", "consumption", "subsidy", "penalty", "scattered", "iron", "other", "special", 
                "team_type", "lie_count", "note")
        
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15, selectmode="browse")
        
        column_config = [
            ("uid", "UID", 80),
            ("start_time", "开始时间", 120),
            ("end_time", "结束时间", 120),
            ("dungeon_name", "副本名", 100),
            ("black_person", "黑本人", 80),
            ("worker", "打工仔", 80),
            ("team_total", "团队总工资", 100),
            ("personal", "个人工资", 80),
            ("consumption", "本场消费", 80),
            ("subsidy", "补贴", 60),
            ("penalty", "罚款", 60),
            ("scattered", "散件金额", 80),
            ("iron", "小铁金额", 80),
            ("other", "其他金额", 80),
            ("special", "特殊金额", 80),
            ("team_type", "团队类型", 80),
            ("lie_count", "躺拍人数", 80),
            ("note", "备注", 100)
        ]
        
        for col_id, heading, width in column_config:
            self.result_tree.heading(col_id, text=heading, anchor="center")
            self.result_tree.column(col_id, width=int(width*SCALE_FACTOR), anchor=tk.CENTER)
        
        vsb = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        hsb = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

    def update_progress(self, value, status=""):
        try:
            self.progress_var.set(value)
            if status:
                self.status_var.set(status)
            self.parent.update_idletasks()
        except Exception:
            pass

    def scan_folder_for_db_files(self, folder_path):
        db_files = []
        try:
            if not os.path.exists(folder_path):
                return []
                
            for file in os.listdir(folder_path):
                if file.endswith('.db'):
                    file_path = os.path.join(folder_path, file)
                    if os.path.isfile(file_path):
                        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                        max_size = self.max_file_size_mb
                        
                        if file_size_mb <= max_size:
                            db_files.append(file_path)
        except Exception:
            pass
        
        return db_files

    def analyze_db_file_optimized(self, db_file, remark):
        try:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            total_records = cursor.execute("SELECT COUNT(*) FROM chatlog").fetchone()[0]
            if total_records == 0:
                conn.close()
                return []
            
            all_records = []
            batch_size = self.batch_size
            
            for offset in range(0, total_records, batch_size):
                cursor.execute(
                    "SELECT time, text, msg FROM chatlog ORDER BY time LIMIT ? OFFSET ?", 
                    (batch_size, offset)
                )
                batch_records = cursor.fetchall()
                all_records.extend(batch_records)
                
                progress = min(50, (offset + len(batch_records)) / total_records * 50)
                self.update_progress(progress, f"读取数据: {os.path.basename(db_file)}")
            
            conn.close()
            
            self.update_progress(60, f"分析记录: {os.path.basename(db_file)}")
            analysis_results = self.analyze_records_optimized(all_records, remark, os.path.basename(db_file))
            
            self.update_progress(100, f"完成分析: {os.path.basename(db_file)}")
            return analysis_results
            
        except Exception:
            return [self.create_empty_result(os.path.basename(db_file), remark)]

    def analyze_records_optimized(self, records, remark, filename):
        start_positions = []
        end_positions = []
        
        for i, (time, text, msg) in enumerate(records):
            start_match = self.patterns['start'].search(text)
            if start_match:
                dungeon_info = start_match.group(1)
                start_positions.append((i, time, text, dungeon_info))
            
            end_match = self.patterns['end'].search(text)
            if end_match:
                dungeon_info = end_match.group(1)
                end_positions.append((i, time, text, dungeon_info))
        
        if not start_positions or not end_positions:
            return [self.create_empty_result(filename, remark)]
        
        start_positions.sort(key=lambda x: x[1])
        end_positions.sort(key=lambda x: x[1])
        
        matched_pairs = self.match_record_pairs(start_positions, end_positions)
        
        all_results = []
        for start_idx, end_idx, start_time, end_time, start_text, end_text, dungeon_info in matched_pairs:
            result = self.analyze_single_record_segment_optimized(
                records, start_idx, end_idx, remark, filename, dungeon_info
            )
            if result:
                all_results.append(result)
        
        if not all_results:
            all_results.append(self.create_empty_result(filename, remark))
        
        return all_results

    def match_record_pairs(self, start_positions, end_positions):
        matched_pairs = []
        used_starts = set()
        used_ends = set()
        
        for start_idx, start_time, start_text, start_dungeon_info in start_positions:
            if start_idx in used_starts:
                continue
                
            possible_ends = [
                (idx, t, txt, dungeon_info) for idx, t, txt, dungeon_info in end_positions 
                if idx > start_idx and idx not in used_ends and dungeon_info == start_dungeon_info
            ]
            
            if possible_ends:
                end_idx, end_time, end_text, end_dungeon_info = min(possible_ends, key=lambda x: x[0])
                matched_pairs.append((
                    start_idx, end_idx, start_time, end_time, start_text, end_text, start_dungeon_info
                ))
                used_starts.add(start_idx)
                used_ends.add(end_idx)
        
        return matched_pairs

    def process_item_purchase_with_consumption(self, item_match, analysis_data, special_items_list, current_worker):
        room_name = item_match.group(1)
        buyer = item_match.group(2)
        gold_text = item_match.group(3)
        item_name = item_match.group(4)
        
        item_price = self.parse_gold_amount(gold_text)
        is_worker_purchase = (buyer == current_worker)
        
        is_special = False
        special_item_name = ""
        
        for special_item in special_items_list:
            if self.is_special_item_match(item_name, special_item):
                is_special = True
                special_item_name = special_item
                break
        
        if is_special:
            analysis_data["special_total"] += item_price
            analysis_data["special_items"].append({
                "item": special_item_name,
                "price": item_price,
                "original_name": item_name,
                "buyer": buyer
            })
            
            if is_worker_purchase:
                analysis_data["special_consumption"] += item_price
        else:
            is_potential_special = self.is_potential_special_item(item_name)
            if is_potential_special:
                return
            
            is_scattered = any(keyword in item_name for keyword in self.fixed_rules["scattered_keywords"])
            is_iron = any(keyword in item_name for keyword in self.fixed_rules["iron_keywords"])
            
            if is_worker_purchase:
                if is_scattered:
                    analysis_data["scattered_total"] += item_price
                    analysis_data["scattered_consumption"] += item_price
                elif is_iron:
                    analysis_data["iron_total"] += item_price
                    analysis_data["iron_consumption"] += item_price
                else:
                    analysis_data["other_total"] += item_price
                    analysis_data["other_consumption"] += item_price
            else:
                if is_scattered:
                    analysis_data["scattered_total"] += item_price
                elif is_iron:
                    analysis_data["iron_total"] += item_price
                else:
                    analysis_data["other_total"] += item_price

    def analyze_single_line_with_consumption(self, text, msg, analysis_data, special_items_list, current_worker):
        if "拍团目前总收入为" in text:
            team_match = self.patterns['team_info'].search(text)
            if team_match:
                analysis_data.update({
                    "black_person": team_match.group(1),
                    "team_total_salary": int(team_match.group(2)),
                    "subsidy_total": int(team_match.group(3)),
                    "actual_distributable": int(team_match.group(4)),
                    "distribution_count": int(team_match.group(5)),
                    "base_salary": int(team_match.group(6))
                })
        
        item_match = self.patterns['item_purchase'].search(text)
        if item_match:
            self.process_item_purchase_with_consumption(item_match, analysis_data, special_items_list, current_worker)
        
        if msg and "你获得：" in msg and "Text_Gold" in msg:
            print(f"\n=== 开始解析个人工资信息 ===")
            print(f"原始msg: {msg}")
        
            cleaned_msg = re.sub(r'\s+', '', msg)
            
            matches = self.patterns['personal_salary_named'].findall(cleaned_msg)
            if len(matches) >= 3:
                gold = silver = copper = 0
                for num, coin_type in matches:
                    if coin_type == "Gold":
                        gold = int(num)
                    elif coin_type == "Silver":
                        silver = int(num)
                    elif coin_type == "Copper":
                        copper = int(num)
                
                if gold > 0:
                    total_copper = gold * 10000 + silver * 100 + copper
                    salary_amount = round(total_copper / 10000)
                    print(f"解析结果 - 金: {gold}, 银: {silver}, 铜: {copper}")
                    print(f"铜钱总数: {total_copper}, 折算金数: {salary_amount}")
                    analysis_data["personal_salaries"].append(salary_amount)
        
        penalty_match = self.patterns['penalty'].search(text)
        if penalty_match:
            penalty_player = penalty_match.group(1)
            penalty_amount = int(penalty_match.group(2))
            
            if penalty_player == current_worker:
                analysis_data["penalty_total"] += penalty_amount

    def analyze_single_record_segment_optimized(self, records, start_idx, end_idx, remark, filename, dungeon_info):
        team_type, dungeon_name, difficulty_note = self.parse_dungeon_info(dungeon_info)
        
        analysis_data = {
            "dungeon_name": dungeon_name,
            "team_type": team_type,
            "difficulty_note": difficulty_note,
            "black_person": "",
            "personal_salaries": [],
            "team_total_salary": 0,
            "subsidy_total": 0,
            "actual_distributable": 0,
            "distribution_count": 0,
            "base_salary": 0,
            "penalty_total": 0,
            "lie_count": 0,
            "scattered_total": 0,
            "iron_total": 0,
            "other_total": 0,
            "special_total": 0,
            "special_items": [],
            "scattered_consumption": 0,
            "iron_consumption": 0,
            "special_consumption": 0,
            "other_consumption": 0,
            "total_consumption": 0,
            "worker": remark
        }
        
        current_dungeon_special_items = self.get_special_items_for_dungeon(dungeon_name)
        
        for i in range(start_idx, end_idx + 1):
            time, text, msg = records[i]
            self.analyze_single_line_with_consumption(text, msg, analysis_data, current_dungeon_special_items, remark)
        
        analysis_data["lie_count"] = self.calculate_lie_count(
            analysis_data["team_type"], 
            analysis_data["distribution_count"]
        )
        
        analysis_data["total_consumption"] = (
            analysis_data["scattered_consumption"] + 
            analysis_data["iron_consumption"] + 
            analysis_data["special_consumption"] + 
            analysis_data["other_consumption"]
        )
        
        return self.calculate_final_result(analysis_data, records, start_idx, end_idx, remark, filename)

    def calculate_lie_count(self, team_type, distribution_count):
        if not distribution_count or distribution_count <= 0:
            return 0
            
        if team_type == "十人本":
            total_players = 10
        elif team_type == "二十五人本":
            total_players = 25
        else:
            return 0
        
        lie_count = total_players - distribution_count
        return max(0, lie_count)

    def parse_dungeon_info(self, dungeon_info):
        team_type = "未知"
        dungeon_name = "未知副本"
        difficulty_note = ""
        
        try:
            if "10人" in dungeon_info:
                team_type = "十人本"
                clean_info = dungeon_info.replace("10人", "")
            elif "25人" in dungeon_info:
                team_type = "二十五人本"
                clean_info = dungeon_info.replace("25人", "")
            else:
                clean_info = dungeon_info
            
            difficulty_patterns = ["普通", "英雄", "挑战", "简单", "困难"]
            found_difficulty = ""
            
            for pattern in difficulty_patterns:
                if pattern in clean_info:
                    found_difficulty = pattern
                    clean_info = clean_info.replace(pattern, "")
                    break
            
            if team_type == "二十五人本" and found_difficulty in ["普通", "英雄"]:
                difficulty_note = found_difficulty
            
            raw_dungeon_name = clean_info.strip()
            dungeon_name = self.find_matching_dungeon(raw_dungeon_name)
            
        except Exception:
            if "10人" in dungeon_info:
                team_type = "十人本"
            elif "25人" in dungeon_info:
                team_type = "二十五人本"
            dungeon_name = self.find_matching_dungeon(dungeon_info)
        
        return team_type, dungeon_name, difficulty_note

    def find_matching_dungeon(self, raw_dungeon_name):
        dungeons = self.load_all_dungeons()
        
        for dungeon in dungeons:
            if dungeon in raw_dungeon_name:
                return dungeon
        
        for dungeon in dungeons:
            if raw_dungeon_name in dungeon:
                return dungeon
        
        return "未知副本"

    def load_all_dungeons(self):
        if hasattr(self, '_cached_dungeons'):
            return self._cached_dungeons
        
        try:
            result = self.main_app.db.execute_query("SELECT name FROM dungeons")
            self._cached_dungeons = [row[0] for row in result]
            return self._cached_dungeons
        except Exception:
            return []

    def get_special_items_for_dungeon(self, dungeon_name):
        try:
            result = self.main_app.db.execute_query(
                "SELECT special_drops FROM dungeons WHERE name = ?", 
                (dungeon_name,)
            )
            if result and result[0][0]:
                items = [item.strip() for item in result[0][0].split(',')]
                return items
            else:
                return []
        except Exception:
            return []

    def calculate_final_result(self, analysis_data, records, start_idx, end_idx, remark, filename):
        personal_salary = max(analysis_data["personal_salaries"]) if analysis_data["personal_salaries"] else 0
        
        note_parts = []
        
        if analysis_data.get("difficulty_note"):
            note_parts.append(analysis_data["difficulty_note"])
        
        if personal_salary == 10:
            personal_salary = 0
            note_parts.append("躺拍")
            
            if analysis_data["penalty_total"] > 0:
                note_parts.append(f"抵消{analysis_data['penalty_total']}金")
        
        note = "，".join(note_parts)
        
        subsidy = 0
        if personal_salary > 0:
            if personal_salary > analysis_data["base_salary"]:
                subsidy = analysis_data["penalty_total"] + (personal_salary - analysis_data["base_salary"])
        
        start_time_str = dt.datetime.fromtimestamp(records[start_idx][0]).strftime('%Y-%m-%d %H:%M:%S')
        end_time_str = dt.datetime.fromtimestamp(records[end_idx][0]).strftime('%Y-%m-%d %H:%M:%S')
        
        analysis_result = {
            "filename": filename,
            "remark": remark,
            "start_time": start_time_str,
            "end_time": end_time_str,
            "dungeon_name": analysis_data["dungeon_name"],
            "black_person": analysis_data["black_person"],
            "worker": remark,
            "team_total_salary": analysis_data["team_total_salary"],
            "personal_salary": personal_salary,
            "subsidy": subsidy,
            "penalty_total": analysis_data["penalty_total"],
            "scattered_total": analysis_data["scattered_total"],
            "iron_total": analysis_data["iron_total"],
            "other_total": analysis_data["other_total"],
            "special_total": analysis_data["special_total"],
            "special_items": analysis_data["special_items"],
            "team_type": analysis_data["team_type"],
            "lie_count": analysis_data["lie_count"],
            "note": note,
            "scattered_consumption": analysis_data["scattered_consumption"],
            "iron_consumption": analysis_data["iron_consumption"],
            "special_consumption": analysis_data["special_consumption"],
            "other_consumption": analysis_data["other_consumption"],
            "total_consumption": analysis_data["total_consumption"]
        }
        
        analysis_result["uid"] = self.generate_uid(analysis_result)
        
        return analysis_result

    def generate_uid(self, analysis_result):
        import hashlib
        
        key_string = (
            f"{analysis_result['start_time']}|"
            f"{analysis_result['end_time']}|"
            f"{analysis_result['dungeon_name']}|"
            f"{analysis_result['black_person']}|"
            f"{analysis_result['worker']}|"
            f"{analysis_result['team_total_salary']}|"
            f"{analysis_result['personal_salary']}|"
            f"{analysis_result['scattered_total']}|"
            f"{analysis_result['iron_total']}|"
            f"{analysis_result['other_total']}|"
            f"{analysis_result['special_total']}|"
            f"{analysis_result['note']}"
        )
        
        hash_object = hashlib.md5(key_string.encode('utf-8'))
        return hash_object.hexdigest()[:8]

    def load_filled_uids(self):
        try:
            result = self.main_app.db.execute_query("SELECT uid FROM filled_uids")
            if result:
                self.filled_uids = {row[0] for row in result}
        except Exception:
            self.create_filled_uids_table()

    def create_filled_uids_table(self):
        try:
            self.main_app.db.execute_update('''
                CREATE TABLE IF NOT EXISTS filled_uids (
                    uid TEXT PRIMARY KEY,
                    fill_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
        except Exception:
            pass

    def save_filled_uid(self, uid):
        try:
            self.main_app.db.execute_update(
                "INSERT OR IGNORE INTO filled_uids (uid) VALUES (?)",
                (uid,)
            )
            self.filled_uids.add(uid)
        except Exception:
            pass

    def load_folder_list(self):
        try:
            result = self.main_app.db.execute_query("SELECT file_path, remark FROM analysis_files")
            if result:
                self.db_folders = {}
                
                folders = {}
                for row in result:
                    file_path, remark = row
                    if remark.startswith("FOLDER:"):
                        folder_path = file_path
                        actual_remark = remark.replace("FOLDER:", "")
                        folders[folder_path] = (actual_remark, [])
                
                for row in result:
                    file_path, remark = row
                    if remark.startswith("FILE:"):
                        folder_path = os.path.dirname(file_path)
                        if folder_path in folders:
                            actual_remark = remark.replace("FILE:", "")
                            folders[folder_path][1].append(file_path)
                
                for folder_path, (remark, file_list) in folders.items():
                    if file_list:
                        self.db_folders[folder_path] = (remark, file_list)
                
                self.refresh_treeview()
        except Exception:
            self.db_folders = {}

    def save_folder_list(self):
        try:
            self.main_app.db.execute_update("DELETE FROM analysis_files")
            
            for folder_path, (remark, file_list) in self.db_folders.items():
                self.main_app.db.execute_update(
                    "INSERT INTO analysis_files (file_path, remark) VALUES (?, ?)",
                    (folder_path, f"FOLDER:{remark}")
                )
                
                for file_path in file_list:
                    self.main_app.db.execute_update(
                        "INSERT INTO analysis_files (file_path, remark) VALUES (?, ?)",
                        (file_path, f"FILE:{remark}")
                    )
            
            messagebox.showinfo("成功", "文件夹列表已保存到数据库")
        except Exception as e:
            messagebox.showerror("错误", f"保存文件夹列表失败: {str(e)}")
        
    def refresh_treeview(self):
        for item in self.file_treeview.get_children():
            self.file_treeview.delete(item)
        
        for folder_path, (remark, file_list) in self.db_folders.items():
            self.file_treeview.insert("", "end", values=(
                folder_path,
                remark
            ))
    
    def on_treeview_select(self, event):
        selection = self.file_treeview.selection()
        if selection:
            item = selection[0]
            values = self.file_treeview.item(item, "values")
            if values:
                self.remark_entry.delete(0, tk.END)
                self.remark_entry.insert(0, values[1])
        
    def add_folder(self):
        folder_path = filedialog.askdirectory(
            title="选择包含.db文件的文件夹"
        )
        
        if not folder_path:
            return
            
        if folder_path in self.db_folders:
            messagebox.showwarning("警告", "该文件夹已添加")
            return
            
        db_files = self.scan_folder_for_db_files(folder_path)
            
        if not db_files:
            messagebox.showwarning("警告", "该文件夹中没有找到.db文件")
            return
            
        remark = self.remark_entry.get()
        self.db_folders[folder_path] = (remark, db_files)
        
        self.refresh_treeview()
        self.remark_entry.delete(0, tk.END)
        self.save_folder_list_silent()
        
        messagebox.showinfo("成功", f"已添加文件夹，找到 {len(db_files)} 个.db文件")
    
    def remove_folder(self):
        selection = self.file_treeview.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个文件夹")
            return
            
        item = selection[0]
        folder_path = self.file_treeview.item(item, 'values')[0]
        
        if folder_path in self.db_folders:
            del self.db_folders[folder_path]
            self.refresh_treeview()
            self.save_folder_list_silent()
    
    def clear_folders(self):
        if not self.db_folders:
            return
            
        if messagebox.askyesno("确认", "确定要清空所有文件夹吗？"):
            self.db_folders = {}
            self.refresh_treeview()
            self.save_folder_list_silent()
    
    def edit_selected_remark(self):
        selection = self.file_treeview.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个文件夹")
            return
            
        item = selection[0]
        folder_path = self.file_treeview.item(item, 'values')[0]
        
        if folder_path not in self.db_folders:
            return
            
        new_remark = self.remark_entry.get() or "打工仔"
        current_remark, file_list = self.db_folders[folder_path]
        
        if new_remark != current_remark:
            self.db_folders[folder_path] = (new_remark, file_list)
            self.refresh_treeview()
            
            children = self.file_treeview.get_children()
            for child in children:
                if self.file_treeview.item(child, 'values')[0] == folder_path:
                    self.file_treeview.selection_set(child)
                    break
            
            self.save_folder_list_silent()
    
    def save_folder_list_silent(self):
        try:
            self.main_app.db.execute_update("DELETE FROM analysis_files")
            
            for folder_path, (remark, file_list) in self.db_folders.items():
                self.main_app.db.execute_update(
                    "INSERT INTO analysis_files (file_path, remark) VALUES (?, ?)",
                    (folder_path, f"FOLDER:{remark}")
                )
                
                for file_path in file_list:
                    self.main_app.db.execute_update(
                        "INSERT INTO analysis_files (file_path, remark) VALUES (?, ?)",
                        (file_path, f"FILE:{remark}")
                    )
        except Exception:
            pass

    def create_empty_result(self, filename, remark):
        return {
            "filename": filename,
            "remark": remark,
            "start_time": "未找到",
            "end_time": "未找到",
            "dungeon_name": "未知副本",
            "black_person": "",
            "worker": remark,
            "team_total_salary": 0,
            "personal_salary": 0,
            "subsidy": 0,
            "penalty_total": 0,
            "scattered_total": 0,
            "iron_total": 0,
            "other_total": 0,
            "special_total": 0,
            "special_items": [],
            "team_type": "未知",
            "lie_count": 0,
            "note": "",
            "scattered_consumption": 0,
            "iron_consumption": 0,
            "special_consumption": 0,
            "other_consumption": 0,
            "total_consumption": 0,
            "uid": "empty"
        }

    def add_result_to_tree(self, result):
        consumption_total = (
            result.get("scattered_consumption", 0) + 
            result.get("iron_consumption", 0) + 
            result.get("special_consumption", 0) + 
            result.get("other_consumption", 0)
        )
        
        self.result_tree.insert("", "end", values=(
            result["uid"],
            result["start_time"],
            result["end_time"],
            result["dungeon_name"],
            result["black_person"],
            result["worker"],
            f"{result['team_total_salary']}金",
            f"{result['personal_salary']}金",
            f"{consumption_total}金",
            f"{result['subsidy']}金",
            f"{result['penalty_total']}金",
            f"{result['scattered_total']}金",
            f"{result['iron_total']}金",
            f"{result['other_total']}金",
            f"{result['special_total']}金",
            result["team_type"],
            result["lie_count"],
            result["note"]
        ))

    def start_analysis(self):
        if not self.db_folders:
            messagebox.showwarning("警告", "请先添加包含.db文件的文件夹")
            return
        
        self.update_progress(0, "开始扫描文件夹...")
        
        updated_folders = {}
        total_files = 0
        
        for folder_path, (remark, old_file_list) in self.db_folders.items():
            self.update_progress(10, f"扫描文件夹: {os.path.basename(folder_path)}")
            
            new_file_list = self.scan_folder_for_db_files(folder_path)
            updated_folders[folder_path] = (remark, new_file_list)
            total_files += len(new_file_list)
        
        self.db_folders = updated_folders
        self.save_folder_list_silent()
        
        if total_files == 0:
            messagebox.showwarning("警告", "所有文件夹中都没有找到.db文件")
            self.update_progress(0, "没有找到.db文件")
            return
        
        self.update_progress(60, "开始分析所有.db文件")
        
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.analysis_results = []
        
        success_count = 0
        duplicate_count = 0
        seen_uids = set()
        
        processed_files = 0
        
        for folder_path, (remark, file_list) in self.db_folders.items():
            for db_file in file_list:
                try:
                    processed_files += 1
                    progress = 10 + (processed_files / total_files) * 80
                    
                    self.update_progress(
                        progress, 
                        f"分析进度: {processed_files}/{total_files} - {os.path.basename(db_file)}"
                    )
                    
                    results = self.analyze_db_file_optimized(db_file, remark)
                    if results:
                        for result in results:
                            uid = result["uid"]
                            
                            if uid in seen_uids or uid in self.filled_uids:
                                duplicate_count += 1
                                continue
                            
                            self.analysis_results.append(result)
                            self.add_result_to_tree(result)
                            seen_uids.add(uid)
                            success_count += 1
                            
                except Exception:
                    pass
        
        if success_count > 0:
            messagebox.showinfo("完成", f"分析完成！成功分析{success_count}个记录段")
        else:
            messagebox.showwarning("警告", "没有成功分析任何记录段")
        
        self.update_progress(0, "分析完成")

    def fill_form(self):
        selected = self.result_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一条分析结果")
            return
        
        item = selected[0]
        values = self.result_tree.item(item, 'values')
        uid = values[0]
        
        result = next((r for r in self.analysis_results if r.get('uid') == uid), None)
        if not result:
            messagebox.showerror("错误", "找不到对应的分析结果")
            return
        
        try:
            self.main_app.dungeon_var.set(result.get("dungeon_name", ""))
            self.main_app.trash_gold_var.set(str(result.get("scattered_total", 0)))
            self.main_app.iron_gold_var.set(str(result.get("iron_total", 0)))
            self.main_app.other_gold_var.set(str(result.get("other_total", 0)))
            self.main_app.total_gold_var.set(str(result.get("team_total_salary", 0)))
            self.main_app.personal_gold_var.set(str(result.get("personal_salary", 0)))
            self.main_app.subsidy_gold_var.set(str(result.get("subsidy", 0)))
            self.main_app.fine_gold_var.set(str(result.get("penalty_total", 0)))
            self.main_app.team_type_var.set(result.get("team_type", ""))
            self.main_app.lie_down_var.set(str(result.get("lie_count", 0)))
            self.main_app.black_owner_var.set(result.get("black_person", ""))
            self.main_app.worker_var.set(result.get("worker", ""))
            self.main_app.note_var.set(result.get("note", ""))
            
            scattered_consumption = result.get("scattered_consumption", 0)
            iron_consumption = result.get("iron_consumption", 0)
            special_consumption = result.get("special_consumption", 0)
            other_consumption = result.get("other_consumption", 0)
            total_consumption = result.get("total_consumption", 0)
            
            self.main_app.scattered_consumption_var.set(str(scattered_consumption))
            self.main_app.iron_consumption_var.set(str(iron_consumption))
            self.main_app.special_consumption_var.set(str(special_consumption))
            self.main_app.other_consumption_var.set(str(other_consumption))
            self.main_app.total_consumption_var.set(str(total_consumption))
            
            self.main_app.special_tree.clear()
            for item_data in result.get("special_items", []):
                self.main_app.special_tree.add_item(item_data.get("item", ""), item_data.get("price", 0))

            special_total = sum(item_data.get("price", 0) for item_data in result.get("special_items", []))
            self.main_app.special_total_var.set(str(special_total))
            
            self.save_filled_uid(uid)
            self.result_tree.delete(item)
            self.analysis_results = [r for r in self.analysis_results if r.get('uid') != uid]
            
            messagebox.showinfo("成功", "分析结果已填充到表单，该记录已从列表中移除")
            
        except Exception as e:
            messagebox.showerror("错误", f"填充表单时出错: {str(e)}")

    def is_special_item_match(self, item_name, special_item):
        clean_special = re.sub(r'（.*?）', '', special_item).strip()
        return clean_special in item_name

    def parse_gold_amount(self, gold_text):
        total = 0
        
        brick_match = re.search(r'(\d+)金砖', gold_text)
        if brick_match:
            total += int(brick_match.group(1)) * 10000
        
        gold_match = re.search(r'(\d+)金(?!砖)', gold_text)
        if gold_match:
            total += int(gold_match.group(1))
        
        return total

    def is_potential_special_item(self, item_name):
        all_special_items = self.load_special_items()
        for special_item in all_special_items:
            if self.is_special_item_match(item_name, special_item):
                return True
        return False

    def load_special_items(self):
        special_items = []
        try:
            result = self.main_app.db.execute_query("SELECT special_drops FROM dungeons")
            for row in result:
                if row[0]:
                    items = [item.strip() for item in row[0].split(',')]
                    special_items.extend(items)
        except Exception:
            pass
        
        return special_items

class JX3DungeonTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("反馈Q群：923399567")
        
        self.initialize_all_attributes()
        
        self.root.withdraw()
        
        self.show_splash_screen()
        
        try:
            locale.setlocale(locale.LC_TIME, '')
        except:
            pass 
        
        app_data_dir = get_app_data_path()
        db_path = os.path.join(app_data_dir, 'jx3_dungeon.db')
        os.makedirs(app_data_dir, exist_ok=True)
        
        try:
            self.db_initialized = False
            self.db_error = None
            self.init_db_in_background(db_path)
        except Exception as e:
            self.hide_splash_screen()
            messagebox.showerror("数据库错误", f"无法初始化数据库: {str(e)}")
            self.root.destroy()
            return
        
        self.setup_basic_ui()
        
        self.wait_for_db_init()
        
        if self.db_error:
            self.hide_splash_screen()
            messagebox.showerror("数据库错误", f"数据库初始化失败: {self.db_error}")
            self.root.destroy()
            return
        
        self.complete_ui_setup()
        
        self.after_ids = []
        self.new_record_ids = set()
        self.cached_dungeons = None
        self.cached_owners = None
        self.cached_workers = None
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.schedule_time_update()
        self.setup_pane_events()
        self.setup_window_tracking()

    def initialize_all_attributes(self):
        self.record_tree = None
        self.worker_stats_tree = None
        self.dungeon_tree = None
        self.weekly_tree = None
        self.record_pane = None
        self.fig = None
        self.ax = None
        self.canvas = None
        self.db_analyzer = None
        self.trash_gold_entry = None
        self.iron_gold_entry = None
        self.other_gold_entry = None
        self.subsidy_gold_entry = None
        self.fine_gold_entry = None
        self.scattered_consumption_entry = None
        self.iron_consumption_entry = None
        self.special_consumption_entry = None
        self.other_consumption_entry = None
        self.total_consumption_entry = None
        self.lie_down_entry = None
        self.total_gold_entry = None
        self.personal_gold_entry = None
        self.note_entry = None
        self.dungeon_combo = None
        self.special_item_combo = None
        self.special_price_entry = None
        self.team_type_combo = None
        self.black_owner_combo = None
        self.worker_combo = None
        self.search_owner_combo = None
        self.search_worker_combo = None
        self.search_dungeon_combo = None
        self.search_item_combo = None
        self.search_team_type_combo = None
        self.start_date_entry = None
        self.end_date_entry = None
        self.weekly_worker_combo = None
        self.add_btn = None
        self.edit_btn = None
        self.update_btn = None
        self.special_tree = None
        self.current_edit_id = None
        self.after_ids = []
        self.new_record_ids = set()
        self.cached_dungeons = None
        self.cached_owners = None
        self.cached_workers = None
        self.db_initialized = False
        self.db_error = None

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

    def validate_and_save(self):
        if not self.dungeon_var.get():
            messagebox.showwarning("提示", "请选择副本")
            return False
            
        if not self.total_gold_var.get().isdigit() or int(self.total_gold_var.get()) <= 0:
            messagebox.showwarning("提示", "请输入有效的团队总工资")
            return False
            
        self.add_record()
        return True

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
        
        try:
            self.db.execute_update('''
                INSERT INTO records (
                    dungeon_id, trash_gold, iron_gold, other_gold, special_auctions, 
                    total_gold, black_owner, worker, time, team_type, lie_down_count, 
                    fine_gold, subsidy_gold, personal_gold, note, is_new,
                    scattered_consumption, iron_consumption, special_consumption, other_consumption, total_consumption
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?, ?, ?)
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
                self.note_var.get() or "",
                GoldCalculator.safe_int(self.scattered_consumption_var.get()),
                GoldCalculator.safe_int(self.iron_consumption_var.get()),
                GoldCalculator.safe_int(self.special_consumption_var.get()),
                GoldCalculator.safe_int(self.other_consumption_var.get()),
                GoldCalculator.safe_int(self.total_consumption_var.get())
            ))
            
            self.load_recent_records(50)
            self.update_stats()
            self.load_black_owner_options()
            
            messagebox.showinfo("成功", "记录已添加")
            self.clear_form()
            self.load_weekly_data()
            
        except Exception as e:
            messagebox.showerror("错误", f"保存记录时出错: {str(e)}")

    def edit_record(self):
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
                personal_gold = ?, note = ?,
                scattered_consumption = ?, iron_consumption = ?, special_consumption = ?, other_consumption = ?, total_consumption = ?
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
            GoldCalculator.safe_int(self.scattered_consumption_var.get()),
            GoldCalculator.safe_int(self.iron_consumption_var.get()),
            GoldCalculator.safe_int(self.special_consumption_var.get()),
            GoldCalculator.safe_int(self.other_consumption_var.get()),
            GoldCalculator.safe_int(self.total_consumption_var.get()),
            self.current_edit_id
        ))
        
        self.load_recent_records(50)
        self.update_stats()
        self.load_black_owner_options()
        messagebox.showinfo("成功", "记录已更新")
        self.clear_form()
        self.load_weekly_data()

    def clear_form(self):
        self.dungeon_var.set("")
        self.trash_gold_var.set("")
        self.iron_gold_var.set("")
        self.other_gold_var.set("")
        self.fine_gold_var.set("")
        self.subsidy_gold_var.set("")
        self.lie_down_var.set("")
        self.team_type_var.set("十人本")
        self.total_gold_var.set("")
        self.personal_gold_var.set("")
        self.black_owner_var.set("")
        self.worker_var.set("")
        self.note_var.set("")
        self.special_tree.clear()
        self.special_total_var.set("")
        self.special_item_var.set("")
        self.special_price_var.set("")
        self.scattered_consumption_var.set("")
        self.iron_consumption_var.set("")
        self.special_consumption_var.set("")
        self.other_consumption_var.set("")
        self.total_consumption_var.set("")
        
        self.add_btn.configure(state=tk.NORMAL)
        self.edit_btn.configure(state=tk.NORMAL)
        self.update_btn.configure(state=tk.DISABLED)
        if hasattr(self, 'current_edit_id'):
            del self.current_edit_id

    def on_dungeon_select(self, event):
        selected = self.dungeon_var.get()
        if not selected:
            return
            
        try:
            result = self.db.execute_query("SELECT special_drops FROM dungeons WHERE name=?", (selected,))
            if result and result[0][0]:
                items = [item.strip() for item in result[0][0].split(',')]
                if hasattr(self, 'special_item_combo') and self.special_item_combo:
                    self.special_item_combo['values'] = items
                    self.special_item_var.set("")
            else:
                if hasattr(self, 'special_item_combo') and self.special_item_combo:
                    self.special_item_combo['values'] = []
                    self.special_item_var.set("")
        except Exception as e:
            print(f"加载副本特殊物品时出错: {e}")

    def update_total_consumption(self, *args):
        scattered = self.scattered_consumption_var.get()
        iron = self.iron_consumption_var.get()
        special = self.special_consumption_var.get()
        other = self.other_consumption_var.get()
        
        total = (
            GoldCalculator.safe_int(scattered) + 
            GoldCalculator.safe_int(iron) + 
            GoldCalculator.safe_int(special) + 
            GoldCalculator.safe_int(other)
        )
        self.total_consumption_var.set(str(total))

    def validate_numeric_input(self, new_value):
        return new_value == "" or new_value.isdigit()

    def show_splash_screen(self):
        self.splash = tk.Toplevel(self.root)
        self.splash.title("正在启动...")
        self.splash.geometry("400x200")
        self.splash.configure(bg='#f0f0f0')
        self.splash.overrideredirect(True)
        self.splash.attributes('-topmost', True)
        
        screen_width = self.splash.winfo_screenwidth()
        screen_height = self.splash.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 200) // 2
        self.splash.geometry(f"400x200+{x}+{y}")
        
        ttk.Label(self.splash, text="JX3DungeonTracker", 
                 font=("PingFang SC", 18, "bold"), background='#f0f0f0').pack(pady=20)
        ttk.Label(self.splash, text="剑网3副本记录工具", 
                 font=("PingFang SC", 12), background='#f0f0f0').pack()
        
        ttk.Label(self.splash, text="正在初始化数据库和界面...", 
                 font=("PingFang SC", 10), background='#f0f0f0').pack(pady=10)
        
        self.splash_progress = ttk.Progressbar(self.splash, mode='indeterminate', length=300)
        self.splash_progress.pack(pady=20)
        self.splash_progress.start()
        
        self.splash.update()
        self.root.update()

    def hide_splash_screen(self):
        if hasattr(self, 'splash') and self.splash:
            self.splash.destroy()
            self.splash = None
        
        self.root.deiconify()

    def init_db_in_background(self, db_path):
        def init_db():
            try:
                time.sleep(3)
                
                self.db = DatabaseManager(db_path)
                
                self.db_initialized = True
                
                self.root.after(500, self.stage_2_data_loading)
                
            except Exception as e:
                self.db_error = str(e)
        
        self.db_thread = threading.Thread(target=init_db, daemon=True)
        self.db_thread.start()

    def stage_2_data_loading(self):
        if not self.db_initialized:
            self.root.after(500, self.stage_2_data_loading)
            return
            
        self.load_dungeon_options()
        self.load_black_owner_options()
        
        self.root.after(1000, self.stage_3_data_loading)

    def stage_3_data_loading(self):
        try:
            self.load_recent_records(50)
            self.root.after(1000, self.stage_4_data_loading)
        except Exception as e:
            print(f"加载记录数据时出错: {e}")
            self.root.after(1000, self.stage_4_data_loading)

    def stage_4_data_loading(self):
        try:
            self.update_stats()
            threading.Thread(target=self.load_remaining_records_background, daemon=True).start()
        except Exception as e:
            print(f"加载统计信息时出错: {e}")

    def wait_for_db_init(self):
        max_wait_time = 15
        start_time = time.time()
        
        def update_splash_status(status):
            if hasattr(self, 'splash') and self.splash:
                for widget in self.splash.winfo_children():
                    if isinstance(widget, ttk.Label) and "正在初始化" in widget.cget("text"):
                        widget.config(text=status)
                        break
        
        while not self.db_initialized and not self.db_error:
            elapsed = time.time() - start_time
            if elapsed > max_wait_time:
                self.db_error = "数据库初始化超时"
                break
            
            if elapsed < 2:
                update_splash_status("正在初始化数据库...")
            elif elapsed < 5:
                update_splash_status("正在加载预设数据...")
            elif elapsed < 8:
                update_splash_status("正在准备界面组件...")
            else:
                update_splash_status("正在完成初始化，请稍候...")
            
            self.splash.update()
            time.sleep(0.1)

    def setup_basic_ui(self):
        self.setup_window()
        self.setup_variables()
        self.setup_styles()
        self.create_main_ui()

    def setup_window(self):
        width, height = int(1700*SCALE_FACTOR), int(1000*SCALE_FACTOR)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        self.root.configure(bg="#f5f5f7")
        self.root.minsize(int(1024*SCALE_FACTOR), int(600*SCALE_FACTOR))

    def setup_variables(self):
        self.trash_gold_var = tk.StringVar(value="")
        self.iron_gold_var = tk.StringVar(value="")
        self.other_gold_var = tk.StringVar(value="")
        self.fine_gold_var = tk.StringVar(value="")
        self.subsidy_gold_var = tk.StringVar(value="")
        self.lie_down_var = tk.StringVar(value="")
        self.team_type_var = tk.StringVar(value="十人本")
        self.total_gold_var = tk.StringVar(value="")
        self.personal_gold_var = tk.StringVar(value="")
        self.special_total_var = tk.StringVar(value="")
        self.note_var = tk.StringVar(value="")
        self.start_date_var = tk.StringVar(value="")
        self.end_date_var = tk.StringVar(value="")
        self.time_var = tk.StringVar(value=get_current_time())
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
        self.weekly_worker_var = tk.StringVar()
        self.weekly_period_var = tk.StringVar()
        self.scattered_consumption_var = tk.StringVar(value="")
        self.iron_consumption_var = tk.StringVar(value="")
        self.special_consumption_var = tk.StringVar(value="")
        self.other_consumption_var = tk.StringVar(value="")
        self.total_consumption_var = tk.StringVar(value="")

    def setup_styles(self):
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
        self.root.option_add("*TCombobox*Listbox*Font", ("PingFang SC", int(10*SCALE_FACTOR)))

    def create_main_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=int(10*SCALE_FACTOR), pady=int(10*SCALE_FACTOR))
        
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, int(8*SCALE_FACTOR)))
        title_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, weight=0)
        
        ttk.Label(title_frame, text="反馈Q群：923399567", 
                 font=("PingFang SC", int(16*SCALE_FACTOR), "bold"), anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=int(10*SCALE_FACTOR))
        
        ttk.Label(title_frame, textvariable=self.time_var, 
                 font=("PingFang SC", int(12*SCALE_FACTOR)), anchor="e"
        ).grid(row=0, column=1, sticky="e")
        
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=int(5*SCALE_FACTOR), pady=int(5*SCALE_FACTOR))
        
        self.record_frame = ttk.Frame(notebook)
        self.stats_frame = ttk.Frame(notebook)
        self.preset_frame = ttk.Frame(notebook)
        self.weekly_frame = ttk.Frame(notebook)
        self.analysis_frame = ttk.Frame(notebook)
        
        notebook.add(self.record_frame, text="副本记录")
        notebook.add(self.stats_frame, text="数据总览")
        notebook.add(self.preset_frame, text="副本预设")
        notebook.add(self.weekly_frame, text="秘境记录")
        notebook.add(self.analysis_frame, text="拍团分析")
        
        loading_label = ttk.Label(self.record_frame, text="正在加载数据，请稍候...", 
                                 font=("PingFang SC", 12))
        loading_label.pack(expand=True)

    def complete_ui_setup(self):
        for widget in self.record_frame.winfo_children():
            widget.destroy()
        
        self.create_record_tab(self.record_frame)
        
        self.root.after(500, self.delayed_ui_setup)
        
        self.setup_events()
        self.hide_splash_screen()
        
        self.root.after(2000, self.ensure_ui_loaded)

    def ensure_ui_loaded(self):
        if hasattr(self, 'dungeon_combo') and self.dungeon_combo:
            self.load_dungeon_options()
            self.load_black_owner_options()
            self.load_recent_records(30)
            self.update_stats()
        else:
            self.root.after(1000, self.ensure_ui_loaded)

    def delayed_ui_setup(self):
        try:
            self.create_stats_tab(self.stats_frame)
            self.root.after(500, lambda: self.create_preset_tab(self.preset_frame))
            self.root.after(1000, lambda: self.create_weekly_tab(self.weekly_frame))
            self.root.after(1500, lambda: self.create_analysis_tab(self.analysis_frame))
            self.root.after(2000, self.safe_load_column_widths)
        except Exception as e:
            print(f"延迟UI设置出错: {e}")

    def start_staggered_loading(self):
        self.load_dungeon_options()
        self.load_black_owner_options()
        self.root.after(1000, lambda: self.load_recent_records(30))
        self.root.after(3000, self.update_stats)
        self.root.after(5000, self.start_background_loading)

    def start_background_loading(self):
        threading.Thread(target=self.background_data_loading, daemon=True).start()

    def background_data_loading(self):
        try:
            self.load_remaining_records_background()
            self.load_weekly_data()
            self.load_dungeon_presets()
        except Exception as e:
            print(f"后台数据加载出错: {e}")

    def create_record_tab(self, parent):
        try:
            pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
            pane.pack(fill=tk.BOTH, expand=True, padx=int(5*SCALE_FACTOR), pady=int(5*SCALE_FACTOR))
            
            form_frame = ttk.LabelFrame(pane, text="副本记录管理", padding=int(8*SCALE_FACTOR), width=int(350*SCALE_FACTOR))
            self.build_record_form(form_frame)
            
            list_frame = ttk.LabelFrame(pane, text="副本记录列表", padding=int(8*SCALE_FACTOR))
            self.build_record_list(list_frame)
            
            pane.add(form_frame, weight=1)
            pane.add(list_frame, weight=2)
            
            self.record_pane = pane
            self.record_pane_name = "record_pane"
            
            self.restore_pane_position(self.record_pane, self.record_pane_name)
            
        except Exception as e:
            print(f"创建记录标签页时出错: {e}")
            raise

    def build_record_form(self, parent):
        dungeon_row = ttk.Frame(parent)
        dungeon_row.pack(fill=tk.X, pady=int(3*SCALE_FACTOR))
        ttk.Label(dungeon_row, text="副本名称:").pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        self.dungeon_combo = ttk.Combobox(dungeon_row, textvariable=self.dungeon_var, 
                                         width=int(25*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.dungeon_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        special_frame = ttk.LabelFrame(parent, text="特殊掉落", padding=int(6*SCALE_FACTOR))
        special_frame.pack(fill=tk.BOTH, pady=(0, int(5*SCALE_FACTOR)), expand=True)
        
        tree_frame = ttk.Frame(special_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=int(3*SCALE_FACTOR))
        self.special_tree = SpecialItemsTree(tree_frame)
        
        add_special_frame = ttk.Frame(special_frame)
        add_special_frame.pack(fill=tk.X, pady=(int(8*SCALE_FACTOR), 0))
        
        ttk.Label(add_special_frame, text="物品:").grid(row=0, column=0, padx=(0, int(5*SCALE_FACTOR)), sticky="w")
        self.special_item_combo = ttk.Combobox(add_special_frame, textvariable=self.special_item_var, 
                                              width=int(22*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.special_item_combo.grid(row=0, column=1, padx=(0, int(5*SCALE_FACTOR)), sticky="ew")
        
        ttk.Label(add_special_frame, text="金额:").grid(row=0, column=2, padx=(int(5*SCALE_FACTOR), int(5*SCALE_FACTOR)), sticky="w")
        self.special_price_entry = ttk.Entry(add_special_frame, textvariable=self.special_price_var, 
                                            width=int(7*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.special_price_entry.grid(row=0, column=3, padx=(0, int(5*SCALE_FACTOR)), sticky="w")
        
        ttk.Button(add_special_frame, text="添加", width=int(6*SCALE_FACTOR), command=self.add_special_item
        ).grid(row=0, column=4, padx=(int(5*SCALE_FACTOR), 0))
        add_special_frame.columnconfigure(1, weight=1)
        
        team_frame = ttk.LabelFrame(parent, text="团队项目", padding=int(6*SCALE_FACTOR))
        team_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_team_fields(team_frame)
        
        personal_frame = ttk.LabelFrame(parent, text="个人项目", padding=int(6*SCALE_FACTOR))
        personal_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_personal_fields(personal_frame)
        
        info_frame = ttk.LabelFrame(parent, text="团队信息", padding=int(6*SCALE_FACTOR))
        info_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_info_fields(info_frame)
        
        gold_frame = ttk.LabelFrame(parent, text="工资信息", padding=int(6*SCALE_FACTOR))
        gold_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_gold_fields(gold_frame)
        
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=int(5*SCALE_FACTOR))
        self.build_form_buttons(btn_frame)

    def build_team_fields(self, parent):
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
        ttk.Label(parent, text="补贴金额:").grid(row=0, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.subsidy_gold_entry = ttk.Entry(parent, textvariable=self.subsidy_gold_var, 
                                        width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.subsidy_gold_entry.grid(row=0, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="罚款金额:").grid(row=0, column=2, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.fine_gold_entry = ttk.Entry(parent, textvariable=self.fine_gold_var, 
                                        width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.fine_gold_entry.grid(row=0, column=3, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="散件消费:").grid(row=1, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.scattered_consumption_entry = ttk.Entry(parent, textvariable=self.scattered_consumption_var, 
                                                width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.scattered_consumption_entry.grid(row=1, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="小铁消费:").grid(row=1, column=2, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.iron_consumption_entry = ttk.Entry(parent, textvariable=self.iron_consumption_var, 
                                            width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.iron_consumption_entry.grid(row=1, column=3, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="特殊消费:").grid(row=2, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.special_consumption_entry = ttk.Entry(parent, textvariable=self.special_consumption_var, 
                                                width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.special_consumption_entry.grid(row=2, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="其他消费:").grid(row=2, column=2, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.other_consumption_entry = ttk.Entry(parent, textvariable=self.other_consumption_var, 
                                            width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.other_consumption_entry.grid(row=2, column=3, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="总消费:").grid(row=3, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.total_consumption_entry = ttk.Entry(parent, textvariable=self.total_consumption_var, 
                                            width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.total_consumption_entry.grid(row=3, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)

    def build_info_fields(self, parent):
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
                                   width=int(30*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.note_entry.grid(row=3, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W, columnspan=3)

    def build_gold_fields(self, parent):
        ttk.Label(parent, text="团队总工资:").grid(row=0, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.total_gold_entry = ttk.Entry(parent, textvariable=self.total_gold_var, 
                                        width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.total_gold_entry.grid(row=0, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        
        ttk.Label(parent, text="个人工资:").grid(row=1, column=0, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)
        self.personal_gold_entry = ttk.Entry(parent, textvariable=self.personal_gold_var, 
                                            width=int(10*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.personal_gold_entry.grid(row=1, column=1, padx=int(3*SCALE_FACTOR), pady=int(2*SCALE_FACTOR), sticky=tk.W)

    def build_form_buttons(self, parent):
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
        search_frame = ttk.Frame(parent)
        search_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, int(5*SCALE_FACTOR)))
        self.build_search_controls(search_frame)
        
        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        
        columns = ("row_num", "dungeon", "time", "team_type", "lie_down", "total", "personal", "consumption", "black_owner", "worker", "note")
        self.record_tree = ttk.Treeview(parent, columns=columns, show="headings", 
                                    selectmode="extended", height=int(10*SCALE_FACTOR))
        
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
        self.setup_context_menu()
        
        self.record_tree.tag_configure('new_record', background='#e6f7ff')

    def build_search_controls(self, parent):
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
        columns = [
            ("row_num", "序号", int(35*SCALE_FACTOR)),
            ("dungeon", "副本名称", int(100*SCALE_FACTOR)),
            ("time", "时间", int(100*SCALE_FACTOR)),
            ("team_type", "团队类型", int(70*SCALE_FACTOR)),
            ("lie_down", "躺拍人数", int(70*SCALE_FACTOR)),
            ("total", "团队总工资", int(100*SCALE_FACTOR)),
            ("personal", "个人工资", int(100*SCALE_FACTOR)),
            ("consumption", "本场消费", int(100*SCALE_FACTOR)),
            ("black_owner", "黑本", int(70*SCALE_FACTOR)),
            ("worker", "打工仔", int(70*SCALE_FACTOR)),
            ("note", "备注", int(100*SCALE_FACTOR))
        ]
        
        for col_id, heading, width in columns:
            self.record_tree.heading(col_id, text=heading, anchor="center")
            self.record_tree.column(col_id, width=width, anchor=tk.CENTER, stretch=(col_id == "note"))

    def setup_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="删除记录", command=self.delete_selected_records)
        self.record_tree.bind("<Button-3>", self.show_record_context_menu)

    def create_stats_tab(self, parent):
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
        error_frame = ttk.Frame(parent)
        error_frame.pack(fill=tk.BOTH, expand=True, pady=int(20*SCALE_FACTOR))
        ttk.Label(error_frame, 
                 text="需要安装matplotlib和numpy库才能显示统计图表\n请运行: pip install matplotlib numpy", 
                 foreground="red", font=("PingFang SC", int(10*SCALE_FACTOR)), anchor="center", justify="center"
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
            card = ttk.LabelFrame(parent, text=title, padding=(int(10*SCALE_FACTOR), int(8*SCALE_FACTOR)))
            card.pack(fill=tk.X, pady=int(5*SCALE_FACTOR))
            ttk.Label(card, textvariable=var, font=("PingFang SC", int(14*SCALE_FACTOR), "bold"), anchor="center"
            ).pack(fill=tk.BOTH, expand=True)

    def build_chart_area(self, parent):
        chart_frame = ttk.LabelFrame(parent, text="每周数据统计", padding=(int(8*SCALE_FACTOR), int(6*SCALE_FACTOR)))
        chart_frame.pack(fill=tk.BOTH, expand=True)
        
        if MATPLOTLIB_AVAILABLE:
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'KaiTi', 'Arial Unicode MS']
            plt.rcParams['axes.unicode_minus'] = False
        
            self.fig, self.ax = plt.subplots(figsize=(int(6*SCALE_FACTOR), int(2.5*SCALE_FACTOR)), dpi=100)
            self.fig.patch.set_facecolor('#f5f5f7')
            self.ax.set_facecolor('#f5f5f7')
        
            plt.rcParams['axes.titlesize'] = int(12*SCALE_FACTOR)
            plt.rcParams['axes.labelsize'] = int(10*SCALE_FACTOR)
            plt.rcParams['xtick.labelsize'] = int(8*SCALE_FACTOR)
            plt.rcParams['ytick.labelsize'] = int(8*SCALE_FACTOR)
            plt.rcParams['legend.fontsize'] = int(10*SCALE_FACTOR)
        
            self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def build_worker_stats(self, parent):
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
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=int(5*SCALE_FACTOR), pady=int(5*SCALE_FACTOR))
        
        list_frame = ttk.LabelFrame(main_frame, text="副本列表", padding=int(8*SCALE_FACTOR))
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, int(5*SCALE_FACTOR)))
        self.build_dungeon_list(list_frame)
        
        form_frame = ttk.LabelFrame(main_frame, text="副本详情", padding=int(8*SCALE_FACTOR))
        form_frame.pack(fill=tk.X, pady=(0, int(5*SCALE_FACTOR)))
        self.build_dungeon_form(form_frame)
        
        form_frame.configure(height=int(200*SCALE_FACTOR))

    def build_dungeon_list(self, parent):
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

    def create_weekly_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=int(10*SCALE_FACTOR), pady=int(10*SCALE_FACTOR))
        
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, int(10*SCALE_FACTOR)))
        
        ttk.Label(control_frame, text="选择打工仔:").pack(side=tk.LEFT, padx=(0, int(5*SCALE_FACTOR)))
        self.weekly_worker_combo = ttk.Combobox(control_frame, textvariable=self.weekly_worker_var, 
                                            width=int(20*SCALE_FACTOR), font=("PingFang SC", int(10*SCALE_FACTOR)))
        self.weekly_worker_combo.pack(side=tk.LEFT, padx=(0, int(15*SCALE_FACTOR)))
        self.weekly_worker_combo.bind("<<ComboboxSelected>>", self.on_weekly_worker_select)
        
        ttk.Label(control_frame, textvariable=self.weekly_period_var, 
                font=("PingFang SC", int(10*SCALE_FACTOR), "bold")).pack(side=tk.LEFT)
        
        ttk.Button(control_frame, text="刷新数据", command=self.load_weekly_data
                ).pack(side=tk.RIGHT, padx=(int(5*SCALE_FACTOR), 0))
        
        tree_frame = ttk.LabelFrame(main_frame, text="本周秘境记录", padding=int(8*SCALE_FACTOR))
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("worker", "dungeon", "note")
        self.weekly_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", 
                                    height=int(15*SCALE_FACTOR), selectmode="browse")
        
        column_config = [
            ("worker", "打工仔", int(500*SCALE_FACTOR)),
            ("dungeon", "副本", int(500*SCALE_FACTOR)),
            ("note", "备注", int(500*SCALE_FACTOR))
        ]
        
        for col_id, heading, width in column_config:
            self.weekly_tree.heading(col_id, text=heading, anchor="center")
            self.weekly_tree.column(col_id, width=width, anchor=tk.CENTER, stretch=(col_id == "note"))
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.weekly_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.weekly_tree.xview)
        self.weekly_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.weekly_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        self.setup_column_resizing(self.weekly_tree)
        
        self.load_weekly_worker_options()
        self.root.after(1500, self.load_weekly_data)

    def create_analysis_tab(self, parent):
        self.db_analyzer = DBAnalyzer(parent, self)

    def setup_events(self):
        self.scattered_consumption_var.trace_add("write", self.update_total_consumption)
        self.iron_consumption_var.trace_add("write", self.update_total_consumption)
        self.special_consumption_var.trace_add("write", self.update_total_consumption)
        self.other_consumption_var.trace_add("write", self.update_total_consumption)
        
        reg = self.root.register(self.validate_numeric_input)
        vcmd = (reg, '%P')
        for entry in [self.trash_gold_entry, self.iron_gold_entry, self.other_gold_entry, 
                    self.fine_gold_entry, self.subsidy_gold_entry, self.lie_down_entry, 
                    self.total_gold_entry, self.personal_gold_entry, self.scattered_consumption_entry, 
                    self.iron_consumption_entry, self.special_consumption_entry, 
                    self.other_consumption_entry]:
            entry.config(validate="key", validatecommand=vcmd)
        
        self.dungeon_combo.bind("<<ComboboxSelected>>", self.on_dungeon_select)
        self.search_dungeon_combo.bind("<<ComboboxSelected>>", self.on_search_dungeon_select)
        
        self.record_tree.bind('<ButtonRelease-1>', self.on_record_click)
        
        self.root.after(2000, self.safe_load_column_widths)
        self.root.after(1000, self.update_time)

    def safe_load_column_widths(self):
        try:
            self.load_column_widths()
        except AttributeError as e:
            print(f"列宽度加载失败，将在1秒后重试: {e}")
            self.root.after(1000, self.safe_load_column_widths)

    def load_column_widths(self):
        trees = []
        
        if hasattr(self, 'record_tree') and self.record_tree and self.record_tree.winfo_exists():
            trees.append((self.record_tree, "record_tree"))
        
        if hasattr(self, 'worker_stats_tree') and self.worker_stats_tree and self.worker_stats_tree.winfo_exists():
            trees.append((self.worker_stats_tree, "worker_stats_tree"))
        
        if hasattr(self, 'dungeon_tree') and self.dungeon_tree and self.dungeon_tree.winfo_exists():
            trees.append((self.dungeon_tree, "dungeon_tree"))
        
        if hasattr(self, 'weekly_tree') and self.weekly_tree and self.weekly_tree.winfo_exists():
            trees.append((self.weekly_tree, "weekly_tree"))
        
        for tree, tree_name in trees:
            try:
                result = self.db.execute_query("SELECT widths FROM column_widths WHERE tree_name = ?", (tree_name,))
                if result and result[0]:
                    widths = json.loads(result[0][0])
                    for col, width in widths.items():
                        if col in tree["columns"]:
                            tree.column(col, width=width)
            except Exception as e:
                print(f"加载列宽度时出错 ({tree_name}): {e}")

    def save_column_widths(self):
        trees = []
        
        if hasattr(self, 'record_tree') and self.record_tree and self.record_tree.winfo_exists():
            trees.append((self.record_tree, "record_tree"))
        
        if hasattr(self, 'worker_stats_tree') and self.worker_stats_tree and self.worker_stats_tree.winfo_exists():
            trees.append((self.worker_stats_tree, "worker_stats_tree"))
        
        if hasattr(self, 'dungeon_tree') and self.dungeon_tree and self.dungeon_tree.winfo_exists():
            trees.append((self.dungeon_tree, "dungeon_tree"))
        
        if hasattr(self, 'weekly_tree') and self.weekly_tree and self.weekly_tree.winfo_exists():
            trees.append((self.weekly_tree, "weekly_tree"))
        
        for tree, tree_name in trees:
            try:
                widths = {col: tree.column(col, "width") for col in tree["columns"]}
                self.db.execute_update('''
                    INSERT OR REPLACE INTO column_widths (tree_name, widths)
                    VALUES (?, ?)
                ''', (tree_name, json.dumps(widths)))
            except Exception as e:
                print(f"保存列宽度时出错 ({tree_name}): {e}")

    def auto_resize_column(self, tree, column_id):
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
        columns = tree["columns"]
        for col in columns:
            tree.heading(col, command=lambda c=col: self.auto_resize_column(tree, c))

    def on_record_click(self, event):
        item = self.record_tree.identify('item', event.x, event.y)
        if not item:
            return
            
        selected = self.record_tree.selection()
        if len(selected) == 1 and item in selected:
            self.fill_form_from_record(item)

    def fill_form_from_record(self, item):
        values = self.record_tree.item(item, 'values')
        if not values:
            return
            
        dungeon_name = values[1]
        time_str = values[2]
        
        record = self.db.execute_query('''
            SELECT r.id, d.name, r.trash_gold, r.iron_gold, r.other_gold, r.special_auctions, 
                r.total_gold, r.black_owner, r.worker, r.time, r.team_type, r.lie_down_count, 
                r.fine_gold, r.subsidy_gold, r.personal_gold, r.note,
                r.scattered_consumption, r.iron_consumption, r.special_consumption, r.other_consumption, r.total_consumption
            FROM records r
            JOIN dungeons d ON r.dungeon_id = d.id
            WHERE d.name = ? AND r.time LIKE ?
        ''', (dungeon_name, f"{time_str}%"))
        
        if not record:
            return
            
        r = record[0]
        self.dungeon_var.set(r[1])
        self.trash_gold_var.set(str(r[2]))
        self.iron_gold_var.set(str(r[3]))
        self.other_gold_var.set(str(r[4]))
        self.total_gold_var.set(str(r[6]))
        self.black_owner_var.set(r[7] or "")
        self.worker_var.set(r[8] or "")
        self.team_type_var.set(r[10])
        self.lie_down_var.set(str(r[11]))
        self.fine_gold_var.set(str(r[12]))
        self.subsidy_gold_var.set(str(r[13]))
        self.personal_gold_var.set(str(r[14]))
        self.note_var.set(r[15] or "")
        self.scattered_consumption_var.set(str(r[16]))
        self.iron_consumption_var.set(str(r[17]))
        self.special_consumption_var.set(str(r[18]))
        self.other_consumption_var.set(str(r[19]))
        self.total_consumption_var.set(str(r[20]))
        
        self.special_tree.clear()
        special_auctions = json.loads(r[5]) if r[5] else []
        for item_data in special_auctions:
            self.special_tree.add_item(item_data['item'], item_data['price'])
        self.special_total_var.set(str(self.special_tree.calculate_total()))
        
        self.add_btn.configure(state=tk.NORMAL)
        self.edit_btn.configure(state=tk.NORMAL)
        self.update_btn.configure(state=tk.DISABLED)
        
        if hasattr(self, 'current_edit_id'):
            del self.current_edit_id

    def load_dungeon_options(self):
        try:
            self.cached_dungeons = None
            self.cached_dungeons = [row[0] for row in self.db.execute_query(
                "SELECT name FROM dungeons ORDER BY name"
            )]
        except Exception as e:
            print(f"加载副本选项出错: {e}")
            self.cached_dungeons = []
        
        if hasattr(self, 'dungeon_combo') and self.dungeon_combo:
            self.dungeon_combo['values'] = self.cached_dungeons
        
        if hasattr(self, 'search_dungeon_combo') and self.search_dungeon_combo:
            self.search_dungeon_combo['values'] = self.cached_dungeons

    def get_all_special_items(self):
        items = set()
        for row in self.db.execute_query("SELECT special_drops FROM dungeons"):
            if row[0]:
                for item in row[0].split(','):
                    items.add(item.strip())
        return list(items)

    def load_recent_records(self, limit=50):
        for item in self.record_tree.get_children():
            self.record_tree.delete(item)
        
        records = self.db.execute_query('''
            SELECT r.id, 
                COALESCE(d.name, '未知副本') as dungeon_name, 
                strftime('%Y-%m-%d %H:%M', r.time), 
                r.team_type, r.lie_down_count, r.total_gold, 
                r.personal_gold, r.black_owner, r.worker, r.note, r.is_new,
                r.total_consumption
            FROM records r
            LEFT JOIN dungeons d ON r.dungeon_id = d.id
            ORDER BY r.time DESC
            LIMIT ?
        ''', (limit,))
        
        total_records = len(records)
        row_num = total_records
        
        for row in records:
            note = row[9] or ""
            if len(note) > 30:
                note = note[:30] + "..."
            
            tags = ()
            if row[10] == 1:
                tags = ("new_record",)
                self.new_record_ids.add(row[0])
            
            self.record_tree.insert("", "end", values=(
                row_num, row[1], row[2], row[3], row[4], 
                f"{row[5]:,}", f"{row[6]:,}", f"{row[11]:,}",
                row[7], row[8], note
            ), tags=tags)
            row_num -= 1

    def load_remaining_records_background(self):
        try:
            remaining_records = self.db.execute_query('''
                SELECT r.id, 
                    COALESCE(d.name, '未知副本') as dungeon_name, 
                    strftime('%Y-%m-%d %H:%M', r.time), 
                    r.team_type, r.lie_down_count, r.total_gold, 
                    r.personal_gold, r.black_owner, r.worker, r.note, r.is_new,
                    r.total_consumption
                FROM records r
                LEFT JOIN dungeons d ON r.dungeon_id = d.id
                WHERE r.id NOT IN (
                    SELECT id FROM records ORDER BY time DESC LIMIT 50
                )
                ORDER BY r.time DESC
            ''')
            
            self.root.after(0, self.append_records_batch, remaining_records)
            
        except Exception as e:
            print(f"后台加载记录时出错: {e}")

    def append_records_batch(self, records):
        if not records:
            return
            
        batch_size = 20
        current_batch = records[:batch_size]
        remaining_records = records[batch_size:]
        
        start_index = len(self.record_tree.get_children())
        
        for row in current_batch:
            note = row[9] or ""
            if len(note) > 30:
                note = note[:30] + "..."
            
            tags = ()
            if row[10] == 1:
                tags = ("new_record",)
                self.new_record_ids.add(row[0])
            
            self.record_tree.insert("", "end", values=(
                start_index + 1, row[1], row[2], row[3], row[4], 
                f"{row[5]:,}", f"{row[6]:,}", f"{row[11]:,}",
                row[7], row[8], note
            ), tags=tags)
            start_index += 1
        
        if remaining_records:
            self.root.after(500, self.append_records_batch, remaining_records)

    def clear_new_record_highlights(self):
        self.db.execute_update("UPDATE records SET is_new = 0 WHERE is_new = 1")
        for item in self.record_tree.get_children():
            self.record_tree.item(item, tags=())

    def load_dungeon_presets(self):
        self.dungeon_tree.delete(*self.dungeon_tree.get_children())
        for row in self.db.execute_query("SELECT name, special_drops FROM dungeons"):
            self.dungeon_tree.insert("", "end", values=(row[0], row[1]))

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
        try:
            if not self.root.winfo_exists():
                return
                
            self.time_var.set(get_current_time())
            after_id = self.root.after(1000, self.update_time)
            self.after_ids.append(after_id)
        except Exception:
            return

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
                r.personal_gold, r.black_owner, r.worker, r.note, r.is_new,
                r.total_consumption
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
            if len(note) > 30:
                note = note[:30] + "..."
            
            tags = ()
            if row[10] == 1:
                tags = ("new_record",)
            
            self.record_tree.insert("", "end", values=(
                row_num, row[1], row[2], row[3], row[4], 
                f"{row[5]:,}", f"{row[6]:,}", f"{row[11]:,}",
                row[7], row[8], note
            ), tags=tags)
            row_num -= 1

    def validate_date(self, date_str):
        if not date_str:
            return True
        try:
            dt.datetime.strptime(date_str, "%Y-%m-%d")
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
        self.load_recent_records(50)

    def show_record_context_menu(self, event):
        row_id = self.record_tree.identify_row(event.y)
        if not row_id:
            return
        if row_id not in self.record_tree.selection():
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
            
        self.load_recent_records(50)
        self.update_stats()
        self.load_black_owner_options()
        messagebox.showinfo("成功", "记录已删除")
        self.load_weekly_data()

    def update_stats(self):
        total_records = self.db.execute_query(
            "SELECT COUNT(*) FROM records WHERE personal_gold > 0"
        )[0][0] or 0
        self.total_records_var.set(str(total_records))
        
        team_total_gold = self.db.execute_query(
            "SELECT SUM(total_gold) FROM records WHERE personal_gold > 0"
        )[0][0] or 0
        self.team_total_gold_var.set(f"{team_total_gold:,}金")
        
        team_max_gold = self.db.execute_query(
            "SELECT MAX(total_gold) FROM records WHERE personal_gold > 0"
        )[0][0] or 0
        self.team_max_gold_var.set(f"{team_max_gold:,}金")
        
        personal_total_gold = self.db.execute_query(
            "SELECT SUM(personal_gold) FROM records WHERE personal_gold > 0"
        )[0][0] or 0
        self.personal_total_gold_var.set(f"{personal_total_gold:,}金")
        
        personal_max_gold = self.db.execute_query(
            "SELECT MAX(personal_gold) FROM records WHERE personal_gold > 0"
        )[0][0] or 0
        self.personal_max_gold_var.set(f"{personal_max_gold:,}金")
        
        self.update_worker_stats()
        
        if MATPLOTLIB_AVAILABLE:
            self.update_chart()

    def update_worker_stats(self):
        for item in self.worker_stats_tree.get_children():
            self.worker_stats_tree.delete(item)
        
        stats = self.db.execute_query('''
            SELECT worker, COUNT(*), SUM(personal_gold), AVG(personal_gold)
            FROM records 
            WHERE worker IS NOT NULL AND worker != '' AND personal_gold > 0
            GROUP BY worker
            ORDER BY SUM(personal_gold) DESC
        ''')
        
        for worker, count, total, avg in stats:
            self.worker_stats_tree.insert("", "end", values=(
                worker, count, f"{int(total):,}金", f"{int(avg):,}金"
            ))

    def update_chart(self):
        if not MATPLOTLIB_AVAILABLE:
            return
            
        try:
            self.ax.clear()
            
            data = self.db.execute_query('''
                SELECT date(time), SUM(total_gold), SUM(personal_gold), COUNT(*)
                FROM records 
                WHERE personal_gold > 0
                GROUP BY date(time)
                ORDER BY date(time) DESC
                LIMIT 30
            ''')
            
            if not data:
                self.ax.text(0.5, 0.5, '暂无数据', transform=self.ax.transAxes, 
                           ha='center', va='center', fontsize=12)
                self.canvas.draw()
                return
            
            dates = [dt.datetime.strptime(row[0], '%Y-%m-%d') for row in reversed(data)]
            team_gold = [row[1] for row in reversed(data)]
            personal_gold = [row[2] for row in reversed(data)]
            counts = [row[3] for row in reversed(data)]
            
            self.ax.plot(dates, team_gold, 'b.-', label='团队总工资', linewidth=2, markersize=6)
            self.ax.plot(dates, personal_gold, 'r.-', label='个人总工资', linewidth=2, markersize=6)
            
            self.ax.set_xlabel('日期')
            self.ax.set_ylabel('金额（金）')
            self.ax.set_title('近30天工资趋势')
            self.ax.legend()
            self.ax.grid(True, alpha=0.3)
            
            self.ax.xaxis.set_major_locator(mdates.WeekdayLocator())
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
            plt.setp(self.ax.xaxis.get_majorticklabels(), rotation=45)
            
            self.fig.tight_layout()
            self.canvas.draw()
            
        except Exception:
            pass

    def add_dungeon(self):
        self.clear_preset_form()

    def edit_dungeon(self):
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        values = self.dungeon_tree.item(selected[0], 'values')
        self.preset_name_var.set(values[0])
        self.preset_drops_var.set(values[1] or "")

    def delete_dungeon(self):
        selected = self.dungeon_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择一个副本")
            return
            
        values = self.dungeon_tree.item(selected[0], 'values')
        dungeon_name = values[0]
        
        if not messagebox.askyesno("确认", f"确定要删除副本 '{dungeon_name}' 吗？"):
            return
            
        result = self.db.execute_query("SELECT COUNT(*) FROM records WHERE dungeon_id = (SELECT id FROM dungeons WHERE name = ?)", (dungeon_name,))
        record_count = result[0][0] if result else 0
        
        if record_count > 0:
            messagebox.showwarning("警告", f"该副本有 {record_count} 条关联记录，无法删除")
            return
            
        self.db.execute_update("DELETE FROM dungeons WHERE name = ?", (dungeon_name,))
        self.load_dungeon_presets()
        self.load_dungeon_options()
        messagebox.showinfo("成功", "副本已删除")

    def save_dungeon(self):
        name = self.preset_name_var.get().strip()
        drops = self.preset_drops_var.get().strip()
        
        if not name:
            messagebox.showwarning("提示", "请输入副本名称")
            return
            
        try:
            self.db.execute_update('''
                INSERT OR REPLACE INTO dungeons (name, special_drops)
                VALUES (?, ?)
            ''', (name, drops))
            
            self.load_dungeon_presets()
            self.load_dungeon_options()
            messagebox.showinfo("成功", "副本已保存")
            self.clear_preset_form()
            
        except Exception as e:
            messagebox.showerror("错误", f"保存副本时出错: {str(e)}")

    def clear_preset_form(self):
        self.preset_name_var.set("")
        self.preset_drops_var.set("")

    def batch_add_items(self):
        items_text = self.batch_items_var.get().strip()
        if not items_text:
            return
            
        current_drops = self.preset_drops_var.get().strip()
        new_items = []
        
        if ',' in items_text:
            new_items = [item.strip() for item in items_text.split(',') if item.strip()]
        else:
            bracket_items = re.findall(r'\[(.*?)\]', items_text)
            if bracket_items:
                new_items = [item.strip() for item in bracket_items if item.strip()]
            else:
                new_items = [items_text.strip()]
        
        if not new_items:
            return
            
        existing_items = []
        if current_drops:
            existing_items = [item.strip() for item in current_drops.split(',')]
        
        for item in new_items:
            if item not in existing_items:
                existing_items.append(item)
        
        self.preset_drops_var.set(', '.join(existing_items))
        self.batch_items_var.set("")

    def load_black_owner_options(self):
        if self.cached_owners is None:
            self.cached_owners = [row[0] for row in self.db.execute_query(
                "SELECT DISTINCT black_owner FROM records WHERE black_owner IS NOT NULL AND black_owner != '' ORDER BY black_owner"
            )]
        
        self.black_owner_combo['values'] = self.cached_owners
        self.search_owner_combo['values'] = self.cached_owners
        
        if self.cached_workers is None:
            self.cached_workers = [row[0] for row in self.db.execute_query(
                "SELECT DISTINCT worker FROM records WHERE worker IS NOT NULL AND worker != '' ORDER BY worker"
            )]
        
        self.worker_combo['values'] = self.cached_workers
        self.search_worker_combo['values'] = self.cached_workers

    def check_database_integrity(self):
        try:
            self.db.execute_query("PRAGMA integrity_check")
        except Exception as e:
            messagebox.showerror("数据库错误", f"数据库完整性检查失败: {str(e)}")

    def repair_database(self):
        try:
            self.db.execute_query("VACUUM")
            messagebox.showinfo("成功", "数据库修复完成")
        except Exception as e:
            messagebox.showerror("错误", f"数据库修复失败: {str(e)}")

    def import_data(self):
        file_path = filedialog.askopenfilename(
            title="选择要导入的数据文件",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if not isinstance(data, list):
                messagebox.showerror("错误", "数据格式不正确")
                return
            
            success_count = 0
            duplicate_count = 0
            skipped_count = 0
            
            for record in data:
                try:
                    dungeon_name = record.get('dungeon_name')
                    if not dungeon_name:
                        skipped_count += 1
                        continue
                    
                    existing_records = self.db.execute_query('''
                        SELECT COUNT(*) FROM records r
                        JOIN dungeons d ON r.dungeon_id = d.id
                        WHERE d.name = ? AND r.time = ? AND r.black_owner = ? AND r.worker = ?
                    ''', (
                        dungeon_name,
                        record.get('time'),
                        record.get('black_owner'),
                        record.get('worker')
                    ))
                    
                    if existing_records and existing_records[0][0] > 0:
                        duplicate_count += 1
                        continue
                    
                    result = self.db.execute_query("SELECT id FROM dungeons WHERE name=?", (dungeon_name,))
                    if not result:
                        self.db.execute_update(
                            "INSERT INTO dungeons (name, special_drops) VALUES (?, ?)", 
                            (dungeon_name, record.get('special_drops', ''))
                        )
                        dungeon_id = self.db.cursor.lastrowid
                    else:
                        dungeon_id = result[0][0]
                    
                    special_auctions = record.get('special_auctions', [])
                    if isinstance(special_auctions, str):
                        try:
                            special_auctions = json.loads(special_auctions)
                        except:
                            special_auctions = []
                    
                    self.db.execute_update('''
                        INSERT INTO records (
                            dungeon_id, trash_gold, iron_gold, other_gold, special_auctions, 
                            total_gold, black_owner, worker, time, team_type, lie_down_count, 
                            fine_gold, subsidy_gold, personal_gold, note, is_new,
                            scattered_consumption, iron_consumption, special_consumption, other_consumption, total_consumption
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?)
                    ''', (
                        dungeon_id,
                        record.get('trash_gold', 0),
                        record.get('iron_gold', 0),
                        record.get('other_gold', 0),
                        json.dumps(special_auctions, ensure_ascii=False),
                        record.get('total_gold', 0),
                        record.get('black_owner'),
                        record.get('worker'),
                        record.get('time', get_current_time()),
                        record.get('team_type', '十人本'),
                        record.get('lie_down_count', 0),
                        record.get('fine_gold', 0),
                        record.get('subsidy_gold', 0),
                        record.get('personal_gold', 0),
                        record.get('note', ''),
                        record.get('scattered_consumption', 0),
                        record.get('iron_consumption', 0),
                        record.get('special_consumption', 0),
                        record.get('other_consumption', 0),
                        record.get('total_consumption', 0)
                    ))
                    success_count += 1
                    
                except Exception as e:
                    print(f"导入单条记录时出错: {e}")
                    skipped_count += 1
                    continue
            
            self.load_recent_records(50)
            self.update_stats()
            self.load_dungeon_options()
            self.load_black_owner_options()
            
            messagebox.showinfo("导入完成", 
                            f"成功导入: {success_count} 条记录\n"
                            f"跳过重复: {duplicate_count} 条记录\n"
                            f"跳过无效: {skipped_count} 条记录")
            self.load_weekly_data()
            
        except Exception as e:
            messagebox.showerror("错误", f"导入数据时出错: {str(e)}")

    def export_data(self):
        file_path = filedialog.asksaveasfilename(
            title="保存数据文件",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            records = self.db.execute_query('''
                SELECT d.name, r.trash_gold, r.iron_gold, r.other_gold, r.special_auctions, 
                    r.total_gold, r.black_owner, r.worker, r.time, r.team_type, r.lie_down_count, 
                    r.fine_gold, r.subsidy_gold, r.personal_gold, r.note
                FROM records r
                JOIN dungeons d ON r.dungeon_id = d.id
                ORDER BY r.time DESC
            ''')
            
            export_data = []
            for record in records:
                export_data.append({
                    'dungeon_name': record[0],
                    'trash_gold': record[1],
                    'iron_gold': record[2],
                    'other_gold': record[3],
                    'special_auctions': json.loads(record[4]) if record[4] else [],
                    'total_gold': record[5],
                    'black_owner': record[6],
                    'worker': record[7],
                    'time': record[8],
                    'team_type': record[9],
                    'lie_down_count': record[10],
                    'fine_gold': record[11],
                    'subsidy_gold': record[12],
                    'personal_gold': record[13],
                    'note': record[14]
                })
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("成功", f"数据已导出到 {file_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出数据时出错: {str(e)}")

    def load_weekly_worker_options(self):
        workers = [row[0] for row in self.db.execute_query(
            "SELECT DISTINCT worker FROM records WHERE worker IS NOT NULL AND worker != '' ORDER BY worker"
        )]
        self.weekly_worker_combo['values'] = [""] + workers

    def on_weekly_worker_select(self, event=None):
        self.load_weekly_data()

    def get_weekly_time_range(self, team_type):
        now = dt.datetime.now()
        
        if team_type == "十人本":
            if now.weekday() == 0 and now.hour < 7:
                start_date = now - timedelta(days=now.weekday() + 3)
                start_date = start_date.replace(hour=7, minute=0, second=0, microsecond=0)
                end_date = now.replace(hour=7, minute=0, second=0, microsecond=0)
            elif now.weekday() < 4 or (now.weekday() == 4 and now.hour < 7):
                start_date = now - timedelta(days=now.weekday())
                start_date = start_date.replace(hour=7, minute=0, second=0, microsecond=0)
                end_date = now.replace(hour=7, minute=0, second=0, microsecond=0)
                if now.weekday() == 4 and now.hour < 7:
                    end_date = start_date + timedelta(days=4)
                else:
                    end_date = start_date + timedelta(days=4)
                    end_date = end_date.replace(hour=7, minute=0, second=0, microsecond=0)
            else:
                start_date = now - timedelta(days=(now.weekday() - 4))
                start_date = start_date.replace(hour=7, minute=0, second=0, microsecond=0)
                end_date = start_date + timedelta(days=3)
                end_date = end_date.replace(hour=7, minute=0, second=0, microsecond=0)
        else:
            if now.weekday() == 0 and now.hour < 7:
                start_date = now - timedelta(days=now.weekday() + 7)
                start_date = start_date.replace(hour=7, minute=0, second=0, microsecond=0)
                end_date = now.replace(hour=7, minute=0, second=0, microsecond=0)
            else:
                start_date = now - timedelta(days=now.weekday())
                start_date = start_date.replace(hour=7, minute=0, second=0, microsecond=0)
                end_date = start_date + timedelta(days=7)
        
        return start_date, end_date

    def load_weekly_data(self):
        for item in self.weekly_tree.get_children():
            self.weekly_tree.delete(item)
        
        selected_worker = self.weekly_worker_var.get()
        
        conditions = []
        params = []
        
        if selected_worker:
            conditions.append("r.worker = ?")
            params.append(selected_worker)
        else:
            conditions.append("r.worker IS NOT NULL AND r.worker != ''")
        
        team_types = ["十人本", "二十五人本"]
        all_records = []
        
        for team_type in team_types:
            start_date, end_date = self.get_weekly_time_range(team_type)
            
            sql = '''
                SELECT r.worker, d.name, r.note, r.time, r.team_type
                FROM records r
                JOIN dungeons d ON r.dungeon_id = d.id
                WHERE r.time >= ? AND r.time < ? AND r.team_type = ?
            '''
            
            if conditions:
                sql += " AND " + " AND ".join(conditions)
            
            sql += " ORDER BY r.time DESC"
            
            records = self.db.execute_query(sql, [start_date.strftime("%Y-%m-%d %H:%M:%S"), 
                                                end_date.strftime("%Y-%m-%d %H:%M:%S"), 
                                                team_type] + params)
            all_records.extend(records)
        
        for record in all_records:
            worker, dungeon, note, time_str, team_type = record
            display_note = note if note else ""
            self.weekly_tree.insert("", "end", values=(worker, f"{dungeon}({team_type})", display_note))
        
        self.update_weekly_period_info()

    def update_weekly_period_info(self):
        now = dt.datetime.now()
        
        ten_start, ten_end = self.get_weekly_time_range("十人本")
        twenty_five_start, twenty_five_end = self.get_weekly_time_range("二十五人本")
        
        ten_period = f"十人本周期: {ten_start.strftime('%m-%d %H:%M')} 至 {ten_end.strftime('%m-%d %H:%M')}"
        twenty_five_period = f"二十五人本周期: {twenty_five_start.strftime('%m-%d %H:%M')} 至 {twenty_five_end.strftime('%m-%d %H:%M')}"
        
        self.weekly_period_var.set(f"{ten_period} | {twenty_five_period}")

    def restore_pane_position(self, pane, pane_name):
        try:
            result = self.db.execute_query("SELECT position FROM pane_positions WHERE pane_name = ?", (pane_name,))
            if result:
                position = result[0][0]
                if position is not None:
                    self.root.after(500, lambda: self.safe_set_pane_position(pane, pane_name, position))
        except Exception:
            pass

    def safe_set_pane_position(self, pane, pane_name, position):
        try:
            if not pane.winfo_exists():
                return
                
            pane.update_idletasks()
            
            pane_width = pane.winfo_width()
            pane_height = pane.winfo_height()
            
            if pane_width <= 1 or pane_height <= 1:
                self.root.after(100, lambda: self.safe_set_pane_position(pane, pane_name, position))
                return
            
            orient = pane.cget('orient')
            if orient == 'horizontal':
                max_position = pane_width - 50
                actual_position = min(position, max_position)
                if actual_position > 30:
                    pane.sash_place(0, actual_position, 0)
            else:
                max_position = pane_height - 50
                actual_position = min(position, max_position)
                if actual_position > 30:
                    pane.sash_place(0, 0, actual_position)
                    
        except Exception:
            pass

    def save_pane_positions(self):
        try:
            if hasattr(self, 'record_pane') and self.record_pane.winfo_exists():
                try:
                    sash_pos = self.record_pane.sashpos(0)
                    if sash_pos is not None and sash_pos > 0:
                        self.db.save_pane_position(self.record_pane_name, sash_pos)
                except Exception:
                    pass
            
            if hasattr(self, 'preset_pane') and self.preset_pane.winfo_exists():
                try:
                    sash_pos = self.preset_pane.sashpos(0)
                    if sash_pos is not None and sash_pos > 0:
                        self.db.save_pane_position(self.preset_pane_name, sash_pos)
                except Exception:
                    pass
                    
        except Exception:
            pass

    def setup_pane_events(self):
        def on_pane_configure(event):
            if hasattr(self, '_pane_save_scheduled'):
                self.root.after_cancel(self._pane_save_scheduled)
            self._pane_save_scheduled = self.root.after(1000, self.save_pane_positions)
        
        if hasattr(self, 'record_pane'):
            self.record_pane.bind('<Configure>', on_pane_configure)
            self.root.after(1000, lambda: self.set_default_pane_positions())
        
        self.schedule_pane_position_check()

    def set_default_pane_positions(self):
        try:
            result = self.db.execute_query("SELECT position FROM pane_positions WHERE pane_name = ?", (self.record_pane_name,))
            if not result:
                default_position = int(350 * SCALE_FACTOR)
                self.record_pane.sash_place(0, default_position, 0)
                self.db.save_pane_position(self.record_pane_name, default_position)
        except Exception:
            pass

    def schedule_pane_position_check(self):
        try:
            if not self.root.winfo_exists():
                return
                
            self.save_pane_positions()
            
            after_id = self.root.after(10000, self.schedule_pane_position_check)
            self.after_ids.append(after_id)
        except Exception:
            return

    def schedule_time_update(self):
        self.update_time()
        after_id = self.root.after(1000, self.schedule_time_update)
        self.after_ids.append(after_id)

    def setup_window_tracking(self):
        def save_window_state():
            try:
                width = self.root.winfo_width()
                height = self.root.winfo_height()
                x = self.root.winfo_x()
                y = self.root.winfo_y()
                maximized = 1 if self.root.state() == 'zoomed' else 0
                
                self.db.execute_update('''
                    INSERT OR REPLACE INTO window_state (width, height, maximized, x, y)
                    VALUES (?, ?, ?, ?, ?)
                ''', (width, height, maximized, x, y))
            except Exception:
                pass
        
        self.root.bind('<Configure>', lambda e: self.root.after(1000, save_window_state))
        
        def on_visibility_change():
            self.root.after(100, save_window_state)
        
        self.root.bind('<Unmap>', lambda e: self.root.after(100, on_visibility_change))
        self.root.bind('<Map>', lambda e: self.root.after(100, on_visibility_change))

    def on_close(self):
        for after_id in self.after_ids:
            try:
                self.root.after_cancel(after_id)
            except Exception:
                pass
        
        try:
            self.save_column_widths()
        except Exception as e:
            print(f"保存列宽度时出错: {e}")
        
        try:
            self.save_pane_positions()
        except Exception as e:
            print(f"保存窗格位置时出错: {e}")
        
        try:
            self.db.close()
        except Exception:
            pass
        
        self.root.destroy()
        
        import os
        os._exit(0)

def main():
    root = tk.Tk()
    app = JX3DungeonTracker(root)
    atexit.register(app.on_close)
    root.mainloop()

if __name__ == "__main__":
    main()