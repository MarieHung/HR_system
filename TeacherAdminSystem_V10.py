import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime
import io

try:
    import msofficecrypt
    HAS_MSOFFICE = True
except ImportError:
    HAS_MSOFFICE = False

class TeacherStatsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("教師資料管理中心 - 行政綜合業務系統 (V10)")
        self.root.geometry("1300x900")

        # 全域資料變數
        self.file1_path = None 
        self.df1_combined = None 
        self.file2_path = None 
        self.df_travel = None 
        
        self.title_col1 = None

        # 主 Notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 建立 9 個主分頁
        tabs = [" 🔍 職稱統計 ", " 📅 到離查詢 ", " 📊 給副校長 ", " 📄 D5 報表 ", " 🛌 留職停薪 ", " 🌏 赴大陸統計 ", " 📝 個資申請單 ", " 📑 人事異動單 ", " 🏫 校區人數 "]
        self.tab_frames = []
        for text in tabs:
            f = ttk.Frame(self.notebook)
            self.notebook.add(f, text=text)
            self.tab_frames.append(f)

        self.setup_tab1(self.tab_frames[0])
        self.setup_tab2(self.tab_frames[1]) # 到離查詢：左右分欄
        self.setup_tab3(self.tab_frames[2])
        self.setup_tab4(self.tab_frames[3])
        self.setup_tab5(self.tab_frames[4])
        self.setup_tab6(self.tab_frames[5])
        self.setup_tab7(self.tab_frames[6])
        self.setup_tab8(self.tab_frames[7])
        self.setup_tab9(self.tab_frames[8])

    def get_xl_safe(self, path, password=None, sheet_name=0):
        """通用讀取加密或一般 Excel 的方法"""
        if not password:
            return pd.read_excel(path, sheet_name=sheet_name)
        
        if not HAS_MSOFFICE:
            raise ImportError("請安裝 msoffice-crypt-tool (pip install msoffice-crypt-tool) 以支援加密 Excel")
            
        temp_io = io.BytesIO()
        with open(path, "rb") as f:
            file = msofficecrypt.OfficeFile(f)
            file.load_key(password=password)
            file.decrypt(temp_io)
        temp_io.seek(0)
        return pd.read_excel(temp_io, sheet_name=sheet_name)

    # --- 共通 UI 工具 ---
    def create_report_ui(self, parent, title, desc, btn_text, command):
        main_f = ttk.Frame(parent, padding="20")
        main_f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_f, text=title, font=('Microsoft JhengHei', 14, 'bold')).pack(pady=(0, 5))
        ttk.Label(main_f, text=desc, foreground="blue").pack(pady=(0, 15))
        btn = ttk.Button(main_f, text=btn_text, command=lambda: command(out_txt))
        btn.pack(pady=5)
        out_txt = tk.Text(main_f, font=('Microsoft JhengHei', 11), bg="#f8f9fa")
        out_txt.pack(fill=tk.BOTH, expand=True, pady=10)
        return out_txt

    # --- 分頁 1: 職稱統計 ---
    def setup_tab1(self, f):
        self.tab1_cat_vars = {} 
        self.tab1_title_vars = {} 
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="全校人員職稱人數統計", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        
        file_f = ttk.LabelFrame(main, text="載入資料 (各系所教師.xlsx)", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="選擇 各系所教師 檔案", command=self.load_file1_v7).pack(side=tk.LEFT)
        self.file_label1 = ttk.Label(file_f, text="尚未載入", foreground="gray")
        self.file_label1.pack(side=tk.LEFT, padx=10)
        
        cat_f = ttk.LabelFrame(main, text="1. 請勾選類別 (對應工作表)", padding=10)
        cat_f.pack(fill=tk.X, pady=5)
        cat_grid = ttk.Frame(cat_f)
        cat_grid.pack(fill=tk.X, padx=5)
        categories = ["新制職員", "助教", "技工工友合計", "教師", "約聘教師", "駐衛警察隊", "專案教學人員", "講座教授", "客座教授", "醫事人員", "稀少性科技人員"]
        for i, cat in enumerate(categories):
            var = tk.BooleanVar()
            self.tab1_cat_vars[cat] = var
            ttk.Checkbutton(cat_grid, text=cat, variable=var, command=self.refresh_tab1_titles).grid(row=i//6, column=i%6, sticky=tk.W, padx=10, pady=2)
            
        self.select_f1 = ttk.LabelFrame(main, text="2. 勾選職稱 (橫向排列)", padding=10)
        self.select_f1.pack(fill=tk.BOTH, expand=True, pady=10)
        self.canvas1 = tk.Canvas(self.select_f1, highlightthickness=0)
        self.scroll1 = ttk.Scrollbar(self.select_f1, command=self.canvas1.yview)
        self.scroll_f1 = ttk.Frame(self.canvas1)
        self.scroll_f1.bind("<Configure>", lambda e: self.canvas1.configure(scrollregion=self.canvas1.bbox("all")))
        self.canvas1.create_window((0,0), window=self.scroll_f1, anchor="nw")
        self.canvas1.configure(yscrollcommand=self.scroll1.set)
        self.canvas1.pack(side="left", fill="both", expand=True)
        self.scroll1.pack(side="right", fill="y")
        
        self.calc_btn1 = ttk.Button(main, text="📊 開始統計人數", command=self.calculate1_v7, state=tk.DISABLED)
        self.calc_btn1.pack(pady=5)
        self.res_text1 = tk.Text(main, height=15, font=('Consolas', 11), bg="#fdfdfd", state=tk.DISABLED)
        self.res_text1.pack(fill=tk.BOTH, expand=True, pady=5)

    def load_file1_v7(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p:
            return
        self.file1_path = p
        self.file_label1.config(text=os.path.basename(p), foreground="black")
        if hasattr(self, 'file_label_app'):
            self.file_label_app.config(text=os.path.basename(p), foreground="black")
        self.refresh_tab1_titles()
        # 自動更新校區統計的工作表清單
        if hasattr(self, 'refresh_tab9_sheets'):
            self.refresh_tab9_sheets()

    def refresh_tab1_titles(self):
        if not self.file1_path:
            return
        sel_cats = [cat for cat, var in self.tab1_cat_vars.items() if var.get()]
        if not sel_cats:
            for w in self.scroll_f1.winfo_children():
                w.destroy()
            self.calc_btn1.config(state=tk.DISABLED)
            return
        try:
            xl = pd.ExcelFile(self.file1_path)
            all_dfs = []
            for cat in sel_cats:
                if cat in xl.sheet_names:
                    all_dfs.append(pd.read_excel(self.file1_path, sheet_name=cat))
            if not all_dfs:
                return
            df = pd.concat(all_dfs, ignore_index=True)
            self.df1_combined = df
            self.title_col1 = next((c for c in df.columns if any(k in str(c) for k in ['職稱', '職別'])), df.columns[0])
            titles = df[self.title_col1].astype(str).str.strip().replace(['nan','None',''], pd.NA).dropna().unique()
            for w in self.scroll_f1.winfo_children():
                w.destroy()
            self.tab1_title_vars = {}
            for i, t in enumerate(sorted(titles)):
                var = tk.BooleanVar()
                self.tab1_title_vars[t] = var
                cb = ttk.Checkbutton(self.scroll_f1, text=t, variable=var)
                cb.grid(row=i//4, column=i%4, sticky=tk.W, padx=10, pady=2)
            self.calc_btn1.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def calculate1_v7(self):
        if self.df1_combined is None:
            return
        sel = [t for t, v in self.tab1_title_vars.items() if v.get()]
        if not sel:
            messagebox.showwarning("提示", "請勾選職稱")
            return
        cnt = self.df1_combined[self.title_col1].astype(str).str.strip().value_counts()
        self.res_text1.config(state=tk.NORMAL)
        self.res_text1.delete(1.0, tk.END)
        total = 0
        self.res_text1.insert(tk.END, f"【職稱人數統計結果】\n")
        self.res_text1.insert(tk.END, f"{'職稱名稱':<25} | {'統計人數':>5}\n" + "-"*40 + "\n")
        for t in sel:
            c = cnt.get(t, 0)
            self.res_text1.insert(tk.END, f"{t:<23} | {c:>5} 名\n")
            total += c
        self.res_text1.insert(tk.END, "-"*40 + f"\n總計人數：{total} 名")
        self.res_text1.config(state=tk.DISABLED)

    # --- 分頁 2: 到離查詢 (左右分欄版) ---
    def setup_tab2(self, f):
        self.tab2_sheet_vars = {} 
        
        main = ttk.Frame(f, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        # 建立左右分欄
        left_panel = ttk.Frame(main, width=350)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        
        right_panel = ttk.Frame(main)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- 左側：控制區域 ---
        ttk.Label(left_panel, text="🔍 查詢條件設定", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5, anchor=tk.W)
        
        # 0. 檔案載入
        file_f = ttk.LabelFrame(left_panel, text="步驟 0：載入檔案", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="📂 選擇 到離名單 檔案", command=self.load_file2_v7).pack(fill=tk.X)
        self.file_label2 = ttk.Label(file_f, text="尚未載入檔案", foreground="gray", wraplength=300)
        self.file_label2.pack(pady=5)

        # 1. 工作表勾選區 (增加滾動條)
        self.sheet_frame2 = ttk.LabelFrame(left_panel, text="步驟 1：勾選工作表 (可多選)", padding=10)
        self.sheet_frame2.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.sheet_canvas = tk.Canvas(self.sheet_frame2, highlightthickness=0)
        self.sheet_scroll = ttk.Scrollbar(self.sheet_frame2, orient="vertical", command=self.sheet_canvas.yview)
        self.sheet_container2 = ttk.Frame(self.sheet_canvas)
        
        self.sheet_container2.bind("<Configure>", lambda e: self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox("all")))
        self.sheet_canvas.create_window((0,0), window=self.sheet_container2, anchor="nw")
        self.sheet_canvas.configure(yscrollcommand=self.sheet_scroll.set)
        
        self.sheet_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sheet_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 2. 搜尋條件區
        sf = ttk.LabelFrame(left_panel, text="步驟 2：輸入條件", padding=10)
        sf.pack(fill=tk.X, pady=5)
        
        grid_f = ttk.Frame(sf)
        grid_f.pack(fill=tk.X)
        
        ttk.Label(grid_f, text="日期:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.date_e = ttk.Entry(grid_f, width=20)
        self.date_e.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(grid_f, text="姓名:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.name_e = ttk.Entry(grid_f, width=20)
        self.name_e.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        self.search_btn2 = ttk.Button(left_panel, text="🚀 執行篩選查詢", command=self.search_data2_v7, state=tk.DISABLED)
        self.search_btn2.pack(fill=tk.X, pady=10)

        # --- 右側：結果顯示區域 ---
        ttk.Label(right_panel, text="📋 查詢結果明細", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5, anchor=tk.W)
        
        table_f = ttk.Frame(right_panel)
        table_f.pack(fill=tk.BOTH, expand=True)
        
        cols = ("date", "name", "unit", "level1", "title", "reason", "remarks")
        self.tree2 = ttk.Treeview(table_f, columns=cols, show='headings')
        vsb2 = ttk.Scrollbar(table_f, orient="vertical", command=self.tree2.yview)
        hsb2 = ttk.Scrollbar(table_f, orient="horizontal", command=self.tree2.xview)
        self.tree2.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        
        for c, h in zip(cols, ["日期", "姓名", "單位", "一級單位", "職稱", "異動原因", "備註"]): 
            self.tree2.heading(c, text=h)
            self.tree2.column(c, width=120, anchor=tk.CENTER)
        
        self.tree2.grid(row=0, column=0, sticky='nsew')
        vsb2.grid(row=0, column=1, sticky='ns')
        hsb2.grid(row=1, column=0, sticky='ew')
        
        table_f.grid_rowconfigure(0, weight=1)
        table_f.grid_columnconfigure(0, weight=1)

    def load_file2_v7(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p:
            return
        self.file2_path = p
        self.file_label2.config(text=os.path.basename(p), foreground="black")
        try:
            xl = pd.ExcelFile(p)
            for w in self.sheet_container2.winfo_children():
                w.destroy()
            self.tab2_sheet_vars = {}
            for i, name in enumerate(xl.sheet_names):
                var = tk.BooleanVar()
                self.tab2_sheet_vars[name] = var
                cb = ttk.Checkbutton(self.sheet_container2, text=name, variable=var)
                cb.pack(fill=tk.X, padx=5, pady=1)
            self.search_btn2.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("讀取 Sheet 失敗", str(e))

    def search_data2_v7(self):
        if not self.file2_path:
            return
        sel_sheets = [s for s, v in self.tab2_sheet_vars.items() if v.get()]
        if not sel_sheets:
            messagebox.showwarning("提示", "請至少勾選一個工作表！")
            return
        d_val = self.date_e.get().strip()
        n_val = self.name_e.get().strip()
        try:
            all_dfs = []
            for s in sel_sheets:
                temp_df = pd.read_excel(self.file2_path, sheet_name=s)
                all_dfs.append(temp_df)
            combined_df = pd.concat(all_dfs, ignore_index=True)
            cols = combined_df.columns
            cd = next((c for c in cols if any(k in str(c) for k in ["日期", "交"])), cols[0])
            cn = next((c for c in cols if any(k in str(c) for k in ["姓名", "憪"])), cols[1])
            ct = next((c for c in cols if any(k in str(c) for k in ["職稱", "職別"])), "職稱")
            cr = next((c for c in cols if any(k in str(c) for k in ["異動原因", "原因"])), "原因")
            mask = pd.Series([True] * len(combined_df))
            if d_val:
                mask &= combined_df[cd].astype(str).str.contains(d_val)
            if n_val:
                mask &= combined_df[cn].astype(str).str.contains(n_val)
            res = combined_df[mask]
            for i in self.tree2.get_children():
                self.tree2.delete(i)
            if res.empty:
                messagebox.showinfo("查詢結果", "找不到符合條件的資料。")
                return
            for _, r in res.iterrows():
                vals = [r.get(cd, ""), r.get(cn, ""), r.get("單位", ""), r.get("一級單位", ""), r.get(ct, ""), r.get(cr, ""), r.get("備註", "")]
                self.tree2.insert("", tk.END, values=vals)
        except Exception as e:
            messagebox.showerror("查詢出錯", str(e))

    # --- 分頁 3: 副校長報表 ---
    def setup_tab3(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="副校長 - 全校各類人員員額統計週報", font=('Microsoft JhengHei', 16, 'bold'), foreground="#1a5276").pack(pady=10)
        manual_frame = ttk.LabelFrame(main, text="手動輸入額外人員數據 (約用、兼任、助教、經理人)", padding=10)
        manual_frame.pack(fill=tk.X, pady=10)
        self.manual_entries = {} 
        items = [("約用人員", "manual"), ("兼任教師", "pt"), ("新制助教", "new_ta"), ("舊制助教", "old_ta"), ("專業經理人", "manager")]
        for i, (label, key) in enumerate(items):
            ttk.Label(manual_frame, text=f"{label} - 總人數:").grid(row=i, column=0, padx=5, pady=5, sticky=tk.E)
            total_e = ttk.Entry(manual_frame, width=8); total_e.insert(0, "0"); total_e.grid(row=i, column=1, padx=5)
            ttk.Label(manual_frame, text="男:").grid(row=i, column=2, padx=5); m_e = ttk.Entry(manual_frame, width=8); m_e.insert(0, "0"); m_e.grid(row=i, column=3, padx=5)
            ttk.Label(manual_frame, text="女:").grid(row=i, column=4, padx=5); f_e = ttk.Entry(manual_frame, width=8); f_e.insert(0, "0"); f_e.grid(row=i, column=5, padx=5)
            self.manual_entries[key] = (total_e, m_e, f_e)
        ttk.Button(main, text="📊 生成全校人員統計報表 (含手動與合併類別)", command=self.gen_vice_president_report).pack(pady=5)
        self.vp_out = tk.Text(main, font=('Microsoft JhengHei', 11), bg="#f4f6f7")
        self.vp_out.pack(fill=tk.BOTH, expand=True, pady=10)

    def gen_vice_president_report(self):
        if not self.file1_path:
            messagebox.showwarning("提示", "請載入檔案 1")
            return
        try:
            manual_data = {k: (int(v[0].get()), int(v[1].get()), int(v[2].get())) for k,v in self.manual_entries.items()}
        except ValueError:
            messagebox.showerror("錯誤", "手動輸入欄位請填入整數數字")
            return
        target_sheets = ["教師", "約聘教師", "專案教學人員", "講座教授", "客座人員", "研究人員", "稀少性科技人員", "醫事人員", "新制職員", "駐衛警察隊", "技工工友合計"]
        try:
            xl = pd.ExcelFile(self.file1_path)
            self.vp_out.delete(1.0, tk.END)
            self.vp_out.insert(tk.END, f"【各類人員員額與性別統計摘要 - 副校長】\n日期：{datetime.now().strftime('%Y-%m-%d')}\n" + "="*75 + "\n")
            grand_total = sum(d[0] for d in manual_data.values())
            grand_m = sum(d[1] for d in manual_data.values())
            grand_f = sum(d[2] for d in manual_data.values())
            admin_staff_total = 0; admin_staff_m = 0; admin_staff_f = 0
            for sheet in target_sheets:
                if sheet in xl.sheet_names:
                    df = pd.read_excel(self.file1_path, sheet_name=sheet)
                    c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                    if c_t:
                        df = df[df[c_t].astype(str).str.strip().replace(['nan','','None'], pd.NA).notna()]
                    c_g = next((c for c in df.columns if "性別" in str(c)), None)
                    m = df[df[c_g].astype(str).str.contains("男")].shape[0] if c_g else 0
                    f = df[df[c_g].astype(str).str.contains("女")].shape[0] if c_g else 0
                    count = len(df)
                    self.vp_out.insert(tk.END, f"● {sheet:<15} | {count:>8} | 男:{m:>4} | 女:{f:>4}\n")
                    grand_total += count; grand_m += m; grand_f += f
                    if sheet in ["新制職員", "醫事人員", "稀少性科技人員"]:
                        admin_staff_total += count; admin_staff_m += m; admin_staff_f += f
                else:
                    self.vp_out.insert(tk.END, f"○ {sheet:<15} | （找不到工作表）\n")
            self.vp_out.insert(tk.END, "-"*75 + "\n")
            disp = {"manual": "約用人員(填報)", "pt": "兼任教師(填報)", "new_ta": "新制助教(填報)", "old_ta": "舊制助教(填報)", "manager": "專業經理人(填報)"}
            for k, (t, m, f) in manual_data.items():
                self.vp_out.insert(tk.END, f"● {disp[k]:<13} | {t:>8} | 男:{m:>4} | 女:{f:>4}\n")
            self.vp_out.insert(tk.END, "-"*75 + f"\n【全校總計彙整】\n   - 職員小計 (新制/醫事/稀少性) ： {admin_staff_total:>5} 名 (男:{admin_staff_m} / 女:{admin_staff_f})\n")
            self.vp_out.insert(tk.END, f"   - 全校人員總計                ： {grand_total} 名 (男：{grand_m} / 女：{grand_f})\n" + "="*75)
        except Exception as e:
            messagebox.showerror("失敗", str(e))

    # --- 其它報表分頁 4-6 ---
    # --- 分頁 4: D5 報表 (V7.7) ---
    def setup_tab4(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="D5 員額異動報表系統", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        
        ttk.Label(main, text="說明：本功能將彙整『各系所教師』中的資料，並產出符合 D5 規範的 Excel 底稿。", foreground="blue").pack(pady=(0, 15))

        # 檔名輸入
        file_name_f = ttk.Frame(main)
        file_name_f.pack(fill=tk.X, pady=5)
        ttk.Label(file_name_f, text="匯出檔案名稱:").pack(side=tk.LEFT, padx=5)
        self.d5_filename_e = ttk.Entry(file_name_f, width=30)
        self.d5_filename_e.insert(0, f"D5報表底稿_{datetime.now().strftime('%m%d')}")
        self.d5_filename_e.pack(side=tk.LEFT, padx=5)
        ttk.Label(file_name_f, text=".xlsx").pack(side=tk.LEFT)

        # 按鈕區
        btn_f = ttk.Frame(main)
        btn_f.pack(fill=tk.X, pady=10)
        ttk.Button(btn_f, text="📊 開始計算D5報表", command=self.gen_d5_logic).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_f, text="📥 匯出Excel底稿", command=self.export_d5_excel).pack(side=tk.LEFT, padx=5)

        self.d5_out = tk.Text(main, font=('Microsoft JhengHei', 11), bg="#f0f4f8")
        self.d5_out.pack(fill=tk.BOTH, expand=True, pady=10)

    def gen_d5_logic(self):
        if not self.file1_path:
            messagebox.showwarning("提示", "請先在『職稱統計』分頁載入『各系所教師』檔案！")
            return
        
        try:
            xl = pd.ExcelFile(self.file1_path)
            all_sheets = xl.sheet_names
            
            # 定義三大類別的容器
            cat_data = {
                "相當簡薦委": [],
                "教師": [],
                "聘任": []
            }

            # 1. 相當簡薦委：新制職員 + 稀少性科技人員
            for s in ["新制職員", "稀少性科技人員"]:
                if s in all_sheets:
                    df = pd.read_excel(self.file1_path, sheet_name=s)
                    # 排除職稱為空白的資料
                    c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                    if c_t:
                        df = df[df[c_t].astype(str).str.strip().replace(['nan','','None','nan '], pd.NA).notna()]
                    
                    # 排除「留職停薪」欄位有紀錄的人員
                    c_unpaid = next((c for c in df.columns if "留職停薪" in str(c)), None)
                    if c_unpaid:
                        # 只有當欄位為空 (NaN 或 空字串) 時才保留
                        mask_not_leave = df[c_unpaid].astype(str).str.strip().replace(['nan','None','','nan '], pd.NA).isna()
                        df = df[mask_not_leave]

                    df['D5類別'] = "相當簡薦委"
                    df['來源工作表'] = s
                    cat_data["相當簡薦委"].append(df)

            # 2. 教師 & 3. 聘任 (處理 助教、教師、研究人員)
            # 注意：排除「專案教學人員」、「約聘教師」
            
            # 教師工作表
            if "教師" in all_sheets:
                df = pd.read_excel(self.file1_path, sheet_name="教師")
                # 排除職稱為空白的資料
                c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                if c_t:
                    df = df[df[c_t].astype(str).str.strip().replace(['nan','','None','nan '], pd.NA).notna()]
                
                # 排除「留職停薪」欄位有紀錄的人員
                c_unpaid = next((c for c in df.columns if "留職停薪" in str(c)), None)
                if c_unpaid:
                    mask_not_leave = df[c_unpaid].astype(str).str.strip().replace(['nan','None','','nan '], pd.NA).isna()
                    df = df[mask_not_leave]

                # 判斷是否為「專業技術人員」
                tech_titles = ["教授級專業技術人員", "副教授級專業技術人員", "助理教授級專業技術人員", "講師級專業技術人員"]
                is_tech_mask = df[c_t].astype(str).str.contains('|'.join(tech_titles)) if c_t else pd.Series([False]*len(df))
                
                df_tech = df[is_tech_mask].copy()
                df_normal = df[~is_tech_mask].copy()

                if not df_normal.empty:
                    df_normal['D5類別'] = "教師"
                    df_normal['來源工作表'] = "教師"
                    cat_data["教師"].append(df_normal)
                
                if not df_tech.empty:
                    df_tech['D5類別'] = "聘任"
                    df_tech['來源工作表'] = "教師(專業技術人員)"
                    cat_data["聘任"].append(df_tech)
            
            # 研究人員 -> 歸類為 聘任 (排除 專案教學人員、約聘教師、客座人員)
            for s in ["研究人員"]:
                if s in all_sheets:
                    df = pd.read_excel(self.file1_path, sheet_name=s)
                    c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                    if c_t:
                        df = df[df[c_t].astype(str).str.strip().replace(['nan','','None','nan '], pd.NA).notna()]
                    
                    df['D5類別'] = "聘任"
                    df['來源工作表'] = s
                    cat_data["聘任"].append(df)

            # 助教工作表 (需區分新制/舊制助教)
            if "助教" in all_sheets:
                df = pd.read_excel(self.file1_path, sheet_name="助教")
                c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                if c_t:
                    df = df[df[c_t].astype(str).str.strip().replace(['nan','','None','nan '], pd.NA).notna()]
                
                # 排除「留職停薪」欄位有紀錄的人員
                c_unpaid = next((c for c in df.columns if "留職停薪" in str(c)), None)
                if c_unpaid:
                    mask_not_leave = df[c_unpaid].astype(str).str.strip().replace(['nan','None','','nan '], pd.NA).isna()
                    df = df[mask_not_leave]

                c_check = next((c for c in df.columns if "舊制講師或助教" in str(c)), None)
                if c_check:
                    is_old_mask = df[c_check].astype(str).str.contains("舊制助教")
                    df_old = df[is_old_mask].copy()
                    df_old['D5類別'] = "教師"; df_old['來源工作表'] = "助教(舊制)"
                    if not df_old.empty: cat_data["教師"].append(df_old)
                    
                    is_new_mask = df[c_check].astype(str).str.contains("新制助教")
                    df_new = df[is_new_mask].copy()
                    df_new['D5類別'] = "聘任"; df_new['來源工作表'] = "助教(新制)"
                    if not df_new.empty: cat_data["聘任"].append(df_new)
                else:
                    df['D5類別'] = "聘任"; df['來源工作表'] = "助教(未區分)"
                    cat_data["聘任"].append(df)

            # 合併所有資料
            final_list = []
            summary_text = "【D5 類別來源明細】\n"
            for name, dfs in cat_data.items():
                if dfs:
                    combined = pd.concat(dfs, ignore_index=True)
                    final_list.append(combined)
                    
                    summary_text += f"● {name} (總計: {len(combined)} 人)\n"
                    for d in dfs:
                        src = d['來源工作表'].iloc[0]
                        summary_text += f"   - 來源: {src:<18} | {len(d):>3} 人\n"
                    
                    # 性別與官等統計摘要
                    c_gender = next((c for c in combined.columns if "性別" in str(c)), None)
                    c_rank = next((c for c in combined.columns if "官等" in str(c)), None)
                    
                    if name == "相當簡薦委" and c_rank and c_gender:
                        summary_text += "     [官等細分統計]\n"
                        # 定義要統計的官等關鍵字
                        target_ranks = ["簡", "薦", "委"]
                        for tr in target_ranks:
                            # 篩選包含該關鍵字的官等
                            mask_r = combined[c_rank].astype(str).str.contains(tr)
                            sub_r = combined[mask_r]
                            if not sub_r.empty:
                                m = sub_r[sub_r[c_gender].astype(str).str.contains("男")].shape[0]
                                f = sub_r[sub_r[c_gender].astype(str).str.contains("女")].shape[0]
                                summary_text += f"       * {tr}等: 共 {len(sub_r):>3} 人 (男:{m} / 女:{f})\n"
                    elif c_gender:
                        m = combined[combined[c_gender].astype(str).str.contains("男")].shape[0]
                        f = combined[combined[c_gender].astype(str).str.contains("女")].shape[0]
                        summary_text += f"     [性別統計] 男:{m} / 女:{f}\n"
                else:
                    summary_text += f"○ {name}: 0 人\n"
                summary_text += "\n"

            if not final_list:
                messagebox.showinfo("提示", "未找到相關工作表資料")
                return

            self.d5_df = pd.concat(final_list, ignore_index=True)
            
            self.d5_out.delete(1.0, tk.END)
            self.d5_out.insert(tk.END, f"【D5 員額分類統計完成】\n資料來源：{os.path.basename(self.file1_path)}\n" + "-"*50 + "\n")
            self.d5_out.insert(tk.END, summary_text)
            self.d5_out.insert(tk.END, "-"*50 + "\n預覽明細(前15筆)：\n")
            # 預覽顯示關鍵欄位
            preview_cols = [c for c in self.d5_df.columns if any(k in str(c) for k in ["姓名", "單位", "職稱", "D5類別", "來源工作表"])]
            self.d5_out.insert(tk.END, self.d5_df[preview_cols].head(15).to_string(index=False))
            
        except Exception as e:
            messagebox.showerror("計算失敗", str(e))

    def export_d5_excel(self):
        if not hasattr(self, 'd5_df'):
            messagebox.showwarning("提示", "請先點擊『開始計算D5報表』！")
            return
        
        out_name = self.d5_filename_e.get().strip()
        if not out_name: messagebox.showwarning("提示", "請輸入檔名"); return
        
        try:
            save_path = f"{out_name}.xlsx"
            # 將不同類別存入不同 Sheet
            with pd.ExcelWriter(save_path) as writer:
                for cat in ["相當簡薦委", "教師", "聘任"]:
                    df_sub = self.d5_df[self.d5_df['D5類別'] == cat]
                    if not df_sub.empty:
                        # 移除輔助欄位後匯出
                        export_cols = [c for c in df_sub.columns if c not in ['D5類別']]
                        df_sub[export_cols].to_excel(writer, sheet_name=cat, index=False)
            
            messagebox.showinfo("成功", f"D5 報表底稿已匯出至：\n{save_path}\n(包含 相當簡薦委、教師、聘任 等工作表)")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))
    # --- 分頁 5: 留職停薪 (V7.7) ---
    def setup_tab5(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="留職停薪人員名單提取系統", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        
        ttk.Label(main, text="說明：本功能將從『各系所教師』中篩選指定類別（新制職員、助教、教師、稀少性科技人員、研究人員）且『留職停薪』欄位有資料的人員。", foreground="blue", wraplength=800).pack(pady=(0, 15))

        # 檔名輸入
        file_name_f = ttk.Frame(main)
        file_name_f.pack(fill=tk.X, pady=5)
        ttk.Label(file_name_f, text="匯出檔案名稱:").pack(side=tk.LEFT, padx=5)
        self.unpaid_filename_e = ttk.Entry(file_name_f, width=30)
        self.unpaid_filename_e.insert(0, f"留職停薪名單_{datetime.now().strftime('%m%d')}")
        self.unpaid_filename_e.pack(side=tk.LEFT, padx=5)
        ttk.Label(file_name_f, text=".xlsx").pack(side=tk.LEFT)

        # 按鈕區
        btn_f = ttk.Frame(main)
        btn_f.pack(fill=tk.X, pady=10)
        ttk.Button(btn_f, text="🔍 開始提取留職停薪名單", command=self.gen_unpaid_leave_logic).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_f, text="📥 匯出 Excel (留職停薪 Sheet)", command=self.export_unpaid_leave_excel).pack(side=tk.LEFT, padx=5)

        self.unpaid_out = tk.Text(main, font=('Microsoft JhengHei', 11), bg="#fdf2f2")
        self.unpaid_out.pack(fill=tk.BOTH, expand=True, pady=10)

    def gen_unpaid_leave_logic(self):
        if not self.file1_path:
            messagebox.showwarning("提示", "請先在『職稱統計』分頁載入『各系所教師』檔案！")
            return
        
        target_sheets = ["新制職員", "助教", "教師", "稀少性科技人員", "研究人員"]
        try:
            xl = pd.ExcelFile(self.file1_path)
            found_dfs = []
            summary_info = "【留職停薪提取明細】\n"
            
            for sheet in target_sheets:
                if sheet in xl.sheet_names:
                    df = pd.read_excel(self.file1_path, sheet_name=sheet)
                    # 尋找包含 "留職停薪" 字樣的欄位
                    col = next((c for c in df.columns if "留職停薪" in str(c)), None)
                    
                    if col:
                        # 篩選該欄位非空且有文字的資料
                        mask = df[col].astype(str).str.strip().replace(['nan', 'None', '', 'nan '], pd.NA).notna()
                        filtered_df = df[mask].copy()
                        if not filtered_df.empty:
                            filtered_df['來源工作表'] = sheet # 註記來源以便辨識
                            found_dfs.append(filtered_df)
                            summary_info += f"● 工作表: {sheet:<12} | 找到 {len(filtered_df):>3} 筆\n"
                        else:
                            summary_info += f"○ 工作表: {sheet:<12} | 無符合紀錄\n"
                    else:
                        summary_info += f"× 工作表: {sheet:<12} | (找不到『留職停薪』欄位)\n"
                else:
                    summary_info += f"× 工作表: {sheet:<12} | (找不到此工作表)\n"
            
            if not found_dfs:
                self.unpaid_out.delete(1.0, tk.END)
                self.unpaid_out.insert(tk.END, summary_info + "\n【提取完成】未在指定工作表中找到任何『留職停薪』紀錄。")
                if hasattr(self, 'unpaid_df'): del self.unpaid_df
                return

            self.unpaid_df = pd.concat(found_dfs, ignore_index=True)
            
            self.unpaid_out.delete(1.0, tk.END)
            self.unpaid_out.insert(tk.END, f"【留職停薪名單提取完成】\n資料來源：{os.path.basename(self.file1_path)}\n\n")
            self.unpaid_out.insert(tk.END, summary_info)
            self.unpaid_out.insert(tk.END, "-"*50 + f"\n總計共：{len(self.unpaid_df)} 筆紀錄\n\n預覽名細：\n")
            
            # 顯示關鍵欄位預覽
            preview_cols = [c for c in self.unpaid_df.columns if any(k in str(c) for k in ["姓名", "職稱", "留職停薪", "來源工作表"])]
            self.unpaid_out.insert(tk.END, self.unpaid_df[preview_cols].to_string(index=False))
            
        except Exception as e:
            messagebox.showerror("提取失敗", str(e))
            
        except Exception as e:
            messagebox.showerror("提取失敗", str(e))

    def export_unpaid_leave_excel(self):
        if not hasattr(self, 'unpaid_df'):
            messagebox.showwarning("提示", "請先點擊『開始提取留職停薪名單』！")
            return
        
        out_name = self.unpaid_filename_e.get().strip()
        if not out_name: messagebox.showwarning("提示", "請輸入檔名"); return
        
        try:
            save_path = f"{out_name}.xlsx"
            # 根據需求，存入一個名為 "留職停薪" 的 sheet
            with pd.ExcelWriter(save_path) as writer:
                self.unpaid_df.to_excel(writer, sheet_name="留職停薪", index=False)
            messagebox.showinfo("成功", f"留職停薪人員名單已匯出至：\n{save_path}")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))
    def setup_tab6(self, f):
        main = ttk.Frame(f, padding="20"); main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="🌏 赴大陸地區 - 主管名單自動化統計中心", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        ttk.Label(main, text="說明：只需選擇一份包含主管名單的 Excel 檔案，系統將自動完成所有統計與報表更新。", foreground="blue").pack(pady=(0, 10))
        btn_f = ttk.LabelFrame(main, text="操作指令", padding=15); btn_f.pack(fill=tk.X, pady=5)
        ttk.Button(btn_f, text="📂 執行主管統計 (職等)", command=self.load_manager_list).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_f, text="👨‍🏫 計算教兼主管", command=self.process_teacher_managers).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_f, text="🧹 清空結果", command=lambda: self.t_out.delete(1.0, tk.END)).pack(side=tk.LEFT, padx=10)
        self.file_label_t = ttk.Label(btn_f, text="尚未選擇檔案", foreground="gray"); self.file_label_t.pack(side=tk.LEFT, padx=20)
        self.t_out = tk.Text(main, font=('Microsoft JhengHei', 11), bg="#fdfefe"); self.t_out.pack(fill=tk.BOTH, expand=True, pady=10)

    def load_manager_list(self):
        p = filedialog.askopenfilename(title="選擇主管名單 Excel 檔案", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not p: return
        target_file = p
        try:
            all_sheets_data = {}
            with pd.ExcelFile(target_file) as xl:
                sheet_name = next((s for s in xl.sheet_names if any(k in s for k in ["主管", "名單", "銝餌恣"])), xl.sheet_names[0])
                for s in xl.sheet_names: all_sheets_data[s] = pd.read_excel(xl, sheet_name=s)
            df = all_sheets_data[sheet_name]
            self.file_label_t.config(text=f"檔案：{os.path.basename(target_file)}", foreground="black")
            def find_col(keywords, default_idx):
                for i, col in enumerate(df.columns):
                    if any(k in str(col) for k in keywords): return i
                if len(df) > 0:
                    for idx, val in enumerate(df.iloc[0]):
                        if any(k in str(val) for k in keywords): return idx
                return default_idx
            name_idx = find_col(["姓名", "憪"], 2); rank_idx = find_col(["職等", "瑞"], 7); job_idx = find_col(["本職", "教師職稱", "單位"], 0)
            df_calc = df.copy()
            df_calc['rank_num'] = pd.to_numeric(df_calc.iloc[:, rank_idx], errors='coerce').fillna(0)
            df_calc['name_str'] = df_calc.iloc[:, name_idx].astype(str).str.strip()
            df_calc['orig_job'] = df_calc.iloc[:, job_idx].astype(str).str.strip().str.upper()
            # 統計全校兼行政職教師 (不論職等，只要教師職稱非 #N/A 且不重複)
            mask_teacher = (df_calc['orig_job'] != '#N/A') & (df_calc['orig_job'] != 'NAN') & (df_calc['rank_num'] > 0)
            df_teacher_mgrs = df_calc[mask_teacher].copy()
            # 依姓名去重，保留一筆
            df_unique_teacher_mgrs = df_teacher_mgrs.drop_duplicates(subset='name_str', keep='first')
            final_teacher_count = len(df_unique_teacher_mgrs)
            
            # 額外統計 11 職等以上之兼行政教師
            df_teacher_11plus = df_unique_teacher_mgrs[df_unique_teacher_mgrs['rank_num'] >= 11]
            final_teacher_11plus_count = len(df_teacher_11plus)

            # 全校主管職等分布統計 (包含所有行政人員)
            df_all_unique = df_calc.sort_values(by='rank_num', ascending=False).drop_duplicates(subset='name_str', keep='first')
            df_all_unique = df_all_unique[~df_all_unique['name_str'].isin(['nan', 'None', '', '姓名', '憪'])]
            rank_counts = df_all_unique['rank_num'].value_counts().sort_index(ascending=False).reset_index()
            rank_counts.columns = ['職等', '人數']
            
            all_sheets_data["主管名單_刪減計算用"] = df_calc[df_calc['rank_num'] > 0]
            all_sheets_data["依職等統計人數"] = rank_counts
            try:
                with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
                    for s_name, s_df in all_sheets_data.items(): s_df.to_excel(writer, sheet_name=s_name, index=False)
            except PermissionError:
                raise PermissionError("無法寫入檔案！請確認：\n1. 檔案是否已在 Excel 中關閉\n2. 請關閉檔案總管的『預覽窗格』\n3. 檔案未被其他程式佔用")
            
            self.t_out.delete(1.0, tk.END)
            self.t_out.insert(tk.END, f"【主管名單自動化統計報告】\n" + "="*60 + "\n")
            self.t_out.insert(tk.END, f"📁 檔案名稱    ： {os.path.basename(target_file)}\n")
            self.t_out.insert(tk.END, f"📄 使用分頁    ： {sheet_name}\n")
            self.t_out.insert(tk.END, "-"*60 + "\n")
            self.t_out.insert(tk.END, f"📊 兼行政職教師總人數 (含10職等以下) ： {final_teacher_count:>5} 名\n")
            self.t_out.insert(tk.END, f"📊 其中 11 職等以上之兼行政職教師     ： {final_teacher_11plus_count:>5} 名\n")
            self.t_out.insert(tk.END, "-"*60 + "\n")
            self.t_out.insert(tk.END, f"【全校主管職等分布統計】\n")
            for _, row_s in rank_counts.iterrows():
                if row_s['職等'] > 0: 
                    self.t_out.insert(tk.END, f"  - 第 {int(row_s['職等']):>2} 職等 ： {int(row_s['人數']):>4} 人\n")
            self.t_out.insert(tk.END, "="*60 + "\n✅ 統計完成！已成功更新 Excel 報表。\n")
            messagebox.showinfo("成功", f"統計完成！\n全校兼行政職教師總數：{final_teacher_count} 名\n(含 11 職等以上：{final_teacher_11plus_count} 名)")
        except PermissionError as pe: messagebox.showerror("檔案存取失敗", str(pe))
        except Exception as e: messagebox.showerror("統計失敗", f"錯誤原因：{str(e)}")

    def process_teacher_managers(self):
        p = filedialog.askopenfilename(title="選擇主管名單 Excel 檔案", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not p: return
        target_file = p
        try:
            all_sheets_data = {}
            with pd.ExcelFile(target_file) as xl:
                # 尋找「主管名單」Sheet
                sheet_name = next((s for s in xl.sheet_names if any(k in s for k in ["主管", "名單"])), xl.sheet_names[0])
                for s in xl.sheet_names: all_sheets_data[s] = pd.read_excel(xl, sheet_name=s)
            
            df_orig = all_sheets_data[sheet_name].copy()
            
            # 尋找關鍵欄位
            def find_col_idx(df, keywords):
                for i, col in enumerate(df.columns):
                    if any(k in str(col) for k in keywords): return col
                return None

            c_name = find_col_idx(df_orig, ["姓名", "憪"])
            c_job = find_col_idx(df_orig, ["教師職稱", "本職"])
            
            if not c_name or not c_job:
                messagebox.showerror("錯誤", "找不到『姓名』或『教師職稱』欄位，請檢查檔案格式。")
                return

            # 1. 刪除教師職稱為 #N/A 的 row
            # 使用 .str 存取器來執行 upper()，並處理可能的缺失值
            mask_na = df_orig[c_job].astype(str).str.upper().str.strip() == "#N/A"
            df_filtered = df_orig[~mask_na].copy()
            
            # 2. 刪除姓名重複的 row (保留第一筆)
            df_filtered = df_filtered.drop_duplicates(subset=c_name, keep='first')
            
            # 將結果存入新 Sheet
            new_sheet_name = "教兼主管名單(已過濾)"
            all_sheets_data[new_sheet_name] = df_filtered
            
            try:
                with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
                    for s_n, s_df in all_sheets_data.items():
                        s_df.to_excel(writer, sheet_name=s_n, index=False)
            except PermissionError:
                raise PermissionError("無法寫入檔案！請關閉 Excel 中的檔案。")

            self.t_out.delete(1.0, tk.END)
            self.t_out.insert(tk.END, f"【教兼主管名單處理報告】\n" + "="*50 + "\n")
            self.t_out.insert(tk.END, f"📁 檔案名稱    ： {os.path.basename(target_file)}\n")
            self.t_out.insert(tk.END, f"📄 原始總數量  ： {len(df_orig):>5} 筆\n")
            self.t_out.insert(tk.END, f"✅ 過濾後數量  ： {len(df_filtered):>5} 筆\n")
            self.t_out.insert(tk.END, f"✨ 新增工作表  ： {new_sheet_name}\n")
            self.t_out.insert(tk.END, "="*50 + "\n")
            messagebox.showinfo("成功", f"處理完成！\n已過濾出 {len(df_filtered)} 筆教兼主管名單。")
            
        except Exception as e:
            messagebox.showerror("處理失敗", str(e))

    def gen_travel_report(self):
        if not hasattr(self, 'df_travel') or self.df_travel is None:
            messagebox.showwarning("提示", "請先載入赴大陸資料檔案！")
            return
        
        try:
            total = len(self.df_travel)
            self.t_out.delete(1.0, tk.END)
            self.t_out.insert(tk.END, f"【赴大陸地區人次統計報告】\n")
            self.t_out.insert(tk.END, f"統計日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
            self.t_out.insert(tk.END, "-"*40 + "\n")
            self.t_out.insert(tk.END, f"總人次統計： {total} 人次\n")
            
            # 嘗試做簡單的單位統計 (假設第一欄是單位)
            cols = self.df_travel.columns
            unit_col = next((c for c in cols if "單位" in str(c) or "系所" in str(c)), cols[0])
            
            if unit_col:
                counts = self.df_travel[unit_col].value_counts().head(10)
                self.t_out.insert(tk.END, "\n[前十大熱門單位統計]\n")
                for u, c in counts.items():
                    self.t_out.insert(tk.END, f" - {str(u):<20}: {c} 人次\n")
            
        except Exception as e:
            messagebox.showerror("統計失敗", str(e))

  # --- 分頁 7: 個資申請單 ---
    def setup_tab7(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="個人資料使用申請 - Excel 快速匯出", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        cat_f = ttk.LabelFrame(main, text="1. 勾選類別", padding=10); cat_f.pack(fill=tk.X, pady=5)
        self.cat_vars = {}
        cat_grid = ttk.Frame(cat_f); cat_grid.pack(fill=tk.X)
        categories = ["新制職員", "助教", "技工工友合計", "教師", "約聘教師", "駐衛警察隊", "專案教學人員", "講座教授", "客座人員", "醫事人員", "稀少性科技人員"]
        for i, cat in enumerate(categories):
            var = tk.BooleanVar(); self.cat_vars[cat] = var
            ttk.Checkbutton(cat_grid, text=cat, variable=var).grid(row=i//6, column=i%6, sticky=tk.W, padx=10, pady=2)
        item_f = ttk.LabelFrame(main, text="2. 勾選項目", padding=10); item_f.pack(fill=tk.X, pady=5)
        self.item_vars = {}
        item_grid = ttk.Frame(item_f); item_grid.pack(fill=tk.X)
        data_items = ["姓名", "服務單位", "一級單位", "身分證統一編號", "職稱", "性別", "薪點", "最高學位", "出生年月日", "到校日期", "聘期起", "聘期迄", "年齡", "服務本校年資", "EMI", "本校電子信箱", "校外電子信箱", "留職停薪", "帶職帶薪","國籍","雙重國籍"]
        for i, item in enumerate(data_items):
            var = tk.BooleanVar(); self.item_vars[item] = var
            ttk.Checkbutton(item_grid, text=item, variable=var).grid(row=i//4, column=i%4, sticky=tk.W, padx=10, pady=2)
        export_f = ttk.Frame(main); export_f.pack(fill=tk.X, pady=5)
        ttk.Label(export_f, text="匯出檔名:").pack(side=tk.LEFT)
        self.app_filename_e = ttk.Entry(export_f, width=30); self.app_filename_e.insert(0, "個資申請名單"); self.app_filename_e.pack(side=tk.LEFT, padx=5)
        ttk.Button(main, text="👁️ 預覽", command=self.preview_app_data).pack(side=tk.LEFT, padx=10)
        ttk.Button(main, text="📥 匯出 Excel", command=self.gen_app_excel).pack(side=tk.LEFT, padx=10)
        self.app_out = tk.Text(main, height=10, font=('Microsoft JhengHei', 10), bg="#fffcf0"); self.app_out.pack(fill=tk.BOTH, expand=True, pady=10)

    def get_app_core_data(self):
        if not self.file1_path: return None, None, None
        sel_cats = [c for c, v in self.cat_vars.items() if v.get()]
        sel_items = [i for i, v in self.item_vars.items() if v.get()]
        if not sel_cats or not sel_items: return None, None, None
        try:
            xl = pd.ExcelFile(self.file1_path); all_dfs = []
            for c in sel_cats:
                if c in xl.sheet_names: all_dfs.append(pd.read_excel(self.file1_path, sheet_name=c))
            if not all_dfs: return None, None, None
            df = pd.concat(all_dfs, ignore_index=True)
            c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), df.columns[0])
            df = df[df[c_t].astype(str).str.strip().replace(['nan','','None'], pd.NA).notna()]
            col_map = {
                "姓名": ["姓名", "憪"], 
                "服務單位": ["單位", "系所"], 
                "身分證統一編號": ["身分證", "統一編號", "ID"], 
                "職稱": [c_t], 
                "聘期起": ["聘期起"],
                "聘期迄": ["聘期迄"],
                "年齡": ["年齡"],
                "服務本校年資": ["服務本校年資"],
                "留職停薪": ["留職停薪"],
                "帶職帶薪": ["帶職帶薪"]
            }
            found = []
            for it in sel_items:
                act = next((c for c in df.columns if any(a in str(c) for a in col_map.get(it, [it]))), None)
                if act: found.append((it, act))
            return df, found, sel_cats
        except Exception: return None, None, None

    def preview_app_data(self):
        df, cols, cats = self.get_app_core_data()
        if df is not None:
            self.app_out.delete(1.0, tk.END)
            for _, r in df.head(20).iterrows(): self.app_out.insert(tk.END, " | ".join([str(r.get(c, "")) for _, c in cols]) + "\n")

    def gen_app_excel(self):
        df, cols, cats = self.get_app_core_data()
        if df is not None:
            try:
                save_p = f"{self.app_filename_e.get()}.xlsx"; df[[c for _, c in cols]].to_excel(save_p, index=False)
                messagebox.showinfo("成功", f"匯出至 {save_p}")
            except Exception as e: messagebox.showerror("錯誤", str(e))


    # --- 分頁 8: 人事異動單 ---
    def setup_tab8(self, f):
        self.tab8_sheet_vars = {}
        self.tab8_file_path = None
        
        main = ttk.Frame(f, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        # 建立左右分欄
        left_panel = ttk.Frame(main, width=350)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        
        right_panel = ttk.Frame(main)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- 左側：控制區域 ---
        ttk.Label(left_panel, text="⚙️ 設定與篩選", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5, anchor=tk.W)
        
        # 檔案載入
        file_f = ttk.LabelFrame(left_panel, text="1. 檔案載入 (總異動)", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        
        btn_f = ttk.Frame(file_f)
        btn_f.pack(fill=tk.X)
        ttk.Button(btn_f, text="📂 選擇檔案", command=self.load_file_changes).pack(side=tk.LEFT)
        ttk.Button(btn_f, text="🔄 Reload", command=self.reload_file_changes).pack(side=tk.LEFT, padx=5)
        
        self.tab8_file_label = ttk.Label(file_f, text="尚未載入檔案", foreground="gray", wraplength=300)
        self.tab8_file_label.pack(pady=5, anchor=tk.W)

        # 工作表勾選區
        self.sheet_frame8 = ttk.LabelFrame(left_panel, text="2. 選擇工作表", padding=10)
        self.sheet_frame8.pack(fill=tk.BOTH, expand=True, pady=5)
        
        toggle_f = ttk.Frame(self.sheet_frame8)
        toggle_f.pack(fill=tk.X, pady=(0,5))
        ttk.Button(toggle_f, text="✅ 全選", command=lambda: self.set_all_sheets_changes(True)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,2))
        ttk.Button(toggle_f, text="❌ 全不選", command=lambda: self.set_all_sheets_changes(False)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2,0))

        self.sheet_canvas8 = tk.Canvas(self.sheet_frame8, highlightthickness=0)
        self.sheet_scroll8 = ttk.Scrollbar(self.sheet_frame8, orient="vertical", command=self.sheet_canvas8.yview)
        self.sheet_container8 = ttk.Frame(self.sheet_canvas8)
        
        self.sheet_container8.bind("<Configure>", lambda e: self.sheet_canvas8.configure(scrollregion=self.sheet_canvas8.bbox("all")))
        self.sheet_canvas8.create_window((0,0), window=self.sheet_container8, anchor="nw")
        self.sheet_canvas8.configure(yscrollcommand=self.sheet_scroll8.set)
        
        self.sheet_canvas8.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sheet_scroll8.pack(side=tk.RIGHT, fill=tk.Y)

        # 搜尋條件
        sf = ttk.LabelFrame(left_panel, text="3. 輸入搜尋條件", padding=10)
        sf.pack(fill=tk.X, pady=5)
        
        grid_f = ttk.Frame(sf)
        grid_f.pack(fill=tk.X)
        
        ttk.Label(grid_f, text="生效日期:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.tab8_date_e = ttk.Entry(grid_f, width=20)
        self.tab8_date_e.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(grid_f, text="姓名:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.tab8_name_e = ttk.Entry(grid_f, width=20)
        self.tab8_name_e.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(grid_f, text="關鍵字:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.tab8_kw_e = ttk.Entry(grid_f, width=20)
        self.tab8_kw_e.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        self.tab8_search_btn = ttk.Button(left_panel, text="🚀 執行篩選查詢", command=self.search_changes_logic)
        self.tab8_search_btn.pack(fill=tk.X, pady=10)

        # --- 右側：結果顯示區域 ---
        ttk.Label(right_panel, text="📋 異動紀錄結果 (依據搜尋條件篩選)", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5, anchor=tk.W)
        
        table_f = ttk.Frame(right_panel)
        table_f.pack(fill=tk.BOTH, expand=True)
        
        # 重新定義顯示欄位，符合使用者要求的 8 欄位格式
        cols = ("姓名", "原單位", "原職稱", "動態", "新單位", "新職稱", "生效日期", "備註")
        self.tree8 = ttk.Treeview(table_f, columns=cols, show='headings')
        vsb8 = ttk.Scrollbar(table_f, orient="vertical", command=self.tree8.yview)
        hsb8 = ttk.Scrollbar(table_f, orient="horizontal", command=self.tree8.xview)
        self.tree8.configure(yscrollcommand=vsb8.set, xscrollcommand=hsb8.set)
        
        # 設定欄位標頭與寬度
        column_widths = {
            "姓名": 80, "原單位": 150, "原職稱": 100, "動態": 120, 
            "新單位": 150, "新職稱": 100, "生效日期": 100, "備註": 150
        }
        for c in cols:
            self.tree8.heading(c, text=c)
            self.tree8.column(c, width=column_widths.get(c, 100), anchor=tk.CENTER)
        
        self.tree8.grid(row=0, column=0, sticky='nsew')
        vsb8.grid(row=0, column=1, sticky='ns')
        hsb8.grid(row=1, column=0, sticky='ew')
        
        table_f.grid_rowconfigure(0, weight=1)
        table_f.grid_columnconfigure(0, weight=1)

    def load_file_changes(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not p:
            return
        self.tab8_file_path = p
        self.tab8_file_label.config(text=os.path.basename(p), foreground="black")
        self.reload_file_changes()

    def reload_file_changes(self):
        if not self.tab8_file_path:
            return
        try:
            xl = pd.ExcelFile(self.tab8_file_path)
            for w in self.sheet_container8.winfo_children():
                w.destroy()
            self.tab8_sheet_vars = {}
            for name in xl.sheet_names:
                var = tk.BooleanVar(value=True) # 預設全選
                self.tab8_sheet_vars[name] = var
                cb = ttk.Checkbutton(self.sheet_container8, text=name, variable=var)
                cb.pack(fill=tk.X, padx=5, pady=1)
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取失敗: {e}")

    def set_all_sheets_changes(self, val):
        if not self.tab8_sheet_vars:
            return
        for v in self.tab8_sheet_vars.values():
            v.set(val)

    def search_changes_logic(self):
        if not self.tab8_file_path:
            messagebox.showwarning("提示", "請先載入檔案")
            return
        
        sel_sheets = [s for s, v in self.tab8_sheet_vars.items() if v.get()]
        if not sel_sheets:
            messagebox.showwarning("提示", "請至少勾選一個工作表")
            return
            
        d_val = self.tab8_date_e.get().strip()
        n_val = self.tab8_name_e.get().strip()
        k_val = self.tab8_kw_e.get().strip()
        
        for i in self.tree8.get_children():
            self.tree8.delete(i)
            
        try:
            fname = os.path.basename(self.tab8_file_path)
            # 判斷是否為特殊格式檔案 (含標題列需跳過者)
            is_special_fmt = any(k in fname for k in ["活頁簿", "總異動", "114"])
            
            for s in sel_sheets:
                if is_special_fmt:
                    # 讀取整張表以判斷結構
                    df_full = pd.read_excel(self.tab8_file_path, sheet_name=s, header=None)
                    if df_full.empty: continue
                    
                    # 判斷資料起始位置 (通常第 5 列開始為資料，即 index 4)
                    start_row = 4
                    if len(df_full) > 4:
                        if pd.isna(df_full.iloc[3, 0]) and not pd.isna(df_full.iloc[4, 0]):
                            start_row = 4
                        elif not pd.isna(df_full.iloc[3, 0]):
                            start_row = 3
                    
                    data_df = df_full.iloc[start_row:].copy()
                    # 排除全空的列 (確保姓名或動態至少有一個有值)
                    data_df = data_df[data_df[0].notna() | data_df[3].notna()]
                    
                    mask = pd.Series(True, index=data_df.index)
                    if n_val:
                        mask &= data_df[0].astype(str).str.contains(n_val, na=False)
                    
                    # 日期篩選 (針對特殊格式合併 7,8,9 欄位)
                    def get_date_str(row):
                        parts = [str(row.get(i, "")).strip() for i in [7, 8, 9] if not pd.isna(row.get(i))]
                        return "".join([p for p in parts if p.isdigit() or (len(p)>0 and p != "nan")])

                    if d_val:
                        mask &= data_df.apply(lambda r: d_val in get_date_str(r), axis=1)

                    if k_val:
                        k_mask = data_df.astype(str).apply(lambda x: x.str.contains(k_val, na=False)).any(axis=1)
                        mask &= k_mask
                    
                    res = data_df[mask]
                    for _, r in res.iterrows():
                        def gv(idx): 
                            val = r.get(idx, "")
                            if pd.isna(val): return "-"
                            if isinstance(val, float) and val.is_integer(): return str(int(val))
                            return str(val).strip().replace("\n", " ")
                        
                        vals = [
                            gv(0),      # 姓名
                            gv(1),      # 原單位
                            gv(2),      # 原職稱
                            gv(3),      # 動態
                            gv(4),      # 新單位
                            gv(6),      # 新職稱
                            get_date_str(r), # 生效日期
                            gv(11)      # 備註
                        ]
                        self.tree8.insert("", tk.END, values=vals)
                    continue 

                # --- 一般通用 Excel 邏輯 ---
                df = pd.read_excel(self.tab8_file_path, sheet_name=s)
                if df.empty: continue
                
                cols = df.columns.astype(str)
                c_name = next((c for c in cols if "姓名" in c), None)
                c_orig_u = next((c for c in cols if "原" in c and "單位" in c), None)
                c_orig_t = next((c for c in cols if "原" in c and "職" in c), None)
                c_status = next((c for c in cols if any(k in c for k in ["動態", "異動", "原因"])), None)
                c_new_u = next((c for c in cols if "新" in c and "單位" in c), None)
                c_new_t = next((c for c in cols if "新" in c and "職" in c), None)
                c_date = next((c for c in cols if any(k in c for k in ["生效", "日期"])), None)
                c_rem = next((c for c in cols if "備註" in c), None)
                
                mask = pd.Series(True, index=df.index)
                if n_val and c_name:
                    mask &= df[c_name].astype(str).str.contains(n_val, na=False)
                if d_val and c_date:
                    mask &= df[c_date].astype(str).str.contains(d_val, na=False)
                if k_val:
                    k_mask = df.astype(str).apply(lambda x: x.str.contains(k_val, na=False)).any(axis=1)
                    mask &= k_mask
                
                res = df[mask]
                for _, r in res.iterrows():
                    def gv_gen(col): return str(r.get(col, "")).replace("nan", "-").strip() if col else "-"
                    vals = [
                        gv_gen(c_name), gv_gen(c_orig_u), gv_gen(c_orig_t), 
                        gv_gen(c_status), gv_gen(c_new_u), gv_gen(c_new_t), 
                        gv_gen(c_date), gv_gen(c_rem)
                    ]
                    self.tree8.insert("", tk.END, values=vals)
                    
        except Exception as e:
            messagebox.showerror("搜尋出錯", f"錯誤原因：{str(e)}\n請檢查檔案結構是否符合規範。")

    # --- 分頁 9: 校區人數統計 ---
    def setup_tab9(self, f):
        self.tab9_sheet_vars = {}
        self.tab9_campus_vars = {}
        
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="校區各類人員人數統計 (各系所教師)", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        
        # 建立左右分欄
        content_f = ttk.Frame(main)
        content_f.pack(fill=tk.BOTH, expand=True)
        
        left_f = ttk.Frame(content_f, width=400)
        left_f.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        
        right_f = ttk.Frame(content_f)
        right_f.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # --- 左側：控制區域 ---
        # 1. 校區選擇
        campus_f = ttk.LabelFrame(left_f, text="1. 勾選校區 (可多選)", padding=10)
        campus_f.pack(fill=tk.X, pady=5)
        campuses = ["林口校區", "校本部", "公館校區"]
        for cp in campuses:
            var = tk.BooleanVar(value=(cp == "林口校區"))
            self.tab9_campus_vars[cp] = var
            ttk.Checkbutton(campus_f, text=cp, variable=var).pack(side=tk.LEFT, padx=10)

        # 2. 工作表選擇
        sheet_labelf = ttk.LabelFrame(left_f, text="2. 勾選工作表 (可多選)", padding=10)
        sheet_labelf.pack(fill=tk.BOTH, expand=True, pady=5)
        
        btn_sheet_f = ttk.Frame(sheet_labelf)
        btn_sheet_f.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(btn_sheet_f, text="✅ 全選", command=lambda: self.set_tab9_sheets(True)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_sheet_f, text="❌ 全不選", command=lambda: self.set_tab9_sheets(False)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_sheet_f, text="🔄 載入/重新整理工作表", command=self.refresh_tab9_sheets).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        self.sheet_canvas9 = tk.Canvas(sheet_labelf, highlightthickness=0)
        self.sheet_scroll9 = ttk.Scrollbar(sheet_labelf, orient="vertical", command=self.sheet_canvas9.yview)
        self.sheet_container9 = ttk.Frame(self.sheet_canvas9)
        self.sheet_container9.bind("<Configure>", lambda e: self.sheet_canvas9.configure(scrollregion=self.sheet_canvas9.bbox("all")))
        self.sheet_canvas9.create_window((0,0), window=self.sheet_container9, anchor="nw")
        self.sheet_canvas9.configure(yscrollcommand=self.sheet_scroll9.set)
        self.sheet_canvas9.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sheet_scroll9.pack(side=tk.RIGHT, fill=tk.Y)

        self.calc_btn9 = ttk.Button(left_f, text="📊 開始統計校區人數", command=self.calculate_campus_stats)
        self.calc_btn9.pack(fill=tk.X, pady=10)

        # --- 右側：結果區域 ---
        ttk.Label(right_f, text="📋 統計結果報告", font=('Microsoft JhengHei', 12, 'bold')).pack(pady=5, anchor=tk.W)
        self.campus_out = tk.Text(right_f, font=('Microsoft JhengHei', 11), bg="#f0f8ff")
        self.campus_out.pack(fill=tk.BOTH, expand=True)

    def set_tab9_sheets(self, val):
        for v in self.tab9_sheet_vars.values():
            v.set(val)

    def refresh_tab9_sheets(self):
        if not self.file1_path:
            return
        try:
            xl = pd.ExcelFile(self.file1_path)
            for w in self.sheet_container9.winfo_children():
                w.destroy()
            self.tab9_sheet_vars = {}
            # 預設勾選常見的統計工作表
            default_targets = ["新制職員", "助教", "教師", "約聘教師", "專案教學人員", "講座教授", "客座人員", "醫事人員", "稀少性科技人員", "技工工友合計"]
            for name in xl.sheet_names:
                var = tk.BooleanVar(value=(name in default_targets))
                self.tab9_sheet_vars[name] = var
                cb = ttk.Checkbutton(self.sheet_container9, text=name, variable=var)
                cb.pack(fill=tk.X, padx=5, pady=1)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取工作表: {e}")

    def calculate_campus_stats(self):
        if not self.file1_path:
            messagebox.showwarning("提示", "請先在『職稱統計』分頁載入『各系所教師』檔案！")
            return
        
        sel_campuses = [cp for cp, v in self.tab9_campus_vars.items() if v.get()]
        sel_sheets = [s for s, v in self.tab9_sheet_vars.items() if v.get()]
        
        if not sel_campuses:
            messagebox.showwarning("提示", "請至少勾選一個校區！")
            return
        if not sel_sheets:
            messagebox.showwarning("提示", "請至少勾選一個工作表！")
            return
        
        try:
            xl = pd.ExcelFile(self.file1_path)
            self.campus_out.delete(1.0, tk.END)
            self.campus_out.insert(tk.END, f"【校區人數統計報告】\n")
            self.campus_out.insert(tk.END, f"資料來源：{os.path.basename(self.file1_path)}\n")
            self.campus_out.insert(tk.END, f"統計日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
            self.campus_out.insert(tk.END, "="*60 + "\n")
            
            grand_total = 0
            
            for target_campus in sel_campuses:
                self.campus_out.insert(tk.END, f"📍 校區：{target_campus}\n" + "-"*40 + "\n")
                campus_total = 0
                
                for sheet in sel_sheets:
                    if sheet in xl.sheet_names:
                        df = pd.read_excel(self.file1_path, sheet_name=sheet)
                        
                        # 過濾無職稱的空列
                        c_title = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                        if c_title:
                            df = df[df[c_title].astype(str).str.strip().replace(['nan','','None'], pd.NA).notna()]
                        
                        # 搜尋校區欄位
                        campus_col = next((c for c in df.columns if "校區" in str(c)), None)
                        if not campus_col:
                            campus_col = next((c for c in df.columns if any(k in str(c) for k in ["地點", "單位", "系所"])), None)
                        
                        count = 0
                        if campus_col:
                            mask = df[campus_col].astype(str).str.contains(target_campus, na=False)
                            count = mask.sum()
                        
                        # 如果精確欄位沒找到，嘗試全表模糊搜尋 (保險機制)
                        if count == 0:
                            mask_any = df.astype(str).apply(lambda x: x.str.contains(target_campus)).any(axis=1)
                            count = mask_any.sum()
                        
                        if count > 0:
                            self.campus_out.insert(tk.END, f" ● {sheet:<15} | {count:>5} 名\n")
                            campus_total += count
                    else:
                        self.campus_out.insert(tk.END, f" ○ {sheet:<15} | (找不到工作表)\n")
                
                self.campus_out.insert(tk.END, f" > {target_campus} 小計：{campus_total} 名\n\n")
                grand_total += campus_total
            
            self.campus_out.insert(tk.END, "="*60 + f"\n總計選定校區總人數：{grand_total} 名\n")
                
        except Exception as e:
            messagebox.showerror("統計失敗", str(e))

if __name__ == "__main__":
    root = tk.Tk(); app = TeacherStatsApp(root); root.mainloop()

  
