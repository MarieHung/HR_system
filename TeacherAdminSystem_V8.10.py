import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime

class TeacherStatsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("教師資料管理中心 - 行政綜合業務系統 (V7.6)")
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

        # 建立 7 個主分頁
        tabs = [" 🔍 職稱統計 ", " 📅 到離查詢 ", " 📊 給副校長 ", " 📄 D5 報表 ", " 🛌 留職停薪 ", " 🌏 赴大陸統計 ", " 📝 個資申請單 "]
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
        
        cols = ("date", "name", "unit", "level1", "reason", "remarks")
        self.tree2 = ttk.Treeview(table_f, columns=cols, show='headings')
        vsb2 = ttk.Scrollbar(table_f, orient="vertical", command=self.tree2.yview)
        hsb2 = ttk.Scrollbar(table_f, orient="horizontal", command=self.tree2.xview)
        self.tree2.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        
        for c, h in zip(cols, ["日期", "姓名", "單位", "一級單位", "異動原因", "備註"]): 
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
                vals = [r.get(cd, ""), r.get(cn, ""), r.get("單位", ""), r.get("一級單位", ""), r.get(cr, ""), r.get("備註", "")]
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
                    
                    df['D5類別'] = "相當簡薦委"
                    df['來源工作表'] = s
                    cat_data["相當簡薦委"].append(df)

            # 2. 教師 & 3. 聘任 (處理 助教、教師、專案教學人員、研究人員)
            # 教師工作表
            if "教師" in all_sheets:
                df = pd.read_excel(self.file1_path, sheet_name="教師")
                # 排除職稱為空白的資料
                c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                if c_t:
                    df = df[df[c_t].astype(str).str.strip().replace(['nan','','None','nan '], pd.NA).notna()]
                
                df['D5類別'] = "教師"
                df['來源工作表'] = "教師"
                cat_data["教師"].append(df)
            
            # 專案教學人員 (專技) & 研究人員 -> 歸類為 聘任
            for s in ["專案教學人員", "研究人員"]:
                if s in all_sheets:
                    df = pd.read_excel(self.file1_path, sheet_name=s)
                    # 排除職稱為空白的資料
                    c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                    if c_t:
                        df = df[df[c_t].astype(str).str.strip().replace(['nan','','None','nan '], pd.NA).notna()]
                    
                    df['D5類別'] = "聘任"
                    df['來源工作表'] = s
                    cat_data["聘任"].append(df)

            # 助教工作表 (需區分新制/舊制助教)
            if "助教" in all_sheets:
                df = pd.read_excel(self.file1_path, sheet_name="助教")
                # 排除職稱為空白的資料
                c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                if c_t:
                    df = df[df[c_t].astype(str).str.strip().replace(['nan','','None','nan '], pd.NA).notna()]
                
                # 尋找「舊制講師或助教」欄位
                c_check = next((c for c in df.columns if "舊制講師或助教" in str(c)), None)
                
                if c_check:
                    # 1. 判斷是否包含「舊制助教」字樣 -> 歸類為 教師
                    is_old_mask = df[c_check].astype(str).str.contains("舊制助教")
                    df_old = df[is_old_mask].copy()
                    df_old['D5類別'] = "教師"
                    df_old['來源工作表'] = "助教(舊制)"
                    if not df_old.empty: cat_data["教師"].append(df_old)
                    
                    # 2. 判斷是否包含「新制助教」字樣 -> 歸類為 聘任
                    is_new_mask = df[c_check].astype(str).str.contains("新制助教")
                    df_new = df[is_new_mask].copy()
                    df_new['D5類別'] = "聘任"
                    df_new['來源工作表'] = "助教(新制)"
                    if not df_new.empty: cat_data["聘任"].append(df_new)
                else:
                    # 若找不到該欄位，預設歸類為聘任
                    df['D5類別'] = "聘任"
                    df['來源工作表'] = "助教(未區分)"
                    cat_data["聘任"].append(df)

            # 合併所有資料
            final_list = []
            summary_text = ""
            for name, dfs in cat_data.items():
                if dfs:
                    combined = pd.concat(dfs, ignore_index=True)
                    final_list.append(combined)
                    
                    if name == "相當簡薦委":
                        # 針對相當簡薦委進行官等與性別統計
                        summary_text += f"● {name:<10}: {len(combined):>4} 人\n"
                        # 偵測欄位
                        c_rank = next((c for c in combined.columns if "官等" in str(c)), None)
                        c_gender = next((c for c in combined.columns if "性別" in str(c)), None)
                        
                        if c_rank and c_gender:
                            # 取得所有唯一的官等
                            ranks = combined[c_rank].astype(str).str.strip().unique()
                            for r in sorted(ranks):
                                if r in ['nan', 'None', '']: continue
                                mask_r = (combined[c_rank].astype(str).str.strip() == r)
                                m_cnt = combined[mask_r & (combined[c_gender].astype(str).str.contains("男"))].shape[0]
                                f_cnt = combined[mask_r & (combined[c_gender].astype(str).str.contains("女"))].shape[0]
                                summary_text += f"   - {r:<6} | 男：{m_cnt:>3} / 女：{f_cnt:>3}\n"
                        else:
                            summary_text += "   (找不到『官等』或『性別』欄位，無法進行細分統計)\n"
                    else:
                        # 針對教師與聘任進行性別統計
                        summary_text += f"● {name:<10}: {len(combined):>4} 人\n"
                        c_gender = next((c for c in combined.columns if "性別" in str(c)), None)
                        if c_gender:
                            m_cnt = combined[combined[c_gender].astype(str).str.contains("男")].shape[0]
                            f_cnt = combined[combined[c_gender].astype(str).str.contains("女")].shape[0]
                            summary_text += f"   - 總計 | 男：{m_cnt:>3} / 女：{f_cnt:>3}\n"
                        else:
                            summary_text += "   (找不到『性別』欄位)\n"
                else:
                    summary_text += f"● {name:<10}:    0 人\n"

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
            self.d5_df.to_excel(save_path, index=False)
            messagebox.showinfo("成功", f"D5 底稿已匯出：\n{save_path}")
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
            
            if not found_dfs:
                self.unpaid_out.delete(1.0, tk.END)
                self.unpaid_out.insert(tk.END, "【提取完成】未在指定工作表中找到任何『留職停薪』欄位有文字的記錄。")
                if hasattr(self, 'unpaid_df'): del self.unpaid_df
                return

            self.unpaid_df = pd.concat(found_dfs, ignore_index=True)
            
            self.unpaid_out.delete(1.0, tk.END)
            self.unpaid_out.insert(tk.END, f"【留職停薪名單提取完成】\n資料來源：{os.path.basename(self.file1_path)}\n共計：{len(self.unpaid_df)} 筆符合記錄\n" + "-"*50 + "\n")
            self.unpaid_out.insert(tk.END, self.unpaid_df.to_string())
            
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
        ttk.Button(main, text="載入 赴大陸資料", command=self.load_file_travel).pack()
        self.file_label_t = ttk.Label(main, text="尚未載入"); self.file_label_t.pack()
        ttk.Button(main, text="開始統計", command=self.gen_travel_report).pack()
        self.t_out = tk.Text(main, font=('Microsoft JhengHei', 11)); self.t_out.pack(fill=tk.BOTH, expand=True)

    def load_file_travel(self):
        p = filedialog.askopenfilename()
        if not p: return
        try:
            self.df_travel = pd.read_excel(p); self.file_label_t.config(text=os.path.basename(p))
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
    def gen_travel_report(self):
        if self.df_travel is not None:
            self.t_out.delete(1.0, tk.END); self.t_out.insert(tk.END, f"赴大陸總人次：{len(self.df_travel)}")

    # --- 分頁 7: 個資申請單 ---
    def setup_tab7(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="個人資料使用申請 - Excel 快速匯出", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        cat_f = ttk.LabelFrame(main, text="1. 勾選類別", padding=10); cat_f.pack(fill=tk.X, pady=5)
        self.cat_vars = {}
        cat_grid = ttk.Frame(cat_f); cat_grid.pack(fill=tk.X)
        categories = ["新制職員", "助教", "技工工友合計", "教師", "約聘教師", "駐衛警察隊", "專案教學人員", "講座教授", "客座教授", "醫事人員", "稀少性科技人員"]
        for i, cat in enumerate(categories):
            var = tk.BooleanVar(); self.cat_vars[cat] = var
            ttk.Checkbutton(cat_grid, text=cat, variable=var).grid(row=i//6, column=i%6, sticky=tk.W, padx=10, pady=2)
        item_f = ttk.LabelFrame(main, text="2. 勾選項目", padding=10); item_f.pack(fill=tk.X, pady=5)
        self.item_vars = {}
        item_grid = ttk.Frame(item_f); item_grid.pack(fill=tk.X)
        data_items = ["姓名", "服務單位", "一級單位", "身分證統一編號", "職稱", "性別", "薪點", "最高學位", "出生年月日", "到校日期", "服務本校年資", "EMI", "本校電子信箱", "校外電子信箱"]
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
            col_map = {"姓名": ["姓名", "憪"], "服務單位": ["單位", "系所"], "身分證統一編號": ["身分證", "統一編號", "ID"], "職稱": [c_t], "服務本校年資": ["年資", "本校年資"]}
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

if __name__ == "__main__":
    root = tk.Tk(); app = TeacherStatsApp(root); root.mainloop()
