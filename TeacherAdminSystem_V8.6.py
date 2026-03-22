import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime

class TeacherStatsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("教師資料管理中心 - 行政綜合業務系統 (V7.5)")
        self.root.geometry("1150x950")

        # 全域資料變數
        self.file1_path = None # 各系所教師.xlsx 的路徑
        self.df1_combined = None # 職稱統計用的合併資料
        self.df2 = None  
        self.df_travel = None 
        
        self.title_col1 = None

        # 主 Notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 建立 7 個主分頁
        tabs = [" 🔍 職稱統計 ", " 📅 到離查詢 ", " 📊 給副校長 ", " 📄 D5 報表 ", " 💰 給主計室 ", " 🌏 赴大陸統計 ", " 📝 個資申請單 "]
        self.tab_frames = []
        for text in tabs:
            f = ttk.Frame(self.notebook)
            self.notebook.add(f, text=text)
            self.tab_frames.append(f)

        self.setup_tab1(self.tab_frames[0])
        self.setup_tab2(self.tab_frames[1])
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
        ttk.Button(file_f, text="選擇檔案", command=self.load_file1_v7).pack(side=tk.LEFT)
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
            
        self.select_f1 = ttk.LabelFrame(main, text="2. 勾選職稱 (清單自動更新)", padding=10)
        self.select_f1.pack(fill=tk.BOTH, expand=True, pady=10)
        self.canvas = tk.Canvas(self.select_f1, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self.select_f1, command=self.canvas.yview)
        self.scroll_f = ttk.Frame(self.canvas)
        self.scroll_f.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scroll_f, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")
        
        self.calc_btn1 = ttk.Button(main, text="📊 開始統計人數", command=self.calculate1_v7, state=tk.DISABLED)
        self.calc_btn1.pack(pady=5)
        
        # 擴大輸出結果視窗 (高度增加至 15)
        self.res_text1 = tk.Text(main, height=15, font=('Consolas', 11), bg="#fdfdfd", state=tk.DISABLED)
        self.res_text1.pack(fill=tk.BOTH, expand=True, pady=5)

    def load_file1_v7(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p:
            return
        self.file1_path = p
        self.file_label1.config(text=os.path.basename(p), foreground="black")
        self.refresh_tab1_titles()

    def refresh_tab1_titles(self):
        if not self.file1_path:
            return
        sel_cats = [cat for cat, var in self.tab1_cat_vars.items() if var.get()]
        if not sel_cats:
            for w in self.scroll_f.winfo_children():
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
            self.df1_combined = pd.concat(all_dfs, ignore_index=True)
            self.title_col1 = next((c for c in self.df1_combined.columns if any(k in str(c) for k in ['職稱', '職別'])), self.df1_combined.columns[0])
            titles = self.df1_combined[self.title_col1].astype(str).str.strip().replace(['nan','None',''], pd.NA).dropna().unique()
            for w in self.scroll_f.winfo_children():
                w.destroy()
            self.tab1_title_vars = {}
            for i, t in enumerate(sorted(titles)):
                var = tk.BooleanVar()
                self.tab1_title_vars[t] = var
                cb = ttk.Checkbutton(self.scroll_f, text=t, variable=var)
                # 使用 Grid 實現橫向排列，每列 4 個
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
        for t in sel:
            c = cnt.get(t, 0)
            self.res_text1.insert(tk.END, f"{t:<23} | {c:>5} 名\n")
            total += c
        self.res_text1.insert(tk.END, "-"*35 + f"\n總計人數：{total} 名")
        self.res_text1.config(state=tk.DISABLED)

    # --- 分頁 2: 到離查詢 ---
    def setup_tab2(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        file_f = ttk.LabelFrame(main, text="1. 載入 到離名單.xlsx", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="選擇檔案", command=self.load_file2).pack(side=tk.LEFT)
        self.file_label2 = ttk.Label(file_f, text="尚未載入", foreground="gray")
        self.file_label2.pack(side=tk.LEFT, padx=10)
        
        sf = ttk.LabelFrame(main, text="2. 查詢條件", padding=10)
        sf.pack(fill=tk.X, pady=5)
        ttk.Label(sf, text="日期:").pack(side=tk.LEFT)
        self.date_e = ttk.Entry(sf, width=10)
        self.date_e.pack(side=tk.LEFT, padx=5)
        ttk.Label(sf, text="姓名:").pack(side=tk.LEFT)
        self.name_e = ttk.Entry(sf, width=10)
        self.name_e.pack(side=tk.LEFT, padx=5)
        self.search_btn = ttk.Button(sf, text="執行查詢", command=self.search_data2, state=tk.DISABLED)
        self.search_btn.pack(side=tk.LEFT, padx=10)
        
        table_f = ttk.Frame(main)
        table_f.pack(fill=tk.BOTH, expand=True, pady=10)
        cols = ("date", "name", "unit", "level1", "reason", "remarks")
        self.tree = ttk.Treeview(table_f, columns=cols, show='headings')
        vsb = ttk.Scrollbar(table_f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        for c, h in zip(cols, ["日期", "姓名", "單位", "一級單位", "異動原因", "備註"]): 
            self.tree.heading(c, text=h)
            self.tree.column(c, width=120, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def load_file2(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p:
            return
        try:
            self.df2 = pd.read_excel(p)
            self.file_label2.config(text=os.path.basename(p), foreground="black")
            self.search_btn.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def search_data2(self):
        if self.df2 is None:
            return
        d, n = self.date_e.get(), self.name_e.get()
        cols = self.df2.columns
        cd = next((c for c in cols if any(k in str(c) for k in ["日期", "交"])), cols[0])
        cn = next((c for c in cols if any(k in str(c) for k in ["姓名", "憪"])), cols[1])
        m = pd.Series([True]*len(self.df2))
        if d:
            m &= self.df2[cd].astype(str).str.contains(d)
        if n:
            m &= self.df2[cn].astype(str).str.contains(n)
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, r in self.df2[m].iterrows():
            self.tree.insert("", tk.END, values=[r.get(c, "") for c in [cd, cn, "單位", "一級單位", "原因", "備註"]])

    # --- 分頁 3: 副校長報表 ---
    def setup_tab3(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="副校長 - 全校各類人員員額統計週報", font=('Microsoft JhengHei', 16, 'bold'), foreground="#1a5276").pack(pady=10)
        
        # 手動輸入區 (擴充：專業經理人)
        manual_frame = ttk.LabelFrame(main, text="手動輸入額外人員數據 (約用、兼任、助教、經理人)", padding=10)
        manual_frame.pack(fill=tk.X, pady=10)
        
        self.manual_entries = {} 
        items = [
            ("約用人員", "manual"), 
            ("兼任教師", "pt"), 
            ("新制助教", "new_ta"), 
            ("舊制助教", "old_ta"),
            ("專業經理人", "manager")
        ]
        for i, (label, key) in enumerate(items):
            ttk.Label(manual_frame, text=f"{label} - 總人數:").grid(row=i, column=0, padx=5, pady=2, sticky=tk.E)
            total_e = ttk.Entry(manual_frame, width=8); total_e.insert(0, "0"); total_e.grid(row=i, column=1, padx=5)
            ttk.Label(manual_frame, text="男:").grid(row=i, column=2, padx=5); m_e = ttk.Entry(manual_frame, width=8); m_e.insert(0, "0"); m_e.grid(row=i, column=3, padx=5)
            ttk.Label(manual_frame, text="女:").grid(row=i, column=4, padx=5); f_e = ttk.Entry(manual_frame, width=8); f_e.insert(0, "0"); f_e.grid(row=i, column=5, padx=5)
            self.manual_entries[key] = (total_e, m_e, f_e)

        ttk.Button(main, text="📊 生成全校人員統計報表 (含手動與新增類別)", command=self.gen_vice_president_report).pack(pady=5)
        self.vp_out = tk.Text(main, font=('Microsoft JhengHei', 11), bg="#f4f6f7")
        self.vp_out.pack(fill=tk.BOTH, expand=True, pady=10)

    def gen_vice_president_report(self):
        if not self.file1_path:
            messagebox.showwarning("提示", "請載入檔案 1")
            return
        manual_data = {}
        try:
            for key, (t_e, m_e, f_e) in self.manual_entries.items():
                manual_data[key] = (int(t_e.get()), int(m_e.get()), int(f_e.get()))
        except ValueError:
            messagebox.showerror("錯誤", "手動輸入欄位請填入整數數字")
            return

        # 指定 11 個要讀取的工作表
        target_sheets = [
            "教師", "約聘教師", "專案教學人員", "講座教授", "客座人員", 
            "研究人員", "稀少性科技人員", "醫事人員", "新制職員", 
            "駐衛警察隊", "技工工友合計"
        ]
        
        try:
            xl = pd.ExcelFile(self.file1_path)
            self.vp_out.delete(1.0, tk.END)
            self.vp_out.insert(tk.END, f"【各類人員員額與性別統計摘要 - 副校長】\n日期：{datetime.now().strftime('%Y-%m-%d')}\n" + "="*75 + "\n")
            
            grand_total = sum(d[0] for d in manual_data.values())
            grand_m = sum(d[1] for d in manual_data.values())
            grand_f = sum(d[2] for d in manual_data.values())
            
            # 特殊彙整統計 (新需求：職員 = 新制職員 + 醫事人員 + 稀少性科技人員)
            admin_staff_total = 0; admin_staff_m = 0; admin_staff_f = 0
            
            for sheet in target_sheets:
                if sheet in xl.sheet_names:
                    df = pd.read_excel(self.file1_path, sheet_name=sheet)
                    c_t = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別"])), None)
                    if c_t: df = df[df[c_t].astype(str).str.strip().replace(['nan','','None'], pd.NA).notna()]
                    c_g = next((c for c in df.columns if "性別" in str(c)), None)
                    m = df[df[c_g].astype(str).str.contains("男")].shape[0] if c_g else 0
                    f = df[df[c_g].astype(str).str.contains("女")].shape[0] if c_g else 0
                    count = len(df)
                    self.vp_out.insert(tk.END, f"● {sheet:<15} | {count:>8} | 男:{m:>4} | 女:{f:>4}\n")
                    
                    grand_total += count; grand_m += m; grand_f += f
                    
                    # 職員小計邏輯
                    if sheet in ["新制職員", "醫事人員", "稀少性科技人員"]:
                        admin_staff_total += count; admin_staff_m += m; admin_staff_f += f
                else:
                    self.vp_out.insert(tk.END, f"○ {sheet:<15} | （找不到工作表）\n")
            
            self.vp_out.insert(tk.END, "-"*75 + "\n")
            # 顯示手動填報項
            display_names = {
                "manual": "約用人員(填報)", "pt": "兼任教師(填報)", 
                "new_ta": "新制助教(填報)", "old_ta": "舊制助教(填報)",
                "manager": "專業經理人(填報)"
            }
            for key, (t, m, f) in manual_data.items():
                self.vp_out.insert(tk.END, f"● {display_names[key]:<13} | {t:>8} | 男:{m:>4} | 女:{f:>4}\n")
            
            self.vp_out.insert(tk.END, "-"*75 + f"\n【全校總計彙整】\n")
            self.vp_out.insert(tk.END, f"   - 職員小計 (新制/醫事/稀少性) ： {admin_staff_total:>5} 名 (男:{admin_staff_m} / 女:{admin_staff_f})\n")
            self.vp_out.insert(tk.END, f"   - 全校人員總計                ： {grand_total} 名 (男：{grand_m} / 女：{grand_f})\n" + "="*75)
        except Exception as e: messagebox.showerror("失敗", str(e))

    # --- 報表分頁 4-6 ---
    def setup_tab4(self, f): self.create_report_ui(f, "D5 報表", "員額明細。", "匯出 D5", self.gen_d5_report)
    def setup_tab5(self, f): self.create_report_ui(f, "給主計室", "月報彙整。", "生成月報", self.gen_accounting_report)
    def setup_tab6(self, f):
        main = ttk.Frame(f, padding="20"); main.pack(fill=tk.BOTH, expand=True)
        ttk.Button(main, text="載入 赴大陸資料", command=self.load_file_travel).pack()
        self.file_label_t = ttk.Label(main, text="尚未載入"); self.file_label_t.pack()
        ttk.Button(main, text="分析統計", command=self.gen_travel_report).pack()
        self.t_out = tk.Text(main, font=('Microsoft JhengHei', 11)); self.t_out.pack(fill=tk.BOTH, expand=True)

    def load_file_travel(self):
        p = filedialog.askopenfilename()
        if not p: return
        try: self.df_travel = pd.read_excel(p); self.file_label_t.config(text=os.path.basename(p))
        except Exception as e: messagebox.showerror("錯誤", str(e))

    # --- 分頁 7: 個資申請單 ---
    def setup_tab7(self, f):
        main = ttk.Frame(f, padding="20"); main.pack(fill=tk.BOTH, expand=True)
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
            df = df[df[c_t].astype(str).str.strip().replace(['nan',''], pd.NA).notna()]
            col_map = {"姓名": ["姓名", "憪"], "服務單位": ["單位", "系所"], "身分證統一編號": ["身分證", "統一編號", "ID"], "職稱": [c_t]}
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
            save_p = f"{self.app_filename_e.get()}.xlsx"; df[[c for _, c in cols]].to_excel(save_p, index=False)
            messagebox.showinfo("成功", f"匯出至 {save_p}")

    def gen_travel_report(self, out=None): pass
    def gen_monday_report(self, out): pass
    def gen_d5_report(self, out): pass
    def gen_accounting_report(self, out): pass

if __name__ == "__main__":
    root = tk.Tk(); app = TeacherStatsApp(root); root.mainloop()
