import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime

class TeacherStatsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("教師資料管理中心 - 行政綜合業務系統")
        self.root.geometry("1000x850")

        # 全域資料變數
        self.df1 = None  # 各系所教師
        self.df2 = None  # 到離名單
        self.df_travel = None # 赴大陸資料
        self.title_col1 = None
        self.cat_col1 = None

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

    # --- 共通報表 UI 生成器 ---
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

    # --- 分頁佈局 ---
    def setup_tab1(self, f):
        self.check_vars = {}
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        file_f = ttk.LabelFrame(main, text="載入 各系所教師.xlsx", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="選擇檔案", command=self.load_file1).pack(side=tk.LEFT)
        self.file_label1 = ttk.Label(file_f, text="尚未載入", foreground="gray")
        self.file_label1.pack(side=tk.LEFT, padx=10)
        
        self.select_f1 = ttk.LabelFrame(main, text="勾選職稱", padding=10)
        self.select_f1.pack(fill=tk.BOTH, expand=True, pady=10)
        self.canvas = tk.Canvas(self.select_f1, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self.select_f1, command=self.canvas.yview)
        self.scroll_f = ttk.Frame(self.canvas)
        self.scroll_f.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scroll_f, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")
        
        self.calc_btn1 = ttk.Button(main, text="開始統計", command=self.calculate1, state=tk.DISABLED)
        self.calc_btn1.pack(pady=5)
        self.res_text1 = tk.Text(main, height=6, font=('Consolas', 10), state=tk.DISABLED)
        self.res_text1.pack(fill=tk.X)

    def setup_tab2(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        file_f = ttk.LabelFrame(main, text="載入 到離名單.xlsx", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="選擇檔案", command=self.load_file2).pack(side=tk.LEFT)
        self.file_label2 = ttk.Label(file_f, text="尚未載入", foreground="gray")
        self.file_label2.pack(side=tk.LEFT, padx=10)
        
        sf = ttk.LabelFrame(main, text="查詢條件", padding=10)
        sf.pack(fill=tk.X, pady=5)
        ttk.Label(sf, text="日期:").pack(side=tk.LEFT)
        self.date_e = ttk.Entry(sf, width=10)
        self.date_e.pack(side=tk.LEFT, padx=5)
        ttk.Label(sf, text="姓名:").pack(side=tk.LEFT)
        self.name_e = ttk.Entry(sf, width=10)
        self.name_e.pack(side=tk.LEFT, padx=5)
        self.search_btn = ttk.Button(sf, text="查詢", command=self.search_data2, state=tk.DISABLED)
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

    def setup_tab3(self, f): self.create_report_ui(f, "每周一給副校長", "統計『到離名單』中所有人員的變動摘要。", "生成週報摘要", self.gen_monday_report)
    def setup_tab4(self, f): self.create_report_ui(f, "D5 報表", "生成符合 D5 格式的員額異動明細。", "匯出 D5 明細", self.gen_d5_report)
    def setup_tab5(self, f): self.create_report_ui(f, "給主計室", "統計載入檔案中所有薪資異動人員。", "生成主計月報", self.gen_accounting_report)
    
    def setup_tab6(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="赴大陸地區人數統計", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        file_f = ttk.LabelFrame(main, text="載入 赴大陸資料.xlsx", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="選擇檔案", command=self.load_file_travel).pack(side=tk.LEFT)
        self.file_label_t = ttk.Label(file_f, text="尚未載入", foreground="gray")
        self.file_label_t.pack(side=tk.LEFT, padx=10)
        ttk.Button(main, text="開始全量統計", command=self.gen_travel_report).pack(pady=10)
        self.t_out = tk.Text(main, font=('Microsoft JhengHei', 11), bg="#fdfefe")
        self.t_out.pack(fill=tk.BOTH, expand=True)

    def setup_tab7(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="個人資料使用申請 - 人員名單彙整", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        
        input_f = ttk.LabelFrame(main, text="請勾選欲申請的人員類別", padding=15)
        input_f.pack(fill=tk.X, pady=10)
        
        self.cat_vars = {}
        cat_frame = ttk.Frame(input_f)
        cat_frame.pack(fill=tk.X, padx=10, pady=5)
        
        categories = ["新制職員", "助教", "技工工友", "教師", "約聘教師", "駐衛警察隊", "專案教學人員", "講座教授", "客座教授"]
        for i, cat in enumerate(categories):
            var = tk.BooleanVar()
            self.cat_vars[cat] = var
            cb = ttk.Checkbutton(cat_frame, text=cat, variable=var)
            cb.grid(row=i//3, column=i%3, sticky=tk.W, padx=15, pady=5)

        ttk.Button(main, text="📝 生成該類別人員名單彙整", command=self.gen_app_form).pack(pady=10)
        self.app_out = tk.Text(main, font=('Microsoft JhengHei', 11), bg="#fffcf0")
        self.app_out.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    # --- 邏輯功能 ---
    def load_file1(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p: return
        try:
            self.df1 = pd.read_excel(p, sheet_name='教師')
            possible_title_cols = [c for c in self.df1.columns if any(k in str(c) for k in ['教授', '講師', '職稱'])]
            if not possible_title_cols: raise ValueError("找不到職稱相關欄位")
            self.title_col1 = possible_title_cols[0]
            
            possible_cat_cols = [c for c in self.df1.columns if any(k in str(c) for k in ['類別', '職別', '身分', '屬性', '人員類別'])]
            self.cat_col1 = possible_cat_cols[0] if possible_cat_cols else self.title_col1

            ut = sorted([str(x).strip() for x in self.df1[self.title_col1].unique() if pd.notna(x)])
            for w in self.scroll_f.winfo_children(): w.destroy()
            for t in ut:
                var = tk.BooleanVar(); self.check_vars[t] = var
                ttk.Checkbutton(self.scroll_f, text=t, variable=var).pack(anchor="w", padx=10)
            self.file_label1.config(text=os.path.basename(p), foreground="black"); self.calc_btn1.config(state=tk.NORMAL)
        except Exception as e: messagebox.showerror("載入失敗", str(e))

    def calculate1(self):
        sel = [t for t, v in self.check_vars.items() if v.get()]
        if not sel: return
        cnt = self.df1[self.title_col1].astype(str).str.strip().value_counts()
        self.res_text1.config(state=tk.NORMAL); self.res_text1.delete(1.0, tk.END)
        for t in sel: self.res_text1.insert(tk.END, f"{t:<23} | {cnt.get(t, 0):>5} 名\n")
        self.res_text1.config(state=tk.DISABLED)

    def load_file2(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p: return
        try: self.df2 = pd.read_excel(p); self.file_label2.config(text=os.path.basename(p), foreground="black"); self.search_btn.config(state=tk.NORMAL)
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def search_data2(self):
        if self.df2 is None: return
        d, n = self.date_e.get(), self.name_e.get()
        cols = self.df2.columns
        cd = next((c for c in cols if any(k in str(c) for k in ["日期", "交"])), cols[0])
        cn = next((c for c in cols if any(k in str(c) for k in ["姓名", "憪"])), cols[1])
        m = pd.Series([True]*len(self.df2))
        if d: m &= self.df2[cd].astype(str).str.contains(d)
        if n: m &= self.df2[cn].astype(str).str.contains(n)
        for i in self.tree.get_children(): self.tree.delete(i)
        for _, r in self.df2[m].iterrows():
            vals = [r.get(c, "") for c in [cd, cn, "單位", "一級單位", "原因", "備註"]]
            self.tree.insert("", tk.END, values=vals)

    def load_file_travel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p: return
        try: self.df_travel = pd.read_excel(p); self.file_label_t.config(text=os.path.basename(p), foreground="black")
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def gen_monday_report(self, out):
        if self.df2 is None: return
        out.delete(1.0, tk.END); out.insert(tk.END, f"【人事異動週報】 {datetime.now().strftime('%Y-%m-%d')}\n" + "="*50 + "\n")
        for _, r in self.df2.iterrows(): out.insert(tk.END, f"• {r.get('姓名', '未知')} ({r.get('單位', '未知')}) - {r.get('原因', '異動')}\n")

    def gen_d5_report(self, out):
        if self.df2 is None: return
        out.delete(1.0, tk.END); out.insert(tk.END, f"【D5 員額變動明細】\n" + "-"*50 + "\n")
        for _, r in self.df2.iterrows(): out.insert(tk.END, f"{r.get('姓名', '')} | {r.get('單位', '')} | {r.get('原因', '')}\n")

    def gen_accounting_report(self, out):
        if self.df2 is None: return
        out.delete(1.0, tk.END); out.insert(tk.END, f"【主計室月報彙整】\n" + "-"*50 + "\n")
        for _, r in self.df2.iterrows(): out.insert(tk.END, f"姓名：{r.get('姓名', '')} | 單位：{r.get('單位', '')}\n")

    def gen_travel_report(self):
        if self.df_travel is None: return
        self.t_out.delete(1.0, tk.END); self.t_out.insert(tk.END, f"【赴大陸地區統計】總計：{len(self.df_travel)} 人次\n")

    def gen_app_form(self):
        sel_cats = [cat for cat, var in self.cat_vars.items() if var.get()]
        if not sel_cats: messagebox.showwarning("提示", "請勾選類別"); return
        if self.df1 is None: messagebox.showwarning("提示", "請載入各系所教師檔案"); return
        
        c_name = next((c for c in self.df1.columns if any(k in str(c) for k in ["姓名", "憪"])), self.df1.columns[1])
        c_unit = next((c for c in self.df1.columns if any(k in str(c) for k in ["單位", "系所"])), self.df1.columns[2])
        mask = self.df1[self.cat_col1].astype(str).apply(lambda x: any(cat in x for cat in sel_cats))
        res = self.df1[mask]

        self.app_out.delete(1.0, tk.END)
        self.app_out.insert(tk.END, f"【個資使用申請名單】類別：{'、'.join(sel_cats)}\n人數：{len(res)} 名\n" + "="*60 + "\n")
        for _, r in res.iterrows():
            self.app_out.insert(tk.END, f"{str(r.get(c_name,'')):<10} | {str(r.get(c_unit,'')):<25} | {str(r.get(self.title_col1,'')):<20}\n")

if __name__ == "__main__":
    root = tk.Tk(); app = TeacherStatsApp(root); root.mainloop()
