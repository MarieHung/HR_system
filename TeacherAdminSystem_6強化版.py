import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime

class TeacherStatsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("教師資料管理中心 - 行政綜合業務系統 (V6.2)")
        self.root.geometry("1150x950")

        # 全域資料變數
        self.df1 = None  
        self.file1_path = None 
        self.df2 = None  
        self.df_travel = None 
        
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

    # --- 各分頁佈局實作 ---
    def setup_tab1(self, f):
        self.check_vars = {}
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        file_f = ttk.LabelFrame(main, text="1. 載入 各系所教師.xlsx", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="選擇檔案", command=self.load_file1).pack(side=tk.LEFT)
        self.file_label1 = ttk.Label(file_f, text="尚未載入", foreground="gray")
        self.file_label1.pack(side=tk.LEFT, padx=10)
        
        self.select_f1 = ttk.LabelFrame(main, text="2. 勾選職稱", padding=10)
        self.select_f1.pack(fill=tk.BOTH, expand=True, pady=10)
        self.canvas = tk.Canvas(self.select_f1, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self.select_f1, command=self.canvas.yview)
        self.scroll_f = ttk.Frame(self.canvas)
        self.scroll_f.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scroll_f, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")
        
        self.calc_btn1 = ttk.Button(main, text="開始統計人數", command=self.calculate1, state=tk.DISABLED)
        self.calc_btn1.pack(pady=5)
        self.res_text1 = tk.Text(main, height=6, font=('Consolas', 10), state=tk.DISABLED)
        self.res_text1.pack(fill=tk.X)

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

    def setup_tab3(self, f): self.create_report_ui(f, "每周一給副校長", "統計『到離名單』摘要。", "生成週報摘要", self.gen_monday_report)
    def setup_tab4(self, f): self.create_report_ui(f, "D5 報表", "生成 D5 格式明細。", "匯出 D5 明細", self.gen_d5_report)
    def setup_tab5(self, f): self.create_report_ui(f, "給主計室", "統計薪資異動人員。", "生成主計月報", self.gen_accounting_report)
    
    def setup_tab6(self, f):
        main = ttk.Frame(f, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main, text="赴大陸地區統計", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
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
        ttk.Label(main, text="個人資料使用申請 - Excel 快速預覽與匯出", font=('Microsoft JhengHei', 14, 'bold')).pack(pady=5)
        
        file_f = ttk.LabelFrame(main, text="0. 載入資料來源 (各系所教師.xlsx)", padding=10)
        file_f.pack(fill=tk.X, pady=5)
        ttk.Button(file_f, text="選擇 各系所教師 檔案", command=self.load_file1).pack(side=tk.LEFT)
        self.file_label_app = ttk.Label(file_f, text="尚未載入", foreground="gray")
        self.file_label_app.pack(side=tk.LEFT, padx=10)

        # 1. 人員類別
        cat_f = ttk.LabelFrame(main, text="1. 請勾選欲申請的人員類別 (工作表)", padding=10)
        cat_f.pack(fill=tk.X, pady=5)
        self.cat_vars = {}
        cat_grid = ttk.Frame(cat_f)
        cat_grid.pack(fill=tk.X, padx=5)
        categories = ["新制職員", "助教", "技工工友", "教師", "約聘教師", "駐衛警察隊", "專案教學人員", "講座教授", "客座教授"]
        for i, cat in enumerate(categories):
            var = tk.BooleanVar()
            self.cat_vars[cat] = var
            ttk.Checkbutton(cat_grid, text=cat, variable=var).grid(row=i//3, column=i%3, sticky=tk.W, padx=10, pady=2)

        # 2. 資料項目
        item_f = ttk.LabelFrame(main, text="2. 請勾選欲申請的資料項目 (欄位)", padding=10)
        item_f.pack(fill=tk.X, pady=5)
        self.item_vars = {}
        item_grid = ttk.Frame(item_f)
        item_grid.pack(fill=tk.X, padx=5)
        data_items = [
            "姓名", "服務單位", "一級單位", "身分證統一編號", "職稱", 
            "性別", "薪點", "最高學位", "出生年月日", "到校日期", 
            "服務本校年資", "帶職帶薪", "留職停薪", "EMI",
            "本校電子信箱", "校外電子信箱"
        ]
        for i, item in enumerate(data_items):
            var = tk.BooleanVar()
            self.item_vars[item] = var
            ttk.Checkbutton(item_grid, text=item, variable=var).grid(row=i//4, column=i%4, sticky=tk.W, padx=10, pady=2)

        # 3. 匯出設定
        export_f = ttk.LabelFrame(main, text="3. 匯出設定", padding=10)
        export_f.pack(fill=tk.X, pady=5)
        ttk.Label(export_f, text="匯出檔案名稱:").pack(side=tk.LEFT, padx=5)
        self.app_filename_e = ttk.Entry(export_f, width=30)
        self.app_filename_e.insert(0, f"個資申請名單_{datetime.now().strftime('%m%d')}")
        self.app_filename_e.pack(side=tk.LEFT, padx=5)
        ttk.Label(export_f, text=".xlsx").pack(side=tk.LEFT)

        # 按鈕區
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="👁️ 預覽檔案內容", command=self.preview_app_data).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="📥 匯出成Excel檔", command=self.gen_app_excel).pack(side=tk.LEFT, padx=10)
        
        self.app_out = tk.Text(main, height=12, font=('Microsoft JhengHei', 10), bg="#fffcf0")
        self.app_out.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    # --- 邏輯功能 ---
    def load_file1(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p:
            return
        self.file1_path = p 
        try:
            xl = pd.ExcelFile(p)
            sheets = xl.sheet_names
            target_sheet = "教師" if "教師" in sheets else sheets[0]
            self.df1 = pd.read_excel(p, sheet_name=target_sheet)
            
            possible_title_cols = [c for c in self.df1.columns if any(k in str(c) for k in ['教授', '講師', '職稱'])]
            self.title_col1 = possible_title_cols[0] if possible_title_cols else self.df1.columns[0]
            possible_cat_cols = [c for c in self.df1.columns if any(k in str(c) for k in ['類別', '職別', '身分'])]
            self.cat_col1 = possible_cat_cols[0] if possible_cat_cols else self.title_col1

            ut = sorted([str(x).strip() for x in self.df1[self.title_col1].unique() if pd.notna(x)])
            for w in self.scroll_f.winfo_children():
                w.destroy()
            for t in ut:
                var = tk.BooleanVar()
                self.check_vars[t] = var
                ttk.Checkbutton(self.scroll_f, text=t, variable=var).pack(anchor="w", padx=10)
            
            self.file_label1.config(text=os.path.basename(p), foreground="black")
            if hasattr(self, 'file_label_app'):
                self.file_label_app.config(text=os.path.basename(p), foreground="black")
            self.calc_btn1.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("載入失敗", str(e))

    def calculate1(self):
        if self.df1 is None:
            return
        sel = [t for t, v in self.check_vars.items() if v.get()]
        if not sel:
            return
        cnt = self.df1[self.title_col1].astype(str).str.strip().value_counts()
        self.res_text1.config(state=tk.NORMAL)
        self.res_text1.delete(1.0, tk.END)
        for t in sel:
            self.res_text1.insert(tk.END, f"{t:<23} | {cnt.get(t, 0):>5} 名\n")
        self.res_text1.config(state=tk.DISABLED)

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
        cr = next((c for c in cols if any(k in str(c) for k in ["異動原因", "原因"])), "異動原因")
        m = pd.Series([True]*len(self.df2))
        if d:
            m &= self.df2[cd].astype(str).str.contains(d)
        if n:
            m &= self.df2[cn].astype(str).str.contains(n)
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, r in self.df2[m].iterrows():
            self.tree.insert("", tk.END, values=[r.get(cd, ""), r.get(cn, ""), r.get("單位", ""), r.get("一級單位", ""), r.get(cr, ""), r.get("備註", "")])

    def load_file_travel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not p:
            return
        try:
            self.df_travel = pd.read_excel(p)
            self.file_label_t.config(text=os.path.basename(p), foreground="black")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    # --- 個資申請單核心邏輯 (V6.2: 移除樞紐檢查) ---
    def get_combined_app_data(self):
        sel_cats = [cat for cat, var in self.cat_vars.items() if var.get()]
        sel_items = [it for it, var in self.item_vars.items() if var.get()]
        if not sel_cats:
            messagebox.showwarning("提示", "請勾選人員類別")
            return None, None, None
        if not sel_items:
            messagebox.showwarning("提示", "請勾選資料項目")
            return None, None, None
        if not self.file1_path:
            messagebox.showwarning("提示", "請先載入各系所教師檔案！")
            return None, None, None

        try:
            xl = pd.ExcelFile(self.file1_path)
            all_dfs = []
            
            for cat in sel_cats:
                if cat in xl.sheet_names:
                    temp_df = pd.read_excel(self.file1_path, sheet_name=cat)
                    all_dfs.append(temp_df)
            
            if not all_dfs:
                messagebox.showerror("錯誤", "在 Excel 中找不到對應的工作表")
                return None, None, None
            
            df = pd.concat(all_dfs, ignore_index=True)
            
            # --- 過濾職稱為空白的人員 (新需求) ---
            # 智慧尋找職稱欄位
            c_title_detect = next((c for c in df.columns if any(k in str(c) for k in ["職稱", "職別", "職級"])), None)
            if c_title_detect:
                # 刪除職稱為空 (NaN 或 僅含空白字串) 的列
                df = df[df[c_title_detect].astype(str).str.strip().replace(['nan', 'None', ''], pd.NA).notna()]
            
            col_map = {
                "姓名": ["姓名", "憪"], "服務單位": ["單位", "系所", "桐"], "一級單位": ["一級單位", "銝蝝雿"],
                "身分證統一編號": ["身分證", "統一編號", "ID"], "職稱": ["職稱", "職別"], "性別": ["性別"],
                "薪點": ["薪點", "俸點"], "最高學位": ["學位", "最高學位"], "出生年月日": ["出生", "生日"], "到校日期": ["到校", "起聘"],
                "服務本校年資": ["年資"], "帶職帶薪": ["帶職帶薪"], "留職停薪": ["留職停薪"], "EMI": ["EMI"],
                "本校電子信箱": ["校內信箱", "本校信箱", "Email", "電子郵件"], 
                "校外電子信箱": ["校外信箱", "個人信箱", "外部信箱"]
            }
            
            found_cols = []
            for it in sel_items:
                actual = next((c for c in df.columns if any(a in str(c) for a in col_map.get(it, [it]))), None)
                if actual:
                    found_cols.append((it, actual))
            
            return df, found_cols, sel_cats
        except Exception as e:
            messagebox.showerror("資料讀取失敗", f"原因：{str(e)}")
            return None, None, None

    def preview_app_data(self):
        df, cols, cats = self.get_combined_app_data()
        if df is None:
            return
        
        self.app_out.delete(1.0, tk.END)
        self.app_out.insert(tk.END, f"【預覽】個資申請彙整 (來源：{', '.join(cats)})\n總計人數：{len(df)} 名\n" + "="*80 + "\n")
        header = " | ".join([it for it, _ in cols])
        self.app_out.insert(tk.END, header + "\n" + "-"*80 + "\n")
        
        for _, r in df.head(50).iterrows():
            self.app_out.insert(tk.END, " | ".join([str(r.get(c, "")).strip() for _, c in cols]) + "\n")
        
        if len(df) > 50:
            self.app_out.insert(tk.END, "\n...(其餘資料已省略，請匯出 Excel 查看完整名單)\n")

    def gen_app_excel(self):
        df, cols, cats = self.get_combined_app_data()
        if df is None:
            return
        out_name = self.app_filename_e.get().strip()
        if not out_name:
            messagebox.showwarning("提示", "請輸入檔名")
            return

        try:
            actual_cols = [c for _, c in cols]
            export_df = df[actual_cols]
            export_df.columns = [it for it, _ in cols]
            save_path = f"{out_name}.xlsx"
            export_df.to_excel(save_path, index=False)
            
            self.app_out.delete(1.0, tk.END)
            self.app_out.insert(tk.END, f"✅ 匯出成功！\n檔案路徑：{os.path.abspath(save_path)}\n資料筆數：{len(export_df)}\n")
            messagebox.showinfo("成功", f"Excel 檔案匯出成功：\n{save_path}")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))

    def gen_monday_report(self, out):
        if self.df2 is None:
            return
        out.delete(1.0, tk.END)
        out.insert(tk.END, f"【人事異動週報摘要】\n")

    def gen_d5_report(self, out):
        if self.df2 is None:
            return
        out.delete(1.0, tk.END)
        out.insert(tk.END, f"【D5 員額明細】\n")

    def gen_accounting_report(self, out):
        if self.df2 is None:
            return
        out.delete(1.0, tk.END)
        out.insert(tk.END, f"【主計室月報彙整】\n")

    def gen_travel_report(self):
        if self.df_travel is None:
            return
        self.t_out.delete(1.0, tk.END)
        self.t_out.insert(tk.END, f"赴陸統計：{len(self.df_travel)} 人次")

if __name__ == "__main__":
    root = tk.Tk()
    app = TeacherStatsApp(root)
    root.mainloop()
