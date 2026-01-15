import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import defaultdict
import platform
import os
import shutil
import re
import subprocess
from PIL import Image
import pytesseract
from category_manager import load_categories, save_categories


# --- tkcalendar (可選) ---
try:
    from tkcalendar import DateEntry
    CALENDAR_ENABLED = True
except ImportError:
    CALENDAR_ENABLED = False
    DateEntry = None # Placeholder so the rest of the code doesn't crash on reference

# --- Tesseract 設定 ---
# 如果 Tesseract 沒有被自動偵測到，請取消下方註解並提供您的 tesseract.exe 的完整路徑
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# --- 全域設定 ---
CSV_FILE = 'records.csv'
INVOICE_DIR = 'invoices'
HEADERS = ['日期', '類型', '類別', '金額', '備註', '圖片']

# --- Matplotlib 中文字體設定 ---
def set_chinese_font():
    """根據作業系統設定 Matplotlib 的中文字體"""
    system = platform.system()
    if system == 'Windows':
        font_name = 'Microsoft JhengHei'
    elif system == 'Darwin':
        font_name = 'PingFang TC'
    else:
        font_name = 'WenQuanYi Zen Hei'
    
    try:
        plt.rcParams['font.sans-serif'] = [font_name]
        plt.rcParams['axes.unicode_minus'] = False
    except Exception as e:
        print(f"⚠️ 警告：設定字體失敗: {e}。中文可能無法正常顯示。  ")

# --- 資料遷移與初始化 ---
def migrate_csv_if_needed():
    """檢查 CSV 格式是否為舊版，如果是，則遷移到新格式"""
    if not os.path.exists(CSV_FILE):
        return # 檔案不存在，不需要遷移

    try:
        with open(CSV_FILE, 'r', encoding='utf-8', newline='') as file:
            reader = csv.reader(file)
            header = next(reader)
        
        # 根據偵測到的標頭長度決定如何遷移
        if len(header) < len(HEADERS):
            print(f"偵測到舊版({len(header)}欄) CSV 檔案，正在進行資料遷移至新版({len(HEADERS)}欄)...")
            backup_path = f"{CSV_FILE}.{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.bak"
            shutil.copy(CSV_FILE, backup_path)
            print(f"舊檔案已備份至: {backup_path}")

            old_records = []
            with open(backup_path, 'r', encoding='utf-8', newline='') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    old_records.append(row)

            with open(CSV_FILE, 'w', encoding='utf-8', newline='') as file:
                writer = csv.DictWriter(file, fieldnames=HEADERS)
                writer.writeheader()
                for record in old_records:
                    new_record = {
                        '日期': record.get('日期'),
                        '類型': record.get('類型', '支出'), # 如果沒有'類型'欄位，預設為'支出'
                        '類別': record.get('類別'),
                        '金額': record.get('金額'),
                        '備註': record.get('備註'),
                        '圖片': record.get('圖片', '') # 新增圖片欄位，預設為空
                    }
                    writer.writerow(new_record)
            messagebox.showinfo("資料遷移", f"您的資料已成功更新到新格式！\n舊檔案已備份至 {backup_path}")

    except Exception as e:
        messagebox.showerror("遷移錯誤", f"資料遷移失敗: {e}")
        return False
    return True

def init_csv():
    """檢查 CSV 檔案是否存在，若不存在則建立"""
    if not os.path.exists(CSV_FILE):
        try:
            with open(CSV_FILE, 'w', encoding='utf-8', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(HEADERS)
            return True
        except Exception as e:
            messagebox.showerror("錯誤", f"建立 CSV 檔案失敗: {e}")
            return False
    return True

# --- 資料讀取/寫入功能 ---
def read_records():
    """讀取所有紀錄"""
    try:
        with open(CSV_FILE, 'r', encoding='utf-8', newline='') as file:
            reader = csv.DictReader(file)
            records = [row for row in reader if any(row.values())]
        for record in records:
            record['金額'] = float(record['金額'])
        return records
    except FileNotFoundError:
        return [] 
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取檔案時發生錯誤: {e}")
        return []

def append_record_to_csv(record_data):
    """將單筆紀錄附加到 CSV 檔案"""
    try:
        with open(CSV_FILE, 'a', encoding='utf-8', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(record_data)
        return True
    except Exception as e:
        messagebox.showerror("錯誤", f"寫入檔案時發生錯誤: {e}")
        return False

# --- 編輯紀錄視窗 ---
class EditRecordWindow(tk.Toplevel):
    def __init__(self, parent, record_to_edit, categories, callback):
        super().__init__(parent)
        self.transient(parent)
        self.title("編輯紀錄")
        self.geometry("400x300")
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        self.original_record = record_to_edit
        self.callback = callback

        # --- Widgets ---
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="日期:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        if CALENDAR_ENABLED:
            self.date_entry = DateEntry(frame, date_pattern='y-mm-dd')
            try:
                date_obj = datetime.datetime.strptime(self.original_record['日期'], '%Y-%m-%d')
                self.date_entry.set_date(date_obj)
            except (ValueError, KeyError):
                pass # DateEntry will default to today if parsing fails
        else:
            self.date_entry = ttk.Entry(frame)
            self.date_entry.insert(0, self.original_record['日期'])
        self.date_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

        ttk.Label(frame, text="類型:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.type_combobox = ttk.Combobox(frame, values=['支出', '收入'], state="readonly")
        self.type_combobox.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        self.type_combobox.set(self.original_record['類型'])

        ttk.Label(frame, text="類別:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_combobox = ttk.Combobox(frame, values=categories)
        self.category_combobox.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        self.category_combobox.set(self.original_record['類別'])

        ttk.Label(frame, text="金額:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.amount_entry = ttk.Entry(frame)
        self.amount_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)
        self.amount_entry.insert(0, self.original_record['金額'])

        ttk.Label(frame, text="備註:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.note_entry = ttk.Entry(frame)
        self.note_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=5)
        self.note_entry.insert(0, self.original_record['備註'])
        
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=5, columnspan=2, pady=20)
        
        save_button = ttk.Button(button_frame, text="儲存變更", command=self.save_changes)
        save_button.pack(side=tk.LEFT, padx=10)
        
        cancel_button = ttk.Button(button_frame, text="取消", command=self.cancel)
        cancel_button.pack(side=tk.LEFT, padx=10)
        
        self.grab_set()
        self.wait_window()

    def save_changes(self):
        """儲存修改後的紀錄"""
        # --- 取得新資料並驗證 ---
        new_date = self.date_entry.get().strip()
        new_type = self.type_combobox.get()
        new_category = self.category_combobox.get().strip()
        new_amount_str = self.amount_entry.get().strip()
        new_note = self.note_entry.get().strip()

        if not (new_date and new_type and new_category and new_amount_str):
            messagebox.showwarning("輸入錯誤", "日期、類型、類別和金額為必填項。", parent=self)
            return
        try:
            datetime.datetime.strptime(new_date, '%Y-%m-%d')
        except ValueError:
            messagebox.showwarning("輸入錯誤", "日期格式錯誤，請使用 YYYY-MM-DD。", parent=self)
            return
        try:
            new_amount = float(new_amount_str)
            if new_amount <= 0:
                messagebox.showwarning("輸入錯誤", "金額必須是正數。", parent=self)
                return
        except ValueError:
            messagebox.showwarning("輸入錯誤", "金額必須是有效的數字。", parent=self)
            return

        new_record = {
            '日期': new_date,
            '類型': new_type,
            '類別': new_category,
            '金額': new_amount,
            '備註': new_note
        }
        
        try:
            all_records = read_records()
            updated_records = []
            found = False
            for record in all_records:
                if record == self.original_record and not found:
                    updated_records.append(new_record)
                    found = True
                else:
                    updated_records.append(record)

            if not found:
                messagebox.showerror("錯誤", "找不到原始紀錄，無法更新。", parent=self)
                return

            with open(CSV_FILE, 'w', encoding='utf-8', newline='') as file:
                writer = csv.DictWriter(file, fieldnames=HEADERS)
                writer.writeheader()
                writer.writerows(updated_records)
            
            messagebox.showinfo("成功", "紀錄已成功更新！", parent=self.master)
            self.callback() # 觸發主視窗的更新
            self.destroy()

        except Exception as e:
            messagebox.showerror("更新錯誤", f"更新過程中發生錯誤: {e}", parent=self)
    
    def cancel(self):
        self.destroy()

# --- 類別維護視窗 ---
class CategoryMaintenanceWindow(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.transient(parent)
        self.title("類別維護")
        self.geometry("400x450")
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.callback = callback

        # --- 內部資料 ---
        self.categories = load_categories()

        # --- UI 元件 ---
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # --- 上方：新增區塊 ---
        add_frame = ttk.LabelFrame(frame, text="新增類別")
        add_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.new_category_entry = ttk.Entry(add_frame)
        self.new_category_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        
        add_button = ttk.Button(add_frame, text="新增", command=self.add_category)
        add_button.pack(side=tk.LEFT, padx=5)

        # --- 中間：列表區塊 ---
        list_frame = ttk.LabelFrame(frame, text="現有類別")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)

        self.category_listbox = tk.Listbox(list_frame)
        self.category_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 填充列表
        self.refresh_listbox()

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.category_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.category_listbox.config(yscrollcommand=scrollbar.set)

        # --- 下方：按鈕區塊 ---
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        delete_button = ttk.Button(button_frame, text="刪除選定類別", command=self.delete_category)
        delete_button.pack(side=tk.LEFT, padx=10)

        save_button = ttk.Button(button_frame, text="儲存並關閉", command=self.save_and_close)
        save_button.pack(side=tk.RIGHT, padx=10)

        self.grab_set()
        self.wait_window()

    def refresh_listbox(self):
        """重新整理 Listbox 中的內容"""
        self.category_listbox.delete(0, tk.END)
        for category in sorted(self.categories):
            self.category_listbox.insert(tk.END, category)

    def add_category(self):
        """新增一個類別"""
        new_cat = self.new_category_entry.get().strip()
        if not new_cat:
            messagebox.showwarning("輸入錯誤", "類別名稱不可為空。", parent=self)
            return
        if new_cat in self.categories:
            messagebox.showwarning("輸入錯誤", "這個類別已經存在。", parent=self)
            return
        
        self.categories.append(new_cat)
        self.refresh_listbox()
        self.new_category_entry.delete(0, tk.END)

    def delete_category(self):
        """刪除選定的類別"""
        selected_indices = self.category_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("操作失敗", "請先從列表中選擇一個要刪除的類別。", parent=self)
            return
        
        selected_category = self.category_listbox.get(selected_indices[0])
        
        confirm = messagebox.askyesno("確認刪除", f"您確定要刪除 '{selected_category}' 這個類別嗎？\n(注意：這不會影響已存在的紀錄)", parent=self)
        if confirm:
            self.categories.remove(selected_category)
            self.refresh_listbox()

    def save_and_close(self):
        """儲存變更並關閉視窗"""
        if save_categories(self.categories):
            messagebox.showinfo("成功", "類別已成功儲存！", parent=self)
            if self.callback:
                self.callback() # 更新主畫面的下拉選單
            self.destroy()
        else:
            messagebox.showerror("儲存失敗", "寫入類別檔案時發生錯誤。", parent=self)


# --- GUI 應用程式 ---
class AccountingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("全功能記帳系統")
        self.root.geometry("1280x768")

        # --- 初始化與設定 ---
        set_chinese_font()
        self.ensure_invoice_dir_exists() # 確保發票資料夾存在
        migrate_csv_if_needed()
        init_csv()
        
        # --- 資料變數 ---
        self.total_income_var = tk.StringVar(value="總收入: 0.00")
        self.total_expense_var = tk.StringVar(value="總支出: 0.00")
        self.balance_var = tk.StringVar(value="結餘: 0.00")
        self.categories = []
        self.current_invoice_path = None # 用於保存當前附加的圖片路徑
        self.attached_image_var = tk.StringVar(value="圖片: 無")
        
        # --- 整體佈局 ---
        main_paned_window = ttk.PanedWindow(root, orient=tk.VERTICAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)

        # --- 上方區塊：輸入與圖表 ---
        top_frame = ttk.Frame(main_paned_window, height=350)
        main_paned_window.add(top_frame, weight=1)

        left_panel = ttk.Frame(top_frame, width=350)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        self.chart_notebook = ttk.Notebook(top_frame)
        self.chart_notebook.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.pie_chart_frame = ttk.Frame(self.chart_notebook)
        self.chart_notebook.add(self.pie_chart_frame, text='支出圓餅圖')
        self.pie_canvas_widget = None

        self.bar_chart_frame = ttk.Frame(self.chart_notebook)
        self.chart_notebook.add(self.bar_chart_frame, text='收支趨勢長條圖')
        self.bar_canvas_widget = None

        # --- 月份總覽 Tab ---
        self.monthly_summary_frame = ttk.Frame(self.chart_notebook)
        self.chart_notebook.add(self.monthly_summary_frame, text='月份總覽')
        self.setup_monthly_summary_tab()


        # --- 下方區塊：紀錄列表 ---
        bottom_frame = ttk.Frame(main_paned_window)
        main_paned_window.add(bottom_frame, weight=2)
        
        # --- 左側：控制面板 ---
        # 新增紀錄區塊
        add_frame = ttk.LabelFrame(left_panel, text="新增紀錄")
        add_frame.pack(fill=tk.X, pady=(10, 5))

        ttk.Label(add_frame, text="日期:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        if CALENDAR_ENABLED:
            self.date_entry = DateEntry(add_frame, date_pattern='y-mm-dd')
        else:
            self.date_entry = ttk.Entry(add_frame)
            self.date_entry.insert(0, datetime.date.today().strftime('%Y-%m-%d'))
        self.date_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        
        ttk.Label(add_frame, text="類型:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.type_combobox = ttk.Combobox(add_frame, values=['支出', '收入'], state="readonly")
        self.type_combobox.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        self.type_combobox.set('支出')

        ttk.Label(add_frame, text="類別:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.category_combobox = ttk.Combobox(add_frame)
        self.category_combobox.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(add_frame, text="金額:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.amount_entry = ttk.Entry(add_frame)
        self.amount_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(add_frame, text="備註:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        self.note_entry = ttk.Entry(add_frame)
        self.note_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)

        save_button = ttk.Button(add_frame, text="儲存紀錄", command=self.save_new_record)
        save_button.grid(row=5, column=0, columnspan=1, pady=(10, 5), sticky=tk.E)

        upload_button = ttk.Button(add_frame, text="附加圖片", command=self.attach_invoice_image)
        upload_button.grid(row=5, column=1, columnspan=1, pady=(10, 5), sticky=tk.W)
        
        # 用於顯示附加圖片名稱的標籤
        attached_label = ttk.Label(add_frame, textvariable=self.attached_image_var)
        attached_label.grid(row=6, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
        
        # 類別維護按鈕
        maint_button = ttk.Button(add_frame, text="類別維護", command=self.open_category_manager)
        maint_button.grid(row=7, column=0, columnspan=2, pady=(10, 5))
        
        # 篩選區塊
        filter_frame = ttk.LabelFrame(left_panel, text="篩選與總覽")
        filter_frame.pack(fill=tk.X, pady=10)

        ttk.Label(filter_frame, text="開始日期:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        if CALENDAR_ENABLED:
            self.start_date_entry = DateEntry(filter_frame, date_pattern='y-mm-dd')
            self.start_date_entry.delete(0, tk.END) # 篩選器預設為空
        else:
            self.start_date_entry = ttk.Entry(filter_frame)
        self.start_date_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(filter_frame, text="結束日期:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        if CALENDAR_ENABLED:
            self.end_date_entry = DateEntry(filter_frame, date_pattern='y-mm-dd')
            self.end_date_entry.delete(0, tk.END) # 篩選器預設為空
        else:
            self.end_date_entry = ttk.Entry(filter_frame)
        self.end_date_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        
        filter_button_frame = ttk.Frame(filter_frame)
        filter_button_frame.grid(row=2, columnspan=2, pady=5)
        ttk.Button(filter_button_frame, text="篩選", command=self.filter_and_refresh_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_button_frame, text="清除篩選", command=self.clear_filter_and_refresh).pack(side=tk.LEFT, padx=5)

        # 總覽區塊
        summary_frame = ttk.Frame(filter_frame)
        summary_frame.grid(row=3, columnspan=2, pady=10, sticky=tk.W)
        ttk.Label(summary_frame, textvariable=self.total_income_var, foreground="green").pack(anchor=tk.W)
        ttk.Label(summary_frame, textvariable=self.total_expense_var, foreground="red").pack(anchor=tk.W)
        ttk.Label(summary_frame, textvariable=self.balance_var, font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        
        # --- 紀錄列表 (Treeview) ---
        tree_frame = ttk.LabelFrame(bottom_frame, text="紀錄列表")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # --- Treeview 控制按鈕 ---
        tree_controls_frame = ttk.Frame(tree_frame)
        tree_controls_frame.pack(fill=tk.X, pady=5)
        
        edit_button = ttk.Button(tree_controls_frame, text="編輯選定紀錄", command=self.edit_record)
        edit_button.pack(side=tk.LEFT, padx=5)
        
        delete_button = ttk.Button(tree_controls_frame, text="刪除選定紀錄", command=self.delete_record)
        delete_button.pack(side=tk.LEFT, padx=5)

        view_image_button = ttk.Button(tree_controls_frame, text="查看圖片", command=self.view_attached_image)
        view_image_button.pack(side=tk.LEFT, padx=5)

        self.tree = ttk.Treeview(tree_frame, columns=HEADERS, show='headings', height=10)
        for col in HEADERS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.W)
        
        self.tree.column('備註', width=200)
        self.tree.column('圖片', width=250)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(fill='both', expand=True)
        
        # --- 初始資料載入 ---
        self.filter_and_refresh_data()

    def ensure_invoice_dir_exists(self):
        """檢查並建立存放發票圖片的資料夾"""
        try:
            if not os.path.exists(INVOICE_DIR):
                os.makedirs(INVOICE_DIR)
        except OSError as e:
            messagebox.showerror("建立資料夾失敗", f"無法建立存放發票的資料夾 '{INVOICE_DIR}':\n{e}")
            self.root.quit()

    def open_category_manager(self):
        """開啟類別維護視窗"""
        CategoryMaintenanceWindow(self.root, self.update_categories)

    def setup_monthly_summary_tab(self):
        """建立月份總覽分頁的內容"""
        headers = ['月份', '總收入', '總支出', '結餘']
        self.monthly_summary_tree = ttk.Treeview(self.monthly_summary_frame, columns=headers, show='headings')

        for col in headers:
            self.monthly_summary_tree.heading(col, text=col)
            self.monthly_summary_tree.column(col, anchor=tk.E) # 文字靠右對齊

        self.monthly_summary_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def update_monthly_summary_view(self, records):
        """更新月份總覽表格的內容"""
        # 清除舊資料
        for item in self.monthly_summary_tree.get_children():
            self.monthly_summary_tree.delete(item)

        monthly_summary = defaultdict(lambda: {'收入': 0, '支出': 0})
        for record in records:
            month = record['日期'][:7]  # YYYY-MM
            monthly_summary[month][record['類型']] += record['金額']

        # 根據月份排序 (最新的在最上面)
        sorted_months = sorted(monthly_summary.keys(), reverse=True)

        for month in sorted_months:
            income = monthly_summary[month]['收入']
            expense = monthly_summary[month]['支出']
            balance = income - expense
            
            # 格式化為貨幣字串
            income_str = f"{income:,.2f}"
            expense_str = f"{expense:,.2f}"
            balance_str = f"{balance:,.2f}"

            self.monthly_summary_tree.insert('', tk.END, values=[month, income_str, expense_str, balance_str])

    def update_categories(self):
        """從檔案讀取類別並更新下拉選單"""
        self.categories = load_categories()
        self.category_combobox['values'] = self.categories

    def filter_and_refresh_data(self):
        """根據日期篩選資料並更新所有相關的 UI 元件"""
        start_date_str = self.start_date_entry.get().strip()
        end_date_str = self.end_date_entry.get().strip()
        
        all_records = read_records()
        filtered_records = []

        try:
            start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d').date() if start_date_str else None
            end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d').date() if end_date_str else None

            if start_date and end_date and start_date > end_date:
                messagebox.showwarning("日期錯誤", "開始日期不能晚於結束日期。")
                return

            if not start_date and not end_date:
                filtered_records = all_records
            else:
                for record in all_records:
                    record_date = datetime.datetime.strptime(record['日期'], '%Y-%m-%d').date()
                    if start_date and record_date < start_date:
                        continue
                    if end_date and record_date > end_date:
                        continue
                    filtered_records.append(record)

        except ValueError:
            messagebox.showwarning("格式錯誤", "日期格式錯誤，請使用 YYYY-MM-DD。")
            filtered_records = all_records # 格式錯誤時顯示全部
            
        self.refresh_records_view(filtered_records)
        self.update_summary_view(filtered_records)
        self.update_categories()
        self.plot_pie(filtered_records)
        self.plot_bar_chart(filtered_records)
        self.update_monthly_summary_view(filtered_records) # 更新月份總覽

    def clear_filter_and_refresh(self):
        """清除日期篩選並重新整理"""
        self.start_date_entry.delete(0, tk.END)
        self.end_date_entry.delete(0, tk.END)
        self.filter_and_refresh_data()
        
    def update_summary_view(self, records):
        """更新總收入、總支出和結餘的顯示"""
        total_income = sum(r['金額'] for r in records if r['類型'] == '收入')
        total_expense = sum(r['金額'] for r in records if r['類型'] == '支出')
        balance = total_income - total_expense
        
        self.total_income_var.set(f"總收入: {total_income:,.2f}")
        self.total_expense_var.set(f"總支出: {total_expense:,.2f}")
        self.balance_var.set(f"結餘: {balance:,.2f}")

    def build_record_from_values(self, values):
        """從 treeview 的 values 列表中建立一個 record 字典"""
        record = dict(zip(HEADERS, values))
        try:
            record['金額'] = float(record['金額'])
        except (ValueError, KeyError):
            record['金額'] = 0.0
        return record

    def edit_record(self):
        """開啟一個新視窗來編輯選定的紀錄"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("編輯失敗", "請先在表格中選擇一筆您想編輯的紀錄。")
            return
        
        selected_item = selected_items[0]
        record_values = self.tree.item(selected_item, 'values')
        record_to_edit = self.build_record_from_values(record_values)
        
        EditRecordWindow(self.root, record_to_edit, self.categories, self.filter_and_refresh_data)

    def delete_record(self):
        """刪除在 Treeview 中選定的紀錄"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("刪除失敗", "請先在表格中選擇一筆您想刪除的紀錄。")
            return
        
        selected_item = selected_items[0]
        record_values_to_delete = self.tree.item(selected_item, 'values')

        confirm = messagebox.askyesno("確認刪除", f"您確定要刪除這筆紀錄嗎？\n\n{record_values_to_delete}")
        if not confirm:
            return

        try:
            all_records = read_records()
            record_to_delete_dict = self.build_record_from_values(record_values_to_delete)

            updated_records = []
            found = False
            for record in all_records:
                # 比較時要確保類型一致，並處理圖片欄位可能不存在的舊資料
                if record['日期'] == record_to_delete_dict['日期'] and \
                   record['類型'] == record_to_delete_dict['類型'] and \
                   record['類別'] == record_to_delete_dict['類別'] and \
                   float(record['金額']) == float(record_to_delete_dict['金額']) and \
                   record['備註'] == record_to_delete_dict['備註'] and \
                   record.get('圖片', '') == record_to_delete_dict.get('圖片', '') and \
                   not found:
                    found = True
                    # 刪除對應的圖片檔案
                    if record_to_delete_dict.get('圖片'):
                        try:
                            os.remove(record_to_delete_dict['圖片'])
                        except OSError as e:
                            print(f"刪除圖片失敗: {e}")
                    continue
                updated_records.append(record)
            
            if not found:
                messagebox.showerror("錯誤", "在檔案中找不到對應的紀錄，無法刪除。")
                return

            with open(CSV_FILE, 'w', encoding='utf-8', newline='') as file:
                writer = csv.DictWriter(file, fieldnames=HEADERS)
                writer.writeheader()
                for rec in updated_records:
                    rec['金額'] = str(rec.get('金額', 0.0))
                    writer.writerow(rec)

            messagebox.showinfo("成功", "紀錄已成功刪除！")
            self.filter_and_refresh_data()

        except Exception as e:
            messagebox.showerror("刪除錯誤", f"刪除過程中發生錯誤: {e}")

    def view_attached_image(self):
        """開啟選定紀錄所附加的圖片"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("操作失敗", "請先在表格中選擇一筆紀錄。")
            return
        
        selected_item = selected_items[0]
        record_values = self.tree.item(selected_item, 'values')
        
        image_path_index = HEADERS.index('圖片')
        image_path = record_values[image_path_index]

        if not image_path or not os.path.exists(image_path):
            messagebox.showinfo("查看圖片", "此紀錄沒有附加圖片，或圖片檔案不存在。")
            return

        try:
            if platform.system() == "Windows":
                os.startfile(image_path)
            elif platform.system() == "Darwin": # macOS
                subprocess.call(("open", image_path))
            else: # Linux
                subprocess.call(("xdg-open", image_path))
        except Exception as e:
            messagebox.showerror("開啟失敗", f"無法開啟圖片檔案: {e}")

    def refresh_records_view(self, records):
        """使用給定的紀錄列表重新整理 Treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 根據日期排序，最新的在最上面
        for record in sorted(records, key=lambda r: r['日期'], reverse=True):
            values = [record.get(h, '') for h in HEADERS]
            self.tree.insert('', tk.END, values=values)

    def save_new_record(self):
        """從 GUI 輸入儲存新紀錄，包含圖片路徑"""
        date_str = self.date_entry.get().strip()
        record_type = self.type_combobox.get()
        category = self.category_combobox.get().strip()
        amount_str = self.amount_entry.get().strip()
        note = self.note_entry.get().strip()
        image_path = self.current_invoice_path if self.current_invoice_path else ''

        if not (date_str and record_type and category and amount_str):
            messagebox.showwarning("輸入錯誤", "日期、類型、類別和金額為必填項。")
            return
        try:
            datetime.datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            messagebox.showwarning("輸入錯誤", "日期格式錯誤，請使用 YYYY-MM-DD。")
            return
        try:
            amount = float(amount_str)
            if amount <= 0:
                messagebox.showwarning("輸入錯誤", "金額必須是正數。")
                return
        except ValueError:
            messagebox.showwarning("輸入錯誤", "金額必須是有效的數字。")
            return

        record_data = [date_str, record_type, category, amount, note, image_path]
        if append_record_to_csv(record_data):
            messagebox.showinfo("成功", "紀錄已成功儲存！")
            # 清空輸入欄位
            self.category_combobox.set('')
            self.amount_entry.delete(0, tk.END)
            self.note_entry.delete(0, tk.END)
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, datetime.date.today().strftime('%Y-%m-%d'))
            self.type_combobox.set('支出')
            
            # 重設附加圖片資訊
            self.current_invoice_path = None
            self.attached_image_var.set("圖片: 無")

            self.filter_and_refresh_data()

    def attach_invoice_image(self):
        """讓使用者選擇圖片檔案，複製到指定資料夾，並記錄其相對路徑"""
        source_path = filedialog.askopenfilename(
            title="選擇一張發票圖片",
            filetypes=[("圖片檔案", "*.jpg *.jpeg *.png *.bmp *.tiff"), ("所有檔案", "*.*")]
        )
        if not source_path:
            return

        try:
            # 建立一個獨特的檔案名稱
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            original_filename = os.path.basename(source_path)
            new_filename = f"{timestamp}_{original_filename}"
            
            dest_path = os.path.join(INVOICE_DIR, new_filename)
            
            # 複製檔案
            shutil.copy(source_path, dest_path)
            
            # 記錄相對路徑並更新UI
            self.current_invoice_path = dest_path
            self.attached_image_var.set(f"圖片: {new_filename}")
            messagebox.showinfo("上傳成功", f"圖片 '{original_filename}' 已附加。")

        except Exception as e:
            messagebox.showerror("上傳失敗", f"處理圖片時發生錯誤: {e}")
            self.current_invoice_path = None
            self.attached_image_var.set("圖片: 無")

    def _clear_chart_canvas(self, canvas_type):
        """清除指定的圖表畫布"""
        if canvas_type == 'pie' and self.pie_canvas_widget:
            self.pie_canvas_widget.destroy()
            self.pie_canvas_widget = None
        elif canvas_type == 'bar' and self.bar_canvas_widget:
            self.bar_canvas_widget.destroy()
            self.bar_canvas_widget = None

    def plot_pie(self, records):
        """在 GUI 中使用給定的紀錄繪製圓餅圖 (僅支出)"""
        self._clear_chart_canvas('pie')
        
        expense_records = [r for r in records if r.get('類型') == '支出']
        
        if not expense_records:
            no_data_label = ttk.Label(self.pie_chart_frame, text="沒有任何支出紀錄可供繪製圖表。", font=("Helvetica", 14))
            no_data_label.pack(expand=True)
            self.pie_canvas_widget = no_data_label
            return

        category_summary = defaultdict(float)
        for record in expense_records:
            category_summary[record['類別']] += record['金額']
        
        fig, ax = plt.subplots(figsize=(8, 6), dpi=100)
        ax.pie(list(category_summary.values()), labels=list(category_summary.keys()), autopct='%1.1f%%', startangle=140, textprops={'fontsize': 10})
        ax.set_title('各類別支出圓餅圖', fontsize=16)
        ax.axis('equal')
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.pie_chart_frame)
        canvas_tk_widget = canvas.get_tk_widget()
        canvas_tk_widget.pack(fill=tk.BOTH, expand=True)
        self.pie_canvas_widget = canvas_tk_widget
        canvas.draw()

    def plot_bar_chart(self, records):
        """在 GUI 中繪製每月收支趨勢長條圖"""
        self._clear_chart_canvas('bar')

        if not records:
            no_data_label = ttk.Label(self.bar_chart_frame, text="沒有任何紀錄可供繪製圖表。", font=("Helvetica", 14))
            no_data_label.pack(expand=True)
            self.bar_canvas_widget = no_data_label
            return
        
        monthly_summary = defaultdict(lambda: {'收入': 0, '支出': 0})
        for record in records:
            month = record['日期'][:7] # YYYY-MM
            monthly_summary[month][record['類型']] += record['金額']
        
        if not monthly_summary:
            no_data_label = ttk.Label(self.bar_chart_frame, text="此期間無紀錄可繪製圖表。", font=("Helvetica", 14))
            no_data_label.pack(expand=True)
            self.bar_canvas_widget = no_data_label
            return
            
        sorted_months = sorted(monthly_summary.keys())
        income_values = [monthly_summary[m]['收入'] for m in sorted_months]
        expense_values = [monthly_summary[m]['支出'] for m in sorted_months]

        fig, ax = plt.subplots(figsize=(10, 6), dpi=100)

        bar_width = 0.35
        x = range(len(sorted_months))
        
        ax.bar([i - bar_width/2 for i in x], income_values, bar_width, label='收入', color='g')
        ax.bar([i + bar_width/2 for i in x], expense_values, bar_width, label='支出', color='r')

        ax.set_ylabel('金額')
        ax.set_title('每月收支趨勢圖')
        ax.set_xticks(x)
        ax.set_xticklabels(sorted_months, rotation=45, ha="right")
        ax.legend()

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.bar_chart_frame)
        canvas_tk_widget = canvas.get_tk_widget()
        canvas_tk_widget.pack(fill=tk.BOTH, expand=True)
        self.bar_canvas_widget = canvas_tk_widget
        canvas.draw()

def main():
    root = tk.Tk()
    app = AccountingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
