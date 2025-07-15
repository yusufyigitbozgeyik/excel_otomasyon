import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import excel_utils
import os

PRIMARY_COLOR = "#1976D2"
BG_COLOR = "#F5F5F5"
BTN_COLOR = "#2196F3"
BTN_TEXT_COLOR = "#FFFFFF"
FONT = ("Segoe UI", 11)
TITLE_FONT = ("Segoe UI", 18, "bold")
DESC_FONT = ("Segoe UI", 10, "italic")

class StockApp:
    def __init__(self, root):
        self.root = root
        root.title("Stok Yönetim Paneli")
        self.center_window(950, 600)
        root.configure(bg=BG_COLOR)
        self.filename = None
        self.create_widgets()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        # Üst başlık ve dosya seçme
        top_frame = tk.Frame(self.root, bg=BG_COLOR)
        top_frame.pack(fill=tk.X, pady=(15, 0))
        tk.Label(top_frame, text="Stok Yönetim Paneli", font=TITLE_FONT, fg=PRIMARY_COLOR, bg=BG_COLOR).pack(side=tk.LEFT, padx=20)
        tk.Button(top_frame, text="Excel Dosyası Seç/Yarat", command=self.select_file, font=FONT, bg=BTN_COLOR, fg=BTN_TEXT_COLOR, relief=tk.FLAT, bd=0, cursor="hand2").pack(side=tk.RIGHT, padx=20)

        # Filtre ve arama
        filter_frame = tk.Frame(self.root, bg=BG_COLOR)
        filter_frame.pack(fill=tk.X, pady=(10, 0))
        tk.Label(filter_frame, text="Kategori:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=(20, 5))
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(filter_frame, textvariable=self.category_var, state="readonly", width=15)
        self.category_combo.pack(side=tk.LEFT)
        self.category_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_table())
        tk.Label(filter_frame, text="Arama:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=(20, 5))
        self.search_var = tk.StringVar()
        tk.Entry(filter_frame, textvariable=self.search_var, font=FONT, width=20).pack(side=tk.LEFT)
        tk.Button(filter_frame, text="Ara", command=self.refresh_table, font=FONT, bg=BTN_COLOR, fg=BTN_TEXT_COLOR, relief=tk.FLAT, bd=0, cursor="hand2").pack(side=tk.LEFT, padx=10)
        tk.Button(filter_frame, text="Tümünü Göster", command=self.clear_filters, font=FONT, bg="#BDBDBD", fg="#222", relief=tk.FLAT, bd=0, cursor="hand2").pack(side=tk.LEFT)

        # Tablo
        table_frame = tk.Frame(self.root, bg=BG_COLOR)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        columns = ["Stok Kodu", "Ürün Adı", "Kategori", "Miktar", "Birim", "Kritik Seviye", "Açıklama", "Son Güncelleme"]
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110 if col != "Açıklama" else 180, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Butonlar
        btn_frame = tk.Frame(self.root, bg=BG_COLOR)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        btn_style = {"font": FONT, "bg": BTN_COLOR, "fg": BTN_TEXT_COLOR, "relief": tk.FLAT, "bd": 0, "width": 16, "height": 2, "cursor": "hand2"}
        tk.Button(btn_frame, text="Ürün Ekle", command=self.add_product, **btn_style).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Ürün Güncelle", command=self.update_product, **btn_style).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Ürün Sil", command=self.delete_product, **btn_style).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Stok Giriş/Çıkış", command=self.stock_in_out, **btn_style).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Yedekle", command=self.backup_file, bg="#43A047").pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Rapor", command=self.show_report, bg="#FFA000").pack(side=tk.LEFT, padx=8)

        # Kritik stok uyarısı
        self.critical_label = tk.Label(self.root, text="", font=("Segoe UI", 11, "bold"), fg="#D32F2F", bg=BG_COLOR)
        self.critical_label.pack(pady=(0, 5))

        # Alt bilgi
        tk.Label(self.root, text="github.com/kullanici_adi", font=("Segoe UI", 9), fg="#888", bg=BG_COLOR).pack(side=tk.BOTTOM, pady=5)

    def select_file(self):
        filename = filedialog.askopenfilename(title="Bir stok Excel dosyası seçin veya yeni oluşturun", filetypes=[("Excel Dosyaları", "*.xlsx")])
        if not filename:
            filename = simpledialog.askstring("Yeni Dosya", "Yeni dosya adı girin (ör: stok.xlsx):", parent=self.root)
            if filename:
                if not filename.endswith('.xlsx'):
                    filename += '.xlsx'
                excel_utils.create_stock_file(filename)
                filename = os.path.abspath(filename)
        if filename and excel_utils.file_exists(filename):
            self.filename = filename
            self.refresh_table()
            self.refresh_categories()
            self.show_critical_warning()

    def refresh_table(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            return
        keyword = self.search_var.get().strip()
        category = self.category_var.get().strip() or None
        products = excel_utils.filter_products(self.filename, keyword=keyword, category=category)
        for row in self.tree.get_children():
            self.tree.delete(row)
        for i, p in enumerate(products):
            self.tree.insert("", "end", iid=i, values=p)
        self.show_critical_warning()

    def refresh_categories(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            self.category_combo['values'] = []
            return
        cats = excel_utils.get_categories(self.filename)
        self.category_combo['values'] = [""] + cats

    def clear_filters(self):
        self.search_var.set("")
        self.category_var.set("")
        self.refresh_table()

    def add_product(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            messagebox.showerror("Hata", "Önce bir dosya seçin veya oluşturun!", parent=self.root)
            return
        fields = ["Stok Kodu", "Ürün Adı", "Kategori", "Miktar", "Birim", "Kritik Seviye", "Açıklama"]
        values = []
        for f in fields:
            v = simpledialog.askstring("Ürün Ekle", f"{f}:", parent=self.root)
            if v is None:
                return
            values.append(v)
        excel_utils.add_product(self.filename, values)
        self.refresh_table()
        self.refresh_categories()
        messagebox.showinfo("Başarılı", "Ürün eklendi!", parent=self.root)

    def update_product(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            messagebox.showerror("Hata", "Önce bir dosya seçin veya oluşturun!", parent=self.root)
            return
        selected = self.tree.focus()
        if not selected:
            messagebox.showerror("Hata", "Güncellenecek ürünü seçin!", parent=self.root)
            return
        old_values = self.tree.item(selected)['values']
        fields = ["Stok Kodu", "Ürün Adı", "Kategori", "Miktar", "Birim", "Kritik Seviye", "Açıklama"]
        new_values = []
        for i, f in enumerate(fields):
            v = simpledialog.askstring("Ürün Güncelle", f"{f}:", initialvalue=old_values[i], parent=self.root)
            if v is None:
                return
            new_values.append(v)
        excel_utils.update_product(self.filename, int(selected), new_values)
        self.refresh_table()
        self.refresh_categories()
        messagebox.showinfo("Başarılı", "Ürün güncellendi!", parent=self.root)

    def delete_product(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            messagebox.showerror("Hata", "Önce bir dosya seçin veya oluşturun!", parent=self.root)
            return
        selected = self.tree.focus()
        if not selected:
            messagebox.showerror("Hata", "Silinecek ürünü seçin!", parent=self.root)
            return
        if messagebox.askyesno("Onay", "Seçili ürünü silmek istediğinize emin misiniz?", parent=self.root):
            excel_utils.delete_product(self.filename, int(selected))
            self.refresh_table()
            self.refresh_categories()
            messagebox.showinfo("Başarılı", "Ürün silindi!", parent=self.root)

    def stock_in_out(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            messagebox.showerror("Hata", "Önce bir dosya seçin veya oluşturun!", parent=self.root)
            return
        selected = self.tree.focus()
        if not selected:
            messagebox.showerror("Hata", "Stok işlemi için bir ürün seçin!", parent=self.root)
            return
        miktar = simpledialog.askfloat("Stok Giriş/Çıkış", "Miktar (pozitif: giriş, negatif: çıkış):", parent=self.root)
        if miktar is None:
            return
        excel_utils.stock_in_out(self.filename, int(selected), miktar)
        self.refresh_table()
        self.show_critical_warning()
        messagebox.showinfo("Başarılı", "Stok güncellendi!", parent=self.root)

    def show_critical_warning(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            self.critical_label.config(text="")
            return
        critical = excel_utils.get_critical_products(self.filename)
        if critical:
            self.critical_label.config(text=f"Kritik seviyede ürün(ler) var! ({len(critical)} ürün)")
        else:
            self.critical_label.config(text="")

    def backup_file(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            messagebox.showerror("Hata", "Önce bir dosya seçin veya oluşturun!", parent=self.root)
            return
        backup_name = excel_utils.backup_file(self.filename)
        messagebox.showinfo("Yedekleme", f"Yedek dosya oluşturuldu:\n{backup_name}", parent=self.root)

    def show_report(self):
        if not self.filename or not excel_utils.file_exists(self.filename):
            messagebox.showerror("Hata", "Önce bir dosya seçin veya oluşturun!", parent=self.root)
            return
        summary = excel_utils.get_summary(self.filename)
        msg = (
            f"Toplam Ürün: {summary['toplam_urun']}\n"
            f"Toplam Stok: {summary['toplam_stok']}\n"
            f"Kritik Ürün Sayısı: {summary['kritik_urun']}"
        )
        messagebox.showinfo("Rapor", msg, parent=self.root)

if __name__ == "__main__":
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop() 