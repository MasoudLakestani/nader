import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import sys

class GoldTransactionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gold Transaction Manager - مدیریت معاملات طلا")
        self.root.geometry("700x600")
        self.root.configure(bg='#f0f0f0')
        
        # Set window icon (optional)
        try:
            self.root.iconbitmap(default='icon.ico')
        except:
            pass
        
        # Center the window on screen
        self.center_window()
        
        # Constants
        self.FILE_NAME = 'transactions.xlsx'
        self.METHQAL_TO_GRAM = 4.3317
        
        # Initialize Excel file
        self.init_file()
        
        # Create GUI
        self.create_widgets()
        
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        pos_x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        pos_y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{pos_x}+{pos_y}')
        
    def init_file(self):
        """Create Excel file if it doesn't exist"""
        try:
            if not os.path.exists(self.FILE_NAME):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Transactions"
                headers = ['type', 'date', 'weight', 'karat', 'price_per_gram', 'price_per_methqal', 'note']
                ws.append(headers)
                wb.save(self.FILE_NAME)
                messagebox.showinfo("Information", f"Excel file created successfully!\nFile location: {os.path.abspath(self.FILE_NAME)}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not create Excel file: {str(e)}")
    
    def create_widgets(self):
        # Main title
        title_label = tk.Label(self.root, text="مدیریت معاملات طلا", 
                              font=("Arial", 16, "bold"), bg='#f0f0f0')
        title_label.pack(pady=20)
        
        # Button frame
        button_frame = tk.Frame(self.root, bg='#f0f0f0')
        button_frame.pack(pady=20)
        
        # Buy button
        buy_button = tk.Button(button_frame, text="ثبت خرید طلا", 
                              command=self.open_buy_dialog,
                              bg='#4CAF50', fg='white', 
                              font=("Arial", 12, "bold"),
                              width=20, height=2)
        buy_button.pack(pady=10)
        
        # Sell button
        sell_button = tk.Button(button_frame, text="ثبت فروش طلا", 
                               command=self.open_sell_dialog,
                               bg='#f44336', fg='white',
                               font=("Arial", 12, "bold"),
                               width=20, height=2)
        sell_button.pack(pady=10)
        
        # Inventory button
        inventory_button = tk.Button(button_frame, text="نمایش موجودی", 
                                    command=self.show_inventory,
                                    bg='#2196F3', fg='white',
                                    font=("Arial", 12, "bold"),
                                    width=20, height=2)
        inventory_button.pack(pady=10)
        
        # View transactions button
        view_button = tk.Button(button_frame, text="نمایش تمام تراکنش‌ها", 
                               command=self.show_all_transactions,
                               bg='#FF9800', fg='white',
                               font=("Arial", 12, "bold"),
                               width=20, height=2)
        view_button.pack(pady=10)
        
        # Exit button
        exit_button = tk.Button(button_frame, text="خروج", 
                               command=self.root.quit,
                               bg='#607D8B', fg='white',
                               font=("Arial", 12, "bold"),
                               width=20, height=2)
        exit_button.pack(pady=10)
        
        # Status label
        self.status_label = tk.Label(self.root, text="آماده برای استفاده", 
                                    font=("Arial", 10), bg='#f0f0f0')
        self.status_label.pack(side=tk.BOTTOM, pady=10)
    
    def open_buy_dialog(self):
        """Open dialog for buying gold"""
        dialog = tk.Toplevel(self.root)
        dialog.title("ثبت خرید طلا")
        dialog.geometry("400x300")
        dialog.configure(bg='#f0f0f0')
        
        # Weight input
        tk.Label(dialog, text="وزن خرید (گرم):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        weight_entry = tk.Entry(dialog, font=("Arial", 10))
        weight_entry.pack(pady=5)
        
        # Karat input
        tk.Label(dialog, text="عیار طلا (مثلاً 750 یا 900):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        karat_entry = tk.Entry(dialog, font=("Arial", 10))
        karat_entry.pack(pady=5)
        
        # Price per methqal input
        tk.Label(dialog, text="قیمت هر مثقال (تومان):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        price_entry = tk.Entry(dialog, font=("Arial", 10))
        price_entry.pack(pady=5)
        
        # Note input
        tk.Label(dialog, text="توضیح (اختیاری):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        note_entry = tk.Entry(dialog, font=("Arial", 10))
        note_entry.pack(pady=5)
        
        # Save button
        def save_buy():
            try:
                weight = float(weight_entry.get())
                karat = int(karat_entry.get())
                price_per_methqal = float(price_entry.get())
                price_per_gram = round(price_per_methqal / self.METHQAL_TO_GRAM, 2)
                note = note_entry.get()
                date = datetime.now().strftime("%Y-%m-%d")
                
                self.save_transaction('buy', date, weight, karat,
                                     price_per_gram, price_per_methqal, note)
                
                messagebox.showinfo("موفق", f"خرید ثبت شد. قیمت هر گرم: {price_per_gram:,.0f} تومان")
                dialog.destroy()
                self.status_label.config(text="خرید جدید ثبت شد")
            except ValueError:
                messagebox.showerror("خطا", "لطفاً مقادیر صحیح وارد کنید")
        
        tk.Button(dialog, text="ثبت خرید", command=save_buy, 
                 bg='#4CAF50', fg='white', font=("Arial", 10, "bold")).pack(pady=20)
    
    def open_sell_dialog(self):
        """Open dialog for selling gold"""
        dialog = tk.Toplevel(self.root)
        dialog.title("ثبت فروش طلا")
        dialog.geometry("400x300")
        dialog.configure(bg='#f0f0f0')
        
        # Weight input
        tk.Label(dialog, text="وزن فروش (گرم):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        weight_entry = tk.Entry(dialog, font=("Arial", 10))
        weight_entry.pack(pady=5)
        
        # Karat input
        tk.Label(dialog, text="عیار طلا (مثلاً 750 یا 900):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        karat_entry = tk.Entry(dialog, font=("Arial", 10))
        karat_entry.pack(pady=5)
        
        # Price per gram input
        tk.Label(dialog, text="قیمت فروش هر گرم (تومان):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        price_entry = tk.Entry(dialog, font=("Arial", 10))
        price_entry.pack(pady=5)
        
        # Note input
        tk.Label(dialog, text="توضیح (اختیاری):", font=("Arial", 10), bg='#f0f0f0').pack(pady=5)
        note_entry = tk.Entry(dialog, font=("Arial", 10))
        note_entry.pack(pady=5)
        
        # Save button
        def save_sell():
            try:
                weight = float(weight_entry.get())
                karat = int(karat_entry.get())
                price_per_gram = float(price_entry.get())
                note = note_entry.get()
                date = datetime.now().strftime("%Y-%m-%d")
                
                self.save_transaction('sell', date, weight, karat,
                                     price_per_gram, '', note)
                
                messagebox.showinfo("موفق", "فروش ثبت شد")
                dialog.destroy()
                self.status_label.config(text="فروش جدید ثبت شد")
            except ValueError:
                messagebox.showerror("خطا", "لطفاً مقادیر صحیح وارد کنید")
        
        tk.Button(dialog, text="ثبت فروش", command=save_sell, 
                 bg='#f44336', fg='white', font=("Arial", 10, "bold")).pack(pady=20)
    
    def show_inventory(self):
        """Show current inventory"""
        total = 0.0
        transactions = self.load_transactions()
        
        for t in transactions:
            if t['type'] == 'buy':
                total += float(t['weight'])
            elif t['type'] == 'sell':
                total -= float(t['weight'])
        
        messagebox.showinfo("موجودی", f"موجودی فعلی طلا: {total:.2f} گرم")
        self.status_label.config(text=f"موجودی: {total:.2f} گرم")
    
    def show_all_transactions(self):
        """Show all transactions in a new window"""
        transactions = self.load_transactions()
        
        if not transactions:
            messagebox.showinfo("اطلاع", "هیچ تراکنشی یافت نشد")
            return
        
        # Create new window
        trans_window = tk.Toplevel(self.root)
        trans_window.title("تمام تراکنش‌ها")
        trans_window.geometry("800x400")
        
        # Create treeview
        tree = ttk.Treeview(trans_window, columns=('type', 'date', 'weight', 'karat', 'price_gram', 'price_methqal', 'note'), show='headings')
        
        # Define headings
        tree.heading('type', text='نوع')
        tree.heading('date', text='تاریخ')
        tree.heading('weight', text='وزن')
        tree.heading('karat', text='عیار')
        tree.heading('price_gram', text='قیمت/گرم')
        tree.heading('price_methqal', text='قیمت/مثقال')
        tree.heading('note', text='توضیح')
        
        # Insert data
        for trans in transactions:
            tree.insert('', 'end', values=(
                'خرید' if trans['type'] == 'buy' else 'فروش',
                trans['date'],
                trans['weight'],
                trans['karat'],
                trans['price_per_gram'],
                trans['price_per_methqal'],
                trans['note']
            ))
        
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(trans_window, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
    
    def save_transaction(self, t_type, date, weight, karat, price_per_gram, price_per_methqal, note):
        """Save transaction to Excel file"""
        wb = openpyxl.load_workbook(self.FILE_NAME)
        ws = wb["Transactions"]
        ws.append([t_type, date, weight, karat, price_per_gram, price_per_methqal, note])
        wb.save(self.FILE_NAME)
    
    def load_transactions(self):
        """Load all transactions from Excel file"""
        if not os.path.exists(self.FILE_NAME):
            return []
        
        wb = openpyxl.load_workbook(self.FILE_NAME)
        ws = wb["Transactions"]
        transactions = []
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            transactions.append({
                'type': row[0],
                'date': row[1],
                'weight': row[2],
                'karat': row[3],
                'price_per_gram': row[4],
                'price_per_methqal': row[5],
                'note': row[6]
            })
        
        return transactions

if __name__ == '__main__':
    root = tk.Tk()
    app = GoldTransactionApp(root)
    root.mainloop()