import os
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, simpledialog
import win32com.client
import win32print
from win32com.client import Dispatch
import sqlite3
from tkinter import ttk
import subprocess
import win32con
from ctypes import Structure, windll, c_int, sizeof, byref

def set_word_printer(word_app, printer_name):
    word_app.ActivePrinter = printer_name

class PrintingVendingMachine:
    def __init__(self, master):
        self.master = master
        master.title("Print Vending Machine")
        master.geometry("1920x1080")

        self.base_cost = 2
        self.total_cost = 0
        self.user_balance = 0

        self.print_color = StringVar()
        self.print_paper_size = StringVar()
        self.quantity = tk.IntVar(value=1)

        self.conn = sqlite3.connect('print_records.db')
        self.create_table()

        self.settings_and_info_frame = tk.Frame(self.master)
        self.settings_and_info_frame.pack(pady=10)

        self.create_widgets()

        self.admin_password = "12345"

        self.total_amount_received = 0 

        self.set_default_values()

    def set_default_values(self):
        self.print_color.set("Choose Color Here")
        self.print_paper_size.set("Choose Size Here")

    def create_table(self):
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute('DROP TABLE IF EXISTS print_records')
            cursor.execute('''
                CREATE TABLE print_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT,
                    color_option TEXT,
                    paper_size_option TEXT,
                    quantity INTEGER,
                    total_cost INTEGER,
                    total_amount_received INTEGER
                )   
            ''')

    def insert_record(self, file_path, quantity, total_cost):
        color_option = self.print_color.get()
        paper_size_option = self.print_paper_size.get()

        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute('''
                INSERT INTO print_records (file_path, color_option, paper_size_option, quantity, total_cost, total_amount_received)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (file_path, color_option, paper_size_option, quantity, total_cost, self.total_amount_received))

    def create_widgets(self):
        file_frame = tk.Frame(self.master, pady=10)
        file_frame.pack(side=tk.TOP)

        font_style = ("Verdana", 13)
        bold_font = ("Verdana", 13, "bold")

        admin_and_reset_frame = tk.Frame(file_frame)
        admin_and_reset_frame.pack(side=tk.TOP, padx=(1800,0), pady=(0,0))

        file_label = tk.Label(file_frame, text="Choose Word File", font=font_style)
        file_label.pack(side=tk.TOP, padx=(0,0), pady=(0,10))


        browse_button = tk.Button(file_frame, text="Browse", command=self.browse_file, font=font_style)
        browse_button.pack(side=tk.LEFT, padx=(500,0), pady=(0.0))

        self.file_entry = tk.Entry(file_frame, width=70, font=font_style)
        self.file_entry.pack(side=tk.LEFT, padx=(10,0))

        admin_button = tk.Button(admin_and_reset_frame, text="Admin", command=self.view_database_file, font=font_style)
        admin_button.pack(padx=5)

        reset_button = tk.Button(admin_and_reset_frame, text="Reset", command=self.reset_selections, font=font_style)
        reset_button.pack(padx=5)

        settings_and_info_frame = tk.Frame(self.master)
        settings_and_info_frame.pack(pady=10, fill=tk.X)

        customization_frame = tk.Frame(settings_and_info_frame, padx=30, pady=30, borderwidth=2, relief=tk.RIDGE)
        customization_frame.grid(row=0, column=0, sticky="e", padx=(500, 20))

        color_frame = tk.Frame(customization_frame, pady=5)
        color_frame.pack(side=tk.TOP)

        color_option_label = tk.Label(color_frame, text="Color Option:", font=font_style)
        color_option_label.pack(side=tk.TOP)

        color_options = ["Colored (+3 PHP)", "Grayscale (+2 PHP)"]
        self.color_dropdown = ttk.Combobox(color_frame, values=color_options, textvariable=self.print_color, font=font_style)
        self.color_dropdown.pack(side=tk.TOP)
        self.color_dropdown.set(color_options[0])

        size_frame = tk.Frame(customization_frame, pady=5)
        size_frame.pack(side=tk.TOP, padx=5, pady=5)   

        size_option_label = tk.Label(size_frame, text="Paper Size Option:", font=font_style)
        size_option_label.pack(side=tk.TOP)

        size_options = ["Short (+1 PHP)", "Long (+2 PHP)"]
        self.size_dropdown = ttk.Combobox(size_frame, values=size_options, textvariable=self.print_paper_size, font=font_style)
        self.size_dropdown.pack(side=tk.TOP)
        self.size_dropdown.set(size_options[0]) 

        quantity_frame = tk.Frame(customization_frame, pady=5)
        quantity_frame.pack(side=tk.TOP, padx=10, pady=5)

        quantity_label = tk.Label(quantity_frame, text="No. of Copies:", font=font_style)
        quantity_label.pack(side=tk.TOP)

        quantity_spinbox = tk.Spinbox(quantity_frame, from_=0, to=10, textvariable=self.quantity, font=font_style)
        quantity_spinbox.pack(side=tk.TOP)

        info_frame = tk.Frame(settings_and_info_frame, padx=30, pady=30, borderwidth=2, relief=tk.RIDGE)
        info_frame.grid(row=0, column=1, sticky="w", padx=(5, 20))  # Use the same padding values as customization_frame

        self.total_cost_label = tk.Label(info_frame, text="Total Cost\n 0", font=font_style)
        self.total_cost_label.pack(pady=5)

        self.num_pages_label = tk.Label(info_frame, text="Number of Pages\n 0", font=font_style)
        self.num_pages_label.pack(pady=5)

        self.user_balance_label = tk.Label(info_frame, text="User Balance\n  0", font=font_style)
        self.user_balance_label.pack(pady=5)

        insert_2_php_button = tk.Button(self.master, text="Insert 2 PHP", command=self.insert_money, font=font_style)
        insert_2_php_button.pack(padx=5, pady=10)

        self.print_button = tk.Button(self.master, text="Print", command=self.print_file, font=font_style)
        self.print_button.pack(padx=5, pady=10)
        self.print_button.config(state=tk.DISABLED) 

        self.print_color.trace_add("write", lambda *args: self.calculate_total_cost())
        self.print_paper_size.trace_add("write", lambda *args: self.calculate_total_cost())
        self.quantity.trace_add("write", lambda *args: self.calculate_total_cost())


    def view_database_file(self):
        entered_password = simpledialog.askstring("Password", "Enter password:", show='*')

        if entered_password == self.admin_password:
            view_window = tk.Toplevel(self.master)
            view_window.title("View Database")


            tree = ttk.Treeview(view_window)
            tree["columns"] = ("ID", "File Path", "Color Option", "Paper Size Option", "Quantity", "Total Cost", "Total Amount Received")

            tree.column("#0", width=0, stretch=tk.NO)
            tree.column("ID", anchor=tk.W, width=200)
            tree.column("File Path", anchor=tk.W, width=150)
            tree.column("Color Option", anchor=tk.W, width=100)
            tree.column("Paper Size Option", anchor=tk.W, width=100)
            tree.column("Quantity", anchor=tk.W, width=80)
            tree.column("Total Cost", anchor=tk.W, width=80)
            tree.column("Total Amount Received", anchor=tk.W, width=80) 

            tree.heading("#0", text="", anchor=tk.W)
            tree.heading("ID", text="ID", anchor=tk.W)
            tree.heading("File Path", text="File Path", anchor=tk.W)
            tree.heading("Color Option", text="Color Option", anchor=tk.W)
            tree.heading("Paper Size Option", text="Paper Size Option", anchor=tk.W)
            tree.heading("Quantity", text="Quantity", anchor=tk.W)
            tree.heading("Total Cost", text="Total Cost", anchor=tk.W) 
            tree.heading("Total Amount Received", text="Total Amount Received", anchor=tk.W)  
        else:
            messagebox.showerror("Authentication Failed", "Incorrect password. Access denied.")

        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute("SELECT * FROM print_records")
            records = cursor.fetchall()

            for record in records:
                tree.insert("", tk.END, values=record)

        total_amount_received = sum(record[5] for record in records) 
        tree.insert("", tk.END, values=["Total Amount Received", "", "", "", "", "", total_amount_received])

        tree.pack(expand=tk.YES, fill=tk.BOTH)

    def reset_selections(self):
        file_path = self.file_entry.get()
        if file_path:
            self.calculate_total_cost()
            self.insert_record("", self.quantity.get(), self.total_cost)

        self.file_entry.delete(0, tk.END)
        self.print_color.set("Choose Color Here")
        self.print_paper_size.set("Choose Size Here")
        self.quantity.set(1) 
        self.num_pages_label.config(text="Number of Pages\n 0")

        self.total_cost = 0
        self.user_balance = 0

        self.total_cost_label.config(text="Total Cost\n  0 PHP")
        self.update_user_balance_label() 
        self.print_button.config(state=tk.DISABLED)

        if self.file_entry.get():
            self.calculate_total_cost()
            self.insert_record("", self.quantity.get(), self.total_cost)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx;*.doc")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

            num_pages = self.get_number_of_pages(file_path)
            self.total_cost = num_pages * self.base_cost * self.quantity.get()

            self.calculate_total_cost()

            self.num_pages_label.config(text=f"Number of Pages\n {num_pages}")

    def insert_money(self):
        file_path = self.file_entry.get()

        if not file_path:
            messagebox.showwarning("Warning", "Please choose a Word file first.")
            return

        self.user_balance += 2
        self.total_amount_received += 2  
        self.calculate_total_cost()  
        self.update_user_balance_label()
        self.total_cost_label.config(text=f"Total Cost\n  {self.total_cost} PHP")

        if self.user_balance >= self.total_cost:
            self.print_button.config(state=tk.NORMAL)
        else:
            self.print_button.config(state=tk.DISABLED)


        if file_path:
            num_pages = self.get_number_of_pages(file_path)
            self.num_pages_label.config(text=f"Number of Pages\n {num_pages}")

    def print_file(self):
        file_path = self.file_entry.get()
        if not file_path:
            messagebox.showerror("Error", "Please choose a Word file first.")
            return

        if self.user_balance < self.total_cost:
            messagebox.showerror("Error", "Insufficient funds. Please insert more money.")
            return

        try:
            self.user_balance -= self.total_cost
            self.update_user_balance_label()

            color_option = self.print_color.get()
            paper_size_option = self.print_paper_size.get()

            for _ in range(self.quantity.get()):
                self.print_word_file(file_path, color_option, paper_size_option)

            self.insert_record(file_path, self.quantity.get(), self.total_cost)

            messagebox.showinfo("Print Successful", f"The document has been sent to the printer. Change: {self.user_balance} Pesos")
        except Exception as e:
            messagebox.showerror("Error", f"Error printing Word file: {e}")

    def print_word_file(self, word_file, color_option, paper_size_option):
        try:
            if "Colored" in color_option:
                printers = ["Canon iP2700 Series", "HP Deskjet Ink Adv 2010 K010"]
            elif "Grayscale" in color_option:
                printers = ["Canon iP2700 Series", "HP Deskjet Ink Adv 2010 K010"]
            else:
                messagebox.showwarning("Unsupported Color Option", "The selected color option is not supported.")
                return

            if "Short" in paper_size_option:
                paper_size_constant = win32con.DMPAPER_LETTER
                selected_printer = "Canon iP2700 Series"
            elif "Long" in paper_size_option:
                paper_size_constant = win32con.DMPAPER_LEGAL
                selected_printer = "HP Deskjet Ink Adv 2010 K010"
            else:
                messagebox.showwarning("Unsupported Paper Size", "Invalid paper size option selected.")
                return

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False 

            set_word_printer(word, selected_printer)

            doc = word.Documents.Open(word_file)
            doc.PrintOut()

            word.Quit()

        except Exception as e:
            messagebox.showerror("Error", f"Error printing Word file: {e}")

    def calculate_total_cost(self, *args):
        file_path = self.file_entry.get()

        if (
            self.print_color.get() != "Choose Color Here"
            and self.print_paper_size.get() != "Choose Size Here"
            and self.quantity.get() > 0
        ):
            color_option = self.print_color.get()
            paper_size_option = self.print_paper_size.get()

            cost_adjustment = 0

            # Price adjustments
            if "Short" in paper_size_option:
                cost_adjustment += 1
            elif "Long" in paper_size_option:
                cost_adjustment += 2

            if "Colored" in color_option:
                cost_adjustment += 3
            elif "Grayscale" in color_option:
                cost_adjustment += 2

            num_pages = self.get_number_of_pages(file_path)

            self.base_cost = 2
            self.total_cost = num_pages * (self.base_cost + cost_adjustment) * self.quantity.get()

            self.total_cost_label.config(text=f"Total Cost\n  {self.total_cost} PHP")

            if self.user_balance >= self.total_cost and file_path:
                self.print_button.config(state=tk.NORMAL)
            else:
                self.print_button.config(state=tk.DISABLED)

    def update_user_balance_label(self):
        self.user_balance_label.config(text=f"User Balance\n {self.user_balance} PHP")

    def get_number_of_pages(self, file_path):
        try:
            word = Dispatch('Word.Application')
            word.Visible = False

            file_path = os.path.abspath(file_path)

            if not os.path.isfile(file_path):
                messagebox.showerror("Error", "File not found. Please check the file path.")
                return 0

            word = word.Documents.Open(file_path)
            word.Repaginate()
            num_of_sheets = word.ComputeStatistics(2)
            word.Close()

            return num_of_sheets
        except Exception as e:
            messagebox.showerror("Error", f"Error counting pages: {e}")
            print(f"Error details: {e}")
            return 0

if __name__ == "__main__":
    root = tk.Tk()
    app = PrintingVendingMachine(root)
    root.mainloop()