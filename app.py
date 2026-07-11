# =========================================================
# NON AIR OPERATIONS ERP - FINAL VERSION
# =========================================================

from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import pandas as pd
import sqlite3

# =========================================================
# DATABASE
# =========================================================

conn = sqlite3.connect("non_air_erp.db")

cursor = conn.cursor()

# =========================================================
# INVOICES TABLE
# =========================================================

cursor.execute("""

CREATE TABLE IF NOT EXISTS invoices (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    date TEXT,
    invoice_no TEXT,
    owner TEXT,
    accounts TEXT,
    service_date TEXT,
    subject TEXT,
    service_type TEXT,
    supplier TEXT,
    supplier_city TEXT,
    no_of_pax INTEGER,
    supplier_invoice_no TEXT,
    po_number TEXT,
    total_amount REAL,
    paid_to_supplier REAL,
    handling_fees REAL,
    vat REAL,
    currency TEXT,
    payment_status TEXT,
    created_by TEXT

)

""")

# =========================================================
# USERS TABLE
# =========================================================

cursor.execute("""

CREATE TABLE IF NOT EXISTS users (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    username TEXT UNIQUE,
    password TEXT

)

""")

conn.commit()

# =========================================================
# LOAD FILES
# =========================================================

vendors_df = pd.read_excel("vendors.xlsx").fillna("")

issuers_df = pd.read_excel("issuers.xlsx").fillna("")

accounts_df = pd.read_excel("accounts.xlsx").fillna("")

if "City" not in vendors_df.columns:

    vendors_df["City"] = ""

vendors = vendors_df["Alias"].astype(str).tolist()

issuers = issuers_df.iloc[:, 0].astype(str).tolist()

accounts = accounts_df.iloc[:, 0].astype(str).tolist()

# =========================================================
# AUTO CREATE USERS
# =========================================================

for issuer in issuers:

    cursor.execute(
        "SELECT * FROM users WHERE username=?",
        (issuer,)
    )

    existing_user = cursor.fetchone()

    if not existing_user:

        cursor.execute(
            "INSERT INTO users (username, password) VALUES (?, ?)",
            (issuer, "1234")
        )

conn.commit()

# =========================================================
# SEARCHABLE COMBOBOX
# =========================================================

EGYPT_GOVERNORATES = [
    "Cairo",
    "Giza",
    "Alexandria",
    "Qalyubia",
    "Port Said",
    "Suez",
    "Damietta",
    "Dakahlia",
    "Sharqia",
    "Gharbia",
    "Monufia",
    "Beheira",
    "Ismailia",
    "Kafr El Sheikh",
    "North Sinai",
    "South Sinai",
    "Beni Suef",
    "Faiyum",
    "Minya",
    "Asyut",
    "Sohag",
    "Qena",
    "Luxor",
    "Aswan",
    "Red Sea",
    "New Valley",
    "Matrouh"
]

def add_hover_effect(widget, normal_color, hover_color):

    widget.bind(
        "<Enter>",
        lambda e: widget.config(bg=hover_color)
    )

    widget.bind(
        "<Leave>",
        lambda e: widget.config(bg=normal_color)
    )

def searchable_combobox(combo, data):

    def check_input(event):

        value = combo.get()

        if value == "":

            combo["values"] = data

        else:

            filtered = []

            for item in data:

                if value.lower() in item.lower():

                    filtered.append(item)

            combo["values"] = filtered

    combo.bind("<KeyRelease>", check_input)

# =========================================================
# WINDOW
# =========================================================

root = Tk()

root.title("Non-Air Operations ERP")

root.state("zoomed")

root.configure(bg="#020617")

# =========================================================
# CURRENT USER
# =========================================================

current_user = ""

# =========================================================
# LOGIN
# =========================================================

def login():

    global current_user

    username = user_entry.get()

    password = pass_entry.get()

    cursor.execute(

        "SELECT * FROM users WHERE username=? AND password=?",

        (username, password)

    )

    user = cursor.fetchone()

    if user:

        current_user = username

        login_frame.destroy()

        open_system()

    else:

        messagebox.showerror(
            "Login Failed",
            "Wrong Username or Password"
        )

# =========================================================
# MAIN SYSTEM
# =========================================================

def open_system():

    # =====================================================
    # SIDEBAR
    # =====================================================

    sidebar = Frame(
        root,
        bg="#020617",
        width=230
    )

    sidebar.pack(side=LEFT, fill=Y)

    # =====================================================
    # LOGO
    # =====================================================

    img = Image.open("logo.png")

    img = img.resize((120, 120))

    logo = ImageTk.PhotoImage(img)

    logo_label = Label(
        sidebar,
        image=logo,
        bg="#020617"
    )

    logo_label.image = logo

    logo_label.pack(pady=20)

    # =====================================================
    # TITLE
    # =====================================================

    Label(
        sidebar,
        text="Non-Air\nOperations ERP",
        font=("Segoe UI", 18, "bold"),
        bg="#020617",
        fg="#00BFFF"
    ).pack(pady=10)

    # =====================================================
    # USER
    # =====================================================

    Label(
        sidebar,
        text=f"Welcome\n{current_user}",
        font=("Segoe UI", 11, "bold"),
        bg="#020617",
        fg="#00FF99"
    ).pack(pady=10)

    # =====================================================
    # CONTENT
    # =====================================================

    content_frame = Frame(
        root,
        bg="#020617"
    )

    content_frame.pack(fill=BOTH, expand=True)

    # =====================================================
    # PAGES
    # =====================================================

    dashboard_page = Frame(content_frame, bg="#020617")

    invoices_page = Frame(content_frame, bg="#020617")

    reports_page = Frame(content_frame, bg="#020617")

    vendors_page = Frame(content_frame, bg="#020617")

    support_page = Frame(content_frame, bg="#020617")

    admin_page = Frame(content_frame, bg="#020617")

    # =====================================================
    # HIDE PAGES
    # =====================================================

    def hide_frames():

        dashboard_page.pack_forget()

        invoices_page.pack_forget()

        reports_page.pack_forget()

        vendors_page.pack_forget()

        support_page.pack_forget()

        admin_page.pack_forget()

    # =====================================================
    # SHOW PAGES
    # =====================================================

    def show_dashboard():

        hide_frames()

        dashboard_page.pack(fill=BOTH, expand=True)

    def show_invoices():

        hide_frames()

        invoices_page.pack(fill=BOTH, expand=True)

    def show_reports():

        hide_frames()

        reports_page.pack(fill=BOTH, expand=True)

        load_reports()

    def show_vendors():

        hide_frames()

        vendors_page.pack(fill=BOTH, expand=True)

    def show_support():

        hide_frames()

        support_page.pack(fill=BOTH, expand=True)

    def show_admin():

        hide_frames()

        admin_page.pack(fill=BOTH, expand=True)

        load_users_table()

    # =====================================================
    # CHANGE PASSWORD
    # =====================================================

    def change_password_window():

        window = Toplevel(root)

        window.title("Change Password")

        window.geometry("400x300")

        window.configure(bg="#1e293b")

        Label(
            window,
            text="New Password",
            font=("Segoe UI", 12, "bold"),
            bg="#1e293b",
            fg="white"
        ).pack(pady=20)

        new_pass_entry = Entry(
            window,
            show="*",
            width=30,
            font=("Segoe UI", 11)
        )

        new_pass_entry.pack(pady=10)

        def save_password():

            new_password = new_pass_entry.get()

            cursor.execute(

                "UPDATE users SET password=? WHERE username=?",

                (new_password, current_user)

            )

            conn.commit()

            messagebox.showinfo(
                "DONE",
                "Password Changed Successfully ✅"
            )

            window.destroy()

        Button(
            window,
            text="SAVE PASSWORD",
            command=save_password,
            bg="#10B981",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat"
        ).pack(pady=20)

    # =====================================================
    # SIDEBAR BUTTON
    # =====================================================

    def sidebar_button(text, command):

        btn = Button(
            sidebar,
            text=text,
            command=command,
            font=("Segoe UI", 10, "bold"),
            bg="#1e293b",
            fg="white",
            activebackground="#0ea5e9",
            activeforeground="white",
            width=20,
            height=2,
            relief="flat",
            cursor="hand2"
        )

        btn.pack(pady=8)

        btn.bind(
            "<Enter>",
            lambda e: btn.config(bg="#334155")
        )

        btn.bind(
            "<Leave>",
            lambda e: btn.config(bg="#1e293b")
        )

    sidebar_button("Dashboard", show_dashboard)

    sidebar_button("Invoices", show_invoices)

    sidebar_button("Reports", show_reports)

    sidebar_button("Vendors", show_vendors)

    sidebar_button("Support", show_support)

    sidebar_button("Change Password", change_password_window)

    if current_user == "Bassma":

        sidebar_button("Admin Panel", show_admin)

    # =====================================================
    # DASHBOARD
    # =====================================================

    dashboard_header = Frame(
        dashboard_page,
        bg="#0f172a",
        highlightbackground="#334155",
        highlightthickness=1
    )

    dashboard_header.pack(fill=X, padx=0, pady=0)

    Label(
        dashboard_header,
        text="Operations Dashboard",
        font=("Segoe UI", 22, "bold"),
        bg="#0f172a",
        fg="white"
    ).pack(side=LEFT, padx=30, pady=18)

    Label(
        dashboard_header,
        text="Brightstar  •  Non-Air Operations",
        font=("Segoe UI", 10),
        bg="#0f172a",
        fg="#64748B"
    ).pack(side=RIGHT, padx=30)

    Label(
        dashboard_page,
        text="KEY PERFORMANCE INDICATORS",
        font=("Segoe UI", 9, "bold"),
        bg="#020617",
        fg="#64748B"
    ).pack(anchor=W, padx=30, pady=(25, 8))

    dashboard_cards = Frame(
        dashboard_page,
        bg="#020617"
    )

    dashboard_cards.pack(padx=30, fill=X)

    def create_card(parent, title, color, subtitle=""):

        card = Frame(
            parent,
            bg="#111827",
            width=250,
            height=130,
            highlightbackground="#1e293b",
            highlightthickness=1
        )

        card.pack(side=LEFT, padx=(0, 16))

        card.pack_propagate(False)

        accent = Frame(
            card,
            bg=color,
            height=3
        )

        accent.pack(side=TOP, fill=X)

        inner = Frame(card, bg="#111827")

        inner.pack(fill=BOTH, expand=True, padx=18, pady=12)

        number = Label(
            inner,
            text="0",
            font=("Segoe UI", 24, "bold"),
            bg="#111827",
            fg=color,
            anchor=W
        )

        number.pack(fill=X, anchor=W)

        Label(
            inner,
            text=title.upper(),
            font=("Segoe UI", 9, "bold"),
            bg="#111827",
            fg="#94A3B8",
            anchor=W
        ).pack(fill=X, anchor=W, pady=(8, 0))

        if subtitle:

            Label(
                inner,
                text=subtitle,
                font=("Segoe UI", 8),
                bg="#111827",
                fg="#475569",
                anchor=W
            ).pack(fill=X, anchor=W)

        return number

    total_sales_label = create_card(
        dashboard_cards,
        "Total Sales",
        "#22C55E",
        "Sum of all invoice totals"
    )

    total_revenue_label = create_card(
        dashboard_cards,
        "Total Revenue",
        "#3B82F6",
        "Sum paid to suppliers"
    )

    invoice_count_label = create_card(
        dashboard_cards,
        "Invoices Count",
        "#F59E0B",
        "Total invoices generated"
    )

    # =====================================================
    # VENDOR ANALYSIS
    # =====================================================

    Label(
        dashboard_page,
        text="VENDOR ANALYSIS",
        font=("Segoe UI", 9, "bold"),
        bg="#020617",
        fg="#64748B"
    ).pack(anchor=W, padx=30, pady=(35, 8))

    vendor_analysis_card = Frame(
        dashboard_page,
        bg="#111827",
        highlightbackground="#1e293b",
        highlightthickness=1
    )

    vendor_analysis_card.pack(padx=30, fill=X)

    vendor_analysis_frame = Frame(
        vendor_analysis_card,
        bg="#111827"
    )

    vendor_analysis_frame.pack(fill=X, padx=20, pady=18)

    Label(
        vendor_analysis_frame,
        text="Select Vendor",
        font=("Segoe UI", 9, "bold"),
        bg="#111827",
        fg="#94A3B8"
    ).pack(anchor=W)

    vendor_select_combo = ttk.Combobox(
        vendor_analysis_frame,
        values=vendors,
        width=45,
        font=("Segoe UI", 10)
    )

    vendor_select_combo.pack(anchor=W, pady=(6, 0))

    searchable_combobox(
        vendor_select_combo,
        vendors
    )

    vendor_cards_frame = Frame(
        dashboard_page,
        bg="#020617"
    )

    vendor_cards_frame.pack(padx=30, pady=(16, 10), fill=X)

    vendor_purchased_label = create_card(
        vendor_cards_frame,
        "Paid To Supplier",
        "#3B82F6"
    )

    vendor_profit_label = create_card(
        vendor_cards_frame,
        "Profit (Handling + VAT)",
        "#22C55E"
    )

    vendor_invoice_count_label = create_card(
        vendor_cards_frame,
        "Invoices With Vendor",
        "#F59E0B"
    )

    def update_vendor_analysis(event=None):

        selected_vendor = vendor_select_combo.get()

        if selected_vendor == "":

            vendor_purchased_label.config(text="0")

            vendor_profit_label.config(text="0")

            vendor_invoice_count_label.config(text="0")

            return

        cursor.execute("""

        SELECT

        COUNT(*),
        IFNULL(SUM(paid_to_supplier),0),
        IFNULL(SUM(total_amount),0)

        FROM invoices

        WHERE supplier = ?

        """, (selected_vendor,))

        count, purchased, total = cursor.fetchone()

        profit = total - purchased

        vendor_purchased_label.config(text=f"{purchased:,.0f}")

        vendor_profit_label.config(text=f"{profit:,.0f}")

        vendor_invoice_count_label.config(text=f"{count}")

    vendor_select_combo.bind(
        "<<ComboboxSelected>>",
        update_vendor_analysis
    )

    # =====================================================
    # INVOICE PAGE
    # =====================================================

    Label(
        invoices_page,
        text="Invoice Registration",
        font=("Segoe UI", 24, "bold"),
        bg="#020617",
        fg="white"
    ).pack(pady=10)

    card = Frame(
        invoices_page,
        bg="#1e293b",
        padx=40,
        pady=20,
        highlightbackground="#00BFFF",
        highlightthickness=1
    )

    card.pack(
        padx=20,
        pady=10,
        fill=X
    )

    Label(
        card,
        text="Vendor Invoice Details",
        font=("Segoe UI", 16, "bold"),
        bg="#1e293b",
        fg="#00BFFF"
    ).grid(
        row=0,
        column=0,
        columnspan=4,
        pady=20
    )

    # =====================================================
    # STYLE
    # =====================================================

    style = ttk.Style()

    style.theme_use("clam")

    style.configure(
        "TCombobox",
        fieldbackground="#334155",
        background="#334155",
        foreground="white"
    )

    # =====================================================
    # HELPERS
    # =====================================================

    def label_widget(text, row, column):

        lbl = Label(
            card,
            text=text,
            font=("Segoe UI", 10, "bold"),
            bg="#1e293b",
            fg="white"
        )

        lbl.grid(
            row=row,
            column=column,
            padx=10,
            pady=6,
            sticky=W
        )

        lbl.bind(
            "<Enter>",
            lambda e: lbl.config(fg="#7DD3FC")
        )

        lbl.bind(
            "<Leave>",
            lambda e: lbl.config(fg="white")
        )

    def entry_widget(row, column):

        entry = Entry(
            card,
            width=40,
            font=("Segoe UI", 10),
            bg="#334155",
            fg="white",
            relief="flat",
            insertbackground="white"
        )

        entry.grid(
            row=row,
            column=column,
            padx=10,
            pady=6
        )

        return entry

    # =====================================================
    # FORM
    # =====================================================

    label_widget("Date", 1, 0)

    date_entry = DateEntry(
        card,
        width=20,
        font=("Segoe UI", 10),
        background="#00BFFF",
        date_pattern="dd/mm/yyyy"
    )

    date_entry.grid(row=1, column=1, sticky=W)

    label_widget("Invoice No", 1, 2)

    invoice_entry = entry_widget(1, 3)

    label_widget("Owner", 2, 0)

    owner_combo = ttk.Combobox(
        card,
        values=issuers,
        width=37,
        font=("Segoe UI", 10)
    )

    owner_combo.grid(row=2, column=1)

    searchable_combobox(
        owner_combo,
        issuers
    )

    label_widget("Accounts", 2, 2)

    accounts_combo = ttk.Combobox(
        card,
        values=accounts,
        width=37,
        font=("Segoe UI", 10)
    )

    accounts_combo.grid(row=2, column=3)

    searchable_combobox(
        accounts_combo,
        accounts
    )

    label_widget("Date of Service", 3, 0)

    service_date_entry = DateEntry(
        card,
        width=20,
        font=("Segoe UI", 10),
        background="#00BFFF",
        date_pattern="dd/mm/yyyy"
    )

    service_date_entry.grid(row=3, column=1, sticky=W)

    label_widget("Subject", 3, 2)

    subject_entry = entry_widget(3, 3)

    label_widget("Type of Service", 4, 0)

    service_type_values = [
        "Hotel",
        "Flight",
        "Transfer",
        "Visa",
        "Insurance",
        "Tour",
        "Restaurant",
        "Cruise",
        "Train",
        "Car Rental",
        "Guide",
        "Entrance Fees",
        "Meet & Assist",
        "Conference / MICE",
        "Other"
    ]

    service_type_combo = ttk.Combobox(
        card,
        values=service_type_values,
        width=37,
        font=("Segoe UI", 10)
    )

    service_type_combo.grid(row=4, column=1)

    searchable_combobox(
        service_type_combo,
        service_type_values
    )

    label_widget("Vendor", 4, 2)

    vendor_combo = ttk.Combobox(
        card,
        values=vendors,
        width=37,
        font=("Segoe UI", 10)
    )

    vendor_combo.grid(row=4, column=3)

    searchable_combobox(
        vendor_combo,
        vendors
    )

    label_widget("City", 5, 0)

    city_combo = ttk.Combobox(
        card,
        values=EGYPT_GOVERNORATES,
        width=37,
        font=("Segoe UI", 10)
    )

    city_combo.grid(row=5, column=1, sticky=W)

    searchable_combobox(
        city_combo,
        EGYPT_GOVERNORATES
    )

    label_widget("No. of PAX", 5, 2)

    pax_entry = entry_widget(5, 3)

    label_widget("Supplier Invoice No", 6, 0)

    supplier_invoice_entry = entry_widget(6, 1)

    label_widget("PO", 6, 2)

    po_entry = entry_widget(6, 3)

    label_widget("Paid To Supplier", 7, 0)

    paid_to_supplier_entry = entry_widget(7, 1)

    label_widget("Handling Fees", 7, 2)

    handling_fees_entry = entry_widget(7, 3)

    label_widget("VAT (14%)", 8, 0)

    vat_label = Label(
        card,
        text="0.00",
        font=("Segoe UI", 10, "bold"),
        bg="#1e293b",
        fg="#F59E0B"
    )

    vat_label.grid(row=8, column=1, sticky=W)

    VAT_RATE = 0.14

    label_widget("Currency", 8, 2)

    currency_combo = ttk.Combobox(
        card,
        values=["EGP", "USD"],
        width=10,
        font=("Segoe UI", 10),
        state="readonly"
    )

    currency_combo.grid(row=8, column=3, sticky=W)

    currency_combo.set("EGP")

    total_result = Label(
        card,
        text="Total Amount : 0",
        font=("Segoe UI", 12, "bold"),
        bg="#1e293b",
        fg="#00FF99"
    )

    total_result.grid(row=9, column=1, columnspan=3, pady=15, sticky=W)

    # =====================================================
    # AUTO VAT CALCULATION
    # =====================================================

    def update_vat(event=None):

        try:

            handling_fees = float(handling_fees_entry.get() or 0)

            vat = handling_fees * VAT_RATE

        except ValueError:

            vat = 0

        vat_label.config(text=f"{vat:.2f}")

    handling_fees_entry.bind("<KeyRelease>", update_vat)

    # =====================================================
    # LOAD VENDOR
    # =====================================================

    def load_vendor(event):

        selected = vendor_combo.get()

        vendor_data = vendors_df[
            vendors_df["Alias"] == selected
        ]

        if not vendor_data.empty:

            city_combo.set(
                vendor_data.iloc[0]["City"]
            )

        else:

            city_combo.set("")

    vendor_combo.bind(
        "<<ComboboxSelected>>",
        load_vendor
    )

    # =====================================================
    # CALCULATE
    # =====================================================

    def calculate():

        try:

            paid_to_supplier = float(paid_to_supplier_entry.get() or 0)

            handling_fees = float(handling_fees_entry.get() or 0)

            vat = handling_fees * VAT_RATE

            vat_label.config(text=f"{vat:.2f}")

            total = paid_to_supplier + handling_fees + vat

            total_result.config(
                text=f"Total Amount : {total:.2f} {currency_combo.get()}"
            )

        except:

            messagebox.showerror(
                "Error",
                "Invalid Amount"
            )

    # =====================================================
    # CLEAR
    # =====================================================

    def clear_form():

        invoice_entry.delete(0, END)

        owner_combo.set("")

        vendor_combo.set("")

        accounts_combo.set("")

        subject_entry.delete(0, END)

        service_type_combo.set("")

        pax_entry.delete(0, END)

        supplier_invoice_entry.delete(0, END)

        po_entry.delete(0, END)

        paid_to_supplier_entry.delete(0, END)

        handling_fees_entry.delete(0, END)

        vat_label.config(text="0.00")

        city_combo.set("")

        total_result.config(text="Total Amount : 0")

        currency_combo.set("EGP")

    # =====================================================
    # DASHBOARD UPDATE
    # =====================================================

    def update_dashboard():

        cursor.execute("""

        SELECT

        COUNT(*),
        IFNULL(SUM(total_amount),0),
        IFNULL(SUM(paid_to_supplier),0)

        FROM invoices

        """)

        count, sales, revenue = cursor.fetchone()

        invoice_count_label.config(text=f"{count}")

        total_sales_label.config(text=f"{sales:,.0f}")

        total_revenue_label.config(text=f"{revenue:,.0f}")

    # =====================================================
    # LOAD REPORTS
    # =====================================================

    def load_reports(search_text=""):

        for row in invoice_table.get_children():

            invoice_table.delete(row)

        if search_text == "":

            cursor.execute("""

            SELECT

            id,
            date,
            invoice_no,
            owner,
            accounts,
            service_date,
            subject,
            service_type,
            supplier,
            supplier_city,
            no_of_pax,
            supplier_invoice_no,
            po_number,
            total_amount,
            paid_to_supplier,
            handling_fees,
            vat,
            currency,
            payment_status

            FROM invoices

            ORDER BY id DESC

            """)

        else:

            cursor.execute("""

            SELECT

            id,
            date,
            invoice_no,
            owner,
            accounts,
            service_date,
            subject,
            service_type,
            supplier,
            supplier_city,
            no_of_pax,
            supplier_invoice_no,
            po_number,
            total_amount,
            paid_to_supplier,
            handling_fees,
            vat,
            currency,
            payment_status

            FROM invoices

            WHERE supplier LIKE ?
            OR accounts LIKE ?
            OR invoice_no LIKE ?

            ORDER BY id DESC

            """, (

                f"%{search_text}%",
                f"%{search_text}%",
                f"%{search_text}%"

            ))

        rows = cursor.fetchall()

        for row in rows:

            if row[18] == "Paid":

                invoice_table.insert(
                    "",
                    END,
                    values=row,
                    tags=("paid",)
                )

            else:

                invoice_table.insert(
                    "",
                    END,
                    values=row,
                    tags=("pending",)
                )

        invoice_table.tag_configure(
            "paid",
            background="#d1fae5"
        )

        invoice_table.tag_configure(
            "pending",
            background="#fef3c7"
        )

    # =====================================================
    # MARK AS PAID
    # =====================================================

    def mark_as_paid():

        selected = invoice_table.focus()

        if not selected:

            messagebox.showwarning(
                "Warning",
                "Please select invoice"
            )

            return

        values = invoice_table.item(
            selected,
            "values"
        )

        invoice_id = values[0]

        cursor.execute(

            "UPDATE invoices SET payment_status='Paid' WHERE id=?",

            (invoice_id,)

        )

        conn.commit()

        load_reports()

        messagebox.showinfo(
            "DONE",
            "Invoice Marked As Paid ✅"
        )

    # =====================================================
    # EXPORT
    # =====================================================

    def export_reports():

        cursor.execute("""

        SELECT *

        FROM invoices

        ORDER BY id DESC

        """)

        rows = cursor.fetchall()

        columns = [

            "ID",
            "Date",
            "Invoice No",
            "Owner",
            "Accounts",
            "Date Of Service",
            "Subject",
            "Type Of Service",
            "Vendor",
            "City",
            "No. of PAX",
            "Supplier Invoice No",
            "PO",
            "Total Amount",
            "Paid To Supplier",
            "Handling Fees",
            "VAT",
            "Currency",
            "Payment Status",
            "Created By"

        ]

        df = pd.DataFrame(
            rows,
            columns=columns
        )

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx")
            ]
        )

        if file_path:

            df.to_excel(
                file_path,
                index=False
            )

            messagebox.showinfo(
                "DONE",
                "Reports Exported Successfully ✅"
            )

    # =====================================================
    # GENERATE
    # =====================================================

    def generate_invoice():

        try:

            paid_to_supplier = float(paid_to_supplier_entry.get() or 0)

            handling_fees = float(handling_fees_entry.get() or 0)

            vat = handling_fees * VAT_RATE

            total = paid_to_supplier + handling_fees + vat

            pax_value = pax_entry.get()

            pax_value = int(pax_value) if pax_value.strip() != "" else 0

            cursor.execute("""

            INSERT INTO invoices (

                date,
                invoice_no,
                owner,
                accounts,
                service_date,
                subject,
                service_type,
                supplier,
                supplier_city,
                no_of_pax,
                supplier_invoice_no,
                po_number,
                total_amount,
                paid_to_supplier,
                handling_fees,
                vat,
                currency,
                payment_status,
                created_by

            )

            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

            """, (

                date_entry.get(),
                invoice_entry.get(),
                owner_combo.get(),
                accounts_combo.get(),
                service_date_entry.get(),
                subject_entry.get(),
                service_type_combo.get(),
                vendor_combo.get(),
                city_combo.get(),
                pax_value,
                supplier_invoice_entry.get(),
                po_entry.get(),
                total,
                paid_to_supplier,
                handling_fees,
                vat,
                currency_combo.get(),
                "Pending",
                current_user

            ))

            conn.commit()

            update_dashboard()

            load_reports()

            messagebox.showinfo(
                "DONE",
                "Invoice Generated Successfully ✅"
            )

            clear_form()

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # =====================================================
    # BUTTONS
    # =====================================================

    button_frame = Frame(
        invoices_page,
        bg="#020617"
    )

    button_frame.pack(pady=15)

    calculate_btn = Button(
        button_frame,
        text="CALCULATE",
        command=calculate,
        font=("Segoe UI", 10, "bold"),
        bg="#F59E0B",
        fg="white",
        width=18,
        height=2,
        relief="flat",
        cursor="hand2"
    )

    calculate_btn.grid(row=0, column=0, padx=10)

    add_hover_effect(calculate_btn, "#F59E0B", "#FBBF24")

    generate_btn = Button(
        button_frame,
        text="GENERATE",
        command=generate_invoice,
        font=("Segoe UI", 10, "bold"),
        bg="#10B981",
        fg="white",
        width=18,
        height=2,
        relief="flat",
        cursor="hand2"
    )

    generate_btn.grid(row=0, column=1, padx=10)

    add_hover_effect(generate_btn, "#10B981", "#34D399")

    clear_btn = Button(
        button_frame,
        text="CLEAR",
        command=clear_form,
        font=("Segoe UI", 10, "bold"),
        bg="#EF4444",
        fg="white",
        width=18,
        height=2,
        relief="flat",
        cursor="hand2"
    )

    clear_btn.grid(row=0, column=2, padx=10)

    add_hover_effect(clear_btn, "#EF4444", "#F87171")

    # =====================================================
    # REPORTS PAGE
    # =====================================================

    Label(
        reports_page,
        text="Reports",
        font=("Segoe UI", 24, "bold"),
        bg="#020617",
        fg="white"
    ).pack(pady=15)

    search_reports_frame = Frame(
        reports_page,
        bg="#020617"
    )

    search_reports_frame.pack(pady=10)

    report_search_entry = Entry(
        search_reports_frame,
        width=50,
        font=("Segoe UI", 10)
    )

    report_search_entry.pack(side=LEFT, padx=10)

    Button(
        search_reports_frame,
        text="SEARCH",
        command=lambda: load_reports(
            report_search_entry.get()
        ),
        font=("Segoe UI", 10, "bold"),
        bg="#00BFFF",
        fg="white",
        relief="flat"
    ).pack(side=LEFT)

    Button(
        search_reports_frame,
        text="MARK AS PAID",
        command=mark_as_paid,
        font=("Segoe UI", 10, "bold"),
        bg="#22c55e",
        fg="white",
        relief="flat"
    ).pack(side=LEFT, padx=10)

    Button(
        search_reports_frame,
        text="EXPORT EXCEL",
        command=export_reports,
        font=("Segoe UI", 10, "bold"),
        bg="#10B981",
        fg="white",
        relief="flat"
    ).pack(side=LEFT, padx=10)

    # =====================================================
    # REPORT TABLE
    # =====================================================

    reports_table_frame = Frame(
        reports_page,
        bg="#020617"
    )

    reports_table_frame.pack(
        fill=BOTH,
        expand=True,
        padx=20,
        pady=10
    )

    columns = (

        "ID",
        "Date",
        "Invoice No",
        "Owner",
        "Accounts",
        "Date Of Service",
        "Subject",
        "Type Of Service",
        "Vendor",
        "City",
        "No. of PAX",
        "Supplier Invoice No",
        "PO",
        "Total Amount",
        "Paid To Supplier",
        "Handling Fees",
        "VAT",
        "Currency",
        "Status"

    )

    invoice_table = ttk.Treeview(
        reports_table_frame,
        columns=columns,
        show="headings"
    )

    for col in columns:

        invoice_table.heading(col, text=col)

        invoice_table.column(
            col,
            width=120,
            anchor=CENTER
        )

    scrollbar = Scrollbar(
        reports_table_frame,
        orient=VERTICAL,
        command=invoice_table.yview
    )

    invoice_table.configure(
        yscroll=scrollbar.set
    )

    scrollbar.pack(side=RIGHT, fill=Y)

    invoice_table.pack(
        fill=BOTH,
        expand=True
    )

    # =====================================================
    # VENDORS PAGE
    # =====================================================

    Label(
        vendors_page,
        text="Vendors",
        font=("Segoe UI", 24, "bold"),
        bg="#020617",
        fg="white"
    ).pack(pady=20)

    vendor_table_frame = Frame(
        vendors_page,
        bg="#020617"
    )

    vendor_table_frame.pack(
        fill=BOTH,
        expand=True,
        padx=20,
        pady=10
    )

    vendors_table = ttk.Treeview(
        vendor_table_frame,
        show="headings"
    )

    vendor_columns = list(vendors_df.columns)

    vendors_table["columns"] = vendor_columns

    for col in vendor_columns:

        vendors_table.heading(col, text=col)

        vendors_table.column(
            col,
            width=170,
            anchor=CENTER
        )

    for _, row in vendors_df.iterrows():

        vendors_table.insert(
            "",
            END,
            values=list(row)
        )

    vendor_scroll = Scrollbar(
        vendor_table_frame,
        orient=VERTICAL,
        command=vendors_table.yview
    )

    vendors_table.configure(
        yscroll=vendor_scroll.set
    )

    vendor_scroll.pack(side=RIGHT, fill=Y)

    vendors_table.pack(
        fill=BOTH,
        expand=True
    )

    # =====================================================
    # ADMIN PANEL PAGE
    # =====================================================

    Label(
        admin_page,
        text="Admin Panel",
        font=("Segoe UI", 24, "bold"),
        bg="#020617",
        fg="white"
    ).pack(pady=20)

    admin_add_card = Frame(
        admin_page,
        bg="#1e293b",
        padx=30,
        pady=20,
        highlightbackground="#00BFFF",
        highlightthickness=1
    )

    admin_add_card.pack(padx=20, pady=10, fill=X)

    Label(
        admin_add_card,
        text="Add New User",
        font=("Segoe UI", 14, "bold"),
        bg="#1e293b",
        fg="#00BFFF"
    ).grid(row=0, column=0, columnspan=4, pady=(0, 15), sticky=W)

    Label(
        admin_add_card,
        text="Username",
        font=("Segoe UI", 10, "bold"),
        bg="#1e293b",
        fg="white"
    ).grid(row=1, column=0, padx=10, sticky=W)

    new_username_entry = Entry(
        admin_add_card,
        width=30,
        font=("Segoe UI", 10),
        bg="#334155",
        fg="white",
        relief="flat",
        insertbackground="white"
    )

    new_username_entry.grid(row=1, column=1, padx=10)

    Label(
        admin_add_card,
        text="Password",
        font=("Segoe UI", 10, "bold"),
        bg="#1e293b",
        fg="white"
    ).grid(row=1, column=2, padx=10, sticky=W)

    new_password_entry = Entry(
        admin_add_card,
        width=30,
        font=("Segoe UI", 10),
        bg="#334155",
        fg="white",
        relief="flat",
        insertbackground="white",
        show="*"
    )

    new_password_entry.grid(row=1, column=3, padx=10)

    def add_user():

        new_username = new_username_entry.get().strip()

        new_password = new_password_entry.get().strip()

        if new_username == "" or new_password == "":

            messagebox.showwarning(
                "Warning",
                "Please enter both username and password"
            )

            return

        try:

            cursor.execute(

                "INSERT INTO users (username, password) VALUES (?, ?)",

                (new_username, new_password)

            )

            conn.commit()

            new_username_entry.delete(0, END)

            new_password_entry.delete(0, END)

            load_users_table()

            messagebox.showinfo(
                "DONE",
                "User Added Successfully ✅"
            )

        except sqlite3.IntegrityError:

            messagebox.showerror(
                "Error",
                "Username already exists"
            )

    add_user_btn = Button(
        admin_add_card,
        text="ADD USER",
        command=add_user,
        font=("Segoe UI", 10, "bold"),
        bg="#10B981",
        fg="white",
        width=18,
        height=1,
        relief="flat",
        cursor="hand2"
    )

    add_user_btn.grid(row=1, column=4, padx=15)

    add_hover_effect(add_user_btn, "#10B981", "#34D399")

    admin_table_frame = Frame(
        admin_page,
        bg="#020617"
    )

    admin_table_frame.pack(
        fill=BOTH,
        expand=True,
        padx=20,
        pady=10
    )

    users_table = ttk.Treeview(
        admin_table_frame,
        columns=("ID", "Username"),
        show="headings"
    )

    users_table.heading("ID", text="ID")

    users_table.heading("Username", text="Username")

    users_table.column("ID", width=80, anchor=CENTER)

    users_table.column("Username", width=300, anchor=CENTER)

    users_scroll = Scrollbar(
        admin_table_frame,
        orient=VERTICAL,
        command=users_table.yview
    )

    users_table.configure(
        yscroll=users_scroll.set
    )

    users_scroll.pack(side=RIGHT, fill=Y)

    users_table.pack(
        fill=BOTH,
        expand=True
    )

    def load_users_table():

        for row in users_table.get_children():

            users_table.delete(row)

        cursor.execute("SELECT id, username FROM users ORDER BY id")

        for row in cursor.fetchall():

            users_table.insert("", END, values=row)

    def delete_user():

        selected = users_table.focus()

        if not selected:

            messagebox.showwarning(
                "Warning",
                "Please select a user"
            )

            return

        values = users_table.item(selected, "values")

        user_id = values[0]

        username = values[1]

        if username == "Bassma":

            messagebox.showerror(
                "Error",
                "Cannot delete the admin account"
            )

            return

        confirm = messagebox.askyesno(
            "Confirm",
            f"Delete user '{username}'?"
        )

        if confirm:

            cursor.execute(
                "DELETE FROM users WHERE id=?",
                (user_id,)
            )

            conn.commit()

            load_users_table()

            messagebox.showinfo(
                "DONE",
                "User Deleted Successfully ✅"
            )

    admin_button_frame = Frame(
        admin_page,
        bg="#020617"
    )

    admin_button_frame.pack(pady=15)

    delete_user_btn = Button(
        admin_button_frame,
        text="DELETE SELECTED USER",
        command=delete_user,
        font=("Segoe UI", 10, "bold"),
        bg="#EF4444",
        fg="white",
        width=24,
        height=2,
        relief="flat",
        cursor="hand2"
    )

    delete_user_btn.pack()

    add_hover_effect(delete_user_btn, "#EF4444", "#F87171")

    # =====================================================
    # START
    # =====================================================

    update_dashboard()

    show_dashboard()

# =========================================================
# LOGIN SCREEN
# =========================================================

login_frame = Frame(
    root,
    bg="#1e293b",
    padx=70,
    pady=60
)

login_frame.place(
    relx=0.5,
    rely=0.5,
    anchor=CENTER
)

# =========================================================
# LOGIN LOGO
# =========================================================

login_img = Image.open("logo.png")

login_img = login_img.resize((150, 150))

login_logo = ImageTk.PhotoImage(login_img)

login_logo_label = Label(
    login_frame,
    image=login_logo,
    bg="#1e293b"
)

login_logo_label.image = login_logo

login_logo_label.pack(pady=10)

# =========================================================
# TITLE
# =========================================================

Label(
    login_frame,
    text="LOGIN",
    font=("Segoe UI", 28, "bold"),
    bg="#1e293b",
    fg="#00BFFF"
).pack(pady=(15, 5))

Label(
    login_frame,
    text="Non-Air Operations ERP System",
    font=("Segoe UI", 10),
    bg="#1e293b",
    fg="#94a3b8"
).pack(pady=(0, 20))

# =========================================================
# USERNAME
# =========================================================

Label(
    login_frame,
    text="Username",
    font=("Segoe UI", 10, "bold"),
    bg="#1e293b",
    fg="white"
).pack()

user_entry = Entry(
    login_frame,
    width=32,
    font=("Segoe UI", 11),
    bg="#334155",
    fg="white",
    relief="flat",
    insertbackground="white"
)

user_entry.pack(pady=10, ipady=5)

# =========================================================
# PASSWORD
# =========================================================

Label(
    login_frame,
    text="Password",
    font=("Segoe UI", 10, "bold"),
    bg="#1e293b",
    fg="white"
).pack()

pass_entry = Entry(
    login_frame,
    show="*",
    width=32,
    font=("Segoe UI", 11),
    bg="#334155",
    fg="white",
    relief="flat",
    insertbackground="white"
)

pass_entry.pack(pady=10, ipady=5)

# =========================================================
# LOGIN BUTTON
# =========================================================

login_btn = Button(
    login_frame,
    text="LOGIN",
    command=login,
    font=("Segoe UI", 11, "bold"),
    bg="#0ea5e9",
    fg="white",
    activebackground="#0284c7",
    activeforeground="white",
    width=22,
    height=2,
    relief="flat",
    cursor="hand2"
)

login_btn.pack(pady=25)

# =========================================================
# FOOTER
# =========================================================

Label(
    login_frame,
    text="Powered By Brightstar",
    font=("Segoe UI", 8),
    bg="#1e293b",
    fg="#64748b"
).pack(pady=(15, 0))

# =========================================================
# RUN
# =========================================================
root.mainloop()
