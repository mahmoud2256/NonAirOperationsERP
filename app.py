import streamlit as st
import psycopg2
import pandas as pd
from datetime import date, datetime

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Non-Air Operations ERP",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# DATABASE CONNECTION
# =========================================================
# ⚠️ حط الباسوورد بتاعك هنا بدل YOUR_PASSWORD_HERE

DB_URL = "postgresql://postgres.iqbdoznbpsefaqqohqvz:YOUR_PASSWORD_HERE@aws-0-eu-west-1.pooler.supabase.com:6543/postgres"

@st.cache_resource
def get_connection():
    return psycopg2.connect(DB_URL)

def get_cursor():
    conn = get_connection()
    conn.autocommit = True
    return conn.cursor()

# =========================================================
# CREATE TABLES
# =========================================================

def create_tables():
    cur = get_cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username TEXT UNIQUE,
        password TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS invoices (
        id SERIAL PRIMARY KEY,
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

    cur.execute("""
    INSERT INTO users (username, password)
    VALUES ('Bassma', '1234')
    ON CONFLICT (username) DO NOTHING
    """)

# =========================================================
# STATIC DATA
# =========================================================

EGYPT_GOVERNORATES = [
    "Cairo", "Giza", "Alexandria", "Qalyubia", "Port Said",
    "Suez", "Damietta", "Dakahlia", "Sharqia", "Gharbia",
    "Monufia", "Beheira", "Ismailia", "Kafr El Sheikh",
    "North Sinai", "South Sinai", "Beni Suef", "Faiyum",
    "Minya", "Asyut", "Sohag", "Qena", "Luxor", "Aswan",
    "Red Sea", "New Valley", "Matrouh"
]

SERVICE_TYPES = [
    "Hotel", "Flight", "Transfer", "Visa", "Insurance",
    "Tour", "Restaurant", "Cruise", "Train", "Car Rental",
    "Guide", "Entrance Fees", "Meet & Assist",
    "Conference / MICE", "Other"
]

VAT_RATE = 0.14

def get_handling_rate(account_name):
    name = account_name.strip().upper()
    if "ASTRA" in name:
        return 0.049
    elif "ACINO" in name:
        return 0.05
    return None

# =========================================================
# LOAD EXCEL FILES
# =========================================================

@st.cache_data
def load_vendors():
    try:
        df = pd.read_excel("vendors.xlsx").fillna("")
        if "City" not in df.columns:
            df["City"] = ""
        return df
    except:
        return pd.DataFrame(columns=["Alias", "City"])

@st.cache_data
def load_issuers():
    try:
        df = pd.read_excel("issuers.xlsx").fillna("")
        return df["Issuer"].tolist()
    except:
        return []

@st.cache_data
def load_accounts():
    try:
        df = pd.read_excel("accounts.xlsx").fillna("")
        return df["Accounts"].tolist()
    except:
        return []

# =========================================================
# CSS STYLING
# =========================================================

st.markdown("""
<style>
    /* ===== GLOBAL ===== */
    .stApp {
        background-color: #F5F6FA;
        color: #1A1A2E;
        font-family: 'Segoe UI', Arial, sans-serif;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    /* ===== SIDEBAR ===== */
    section[data-testid="stSidebar"] {
        background-color: #1A3A5C;
        border-right: none;
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    section[data-testid="stSidebar"] .stButton > button {
        background-color: transparent;
        color: white !important;
        border: 1px solid rgba(255,255,255,0.2);
        border-radius: 4px;
        font-weight: 500;
        text-align: left;
        padding: 8px 16px;
        transition: all 0.2s;
    }

    section[data-testid="stSidebar"] .stButton > button:hover {
        background-color: rgba(255,255,255,0.15);
        border-color: rgba(255,255,255,0.4);
    }

    /* ===== METRIC CARDS ===== */
    .metric-card {
        background: white;
        border: 1px solid #E0E4EA;
        border-radius: 6px;
        padding: 20px 24px;
        margin: 8px 0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }

    .metric-number {
        font-size: 32px;
        font-weight: 700;
        margin: 0;
        line-height: 1.2;
    }

    .metric-label {
        font-size: 11px;
        color: #6B7280;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        margin: 8px 0 0 0;
        font-weight: 600;
    }

    .section-label {
        font-size: 11px;
        color: #6B7280;
        font-weight: 700;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin: 28px 0 10px 0;
        padding-bottom: 6px;
        border-bottom: 2px solid #E0E4EA;
    }

    /* ===== FORM ===== */
    div[data-testid="stForm"] {
        background: white;
        border: 1px solid #E0E4EA;
        border-radius: 6px;
        padding: 24px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }

    /* ===== INPUTS ===== */
    .stTextInput > div > div > input,
    .stSelectbox > div > div,
    .stNumberInput > div > div > input {
        background-color: white;
        border: 1px solid #D1D5DB;
        border-radius: 4px;
        color: #1A1A2E;
        font-size: 14px;
    }

    /* ===== BUTTONS ===== */
    .stButton > button {
        border-radius: 4px;
        font-weight: 600;
        font-size: 13px;
        border: 1px solid #1A3A5C;
        background-color: #1A3A5C;
        color: white;
        padding: 8px 20px;
        transition: all 0.2s;
    }

    .stButton > button:hover {
        background-color: #14304F;
        border-color: #14304F;
    }

    /* ===== DATAFRAME ===== */
    .stDataFrame {
        border: 1px solid #E0E4EA;
        border-radius: 6px;
    }

    /* ===== PAGE TITLE ===== */
    h1 {
        color: #1A3A5C !important;
        font-size: 26px !important;
        font-weight: 700 !important;
        border-bottom: 3px solid #1A3A5C;
        padding-bottom: 10px;
        margin-bottom: 4px !important;
    }

    h2, h3 {
        color: #1A3A5C !important;
    }

    /* ===== ALERTS ===== */
    .stSuccess {
        background-color: #ECFDF5;
        border-left: 4px solid #10B981;
        color: #065F46;
    }

    .stError {
        background-color: #FEF2F2;
        border-left: 4px solid #EF4444;
        color: #991B1B;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# SESSION STATE
# =========================================================

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.current_user = ""
    st.session_state.page = "Dashboard"

# =========================================================
# LOGIN PAGE
# =========================================================

def show_login():
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.image("logo.png", width=250
                )
        st.markdown("## Non-Air Operations ERP")
        st.markdown("##### Bright Star")
        st.markdown("<br>", unsafe_allow_html=True)

        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Login", use_container_width=True)

            if submit:
                cur = get_cursor()
                cur.execute(
                    "SELECT * FROM users WHERE username=%s AND password=%s",
                    (username, password)
                )
                user = cur.fetchone()
                if user:
                    st.session_state.logged_in = True
                    st.session_state.current_user = username
                    st.rerun()
                else:
                    st.error("Invalid username or password")

# =========================================================
# DASHBOARD PAGE
# =========================================================

def show_dashboard():
    st.markdown("# Operations Dashboard")
    st.markdown("**Bright Star  •  Non-Air Operations**")

    cur = get_cursor()
    cur.execute("""
    SELECT COUNT(*), COALESCE(SUM(total_amount),0), COALESCE(SUM(paid_to_supplier),0)
    FROM invoices
    """)
    count, total_sales, total_revenue = cur.fetchone()

    st.markdown('<p class="section-label">Key Performance Indicators</p>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div class="metric-card" style="border-top: 3px solid #22C55E;">
            <p class="metric-number" style="color:#1A7A4A;">{total_sales:,.0f}</p>
            <p class="metric-label">Total Sales</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card" style="border-top: 3px solid #3B82F6;">
            <p class="metric-number" style="color:#1A3A5C;">{total_revenue:,.0f}</p>
            <p class="metric-label">Total Revenue (Paid to Supplier)</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card" style="border-top: 3px solid #F59E0B;">
            <p class="metric-number" style="color:#B45309;">{count}</p>
            <p class="metric-label">Invoices Count</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<p class="section-label">Vendor Analysis</p>', unsafe_allow_html=True)

    vendors_df = load_vendors()
    vendors = vendors_df["Alias"].tolist()

    selected_vendor = st.selectbox("Select Vendor", [""] + vendors)

    if selected_vendor:
        cur.execute("""
        SELECT COUNT(*), COALESCE(SUM(paid_to_supplier),0), COALESCE(SUM(total_amount),0)
        FROM invoices WHERE supplier = %s
        """, (selected_vendor,))
        v_count, v_purchased, v_total = cur.fetchone()
        v_profit = v_total - v_purchased

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
            <div class="metric-card" style="border-top: 3px solid #3B82F6;">
                <p class="metric-number" style="color:#1A3A5C;">{v_purchased:,.0f}</p>
                <p class="metric-label">Paid To Supplier</p>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="metric-card" style="border-top: 3px solid #22C55E;">
                <p class="metric-number" style="color:#1A7A4A;">{v_profit:,.0f}</p>
                <p class="metric-label">Profit (Handling + VAT)</p>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="metric-card" style="border-top: 3px solid #F59E0B;">
                <p class="metric-number" style="color:#B45309;">{v_count}</p>
                <p class="metric-label">Invoices With Vendor</p>
            </div>
            """, unsafe_allow_html=True)

# =========================================================
# INVOICES PAGE
# =========================================================

def show_invoices():
    st.markdown("# Invoice Registration")
    st.markdown("##### Vendor Invoice Details")

    vendors_df = load_vendors()
    vendors = vendors_df["Alias"].tolist()
    issuers = load_issuers()
    accounts = load_accounts()

    # قراءة الحقول خارج الفورم لتفعيل الحساب الفوري
    col1, col2 = st.columns(2)

    with col1:
        inv_date = st.date_input("Date", value=date.today())
        owner = st.selectbox("Owner", [""] + issuers)
        service_date = st.date_input("Date of Service", value=date.today())
        service_type = st.selectbox("Type of Service", [""] + SERVICE_TYPES)
        city = st.selectbox("City", [""] + EGYPT_GOVERNORATES)
        supplier_inv_no = st.text_input("Supplier Invoice No")
        paid_to_supplier = st.number_input("Paid To Supplier", min_value=0.0, step=0.01, key="paid")

    with col2:
        invoice_no = st.text_input("Invoice No")
        accounts_val = st.selectbox("Accounts", [""] + accounts)
        subject = st.text_input("Subject")
        vendor = st.selectbox("Vendor", [""] + vendors)

        vendor_city = ""
        if vendor:
            v_data = vendors_df[vendors_df["Alias"] == vendor]
            if not v_data.empty:
                vendor_city = v_data.iloc[0]["City"]

        pax = st.number_input("No. of PAX", min_value=0, step=1)
        po = st.text_input("PO")

        rate = get_handling_rate(accounts_val) if accounts_val else None
        if rate:
            handling_fees = paid_to_supplier * rate
            st.markdown(f"""
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:6px; padding:12px; margin:8px 0;">
                <span style="color:#1A3A5C; font-size:12px; font-weight:600;">HANDLING FEES ({rate*100:.1f}%)</span><br>
                <span style="color:#1A3A5C; font-size:22px; font-weight:700;">{handling_fees:,.2f}</span>
            </div>
            """, unsafe_allow_html=True)
        else:
            handling_fees = st.number_input("Handling Fees", min_value=0.0, step=0.01, key="handling")

        currency = st.selectbox("Currency", ["EGP", "USD"])

    vat = handling_fees * VAT_RATE
    total = paid_to_supplier + handling_fees + vat

    st.markdown(f"""
    <div style="background:#F0FDF4; border:1px solid #86EFAC; border-radius:6px; padding:16px; margin:16px 0; display:flex; gap:40px;">
        <div>
            <span style="color:#6B7280; font-size:12px; font-weight:600;">VAT (14%)</span><br>
            <span style="color:#B45309; font-size:20px; font-weight:700;">{vat:,.2f}</span>
        </div>
        <div>
            <span style="color:#6B7280; font-size:12px; font-weight:600;">TOTAL AMOUNT</span><br>
            <span style="color:#166534; font-size:24px; font-weight:700;">{total:,.2f} {currency}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("✅ GENERATE INVOICE", use_container_width=True, type="primary"):
        if not invoice_no or not vendor or not accounts_val:
            st.error("Please fill Invoice No, Vendor, and Accounts")
        else:
            cur = get_cursor()
            cur.execute("""
            INSERT INTO invoices (
                date, invoice_no, owner, accounts, service_date,
                subject, service_type, supplier, supplier_city,
                no_of_pax, supplier_invoice_no, po_number,
                total_amount, paid_to_supplier, handling_fees,
                vat, currency, payment_status, created_by
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                str(inv_date), invoice_no, owner, accounts_val,
                str(service_date), subject, service_type, vendor,
                vendor_city or city, pax, supplier_inv_no, po,
                total, paid_to_supplier, handling_fees, vat,
                currency, "Pending", st.session_state.current_user
            ))
            st.success("✅ Invoice Generated Successfully!")
            st.balloons()

# =========================================================
# REPORTS PAGE
# =========================================================

def show_reports():
    st.markdown("# Reports")

    search = st.text_input("🔍 Search by Vendor, Account, or Invoice No")

    cur = get_cursor()

    if search:
        cur.execute("""
        SELECT id, date, invoice_no, owner, accounts, service_date,
               subject, service_type, supplier, supplier_city, no_of_pax,
               supplier_invoice_no, po_number, total_amount, paid_to_supplier,
               handling_fees, vat, currency, payment_status
        FROM invoices
        WHERE supplier ILIKE %s OR accounts ILIKE %s OR invoice_no ILIKE %s
        ORDER BY id DESC
        """, (f"%{search}%", f"%{search}%", f"%{search}%"))
    else:
        cur.execute("""
        SELECT id, date, invoice_no, owner, accounts, service_date,
               subject, service_type, supplier, supplier_city, no_of_pax,
               supplier_invoice_no, po_number, total_amount, paid_to_supplier,
               handling_fees, vat, currency, payment_status
        FROM invoices ORDER BY id DESC
        """)

    rows = cur.fetchall()
    columns = [
        "ID", "Date", "Invoice No", "Owner", "Accounts", "Date of Service",
        "Subject", "Type of Service", "Vendor", "City", "PAX",
        "Supplier Inv No", "PO", "Total Amount", "Paid to Supplier",
        "Handling Fees", "VAT", "Currency", "Status"
    ]

    if rows:
        df = pd.DataFrame(rows, columns=columns)
        st.dataframe(df, use_container_width=True, height=500)

        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Export to Excel",
            csv,
            "invoices_report.csv",
            "text/csv"
        )
    else:
        st.info("No invoices found")

# =========================================================
# ADMIN PANEL PAGE
# =========================================================

def show_admin():
    st.markdown("# Admin Panel")

    cur = get_cursor()

    st.markdown("### Add New User")
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        new_username = st.text_input("Username")
    with col2:
        new_password = st.text_input("Password", type="password")
    with col3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("ADD USER", use_container_width=True):
            if new_username and new_password:
                try:
                    cur.execute(
                        "INSERT INTO users (username, password) VALUES (%s, %s)",
                        (new_username, new_password)
                    )
                    st.success(f"✅ User '{new_username}' added!")
                except:
                    st.error("Username already exists")
            else:
                st.warning("Please fill both fields")

    st.markdown("### Current Users")
    cur.execute("SELECT id, username FROM users ORDER BY id")
    users = cur.fetchall()

    if users:
        df = pd.DataFrame(users, columns=["ID", "Username"])
        st.dataframe(df, use_container_width=True)

        del_username = st.selectbox(
            "Select user to delete",
            [u[1] for u in users if u[1] != "Bassma"]
        )
        if st.button("🗑️ DELETE SELECTED USER", type="primary"):
            cur.execute("DELETE FROM users WHERE username=%s", (del_username,))
            st.success(f"✅ User '{del_username}' deleted!")
            st.rerun()

# =========================================================
# MAIN APP
# =========================================================

try:
    create_tables()
except Exception as e:
    st.error(f"Database connection error: {e}")
    st.info("Please check your database password in the code")
    st.stop()

if not st.session_state.logged_in:
    show_login()
else:
    with st.sidebar:
        try:
            st.image("logo.png", width=200)
        except:
            pass

        st.markdown(f"### Non-Air Operations ERP")
        st.markdown(f"**Welcome, {st.session_state.current_user}**")
        st.markdown("---")

        pages = ["Dashboard", "Invoices", "Reports", "Vendors"]
        if st.session_state.current_user == "Bassma":
            pages.append("Admin Panel")

        for page in pages:
            if st.button(page, use_container_width=True, key=f"nav_{page}"):
                st.session_state.page = page

        st.markdown("---")
        if st.button("Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.current_user = ""
            st.rerun()

    if st.session_state.page == "Dashboard":
        show_dashboard()
    elif st.session_state.page == "Invoices":
        show_invoices()
    elif st.session_state.page == "Reports":
        show_reports()
    elif st.session_state.page == "Vendors":
        vendors_df = load_vendors()
        st.markdown("# Vendors")
        st.dataframe(vendors_df, use_container_width=True, height=500)
    elif st.session_state.page == "Admin Panel":
        show_admin()
