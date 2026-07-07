import streamlit as st
import psycopg2
import pandas as pd
from datetime import date, datetime
import io
import urllib.parse

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    REPORTLAB_AVAILABLE = True
except:
    REPORTLAB_AVAILABLE = False

SUPPORT_WHATSAPP = "201009538924"

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

DB_URL = "postgresql://postgres.iqbdoznbpsefaqqohqvz:Mmooddyy87A@aws-0-eu-west-1.pooler.supabase.com:6543/postgres"

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

    /* ===== HIDE GITHUB ICON ===== */
    .stToolbarActions [data-testid="stBaseButton-github"],
    [data-testid="stToolbar"] a[href*="github"],
    .stAppToolbar a[href*="github"] {
        display: none !important;
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

    st.markdown("""
    <style>
        .stApp { background: linear-gradient(135deg, #0f172a 0%, #1A3A5C 100%) !important; }
        .main .block-container { padding-top: 0 !important; }

        .login-container {
            max-width: 420px;
            margin: 60px auto;
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 12px;
            padding: 40px;
            backdrop-filter: blur(10px);
        }

        .login-title {
            color: white;
            font-size: 24px;
            font-weight: 700;
            text-align: center;
            margin: 16px 0 4px 0;
        }

        .login-subtitle {
            color: #93C5FD;
            font-size: 13px;
            text-align: center;
            margin-bottom: 28px;
        }

        .stTextInput label { color: #CBD5E1 !important; font-size: 13px !important; }
        .stTextInput input {
            background: rgba(255,255,255,0.08) !important;
            border: 1px solid rgba(255,255,255,0.2) !important;
            color: white !important;
            border-radius: 6px !important;
        }
        .stTextInput input:focus {
            border-color: #3B82F6 !important;
            box-shadow: 0 0 0 2px rgba(59,130,246,0.3) !important;
        }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)

        try:
            st.image("logo.png", use_container_width=True)
        except:
            pass

        st.markdown("""
        <div class="login-title">Non-Air Operations ERP</div>
        <div class="login-subtitle">Bright Star • Non-Air Operations</div>
        """, unsafe_allow_html=True)

        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("🔐  Sign In", use_container_width=True)

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
                    st.error("❌ Invalid username or password")

# =========================================================
# DASHBOARD PAGE
# =========================================================

def generate_invoice_pdf(data):

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm
    )

    styles = getSampleStyleSheet()
    elements = []

    # Title
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Normal'],
        fontSize=18,
        textColor=colors.HexColor('#1A3A5C'),
        spaceAfter=6,
        fontName='Helvetica-Bold'
    )

    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#6B7280'),
        spaceAfter=16,
        fontName='Helvetica'
    )

    elements.append(Paragraph("Non-Air Operations ERP", title_style))
    elements.append(Paragraph("Bright Star • Vendor Invoice", subtitle_style))
    elements.append(Spacer(1, 0.3*cm))

    # Invoice details table
    table_data = [
        ["FIELD", "VALUE"],
        ["Invoice No", data.get("invoice_no", "")],
        ["Date", data.get("date", "")],
        ["Owner", data.get("owner", "")],
        ["Accounts", data.get("accounts", "")],
        ["Date of Service", data.get("service_date", "")],
        ["Subject", data.get("subject", "")],
        ["Type of Service", data.get("service_type", "")],
        ["Vendor", data.get("supplier", "")],
        ["City", data.get("supplier_city", "")],
        ["No. of PAX", str(data.get("no_of_pax", ""))],
        ["Supplier Invoice No", data.get("supplier_invoice_no", "")],
        ["PO", data.get("po_number", "")],
        ["Currency", data.get("currency", "EGP")],
    ]

    table = Table(table_data, colWidths=[6*cm, 11*cm])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A3A5C')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F8FAFC'), colors.white]),
        ('TEXTCOLOR', (0, 1), (0, -1), colors.HexColor('#374151')),
        ('TEXTCOLOR', (1, 1), (1, -1), colors.HexColor('#1A1A2E')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E4EA')),
        ('ROWHEIGHT', (0, 0), (-1, -1), 20),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 0.5*cm))

    # Financial summary table
    fin_data = [
        ["FINANCIAL SUMMARY", ""],
        ["Paid To Supplier", f"{data.get('paid_to_supplier', 0):,.2f} {data.get('currency', 'EGP')}"],
        ["Handling Fees", f"{data.get('handling_fees', 0):,.2f} {data.get('currency', 'EGP')}"],
        ["VAT (14%)", f"{data.get('vat', 0):,.2f} {data.get('currency', 'EGP')}"],
        ["TOTAL AMOUNT", f"{data.get('total_amount', 0):,.2f} {data.get('currency', 'EGP')}"],
    ]

    fin_table = Table(fin_data, colWidths=[6*cm, 11*cm])
    fin_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A3A5C')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('SPAN', (0, 0), (1, 0)),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 4), (-1, 4), colors.HexColor('#EFF6FF')),
        ('FONTNAME', (0, 4), (-1, 4), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 4), (-1, 4), 11),
        ('TEXTCOLOR', (0, 4), (-1, 4), colors.HexColor('#1A3A5C')),
        ('FONTNAME', (0, 1), (0, 3), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 3), 9),
        ('ROWBACKGROUNDS', (0, 1), (-1, 3), [colors.HexColor('#F8FAFC'), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E4EA')),
        ('ROWHEIGHT', (0, 0), (-1, -1), 22),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
    ]))

    elements.append(fin_table)
    elements.append(Spacer(1, 0.5*cm))

    # Footer
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor('#9CA3AF'),
        fontName='Helvetica'
    )
    elements.append(Paragraph(
        f"Generated by Non-Air Operations ERP  •  {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        footer_style
    ))

    doc.build(elements)
    buffer.seek(0)
    return buffer
    st.markdown("# Operations Dashboard")
    st.markdown("**Kanoo Travel  •  Non-Air Operations**")

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

    # ── CUSTOMER ANALYSIS ──────────────────────────────────
    st.markdown('<p class="section-label">Customer Analysis</p>', unsafe_allow_html=True)

    accounts = load_accounts()

    selected_account = st.selectbox("Select Customer (Account)", [""] + accounts)

    if selected_account:
        cur.execute("""
        SELECT COUNT(*),
               COALESCE(SUM(total_amount),0),
               COALESCE(SUM(paid_to_supplier),0),
               COALESCE(SUM(handling_fees),0),
               COALESCE(SUM(vat),0)
        FROM invoices WHERE accounts = %s
        """, (selected_account,))
        c_count, c_total, c_paid, c_handling, c_vat = cur.fetchone()
        c_profit = c_handling + c_vat

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"""
            <div class="metric-card" style="border-top: 3px solid #1A3A5C;">
                <p class="metric-number" style="color:#1A3A5C;">{c_count}</p>
                <p class="metric-label">Invoices Count</p>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="metric-card" style="border-top: 3px solid #22C55E;">
                <p class="metric-number" style="color:#1A7A4A;">{c_total:,.0f}</p>
                <p class="metric-label">Total Amount</p>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="metric-card" style="border-top: 3px solid #3B82F6;">
                <p class="metric-number" style="color:#1A3A5C;">{c_paid:,.0f}</p>
                <p class="metric-label">Paid To Supplier</p>
            </div>
            """, unsafe_allow_html=True)
        with col4:
            st.markdown(f"""
            <div class="metric-card" style="border-top: 3px solid #F59E0B;">
                <p class="metric-number" style="color:#B45309;">{c_profit:,.0f}</p>
                <p class="metric-label">Profit (Handling + VAT)</p>
            </div>
            """, unsafe_allow_html=True)

# =========================================================
# INVOICES PAGE
# =========================================================

def show_invoices():

    vendors_df = load_vendors()
    vendors = vendors_df["Alias"].tolist()
    issuers = load_issuers()
    accounts = load_accounts()

    st.markdown("""
    <div style="background:#1A3A5C; padding:16px 24px; border-radius:6px; margin-bottom:24px;">
        <span style="color:white; font-size:20px; font-weight:700; letter-spacing:0.5px;">
            📄 Invoice Registration
        </span>
        <span style="color:#93C5FD; font-size:13px; margin-left:16px;">
            Vendor Invoice Details
        </span>
    </div>
    """, unsafe_allow_html=True)

    # ── ROW 1 ──────────────────────────────────────────────
    st.markdown('<p style="font-size:11px;font-weight:700;color:#6B7280;letter-spacing:1.5px;text-transform:uppercase;border-bottom:1px solid #E0E4EA;padding-bottom:6px;margin-bottom:12px;">General Information</p>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        inv_date = st.date_input("Date", value=date.today())
    with c2:
        invoice_no = st.text_input("Invoice No")
    with c3:
        owner = st.selectbox("Owner", [""] + issuers)
    with c4:
        accounts_val = st.selectbox("Accounts", [""] + accounts)

    # ── ROW 2 ──────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        service_date = st.date_input("Date of Service", value=date.today())
    with c2:
        subject = st.text_input("Subject")
    with c3:
        service_type = st.selectbox("Type of Service", [""] + SERVICE_TYPES)
    with c4:
        vendor = st.selectbox("Vendor", [""] + vendors)

    # ── ROW 3 ──────────────────────────────────────────────
    st.markdown('<p style="font-size:11px;font-weight:700;color:#6B7280;letter-spacing:1.5px;text-transform:uppercase;border-bottom:1px solid #E0E4EA;padding-bottom:6px;margin:16px 0 12px 0;">Service Details</p>', unsafe_allow_html=True)

    vendor_city = ""
    if vendor:
        v_data = vendors_df[vendors_df["Alias"] == vendor]
        if not v_data.empty:
            vendor_city = v_data.iloc[0]["City"]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        city = st.selectbox("City", [""] + EGYPT_GOVERNORATES,
                           index=(EGYPT_GOVERNORATES.index(vendor_city) + 1)
                           if vendor_city in EGYPT_GOVERNORATES else 0)
    with c2:
        pax = st.number_input("No. of PAX", min_value=0, step=1)
    with c3:
        supplier_inv_no = st.text_input("Supplier Invoice No")
    with c4:
        po = st.text_input("PO")

    # ── ROW 4: AMOUNTS ─────────────────────────────────────
    st.markdown('<p style="font-size:11px;font-weight:700;color:#6B7280;letter-spacing:1.5px;text-transform:uppercase;border-bottom:1px solid #E0E4EA;padding-bottom:6px;margin:16px 0 12px 0;">Financial Details</p>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        paid_to_supplier = st.number_input("Paid To Supplier", min_value=0.0, step=0.01, key="paid")
    with c2:
        rate = get_handling_rate(accounts_val) if accounts_val else None
        if rate:
            handling_fees = paid_to_supplier * rate
            st.markdown(f"""
            <div style="margin-top:4px;">
                <label style="font-size:14px;color:#374151;font-weight:500;">Handling Fees ({rate*100:.1f}%)</label>
                <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:6px;padding:10px 14px;margin-top:8px;">
                    <span style="font-size:18px;font-weight:700;color:#1A3A5C;">{handling_fees:,.2f}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            handling_fees = st.number_input("Handling Fees", min_value=0.0, step=0.01, key="handling")
    with c3:
        currency = st.selectbox("Currency", ["EGP", "USD"])
    with c4:
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    # ── TOTALS BAR ─────────────────────────────────────────
    vat = handling_fees * VAT_RATE
    total = paid_to_supplier + handling_fees + vat

    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #1A3A5C 0%, #1e4976 100%);
        border-radius: 8px;
        padding: 20px 28px;
        margin: 20px 0;
        display: flex;
        gap: 60px;
        align-items: center;
    ">
        <div>
            <div style="color:#93C5FD; font-size:11px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; margin-bottom:4px;">VAT (14%)</div>
            <div style="color:white; font-size:22px; font-weight:700; font-family:'Segoe UI',Arial,sans-serif;">{vat:,.2f}</div>
        </div>
        <div style="width:1px; background:rgba(255,255,255,0.2); height:40px;"></div>
        <div>
            <div style="color:#93C5FD; font-size:11px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; margin-bottom:4px;">TOTAL AMOUNT</div>
            <div style="color:#4ADE80; font-size:28px; font-weight:800; font-family:'Segoe UI',Arial,sans-serif;">{total:,.2f} <span style="font-size:16px; color:#86EFAC;">{currency}</span></div>
        </div>
        <div style="width:1px; background:rgba(255,255,255,0.2); height:40px;"></div>
        <div>
            <div style="color:#93C5FD; font-size:11px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; margin-bottom:4px;">PAID TO SUPPLIER</div>
            <div style="color:white; font-size:22px; font-weight:700; font-family:'Segoe UI',Arial,sans-serif;">{paid_to_supplier:,.2f}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── GENERATE BUTTON ────────────────────────────────────
    col_btn, col_empty = st.columns([1, 3])
    with col_btn:
        generate = st.button("✅  GENERATE INVOICE", use_container_width=True, type="primary")

    if generate:
        if not invoice_no or not vendor or not accounts_val:
            st.error("⚠️ Please fill Invoice No, Vendor, and Accounts")
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

            invoice_data = {
                "invoice_no": invoice_no,
                "date": str(inv_date),
                "owner": owner,
                "accounts": accounts_val,
                "service_date": str(service_date),
                "subject": subject,
                "service_type": service_type,
                "supplier": vendor,
                "supplier_city": vendor_city or city,
                "no_of_pax": pax,
                "supplier_invoice_no": supplier_inv_no,
                "po_number": po,
                "total_amount": total,
                "paid_to_supplier": paid_to_supplier,
                "handling_fees": handling_fees,
                "vat": vat,
                "currency": currency,
            }

            col_pdf, col_wa, col_empty = st.columns([1, 1, 2])

            with col_pdf:
                if REPORTLAB_AVAILABLE:
                    pdf_buffer = generate_invoice_pdf(invoice_data)
                    st.download_button(
                        label="📄 Download PDF",
                        data=pdf_buffer,
                        file_name=f"Invoice_{invoice_no}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                else:
                    st.warning("PDF not available")

            with col_wa:
                wa_message = (
                    f"✅ New Invoice Generated\n\n"
                    f"Invoice No: {invoice_no}\n"
                    f"Date: {inv_date}\n"
                    f"Vendor: {vendor}\n"
                    f"Account: {accounts_val}\n"
                    f"Total: {total:,.2f} {currency}\n"
                    f"VAT: {vat:,.2f}\n"
                    f"Handling: {handling_fees:,.2f}\n\n"
                    f"Generated by Non-Air Operations ERP"
                )
                wa_url = f"https://wa.me/{SUPPORT_WHATSAPP}?text={urllib.parse.quote(wa_message)}"
                st.markdown(
                    f'<a href="{wa_url}" target="_blank">'
                    f'<button style="width:100%;background:#25D366;color:white;border:none;'
                    f'padding:8px;border-radius:4px;font-weight:600;font-size:14px;cursor:pointer;">'
                    f'📱 Send WhatsApp</button></a>',
                    unsafe_allow_html=True
                )

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

        # Logo
        try:
            st.image("logo.png", use_container_width=True)
        except:
            pass

        st.markdown(f"""
        <div style="
            text-align: center;
            padding: 12px 0 8px 0;
            border-bottom: 1px solid rgba(255,255,255,0.15);
            margin-bottom: 16px;
        ">
            <div style="
                color: white;
                font-size: 15px;
                font-weight: 700;
                letter-spacing: 0.5px;
                line-height: 1.4;
            ">Non-Air Operations ERP</div>
            <div style="
                color: #93C5FD;
                font-size: 12px;
                margin-top: 4px;
            ">Welcome, {st.session_state.current_user}</div>
        </div>
        """, unsafe_allow_html=True)

        pages = ["Dashboard", "Invoices", "Reports", "Vendors"]
        if st.session_state.current_user == "Bassma":
            pages.append("Admin Panel")
        pages.append("Support")

        for page in pages:
            if st.button(page, use_container_width=True, key=f"nav_{page}"):
                st.session_state.page = page

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

        st.markdown("""
        <div style="
            border-top: 1px solid rgba(255,255,255,0.15);
            padding-top: 12px;
        "></div>
        """, unsafe_allow_html=True)

        if st.button("🚪 Logout", use_container_width=True):
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
    elif st.session_state.page == "Support":
        st.markdown("""
        <div style="background:#1A3A5C; padding:16px 24px; border-radius:6px; margin-bottom:24px;">
            <span style="color:white; font-size:20px; font-weight:700;">🎧 Technical Support</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<p style="font-size:11px;font-weight:700;color:#6B7280;letter-spacing:1.5px;text-transform:uppercase;border-bottom:1px solid #E0E4EA;padding-bottom:6px;margin-bottom:20px;">Contact Information</p>', unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(f"""
            <div class="metric-card" style="border-top:3px solid #1A3A5C;">
                <p style="font-size:13px;color:#6B7280;font-weight:600;margin:0;">RESPONSIBLE PERSON</p>
                <p style="font-size:22px;font-weight:700;color:#1A3A5C;margin:8px 0 0 0;">Mody</p>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div class="metric-card" style="border-top:3px solid #22C55E;">
                <p style="font-size:13px;color:#6B7280;font-weight:600;margin:0;">PHONE NUMBER</p>
                <p style="font-size:22px;font-weight:700;color:#1A7A4A;margin:8px 0 0 0;">010000000000</p>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div class="metric-card" style="border-top:3px solid #F59E0B;">
                <p style="font-size:13px;color:#6B7280;font-weight:600;margin:0;">EMAIL</p>
                <p style="font-size:18px;font-weight:700;color:#B45309;margin:8px 0 0 0;">Bright.Star@moonstone.com</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("""
        <div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:20px 24px;">
            <p style="color:#1A3A5C;font-weight:700;font-size:15px;margin:0 0 8px 0;">📋 Support Hours</p>
            <p style="color:#374151;margin:0;font-size:14px;">Sunday – Thursday: 9:00 AM – 5:00 PM</p>
            <p style="color:#374151;margin:4px 0 0 0;font-size:14px;">Friday – Saturday: Closed</p>
        </div>
        """, unsafe_allow_html=True)

