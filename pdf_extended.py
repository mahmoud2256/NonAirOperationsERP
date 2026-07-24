"""
Extended invoice PDF generator — adds a second table with the extra
Kanoo-format fields (Traveller, Requester, Ticket Number, Service
Multiplier, Airline Code, Taxes, S.FEE, Tax on Commission, Penalty,
etc.) underneath your existing invoice details / financial summary
tables.

HOW TO USE:
Call `generate_invoice_pdf_extended(data)` instead of your current
`generate_invoice_pdf(data)` when you want the fuller invoice. `data`
is the same dict you already build in show_invoices(), just also
include any of the new keys below if you collected them (any missing
key is shown as "-").
"""

import io
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm


EXTRA_FIELDS_LAYOUT = [
    ("Traveller Name", "traveller_name"),
    ("Customer ID", "customer_id"),
    ("Requester", "requester"),
    ("Ticket Number", "ticket_number"),
    ("Airline Code", "airline_code"),
    ("Dep Date", "dep_date"),
    ("Return Date", "return_date"),
    ("Origin", "origin"),
    ("Destination", "destination"),
    ("Trip Code", "trip_code"),
    ("Cost Centre", "cost_centre"),
    ("Business Unit", "business_unit"),
    ("Confirmation No", "confirmation_no"),
]

EXTRA_FINANCIAL_LAYOUT = [
    ("Taxes", "taxes"),
    ("S.FEE", "s_fee"),
    ("Tax on Commission", "tax_on_commission"),
    ("Penalty", "penalty"),
]


def generate_invoice_pdf_extended(data: dict) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=1.5 * cm, leftMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    elements = []

    title_style = ParagraphStyle(
        "Title", parent=styles["Normal"], fontSize=18,
        textColor=colors.HexColor("#1A3A5C"), spaceAfter=6,
        fontName="Helvetica-Bold",
    )
    subtitle_style = ParagraphStyle(
        "Subtitle", parent=styles["Normal"], fontSize=11,
        textColor=colors.HexColor("#6B7280"), spaceAfter=16,
        fontName="Helvetica",
    )

    elements.append(Paragraph("Non-Air Operations ERP", title_style))
    elements.append(Paragraph("Bright Star • Vendor Invoice", subtitle_style))
    elements.append(Spacer(1, 0.3 * cm))

    def styled_table(rows, header_span_two=True):
        t = Table(rows, colWidths=[6 * cm, 11 * cm])
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A3A5C")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 10),
            ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 1), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#F8FAFC"), colors.white]),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E0E4EA")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ]
        if header_span_two:
            style.append(("SPAN", (0, 0), (1, 0)))
        t.setStyle(TableStyle(style))
        return t

    # Core invoice details (same as your current table)
    core_rows = [
        ["FIELD", "VALUE"],
        ["Invoice No", data.get("invoice_no", "")],
        ["Date", data.get("date", "")],
        ["Owner", data.get("owner", "")],
        ["Accounts", data.get("accounts", "")],
        ["Subject", data.get("subject", "")],
        ["Type of Service", data.get("service_type", "")],
        ["Vendor", data.get("supplier", "")],
        ["City", data.get("supplier_city", "")],
        ["No. of PAX", str(data.get("no_of_pax", ""))],
        ["Currency", data.get("currency", "EGP")],
    ]
    elements.append(styled_table(core_rows, header_span_two=False))
    elements.append(Spacer(1, 0.4 * cm))

    # Extra trip/travel fields (only shown if at least one is present)
    extra_rows = [["EXTRA DETAILS", ""]]
    has_extra = False
    for label, key in EXTRA_FIELDS_LAYOUT:
        val = data.get(key)
        if val:
            has_extra = True
        extra_rows.append([label, str(val) if val else "-"])
    if has_extra:
        elements.append(styled_table(extra_rows))
        elements.append(Spacer(1, 0.4 * cm))

    # Financial summary — original 4 lines + the new extra financial fields
    fin_rows = [
        ["FINANCIAL SUMMARY", ""],
        ["Paid To Supplier", f"{data.get('paid_to_supplier', 0):,.2f} {data.get('currency', 'EGP')}"],
        ["Handling Fees", f"{data.get('handling_fees', 0):,.2f} {data.get('currency', 'EGP')}"],
        ["VAT (14%)", f"{data.get('vat', 0):,.2f} {data.get('currency', 'EGP')}"],
    ]
    for label, key in EXTRA_FINANCIAL_LAYOUT:
        val = data.get(key)
        if val:
            fin_rows.append([label, f"{float(val):,.2f} {data.get('currency', 'EGP')}"])
    fin_rows.append(["TOTAL AMOUNT", f"{data.get('total_amount', 0):,.2f} {data.get('currency', 'EGP')}"])

    fin_table = Table(fin_rows, colWidths=[6 * cm, 11 * cm])
    last_row = len(fin_rows) - 1
    fin_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A3A5C")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("SPAN", (0, 0), (1, 0)),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("BACKGROUND", (0, last_row), (-1, last_row), colors.HexColor("#EFF6FF")),
        ("FONTNAME", (0, last_row), (-1, last_row), "Helvetica-Bold"),
        ("FONTSIZE", (0, last_row), (-1, last_row), 11),
        ("TEXTCOLOR", (0, last_row), (-1, last_row), colors.HexColor("#1A3A5C")),
        ("FONTNAME", (0, 1), (0, last_row - 1), "Helvetica-Bold"),
        ("ROWBACKGROUNDS", (0, 1), (-1, last_row - 1), [colors.HexColor("#F8FAFC"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E0E4EA")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    elements.append(fin_table)
    elements.append(Spacer(1, 0.5 * cm))

    footer_style = ParagraphStyle(
        "Footer", parent=styles["Normal"], fontSize=8,
        textColor=colors.HexColor("#9CA3AF"), fontName="Helvetica",
    )
    elements.append(Paragraph(
        f"Generated by Non-Air Operations ERP  •  {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        footer_style,
    ))

    doc.build(elements)
    buffer.seek(0)
    return buffer
