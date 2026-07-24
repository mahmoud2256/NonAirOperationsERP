"""
Maps the external invoice export headers (the Kanoo-style spreadsheet
you shared) to database column names used in the `invoices` table.

Left  = exact header text as it appears in the source file
Right = db column name (snake_case)

Fields that already exist in your current schema (date, invoice_no,
supplier, service_type, currency, vat, total_amount, po_number,
gross_amount, hidden_commission) are mapped to those SAME existing
columns so nothing is duplicated.
"""

COLUMN_MAPPING = {
    "Client Name": "accounts",                 # maps to existing "accounts" column
    "Traveller Name": "traveller_name",
    "Invoice date": "date",                    # existing column
    "Invoice no.": "invoice_no",                # existing column
    "Customer ID": "customer_id",
    "PO NUM": "po_number",                      # existing column
    "Requester": "requester",
    "service type": "service_type",             # existing column
    "Ticket Number": "ticket_number",
    "Service Multiplier": "service_multiplier",
    "Airline Code": "airline_code",
    "Status": "ext_status",
    "Creator": "creator",
    "Issuer": "owner",                          # existing column (Owner)
    "Invoice currency": "currency",             # existing column
    "Dep Date": "dep_date",
    "Return": "return_date",
    "Basic Fare": "gross_amount",               # existing column
    "Taxes": "taxes",
    "S.FEE": "s_fee",
    "SUP Commission": "hidden_commission",      # existing column
    "Tax on Commission": "tax_on_commission",
    "VAT": "vat",                               # existing column
    "Penalty": "penalty",
    "TOTAL": "total_amount",                    # existing column
    "Invoice Client Type": "invoice_client_type",
    "Form of payment code": "form_of_payment_code",
    "Service code": "service_code",
    "Service category": "service_category",
    "Service product type": "service_product_type",
    "Service product name": "service_product_name",
    "Service source": "service_source",
    "Service description": "service_description",
    "Ext. System Id": "ext_system_id",
    "Origin City": "origin_city",
    "Destination City": "destination_city",
    "Order Type": "order_type",
    "Order Code": "order_code",
    "Operational status (code)": "operational_status_code",
    "Financial status (code)": "financial_status_code",
    "Origin": "origin",
    "Destination": "destination",
    "Emp ID": "emp_id",
    "Cost Centre": "cost_centre",
    "LPO NUM": "lpo_num",
    "Business Unit": "business_unit",
    "Ordered By": "ordered_by",
    "Trip Code": "trip_code",
    "Delivered To": "delivered_to",
    "TYPE OF TRAVEL": "type_of_travel",
    "Transaction Type": "transaction_type",
    "No of Nights": "no_of_nights",
    "Classification": "classification",
    "SBU Code": "sbu_code",
    "PROJECT CODE": "project_code",
    "Ticket Type": "ticket_type",
    "Bill or No Bill": "bill_or_no_bill",
    "FRN Num": "frn_num",
    "TR NUM": "tr_num",
    "Online Booking": "online_booking",
    "Authorized By": "authorized_by",
    "MS Trip purpose": "ms_trip_purpose",
    "Dept Id": "dept_id",
    "Published Fare": "published_fare",
    "Class of Travel": "class_of_travel",
    "Air Reason Code": "air_reason_code",
    "Supplier ID": "supplier_id",
    "Company ID": "company_id",
    "Trip Purpose code": "trip_purpose_code",
    "TA NUM": "ta_num",
    "Confirmation No": "confirmation_no",
    "Airline Type": "airline_type",
    "EMBASSY NAME": "embassy_name",
    "VISA Type": "visa_type",
    "Event Type": "event_type",
    "Event No of Guests": "event_no_of_guests",
    "Voucher No./EInv Internal ID": "voucher_no",
    "Default currency exchange rate": "exchange_rate",
    "Outlet code": "outlet_code",
    "Lowest Fare": "lowest_fare",
    "Fare wo deal": "fare_wo_deal",
    "Actual Responsible User": "actual_responsible_user",
    "Class Code": "class_code",
    "Class Name": "class_name",
    "Request Date": "request_date",
    "Request": "request_notes",
}

# Columns that hold money amounts (need comma/number cleanup on import)
NUMERIC_COLUMNS = [
    "gross_amount", "taxes", "s_fee", "hidden_commission",
    "tax_on_commission", "vat", "penalty", "total_amount",
    "service_multiplier", "no_of_nights", "event_no_of_guests",
    "exchange_rate", "published_fare", "lowest_fare", "fare_wo_deal",
]

# Columns that hold dates (need dd-mm-yyyy -> yyyy-mm-dd conversion)
DATE_COLUMNS = ["date", "dep_date", "return_date", "request_date"]
