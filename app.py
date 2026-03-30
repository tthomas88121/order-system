import os
from io import BytesIO
from flask import Flask, render_template, request, send_file, redirect, url_for
from supabase import create_client, Client
from dotenv import load_dotenv
from openpyxl import Workbook

print("STEP 1: app starting...")

# 🔥 force load .env
load_dotenv(dotenv_path=".env")

app = Flask(__name__)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

print("STEP 2: URL =", SUPABASE_URL)
print("STEP 3: KEY exists =", bool(SUPABASE_KEY))

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("Missing SUPABASE_URL or SUPABASE_KEY")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# ✅ Product categories
PRODUCT_CATEGORIES = {
    "磁性尺(MagLine)": [
        "MSK5000AS","MSK5000","LEC100","LEC160","LEC200",
        "MSA111C","MSA213C","MSK200/1","LE200","LE100/1",
        "MSK1000","MA503AS","MA564","MA523/1","MA504/1",
        "OSK20","TS20","MB160","MB500/1","MBA111",
        "MB200/1","MBA213","MB100/1","PSI500","MBR500","MBR200"
    ],
    "定位驅動器(DriveLine)": [
        "AG05","WG05","AG03","AG25","AG06","AG24","ETC5000","AG26"
    ],
    "拉繩編碼器(LinearLine)": [
        "SG5","SG10","SG20","SG30","SG60","SGH10","SGH25","SGHF50"
    ],
    "旋轉編碼器(RotoLine)": [
        "IG06","IH58S15","IG07","IH5828"
    ],
    "位置指示器(PositionLine)": [
        "DA09S","DE10","AP10","AP10T","AP20","AP05",
        "GS04","DE04","KP09","KP04",
        "SZ80/1","DE10P","AP10S","DA05/1",
        "AP20S","DA08","MS500H PL","DK01","KP09P"
    ]
}


@app.route("/")
def home():
    return redirect(url_for("order_page"))


@app.route("/order")
def order_page():
    return render_template(
        "order_form.html",
        product_categories=PRODUCT_CATEGORIES
    )


@app.route("/submit-order", methods=["POST"])
def submit_order():
    customer_name = request.form.get("customer_name", "").strip()
    company_name = request.form.get("company_name", "").strip()
    phone = request.form.get("phone", "").strip()
    email = request.form.get("email", "").strip()
    category = request.form.get("category", "").strip()
    part_number = request.form.get("part_number", "").strip()
    quantity = request.form.get("quantity", "").strip()
    note = request.form.get("note", "").strip()

    if not customer_name or not category or not part_number or not quantity:
        return "Missing required fields", 400

    if category not in PRODUCT_CATEGORIES:
        return "Invalid category", 400

    if part_number not in PRODUCT_CATEGORIES[category]:
        return "Invalid part number", 400

    try:
        quantity_value = int(quantity)
        if quantity_value < 1:
            return "Quantity must be at least 1", 400
    except ValueError:
        return "Quantity must be a number", 400

    data = {
        "customer_name": customer_name,
        "company_name": company_name,
        "phone": phone,
        "email": email,
        "category": category,
        "part_number": part_number,
        "product_name": f"{category} | {part_number}",
        "quantity": quantity_value,
        "note": note,
    }

    supabase.table("orders").insert(data).execute()

    return f"""
    <html>
    <body style="font-family: Arial; text-align:center; padding:50px;">
        <h2 style="color:#0047c6;">Order Submitted</h2>
        <p>{customer_name} - {part_number} ({quantity_value})</p>
        <a href="/order">Back</a>
    </body>
    </html>
    """


@app.route("/admin/orders")
def view_orders():
    result = supabase.table("orders").select("*").order("created_at", desc=True).execute()
    orders = result.data

    html = "<h2>Orders</h2><table border='1' cellpadding='8'>"
    html += "<tr><th>Customer</th><th>Part</th><th>Qty</th></tr>"

    for order in orders:
        html += f"""
        <tr>
            <td>{order.get('customer_name')}</td>
            <td>{order.get('part_number')}</td>
            <td>{order.get('quantity')}</td>
        </tr>
        """

    html += "</table>"
    return html


@app.route("/export-excel")
def export_excel():
    result = supabase.table("orders").select("*").execute()
    orders = result.data

    wb = Workbook()
    ws = wb.active

    ws.append(["Customer", "Part", "Qty"])

    for order in orders:
        ws.append([
            order.get("customer_name"),
            order.get("part_number"),
            order.get("quantity")
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="orders.xlsx")


# 🔥 FORCE RUN
if __name__ == "__main__":
    print("🚀 Flask is starting...")
    app.run(host="127.0.0.1", port=5000, debug=True)