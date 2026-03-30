import os
from io import BytesIO
from flask import Flask, render_template, request, send_file
from supabase import create_client, Client
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()

app = Flask(__name__)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("Missing SUPABASE_URL or SUPABASE_KEY")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

@app.route("/")
def home():
    return render_template("order_form.html")

@app.route("/submit-order", methods=["POST"])
def submit_order():
    data = {
        "company_name": request.form.get("company_name"),
        "contact_name": request.form.get("contact_name"),
        "email": request.form.get("email"),
        "phone": request.form.get("phone"),
        "product_name": request.form.get("product_name"),
        "quantity": int(request.form.get("quantity")),
        "note": request.form.get("note"),
    }

    supabase.table("orders").insert(data).execute()

    return """
    <h2>✅ Order submitted successfully!</h2>
    <a href="/">Back</a>
    """

@app.route("/admin/orders")
def view_orders():
    result = supabase.table("orders").select("*").order("created_at", desc=True).execute()
    orders = result.data

    html = "<h2>Order List</h2><a href='/export-excel'>Export Excel</a><br><br>"
    html += "<table border='1' cellpadding='8'>"
    html += """
    <tr>
        <th>ID</th>
        <th>Company</th>
        <th>Contact</th>
        <th>Email</th>
        <th>Phone</th>
        <th>Product</th>
        <th>Qty</th>
        <th>Note</th>
        <th>Created At</th>
    </tr>
    """

    for order in orders:
        html += f"""
        <tr>
            <td>{order.get('id')}</td>
            <td>{order.get('company_name')}</td>
            <td>{order.get('contact_name')}</td>
            <td>{order.get('email')}</td>
            <td>{order.get('phone')}</td>
            <td>{order.get('product_name')}</td>
            <td>{order.get('quantity')}</td>
            <td>{order.get('note')}</td>
            <td>{order.get('created_at')}</td>
        </tr>
        """

    html += "</table>"
    return html

@app.route("/export-excel")
def export_excel():
    result = supabase.table("orders").select("*").order("created_at", desc=True).execute()
    orders = result.data

    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    ws.append([
        "ID", "Company Name", "Contact Name", "Email", "Phone",
        "Product Name", "Quantity", "Note", "Created At"
    ])

    for order in orders:
        ws.append([
            order.get("id"),
            order.get("company_name"),
            order.get("contact_name"),
            order.get("email"),
            order.get("phone"),
            order.get("product_name"),
            order.get("quantity"),
            order.get("note"),
            order.get("created_at")
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="orders.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)