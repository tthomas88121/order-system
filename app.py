import os
from io import BytesIO
from flask import Flask, render_template, request, send_file
from supabase import create_client, Client
from dotenv import load_dotenv
from openpyxl import Workbook

# 🔥 強制指定 .env 路徑（避免讀不到）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")
load_dotenv(dotenv_path=ENV_PATH)

app = Flask(__name__)

# 🔑 讀取環境變數
SUPABASE_URL = "https://vysechxapfowfddpfnyf.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ5c2VjaHhhcGZvd2ZkZHBmbnlmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ4MzYwNjEsImV4cCI6MjA5MDQxMjA2MX0._M27e3qJNgwaRFrXcJuHFsNgPOgz2wWSDF2ek-SZdNU"

# 🧪 Debug（先確認有讀到）
print("ENV PATH =", ENV_PATH)
print("SUPABASE_URL =", SUPABASE_URL)
print("SUPABASE_KEY loaded =", bool(SUPABASE_KEY))

# ❗ 如果沒讀到就直接報錯
if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("❌ 沒有成功讀到 .env，請檢查 SUPABASE_URL 和 SUPABASE_KEY")

# 🔗 建立 Supabase 連線
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


# ======================
# 🏠 首頁（訂單表單）
# ======================
@app.route("/")
def home():
    return render_template("order_form.html")


# ======================
# 📥 提交訂單
# ======================
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


# ======================
# 📋 查看訂單
# ======================
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


# ======================
# 📊 匯出 Excel
# ======================
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


# ======================
# 🚀 啟動
# ======================
if __name__ == "__main__":
    app.run(debug=True)