from flask import Flask, render_template_string, request, redirect, session, url_for
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "your_secret_key"

EXCEL_FILE = r"E:\bank data11111111111111\sandogh\base_hor.xlsm"
UPLOAD_FOLDER = os.path.dirname(EXCEL_FILE)
ALLOWED_EXTENSIONS = {'xls', 'xlsm'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

BASE_HTML = '''
<!doctype html>
<html lang="fa">
<head>
    <meta charset="UTF-8">
    <title>صندوق قرض‌الحسنه</title>
    <link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazir-font@v30.1.0/dist/font-face.css" rel="stylesheet">
    <style>
        body {
            direction: rtl;
            font-family: 'Vazir', Tahoma, sans-serif;
            background-image: url('https://images.unsplash.com/photo-1617957742303-b9e5a986d718?auto=format&fit=crop&w=1400&q=80');
            background-size: cover;
            background-attachment: fixed;
            color: #fff;
            padding: 0;
            margin: 0;
        }
        .container {
            background-color: rgba(0,0,0,0.7);
            margin: 40px auto;
            padding: 30px;
            max-width: 900px;
            border-radius: 15px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #ccc;
            padding: 10px;
        }
        th {
            background-color: #343a40;
            color: #f8f9fa;
        }
        tr:nth-child(even) {
            background-color: #2e2e2e;
        }
        tr:hover {
            background-color: #495057;
        }
        h2, h4 {
            color: #ffc107;
        }
        button, input[type="submit"] {
            background-color: #17a2b8;
            border: none;
            color: white;
            padding: 10px 18px;
            font-size: 14px;
            cursor: pointer;
            border-radius: 6px;
        }
        input[type="text"], input[type="password"], input[type="file"] {
            padding: 7px;
            margin: 6px 0;
            border: none;
            border-radius: 6px;
            width: 100%;
        }
        label {
            color: #f8f9fa;
            font-weight: bold;
        }
        a {
            text-decoration: none;
            color: #ffc107;
        }
        a:hover {
            color: #fff;
        }
        ul {
            line-height: 1.8;
        }
        @media (max-width: 768px) {
            table, thead, tbody, th, td, tr {
                display: block;
                width: 100%;
            }
            thead tr { display: none; }
            tr {
                margin-bottom: 15px;
                border-bottom: 2px solid #ccc;
                padding-bottom: 10px;
            }
            td {
                text-align: right;
                padding-right: 50%;
                position: relative;
                border: none !important;
            }
            td::before {
                content: attr(data-label);
                position: absolute;
                right: 10px;
                top: 8px;
                font-weight: bold;
                white-space: nowrap;
                color: #ffc107;
            }
        }
    </style>
</head>
<body>
    <div class="container">
    {% if session.get('logged_in') %}
        <div style="text-align:left;"><a href="{{ url_for('logout') }}">🔒 خروج</a></div>
    {% endif %}
    {{ content|safe }}
    </div>
</body>
</html>
'''

# 📌 توابع کمکی
def read_excel_data():
    try:
        df_report = pd.read_excel(EXCEL_FILE, sheet_name="report", header=6)
        df_invoice = pd.read_excel(EXCEL_FILE, sheet_name="invoice", header=1)
        return df_report, df_invoice
    except Exception as e:
        return str(e), None

def format_money(val):
    try:
        val = float(val)
        formatted = f"{val:,.0f} ریال"
        return to_persian_numbers(formatted)
    except:
        return "-"

def to_persian_numbers(text):
    en_to_fa = str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')
    return str(text).translate(en_to_fa)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

LOGIN_FORM = '''
<h2>ورود به سامانه</h2>
<form method=post>
    <label>نام یا شماره حساب:</label>
    <input type=text name=username>
    <label>رمز:</label>
    <input type=password name=password><br>
    <input type=submit value="ورود">
</form>
<hr>
{{ result|safe }}
'''

# 📍 روت‌ها
@app.route("/", methods=["GET", "POST"])
def index():
    if session.get("logged_in"):
        if session.get("is_admin"):
            return redirect(url_for("admin_panel"))
        else:
            return redirect(url_for("user_panel"))

    result = ""
    if request.method == "POST":
        username = request.form.get("username").strip()
        password = request.form.get("password").strip()

        if username == "admin" and password == "admin123":
            session["logged_in"] = True
            session["is_admin"] = True
            return redirect(url_for("admin_panel"))

        df_report, df_invoice = read_excel_data()
        if isinstance(df_report, str):
            result = f"<b>خطا در خواندن فایل:</b> {df_report}"
        else:
            for _, row in df_report.iterrows():
                account = str(row.get("شماره حساب", "")).strip().replace(".0", "")
                password_excel = str(row.get("رمز", "")).strip().replace(".0", "")
                name = str(row.get("نام و نام خانوادگي", "")).strip()

                if (username == account or username == name) and password == password_excel:
                    session["logged_in"] = True
                    session["is_admin"] = False
                    session["account"] = account
                    session["name"] = name
                    return redirect(url_for("user_panel"))

            result = "<p style='color:red;'>نام کاربری یا رمز عبور اشتباه است.</p>"

    return render_template_string(BASE_HTML, content=LOGIN_FORM, result=result)

@app.route("/user")
def user_panel():
    if not session.get("logged_in") or session.get("is_admin"):
        return redirect(url_for("index"))

    df_report, _ = read_excel_data()
    if isinstance(df_report, str):
        return render_template_string(BASE_HTML, content=f"<b>خطا:</b> {df_report}")

    account = session.get("account")
    name = session.get("name")
    row = df_report[df_report["شماره حساب"].astype(str).str.replace(".0", "") == account].iloc[0]

    content = f"<h2>خوش آمدید {name}</h2><p><b>شماره حساب:</b> {account}</p><h4>اطلاعات مالی:</h4><ul>"
    financial_data = row.iloc[6:]
    for col, val in financial_data.items():
        if pd.notna(val) and str(col).strip() != "" and not str(col).startswith("Unnamed") and "nan" not in str(col):
            content += f"<li>{col}: {format_money(val)}</li>"
    content += "</ul><br><a href='/transactions'><button>مشاهده تراکنش‌ها</button></a>"

    return render_template_string(BASE_HTML, content=content)

@app.route("/transactions")
def transactions():
    if not session.get("logged_in") or session.get("is_admin"):
        return redirect(url_for("index"))

    df_report, df_invoice = read_excel_data()
    if isinstance(df_invoice, str):
        return render_template_string(BASE_HTML, content=f"<b>خطا:</b> {df_invoice}")

    account = session.get("account")
    name = session.get("name")
    user_invoice = df_invoice[df_invoice["شماره حساب"] == int(account)]

    search_query = request.args.get("search", "").strip()
    if search_query:
        user_invoice = user_invoice[user_invoice.astype(str).apply(lambda row: row.str.contains(search_query, case=False, na=False)).any(axis=1)]

    important_cols = ['تاریخ', 'برداشت', 'واریز', 'ماهیانه', 'قسط وام', 'وام', 'وام فوری', 'برگشت وام فوری', 'ش کارت']


    content = f'''
    <h2>تراکنش‌های {name}</h2>
    <a href="/user">⬅️ بازگشت به داشبورد</a>
    <form method="get">
        <input type="text" name="search" placeholder="جستجو..." value="{search_query}">
        <input type="submit" value="جستجو">
    </form>
    '''

    if not user_invoice.empty:
        content += "<table><tr>" + "".join(f"<th>{col}</th>" for col in important_cols) + "</tr>"
        for _, row in user_invoice.iterrows():
            content += "<tr>" + "".join(
                f"<td data-label='{col}'>" + (
                    format_money(row.get(col)) if col not in ['توضیحات', 'تاریخ'] and pd.notna(row.get(col)) else (
                        to_persian_numbers(row.get(col)) if pd.notna(row.get(col)) else "-"
                    )
                ) + "</td>" for col in important_cols) + "</tr>"
        content += "</table>"
    else:
        content += "<p>تراکنشی یافت نشد.</p>"

    return render_template_string(BASE_HTML, content=content)

@app.route("/admin")
def admin_panel():
    if not session.get("logged_in") or not session.get("is_admin"):
        return redirect(url_for("index"))

    df_report, _ = read_excel_data()
    if isinstance(df_report, str):
        return render_template_string(BASE_HTML, content=f"<b>خطا:</b> {df_report}")

    content = '''
    <h2>پنل مدیر</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <label>بارگذاری فایل Excel:</label>
        <input type="file" name="file">
        <input type="submit" value="بارگذاری">
    </form>
    <table><tr><th>نام</th><th>شماره حساب</th>'''

    financial_columns = df_report.columns[6:]
    for col in financial_columns:
        if pd.notna(col) and str(col).strip() != "" and not str(col).startswith("Unnamed") and "nan" not in str(col):
            content += f"<th>{col}</th>"
    content += "</tr>"

    for _, row in df_report.iterrows():
        content += f"<tr><td>{row.get('نام و نام خانوادگي')}</td><td>{row.get('شماره حساب')}</td>"
        for col in financial_columns:
            if pd.notna(col) and str(col).strip() != "" and not str(col).startswith("Unnamed") and "nan" not in str(col):
                val = row.get(col)
                content += f"<td>{format_money(val)}</td>"
        content += "</tr>"

    content += "</table>"
    return render_template_string(BASE_HTML, content=content)

@app.route("/upload", methods=["POST"])
def upload_excel():
    if not session.get("logged_in") or not session.get("is_admin"):
        return redirect(url_for("index"))

    file = request.files.get("file")
    if not file or file.filename == '':
        return redirect(url_for("admin_panel"))

    if allowed_file(file.filename):
        filename = "base_hor.xlsm"
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(filepath)
    return redirect(url_for("admin_panel"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)
