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
    <link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazir-font@v30.1.0/dist/font-face.css" rel="stylesheet" />
    <style>
        body {
            direction: rtl;
            font-family: 'Vazir', sans-serif;
            background: linear-gradient(to bottom, #e6f0ff, #ffffff);
            background-attachment: fixed;
            margin: 0;
            padding: 0;
            color: #333;
        }
        header {
            background-color: #007bff;
            color: white;
            padding: 15px;
            text-align: center;
            font-size: 22px;
            font-weight: bold;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }
        .container {
            max-width: 1000px;
            margin: auto;
            padding: 30px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 15px;
            background: white;
            border-radius: 8px;
            overflow: hidden;
        }
        th, td {
            border: 1px solid #dee2e6;
            padding: 10px;
            text-align: center;
        }
        th {
            background-color: #007bff;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f6fc;
        }
        tr:hover {
            background-color: #e9f1ff;
        }
        .card {
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            margin-bottom: 20px;
        }
        input[type="text"], input[type="password"], input[type="file"] {
            width: 100%;
            padding: 10px;
            margin: 6px 0;
            border: 1px solid #ced4da;
            border-radius: 6px;
        }
        input[type="submit"], button {
            background-color: #28a745;
            color: white;
            padding: 10px 25px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            margin-top: 10px;
        }
        input[type="submit"]:hover, button:hover {
            background-color: #218838;
        }
        a {
            color: #007bff;
            text-decoration: none;
        }
        a:hover {
            color: #0056b3;
        }
        .logout {
            position: absolute;
            left: 20px;
            top: 20px;
        }
        @media (max-width: 768px) {
            table, thead, tbody, th, td, tr {
                display: block;
                width: 100%;
            }
            thead tr {
                display: none;
            }
            td {
                position: relative;
                padding-right: 50%;
                text-align: right;
                border: none;
            }
            td::before {
                content: attr(data-label);
                position: absolute;
                right: 10px;
                top: 10px;
                font-weight: bold;
                color: #666;
            }
        }
    </style>
</head>
<body>
    <header>
        سامانه صندوق قرض‌الحسنه
        {% if session.get('logged_in') %}
            <div class="logout"><a href="{{ url_for('logout') }}" style="color:white;">خروج</a></div>
        {% endif %}
    </header>
    <div class="container">
        {{ content|safe }}
    </div>
</body>
</html>
'''

# توابع کمکی (خواندن اکسل، تبدیل اعداد و ...)

def read_excel_data():
    try:
        df_report = pd.read_excel(EXCEL_FILE, sheet_name="report", header=6)
        df_invoice = pd.read_excel(EXCEL_FILE, sheet_name="invoice", header=1)
        return df_report, df_invoice
    except Exception as e:
        return str(e), None

def to_persian_numbers(text):
    return str(text).translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))

def format_money(val):
    try:
        val = float(val)
        return to_persian_numbers(f"{val:,.0f} ریال")
    except:
        return "-"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# صفحات

@app.route("/", methods=["GET", "POST"])
def index():
    if session.get("logged_in"):
        return redirect(url_for("admin_panel" if session.get("is_admin") else "user_panel"))

    result = ""
    if request.method == "POST":
        username = request.form.get("username").strip()
        password = request.form.get("password").strip()

        if username == "admin" and password == "admin123":
            session["logged_in"] = True
            session["is_admin"] = True
            return redirect(url_for("admin_panel"))

        df_report, _ = read_excel_data()
        if isinstance(df_report, str):
            result = f"<p style='color:red;'>{df_report}</p>"
        else:
            for _, row in df_report.iterrows():
                account = str(row.get("شماره حساب", "")).strip().replace(".0", "")
                name = str(row.get("نام و نام خانوادگي", "")).strip()
                password_excel = str(row.get("رمز", "")).strip().replace(".0", "")

                if (username == account or username == name) and password == password_excel:
                    session["logged_in"] = True
                    session["is_admin"] = False
                    session["account"] = account
                    session["name"] = name
                    return redirect(url_for("user_panel"))

            result = "<p style='color:red;'>نام کاربری یا رمز عبور اشتباه است.</p>"

    content = f'''
    <div class="card">
        <h2>ورود</h2>
        <form method="post">
            <label>نام یا شماره حساب:</label>
            <input type="text" name="username">
            <label>رمز عبور:</label>
            <input type="password" name="password">
            <input type="submit" value="ورود">
        </form>
        {result}
    </div>
    '''
    return render_template_string(BASE_HTML, content=content)

@app.route("/user")
def user_panel():
    if not session.get("logged_in") or session.get("is_admin"):
        return redirect(url_for("index"))

    df_report, _ = read_excel_data()
    account = session.get("account")
    name = session.get("name")
    row = df_report[df_report["شماره حساب"].astype(str).str.replace(".0", "") == account].iloc[0]

    content = f'''
    <div class="card">
        <h2>خوش آمدید {name}</h2>
        <p><b>شماره حساب:</b> {to_persian_numbers(account)}</p>
        <h4>اطلاعات مالی:</h4>
        <ul>
    '''
    for col, val in row.iloc[6:].items():
        if pd.notna(val) and str(col).strip() != "" and not str(col).startswith("Unnamed"):
            content += f"<li>{col}: {format_money(val)}</li>"
    content += "</ul><br>"
    content += f'<a href="{url_for("transactions")}"><button>مشاهده تراکنش‌ها</button></a></div>'

    return render_template_string(BASE_HTML, content=content)

@app.route("/transactions")
def transactions():
    if not session.get("logged_in") or session.get("is_admin"):
        return redirect(url_for("index"))

    _, df_invoice = read_excel_data()
    account = int(session.get("account"))
    name = session.get("name")
    search_query = request.args.get("search", "").strip()

    user_invoice = df_invoice[df_invoice["شماره حساب"] == account]
    if search_query:
        user_invoice = user_invoice[user_invoice.astype(str).apply(lambda row: row.str.contains(search_query, case=False, na=False)).any(axis=1)]

    cols = ['تاریخ', 'برداشت', 'واریز', 'ماهیانه', 'قسط وام', 'وام', 'وام فوری', 'برگشت وام فوری', 'توضیحات']

    content = f'''
    <div class="card">
        <h2>تراکنش‌های {name}</h2>
        <form method="get">
            <input type="text" name="search" placeholder="جستجو..." value="{search_query}">
            <input type="submit" value="جستجو">
        </form>
    '''

    if not user_invoice.empty:
        content += "<table><tr>" + "".join([f"<th>{col}</th>" for col in cols]) + "</tr>"
        for _, row in user_invoice.iterrows():
            content += "<tr>" + "".join([
                f"<td data-label='{col}'>{format_money(row[col]) if col not in ['توضیحات', 'تاریخ'] else to_persian_numbers(row[col])}</td>" if pd.notna(row[col]) else "<td>-</td>"
                for col in cols
            ]) + "</tr>"
        content += "</table>"
    else:
        content += "<p>تراکنشی یافت نشد.</p>"
    content += "</div>"

    return render_template_string(BASE_HTML, content=content)

@app.route("/admin")
def admin_panel():
    if not session.get("logged_in") or not session.get("is_admin"):
        return redirect(url_for("index"))

    df_report, _ = read_excel_data()
    financial_cols = [c for c in df_report.columns[6:] if pd.notna(c) and not str(c).startswith("Unnamed")]

    content = '''
    <div class="card">
        <h2>پنل مدیریت</h2>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <label>بارگذاری فایل Excel جدید:</label>
            <input type="file" name="file">
            <input type="submit" value="بارگذاری">
        </form>
        <br>
        <table>
            <tr><th>نام</th><th>شماره حساب</th>''' + "".join([f"<th>{col}</th>" for col in financial_cols]) + "</tr>"

    for _, row in df_report.iterrows():
        content += f"<tr><td>{row['نام و نام خانوادگي']}</td><td>{to_persian_numbers(row['شماره حساب'])}</td>"
        for col in financial_cols:
            content += f"<td>{format_money(row[col])}</td>"
        content += "</tr>"

    content += "</table></div>"
    return render_template_string(BASE_HTML, content=content)

@app.route("/upload", methods=["POST"])
def upload_excel():
    if not session.get("logged_in") or not session.get("is_admin"):
        return redirect(url_for("index"))

    file = request.files.get("file")
    if file and allowed_file(file.filename):
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], "base_hor.xlsm")
        file.save(filepath)
    return redirect(url_for("admin_panel"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)

