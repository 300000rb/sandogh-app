from flask import Flask, render_template_string, request, redirect, session, url_for
import pandas as pd
import os
from werkzeug.utils import secure_filename
app = Flask(__name__)
app.secret_key = "your_secret_key"
EXCEL_FILE = r"E:\bank data11111111111111\sandogh\base_hor.xlsm"
UPLOAD_FOLDER = os.path.dirname(EXCEL_FILE)  # یعنی همان پوشه فایل اکسل
ALLOWED_EXTENSIONS = {'xls', 'xlsm'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


BASE_HTML = '''
<!doctype html>
<html lang="fa">
<head>
    <meta charset="UTF-8">
    <title>صندوق</title>
    <link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazir-font@v30.1.0/dist/font-face.css" rel="stylesheet" type="text/css" />
    <style>
        body {
            direction: rtl;
            font-family: 'Vazir', Tahoma, sans-serif;
            padding: 20px;
            background-color: #f8f9fa;
            color: #212529;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 10px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
        }
        th {
            background-color: #007bff;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        tr:hover {
            background-color: #e9ecef;
        }
        button, input[type="submit"] {
            background-color: #28a745;
            border: none;
            color: white;
            padding: 8px 16px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 14px;
            margin: 4px 2px;
            cursor: pointer;
            border-radius: 5px;
        }
        input[type="text"], input[type="password"], input[type="file"] {
            padding: 5px;
            margin: 5px 0;
            border: 1px solid #ccc;
            border-radius: 5px;
        }
        a {
            text-decoration: none;
            color: #007bff;
        }
        a:hover {
            color: #0056b3;
        }
   @media (max-width: 768px) {
        table, thead, tbody, th, td, tr {
            display: block;
            width: 100%;
        }

        thead tr {
            display: none;
        }

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
            color: #666;
        }
    } 
    </style>
</head>
<body>
    {% if session.get('logged_in') %}
        <div style="text-align:left;"><a href="{{ url_for('logout') }}">خروج</a></div>
    {% endif %}
    {{ content|safe }}
</body>
</html>
'''

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
<h2>ورود</h2>
<form method=post>
    <label>نام یا شماره حساب:</label><br>
    <input type=text name=username><br>
    <label>رمز:</label><br>
    <input type=password name=password><br><br>
    <input type=submit value="ورود">
</form>
<hr>
{{ result|safe }}
'''
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

    content = f"<h2 style='color:green;'>خوش آمدید {name}</h2>"
    content += f"<p><b>شماره حساب:</b> {account}</p>"
    content += "<h4>اطلاعات مالی:</h4><ul>"

    financial_data = row.iloc[6:]
    for col, val in financial_data.items():
        if pd.notna(val) and str(col).strip() != "" and not str(col).startswith("Unnamed") and "nan" not in str(col):
            content += f"<li>{col}: {format_money(val)}</li>"
    content += "</ul>"

    content += f'<br><a href="{url_for("transactions")}"><button>مشاهده تراکنش‌ها</button></a>'

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

    important_cols = ['تاریخ', 'برداشت', 'واریز', 'ماهیانه', 'قسط وام', 'وام', 'وام فوری', 'برگشت وام فوری', 'توضیحات']

    content = f'''
    <h2 style="color:green;">تراکنش‌های {name}</h2>
    <div style="margin-bottom: 15px;">
        <a href="{url_for('user_panel')}">⬅️ بازگشت به داشبورد</a>
    </div>
    <form method="get" style="margin-bottom:10px;">
        <input type="text" name="search" placeholder="جستجو..." value="{search_query}">
        <input type="submit" value="جستجو">
    </form>
    '''

    if not user_invoice.empty:
        content += "<table border='1' style='direction: rtl; text-align: center;'><tr>"
        for col in important_cols:
            content += f"<th>{col}</th>"
        content += "</tr>"

        for _, row in user_invoice.iterrows():
            content += "<tr>"
            for col in important_cols:
                val = row.get(col)
                value = format_money(val) if col not in ['توضیحات', 'تاریخ'] and pd.notna(val) else (to_persian_numbers(val) if pd.notna(val) else "-")
                content += f"<td data-label='{col}'>{value}</td>"
            content += "</tr>"
        content += "</table>"
    else:
        content += "<p>تراکنشی یافت نشد.</p>"

    return render_template_string(BASE_HTML, content=content, to_persian_numbers=to_persian_numbers)
@app.route("/admin")
def admin_panel():
    if not session.get("logged_in") or not session.get("is_admin"):
        return redirect(url_for("index"))

    df_report, _ = read_excel_data()
    if isinstance(df_report, str):
        return render_template_string(BASE_HTML, content=f"<b>خطا:</b> {df_report}")

    content = '''
    <h2>پنل مدیر</h2>
    <form action="/upload" method="post" enctype="multipart/form-data" style="margin-bottom: 20px;">
        <label>بارگذاری فایل Excel:</label>
        <input type="file" name="file">
        <input type="submit" value="بارگذاری">
    </form>
    <table border='1' style='text-align:center; direction: rtl;'>
        <tr>
            <th>نام</th><th>شماره حساب</th>'''

    # افزودن همان ستون‌های اطلاعاتی که کاربر عادی می‌بیند
    financial_columns = df_report.columns[6:]
    for col in financial_columns:
        if pd.notna(col) and str(col).strip() != "" and not str(col).startswith("Unnamed") and "nan" not in str(col) and str(col) not in ["مقایسه", "شماره حساب2", "299465200"]:
            content += f"<th>{col}</th>"
    content += "</tr>"

    for _, row in df_report.iterrows():
        name = row.get("نام و نام خانوادگي", "")
        acc = row.get("شماره حساب", "")
        content += f"<tr><td>{name}</td><td>{acc}</td>"

        for col in financial_columns:
            if pd.notna(col) and str(col).strip() != "" and not str(col).startswith("Unnamed") and "nan" not in str(col) and str(col) not in ["مقایسه", "شماره حساب2", "299465200"]:
                val = row.get(col)
                content += f"<td>{format_money(val)}</td>"
        content += "</tr>"

    content += "</table>"

    return render_template_string(BASE_HTML, content=content)
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))
@app.route('/upload', methods=['POST'])
def upload_excel():
    if not session.get("logged_in") or not session.get("is_admin"):
        return redirect(url_for("index"))

    file = request.files.get('file')
    if not file or file.filename == '':
        return redirect(url_for('admin_panel'))

    if allowed_file(file.filename):
        filename = "base_hor.xlsm"  # نام فایل هدف را ثابت نگه می‌داریم
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        message = "<p style='color:green;'>فایل با موفقیت بارگذاری شد.</p>"
    else:
        message = "<p style='color:red;'>فقط فایل‌های Excel با فرمت .xls یا .xlsm مجاز هستند.</p>"

    session['upload_message'] = message
    return redirect(url_for('admin_panel'))
if __name__ == "__main__":
    app.run(debug=True)
