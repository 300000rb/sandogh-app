from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    return "سلام! اپلیکیشن روی Vercel اجرا شد."

# نیازی به if __name__ == "__main__": نیست چون Vercel از آن استفاده نمی‌کند
