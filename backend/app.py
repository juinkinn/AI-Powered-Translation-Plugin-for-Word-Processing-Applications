from flask import Flask
from routes.translation import translate_bp  

app = Flask(__name__)

# Đăng ký Blueprint
app.register_blueprint(translate_bp)

@app.route('/')
def home():
    return "Flask server is running!"

if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0", port=5000)
