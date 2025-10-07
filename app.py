from flask import Flask, render_template, send_from_directory
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/template.docx')
def get_template():
    """Отдаем шаблон файла для загрузки в браузер"""
    return send_from_directory('.', 'template.docx')

if __name__ == '__main__':
    app.run(debug=True)
