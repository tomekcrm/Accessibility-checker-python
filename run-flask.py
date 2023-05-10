import os
from flask import Flask, request, send_file

app = Flask(__name__)

@app.route('/')
def list_reports():
    reports = [f for f in os.listdir() if f.startswith('report') and f.endswith('.zip')]
    if len(reports) != 0:
        links = [f'<a href="/download-report?filename={f}">{f}</a>' for f in reports]
        return '<br>'.join(links)
    return "No reports to show"

@app.route('/download-report')
def download_report():
    filename = request.args.get('filename', None)
    if not filename:
        return 'No filename specified'
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host=os.getenv('SERVER_IP'), port=5000)
