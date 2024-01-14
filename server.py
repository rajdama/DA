from flask import Flask, render_template, request, send_file
import os
from script import report

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            return 'No file part'

        file = request.files['file']
        # If the user does not select a file, the browser submits an empty file without a filename
        if file.filename == '':
            return 'No selected file'

        if file:
            report(file)
            response = send_file('output_file.xlsx', as_attachment=True)
            return response
    
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
