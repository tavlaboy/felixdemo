from flask import Flask, render_template, request, redirect
import pandas as pd
from functions import removeAddress, sum_by_mapping  # Import your new function
import os
# Configure Flask to use the current directory for both templates and static files.
app = Flask(__name__, template_folder=".", static_folder=".", static_url_path="")

@app.route('/', methods=['GET', 'POST'])
@app.route('/index.html', methods=['GET', 'POST'])
def index():
    table_html = None
    second_table_html = None

    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']

        if file.filename == '':
            return redirect(request.url)

        if file:
            file_path = 'uploaded_file.xlsx'
            file.save(file_path)

            # Process DataFrame
            modified_df = removeAddress(file_path)
            table_html = modified_df.to_html(classes="table table-bordered", index=False)

            # Use the grouping function to get the second DataFrame
            second_df = sum_by_mapping(modified_df)
            second_table_html = second_df.to_html(classes="table table-striped", index=False)

    return render_template('index.html', table=table_html, second_table=second_table_html)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))  # Default to 10000 if PORT is not set
    app.run(debug=True, host="0.0.0.0", port=port)
