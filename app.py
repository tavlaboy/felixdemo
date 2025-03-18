from flask import Flask, render_template, request, redirect, send_file, url_for
import pandas as pd

import os

from functions import removeAddress, sum_by_mapping  # Import your processing functions


from functions import removeAddress, sum_by_mapping  # Import your new function
import os
# Configure Flask to use the current directory for both templates and static files.
>>>>>>> a93f2f88ed82237aac57fd91691b29e30356854d
app = Flask(__name__, template_folder=".", static_folder=".", static_url_path="")

# Define a dedicated folder for uploaded and processed files
UPLOAD_FOLDER = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)  # Create the folder if it doesn't exist

@app.route('/', methods=['GET', 'POST'])
@app.route('/index.html', methods=['GET', 'POST'])
def index():
    table_html = None
    second_table_html = None
    download_link = None

    if request.method == 'POST':
        # Check if there's a file in the request
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)

        if file:
            # Save input file and define output path
            file_path = os.path.join(UPLOAD_FOLDER, 'uploaded_file.xlsx')
            output_path = os.path.join(UPLOAD_FOLDER, 'processed_file.xlsx')
            file.save(file_path)

            # Process DataFrame
            modified_df = removeAddress(file_path)
            table_html = modified_df.to_html(classes="table table-bordered", index=False)

            # 2nd DataFrame
            second_df = sum_by_mapping(modified_df)
            second_table_html = second_df.to_html(classes="table table-striped", index=False)

            # Save two sheets in the processed file
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                modified_df.to_excel(writer, sheet_name='Processed Data', index=False)
                second_df.to_excel(writer, sheet_name='Summary Table', index=False)

            # Make the download link
            download_link = url_for('download_file')

    return render_template(
        'index.html',
        table=table_html,
        second_table=second_table_html,
        download_link=download_link
    )

@app.route('/download')
def download_file():
    file_path = os.path.join(UPLOAD_FOLDER, "processed_file.xlsx")
    if not os.path.exists(file_path):
        return "File wasn't available on site. Try reprocessing the file.", 404
    return send_file(file_path, as_attachment=True, download_name="processed_file.xlsx")

@app.route('/reset', methods=['GET'])
def reset():
    """
    Clear all files in the 'downloads' folder and reset the page.
    """
    # 1) Delete files in the 'downloads/' folder
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

    # 2) Redirect back to the main page (no tables shown)
    return redirect(url_for('index'))


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))  # Use Render's PORT environment variable
    app.run(debug=True, host="0.0.0.0", port=port)
