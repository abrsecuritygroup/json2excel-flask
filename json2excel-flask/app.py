import json
import pandas as pd
import openpyxl
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
from openpyxl.styles import PatternFill, Font

app = Flask(__name__)

# Set global variables. Allow only JSON files
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
ALLOWED_EXTENSIONS = {"json"}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)


def allowed_file(filename):
    """
    Checks if the file extension is allowed. Returns True or False
    """
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def json_to_dataframe(json_data):
    """
    Converts JSON data to a structured Pandas DF by iterating through the
    JSON data where each key represents a data source (breach_name) and its
    value (records) is a list of dictionaries containing breach details
    """

    all_records = []

    for breach_name, records in json_data.items():
        for record in records:
            record["Breach Name"] = breach_name
            all_records.append(record)

    df = pd.DataFrame(all_records)

    # Ensure "Breach Name" is the first column. This is a specific need for ABR
    columns = ["Breach Name"] + [col for col in df.columns if col != "Breach Name"]
    return df[columns]


def convert_json_to_excel(json_path, output_path):
    """
    Converts the JSON data into an Excel file.
    """
    with open(json_path, "r", encoding="utf-8") as file:
        json_data = json.load(file)

if __name__ == "__main__":
    app.run()

    # Convert the JSON data to a Pandas DF
    df = json_to_dataframe(json_data)

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Breach Data")

    except Exception as e:
        print(f"Error: {e}")

        workbook = writer.book
        worksheet = writer.sheets["Breach Data"]

        # Sets the length of each column to the length of the cell value
        for col in worksheet.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col) + 2
            worksheet.column_dimensions[col[0].column_letter].width = max_length

        # Sets the font and fill color of the first row
        header_font = Font(bold=True)
        fill_color = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

        for cell in worksheet[1]:  # First row (headers)
            cell.font = header_font
            cell.fill = fill_color


@app.route("/", methods=["GET", "POST"])
def upload_file():
    """
    Handles POST and GET requests
    """

    if request.method == "POST":

        if "file" not in request.files:
            return "No file part"
        file = request.files["file"]

        if file.filename == '':
            return "No selected file"

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            json_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(json_path)

            excel_filename = filename.rsplit(".", 1)[0] + ".xlsx"
            excel_path = os.path.join(PROCESSED_FOLDER, excel_filename)
            convert_json_to_excel(json_path, excel_path)

            return render_template("index.html", download_link=f"/download/{excel_filename}")
    return render_template("index.html", download_link=None)


@app.route("/download/<filename>")
def download_file(filename):
    """
    Downloads the file to your computer
    """
    return send_file(os.path.join(PROCESSED_FOLDER, filename), as_attachment=True)
