from flask import Flask, request, render_template_string
import os
import pdfplumber
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
WORK_FILE = "port_summary_working.xlsx"
FINAL_FILE = "port_summary_final.xlsx"

# create folders if not exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- HTML page with small design ---
html_page = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Port Summary Tool</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f9f9f9;
            color: #333;
            padding: 30px;
        }
        h2 {
            color: #1f4e79;
            text-align: center;
            margin-bottom: 30px;
        }
        h3 {
            color: #1f78b4;
            margin-top: 20px;
            margin-bottom: 15px;
        }
        form {
            background-color: #fff;
            padding: 20px 25px;
            border-radius: 10px;
            box-shadow: 0px 3px 8px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            max-width: 450px;
        }
        input[type="file"] {
            margin-top: 10px;
        }
        button {
            background-color: #1f78b4;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 15px;
            margin-top: 10px;
        }
        button:hover {
            background-color: #155d8b;
        }
        p {
            background-color: #e7f3ff;
            border-left: 4px solid #1f78b4;
            padding: 10px 12px;
            border-radius: 5px;
            max-width: 450px;
        }
        .container {
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        @media (max-width: 500px) {
            form, p {
                width: 90%;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>Port Summary Tool 🌐</h2>

        <h3>1️⃣ Upload PDF (updates Excel automatically)</h3>
        <form method="POST" enctype="multipart/form-data">
            <input type="file" name="pdf_file" accept=".pdf" required><br>
            <button type="submit">Upload & Update Excel</button>
        </form>

        <h3>2️⃣ Create Final Excel (clean + no duplicates)</h3>
        <form action="/finalize" method="POST">
            <button type="submit">Generate Final Excel</button>
        </form>

        {% if message %}
            <p>{{ message }}</p>
        {% endif %}
    </div>
</body>
</html>
"""

# --- PDF parsing ---
def extract_records(filepath):
    records = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        if not row or len(row) < 3:
                            continue
                        src = (row[0] or "").strip()
                        dst = (row[1] or "").strip()
                        port = (row[2] or "").strip()
                        if port.isdigit():
                            records.append((src, dst, port))
            else:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.splitlines():
                    parts = line.split()
                    if len(parts) >= 3 and parts[2].isdigit():
                        records.append((parts[0], parts[1], parts[2]))
    return records

# --- Main upload route ---
@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    message = ""
    if request.method == "POST":
        file = request.files["pdf_file"]
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        records = extract_records(filepath)
        if not records:
            return render_template_string(html_page, message="No valid port data found in PDF.")

        new_df = pd.DataFrame(records, columns=["Source IP", "Destination IP", "Destination Port"])
        output_path = os.path.join(OUTPUT_FOLDER, WORK_FILE)

        if os.path.exists(output_path):
            old_df = pd.read_excel(output_path)
            combined = pd.concat([old_df, new_df], ignore_index=True)
            combined.drop_duplicates(subset=["Destination Port"], keep="first", inplace=True)
        else:
            combined = new_df.drop_duplicates(subset=["Destination Port"], keep="first")

        combined.to_excel(output_path, index=False)
        message = "Excel updated continuously ✔"

    return render_template_string(html_page, message=message)

# --- Finalize route ---
@app.route("/finalize", methods=["POST"])
def finalize_excel():
    work_path = os.path.join(OUTPUT_FOLDER, WORK_FILE)
    final_path = os.path.join(OUTPUT_FOLDER, FINAL_FILE)

    if not os.path.exists(work_path):
        return render_template_string(html_page, message="No working Excel found yet.")

    df = pd.read_excel(work_path)
    df.drop_duplicates(subset=["Destination Port"], keep="first", inplace=True)
    df.to_excel(final_path, index=False)

    return render_template_string(
        html_page,
        message=f"Final Excel created successfully ✔ Location: {final_path}"
    )

if __name__ == "__main__":
    app.run(debug=True)
