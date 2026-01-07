from flask import Flask, request, render_template_string
import os
import pdfplumber
import pandas as pd
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
EXCEL_FILE = "port_summary.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

html_page = """
<h2>Upload PDF for Port Summary ➜ Excel</h2>
<form method="POST" enctype="multipart/form-data">
  <input type="file" name="pdf_file" accept=".pdf" required>
  <button type="submit">Upload & Process</button>
</form>
<p>{{ message }}</p>
"""

@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    message = ""
    if request.method == "POST":
        file = request.files["pdf_file"]
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        records = []

        # --- Read PDF ---
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

        if not records:
            message = "No valid port data found in PDF."
            return render_template_string(html_page, message=message)

        # --- Remove duplicates within this PDF ---
        unique = {}
        for src, dst, port in records:
            if port not in unique:
                unique[port] = {
                    "Source IP": src,
                    "Destination IP": dst,
                    "Destination Port": port,
                    "Source PDF": file.filename,
                    "Uploaded On": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }

        new_df = pd.DataFrame(unique.values())

        output_path = os.path.join(OUTPUT_FOLDER, EXCEL_FILE)

        # --- Append to existing Excel ---
        if os.path.exists(output_path):
            old_df = pd.read_excel(output_path)

            combined_df = pd.concat([old_df, new_df], ignore_index=True)

            combined_df.drop_duplicates(
                subset=["Destination Port"],
                keep="first",
                inplace=True
            )
        else:
            combined_df = new_df

        # --- Save Excel ---
        combined_df.to_excel(output_path, index=False)

        message = f"Success! Excel updated: {output_path}"

    return render_template_string(html_page, message=message)


if __name__ == "__main__":
    app.run(debug=True)
