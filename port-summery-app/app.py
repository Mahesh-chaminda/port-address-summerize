from flask import Flask, request, render_template_string
import os
import pdfplumber
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

html_page = """
<h2>Upload PDF for Port Summary ➜ Excel</h2>
<form method="POST" enctype="multipart/form-data">
  <input type="file" name="pdf_file" accept=".pdf" required>
  <button type="submit">Upload & Process</button>
</form>
"""

@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    if request.method == "POST":
        file = request.files["pdf_file"]
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        records = []

        # Read tables/text from PDF
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()

                # If the PDF page contains tables
                if tables:
                    for table in tables:
                        for row in table:
                            if not row:
                                continue

                            # Expecting columns like:
                            # Source IP | Destination IP | Destination Port
                            src = (row[0] or "").strip()
                            dst = (row[1] or "").strip()
                            port = (row[2] or "").strip()

                            if port.isdigit():
                                records.append((src, dst, port))
                else:
                    # If not table, fall back to text parsing
                    lines = page.extract_text().splitlines()
                    for line in lines:
                        parts = line.split()
                        if len(parts) >= 3:
                            src, dst, port = parts[0], parts[1], parts[2]
                            if port.isdigit():
                                records.append((src, dst, port))

        # --- Remove duplicate ports (keep first occurrence) ---
        unique = {}
        for src, dst, port in records:
            if port not in unique:
                unique[port] = {"Source IP": src, "Destination IP": dst, "Destination Port": port}

        df = pd.DataFrame(unique.values())

        output_file = os.path.join(OUTPUT_FOLDER, "port_summary.xlsx")
        df.to_excel(output_file, index=False)

        return f"Done! Excel created at: {output_file}"

    return render_template_string(html_page)


if __name__ == "__main__":
    app.run(debug=True)
