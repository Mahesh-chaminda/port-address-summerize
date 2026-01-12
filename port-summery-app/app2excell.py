from flask import Flask, request, render_template_string
import os
import pandas as pd

app = Flask(__name__)

# -------------------------------
# Folders & Files
# -------------------------------
UPLOAD_FOLDER = "excel_uploads"
OUTPUT_FOLDER = "outputs"
FINAL_FILE = "port_summary_final.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# -------------------------------
# HTML Page
# -------------------------------
html_page = """
<h2>📊 Excel Port Summary Tool</h2>

<h3>1️⃣ Upload Excel Files</h3>
<form method="POST" enctype="multipart/form-data">
  <input type="file" name="excel_files" multiple accept=".xlsx" required>
  <button type="submit">Upload & Update Final Excel</button>
</form>

<p>{{ message }}</p>
"""

# -------------------------------
# Helper
# -------------------------------
def clean_port(port):
    """Check if port is valid (1-65535). If not, return 'ALERT'."""
    try:
        port_int = int(str(port).strip())
        if 1 <= port_int <= 65535:
            return port_int
        else:
            return "ALERT"
    except:
        return "ALERT"

# -------------------------------
# Upload & Update Final Excel
# -------------------------------
@app.route("/", methods=["GET", "POST"])
def upload_excel():
    message = ""
    if request.method == "POST":
        files = request.files.getlist("excel_files")
        if not files:
            return render_template_string(html_page, message="No files selected.")

        # Combine all uploaded Excel files
        all_records = []
        for file in files:
            try:
                df = pd.read_excel(file)
                required_cols = ["Source Type", "Destination IP", "Destination Port"]
                if not all(col in df.columns for col in required_cols):
                    continue
                df = df[required_cols]
                df["Destination IP"] = df["Destination IP"].astype(str).str.strip()
                df["Source Type"] = df["Source Type"].astype(str).str.strip()
                df["Destination Port"] = df["Destination Port"].apply(clean_port)
                all_records.append(df)
            except Exception as e:
                print(f"Error reading {file.filename}: {e}")

        if not all_records:
            return render_template_string(html_page, message="No valid Excel data found.")

        new_df = pd.concat(all_records, ignore_index=True)
        new_df = new_df.drop_duplicates(subset=["Destination IP", "Destination Port"], keep="first")

        final_path = os.path.join(OUTPUT_FOLDER, FINAL_FILE)

        # -------------------------------
        # Final Excel exists?
        # -------------------------------
        if os.path.exists(final_path):
            final_df = pd.read_excel(final_path)
            original_columns = final_df.columns.tolist()

            # Identify truly new rows
            existing_set = set(zip(final_df["Destination IP"], final_df["Destination Port"]))
            new_unique_df = new_df[~new_df.apply(lambda row: (row["Destination IP"], row["Destination Port"]) in existing_set, axis=1)]

            if new_unique_df.empty:
                message = "ℹ️ No new rows to add. Final Excel unchanged."
            else:
                # Append new rows at the bottom WITHOUT sorting old rows
                combined = pd.concat([final_df, new_unique_df], ignore_index=True)
                combined = combined.drop_duplicates(subset=["Destination IP", "Destination Port"], keep="first")
                combined = combined.reindex(columns=original_columns)  # preserve column order
                combined.to_excel(final_path, index=False)
                message = f"✅ {len(new_unique_df)} new row(s) added to final Excel."
        else:
            # First upload: save everything as final Excel
            original_columns = new_df.columns.tolist()
            new_df = new_df.reindex(columns=original_columns)
            new_df.to_excel(final_path, index=False)
            message = f"✅ Final Excel created with {len(new_df)} rows."

    return render_template_string(html_page, message=message)

# -------------------------------
# Run App
# -------------------------------
if __name__ == "__main__":
    app.run(debug=True)
