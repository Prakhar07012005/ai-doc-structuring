import os
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from utils.parser import parse_pdf
from utils.extractor import extract_rows, export_to_excel

app = Flask(__name__)
app.secret_key = "replace-this-with-a-secret-for-demo"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Direct run mode for local path testing (no upload)
def direct_run(input_path, output_path):
    text = parse_pdf(input_path)
    rows = extract_rows(text)
    export_to_excel(rows, output_path)
    print(f"Generated Excel at: {output_path}")

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            flash("Please choose a PDF file.")
            return redirect(url_for("index"))
        if not file.filename.lower().endswith(".pdf"):
            flash("Only PDF files are supported.")
            return redirect(url_for("index"))

        in_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(in_path)

        # Parse + Extract + Export
        text = parse_pdf(in_path)
        rows = extract_rows(text)

        out_path = os.path.join(UPLOAD_FOLDER, "Output.xlsx")
        export_to_excel(rows, out_path)

        return send_file(out_path, as_attachment=True)

    return render_template("upload.html")

if __name__ == "__main__":
    # Option A: run Flask app
    # app.run(host="0.0.0.0", port=5000, debug=True)

    # Option B: direct run with your absolute path (uncomment and run once)
    # input_pdf = "/Users/prakhar/Desktop/ai-doc-structuring/Data Input.pdf"
    # output_xlsx = "/Users/prakhar/Desktop/ai-doc-structuring/Output.xlsx"
    # direct_run(input_pdf, output_xlsx)

    # Default: run Flask
    app.run(host="0.0.0.0", port=5001, debug=True)
