from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from mailer import send_campaign
import os
import tempfile

app = Flask(__name__)
app.secret_key = "change-me"
app.config["UPLOAD_FOLDER"] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {"xls", "xlsx"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        subject = request.form.get("subject")
        template_base = request.form.get("template_base") or None
        sheet = request.form.get("sheet") or None
        send_at = request.form.get("send_at") or "now"
        account = request.form.get("account") or None
        language_column = request.form.get("language_column") or "language"
        dry_run = bool(request.form.get("dry_run"))

        file = request.files.get("leads")
        leads_path = None
        if file and file.filename and allowed_file(file.filename):
            fname = secure_filename(file.filename)
            path = os.path.join(app.config["UPLOAD_FOLDER"], fname)
            file.save(path)
            leads_path = path
        try:
            send_campaign(
                excel_path=leads_path,
                subject_line=subject,
                template_base=template_base,
                sheet_name=sheet,
                send_at=send_at,
                account=account,
                language_column=language_column,
                dry_run=dry_run,
            )
            flash("Campaign executed successfully!", "success")
        except Exception as e:
            flash(f"Error: {e}", "danger")
        return redirect(url_for("index"))

    return render_template("webform.html")


if __name__ == "__main__":
    app.run(debug=True)
