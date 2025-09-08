import datetime
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from calendar import month_name
import os
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from collections import defaultdict, OrderedDict
from datetime import datetime, timedelta
from flask import render_template

app = Flask(__name__)
app.secret_key = "your-secret-key"

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_path = os.environ.get("GOOGLE_CREDS_PATH", "credentials.json")
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, SCOPE)
client = gspread.authorize(creds)
SHEET_NAME = client.open("Leave Application (Responses)")

# Columns in Form Responses 1 (keep the exact header text as in the sheet)
COLUMNS = [
    "Timestamp", "Email Address", "Name", "Leave From Date", "Leave To Date",
    "Half Day / Full Day", "Type of leave", "Reason for leave", "Apprved/Rejected",
    "With pay / Without pay", "comment", "month"
]

TS_FORMAT = "%m/%d/%Y %H:%M:%S"  # unify timestamp format app-wide


# --- Email settings ---
#EMAIL_HOST = os.getenv("EMAIL_HOST", "smtp.gmail.com")
#EMAIL_PORT = int(os.getenv("EMAIL_PORT", "465"))  # 465=SSL, 587=STARTTLS
#EMAIL_USER = os.getenv("mgmt.akutoffice@gmail.com")              # e.g. youraddress@gmail.com
#EMAIL_PASS = os.getenv("EMAIL_PASS")              # App Password if Gmail with 2FA

#FROM_NAME  = os.getenv("FROM_NAME", "Leave Bot")

EMAIL_HOST = "smtp.gmail.com"
EMAIL_PORT = 465  # 465 for SSL, 587 for STARTTLS
EMAIL_USER = "mgmt.akutoffice@gmail.com"        # replace with your sender email
EMAIL_PASS = "yqgb fgik xtzb cqbh"
FROM_NAME  = "Leave Management System"
# Comma-separated extra recipients (optional): "boss@firm.com, hr@firm.com"
NOTIFY_EMAILS = [e.strip() for e in os.getenv("NOTIFY_EMAILS", "").split(",") if e.strip()]


def get_users_sheet():
    client = gspread.authorize(
        ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
    )
    return client.open(SHEET_NAME).worksheet("Users")

def get_user_row_by_email(email):
    ws = get_users_sheet()
    records = ws.get_all_records()
    email_l = (email or "").strip().lower()
    for r in records:
        if (r.get("Email", "").strip().lower() == email_l):
            return r
    return None

def get_sheet():
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME).worksheet("Form Responses 1")

def to_date(s: str):
    return datetime.strptime(s, "%m/%d/%Y").date()

def month_key(d):
    return d.strftime("%b-%y")  # e.g., 'Aug-25'

def daterange(d1, d2):
    cur = d1
    while cur <= d2:
        yield cur
        cur += timedelta(days=1)

# robust parser for dates coming from Google Sheet
DATE_FORMATS = [
    "%m/%d/%Y",   # 08/22/2025 (your /apply route writes this)
    "%m-%d-%Y",   # 08-22-2025
    "%d/%m/%Y",   # 22/08/2025
    "%d-%m-%Y",   # 22-08-2025
    "%Y-%m-%d",   # 2025-08-22
    "%d-%b-%Y",   # 22-Aug-2025
]

def try_parse_date(value):
    """Try multiple formats; return a date or None if not parseable."""
    if not value:
        return None
    s = str(value).strip()
    # strip time if present (e.g. "08/22/2025 00:00:00")
    if " " in s:
        s = s.split(" ")[0]
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def get_leaves_for_current_user():
    """Fetch all leave rows for the logged-in user from Google Sheet."""
    sheet = get_sheet()
    records = sheet.get_all_records()

    uemail = (current_user.id or "").strip().lower()
    uname  = (current_user.name or "").strip().lower()

    def mine(r):
        row_email = r.get("Email Address", "").strip().lower()
        row_name  = r.get("Name", "").strip().lower()
        return (row_email == uemail) or (uname and uname in row_name)

    my_leaves = [r for r in records if mine(r)]

    # Sort by timestamp (latest first)
    def ts(row):
        try:
            return datetime.strptime(row.get("Timestamp",""), TS_FORMAT)
        except Exception:
            return datetime.min

    my_leaves.sort(key=ts, reverse=True)
    return my_leaves


def per_month_days(row):
    """
    Split leave days across months.
    - Half-day only when single-day and 'Half' in session.
    - Otherwise count each calendar day (inclusive).
    Returns {} if dates are missing/unparseable.
    """
    d1 = try_parse_date(row.get("Leave From Date"))
    d2 = try_parse_date(row.get("Leave To Date"))
    if not d1 or not d2:
        return {}

    session = str(row.get("Half Day / Full Day", "")).strip().lower()
    if d1 == d2 and session.startswith("half"):
        return {month_key(d1): 0.5}

    out = defaultdict(float)
    cur = d1
    while cur <= d2:
        out[month_key(cur)] += 1.0
        cur += timedelta(days=1)
    return out


class User(UserMixin):
    def __init__(self, email, name, role):
        self.id = email
        self.name = name
        self.role = (role or "").strip().lower()  # normalize

@login_manager.user_loader
def load_user(user_id):
    r = get_user_row_by_email(user_id)
    if not r:
        return None
    role = (r.get("Role", "employee") or "").strip().lower()
    return User(r["Email"], r.get("Name", ""), role)

@app.route("/")
@login_required
def index():
    # Admins land on pending approvals (belt & suspenders)
    if (current_user.role or "").lower() == "admin":
        return redirect(url_for("admin_pending"))

    sheet = get_sheet()
    records = sheet.get_all_records()

    uemail = (current_user.id or "").strip().lower()
    uname  = (current_user.name or "").strip().lower()

    def mine(r):
        row_email = r.get("Email Address", "").strip().lower()
        row_name  = r.get("Name", "").strip().lower()
        return (row_email == uemail) or (uname and uname in row_name)  # partial name match

    my_leaves = [r for r in records if mine(r)]

    # Sort by application time (latest first)
    def ts(row):
        try:
            return datetime.strptime(row.get("Timestamp",""), TS_FORMAT)
        except Exception:
            return datetime.min

    my_leaves.sort(key=ts, reverse=True)
    return render_template("index_gform.html", leaves=my_leaves)


@app.route("/employee")
@login_required
def employee_dashboard():
    leaves_all = get_leaves_for_current_user()

    # Count how many are Approved (tolerant to case/whitespace)
    def is_approved(v): 
        return str(v or "").strip().lower().startswith("approved")

    approved_rows = sum(1 for l in leaves_all if is_approved(l.get("Apprved/Rejected")))
    # Keep only Approved for the summary (change to `leaves_all` to test without filter)
    leaves = [l for l in leaves_all if is_approved(l.get("Apprved/Rejected"))]

    total_rows = len(leaves_all)
    monthly_days = defaultdict(float)
    parsed_rows = 0
    skipped_rows = 0

    for row in leaves:
        chunk = per_month_days(row)  # {} if bad/missing dates
        if chunk:
            parsed_rows += 1
            for mk, dcount in chunk.items():
                monthly_days[mk] += dcount
        else:
            skipped_rows += 1

    # Sort months (latest first)
    def parse_mk(mk):  # 'Aug-25' -> datetime
        return datetime.strptime(mk, "%b-%y")
    monthly_days_sorted = OrderedDict(
        sorted(monthly_days.items(), key=lambda kv: parse_mk(kv[0]), reverse=True)
    )

    return render_template(
        "index.html",
        leaves=leaves,
        monthly_days=monthly_days_sorted,
        total_rows=total_rows,
        approved_rows=approved_rows,
        parsed_rows=parsed_rows,
        skipped_rows=skipped_rows,
    )



@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].strip()
        password = request.form["password"]
        r = get_user_row_by_email(email)
        if r and r.get("Password") == password:
            role = (r.get("Role", "employee") or "").strip().lower()
            user = User(r["Email"], r.get("Name",""), role)
            login_user(user)
            return redirect(url_for("admin_pending") if user.role == "admin" else url_for("index"))
        flash("Invalid credentials", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    logout_user()
    return redirect(url_for("login"))

def get_admin_emails_from_users_sheet():
    try:
        ws = get_users_sheet()
        records = ws.get_all_records()
        return [
            r.get("Email", "").strip()
            for r in records
            if (r.get("Role", "") or "").strip().lower() == "admin"
               and r.get("Email", "").strip()
        ]
    except Exception:
        return []
    
def get_notification_recipients():
    # Union of admins from Users sheet + any extra NOTIFY_EMAILS
    admins = get_admin_emails_from_users_sheet()
    # De-duplicate while preserving order
    seen, out = set(), []
    for e in admins + NOTIFY_EMAILS:
        if e and e not in seen:
            out.append(e); seen.add(e)
    return out

def send_email(to_list, subject, html_body):
    if not to_list:
        return

    msg = MIMEText(html_body, "html", "utf-8")
    msg["Subject"] = subject
    msg["From"] = formataddr((FROM_NAME, EMAIL_USER))
    msg["To"] = ", ".join(to_list)

    if EMAIL_PORT == 465:
        # SSL
        with smtplib.SMTP_SSL(EMAIL_HOST, EMAIL_PORT) as s:
            s.login(EMAIL_USER, EMAIL_PASS)
            s.send_message(msg)
    else:
        # STARTTLS (e.g., port 587)
        with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as s:
            s.starttls()
            s.login(EMAIL_USER, EMAIL_PASS)
            s.send_message(msg)


@app.route("/apply", methods=["GET", "POST"])
@login_required
def apply_leave():
    if request.method == "POST":
        sheet = get_sheet()
        now = datetime.now().strftime(TS_FORMAT)
        from_raw = request.form["from_date"]
        to_raw = request.form["to_date"]

        from_date = datetime.strptime(from_raw, "%Y-%m-%d").strftime("%m/%d/%Y")
        to_date = datetime.strptime(to_raw, "%Y-%m-%d").strftime("%m/%d/%Y")

        data = [
            now,
            current_user.id,
            current_user.name,
            from_date,
            to_date,
            request.form["session"],
            request.form["leave_type"],
            request.form["reason"],
            "",
            "",  # With pay / Without pay
            "",  # comment
            datetime.now().strftime("%b-%Y")
        ]
        sheet.append_row(data)

        try:
            recipients = get_notification_recipients()
            if recipients:
                subject = f"New Leave Request: {current_user.name} ({from_date} → {to_date})"
                # You can adjust BASE URL if needed; relative link is fine if emails are read in the same network/domain
                html = f"""
                <div style="font-family:Arial,sans-serif">
                  <p><b>{current_user.name}</b> submitted a new leave request.</p>
                  <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse">
                    <tr><th align="left">Applied On</th><td>{now}</td></tr>
                    <tr><th align="left">Employee</th><td>{current_user.name} &lt;{current_user.id}&gt;</td></tr>
                    <tr><th align="left">From</th><td>{from_date}</td></tr>
                    <tr><th align="left">To</th><td>{to_date}</td></tr>
                    <tr><th align="left">Session</th><td>{request.form['session']}</td></tr>
                    <tr><th align="left">Type</th><td>{request.form['leave_type']}</td></tr>
                    <tr><th align="left">Reason</th><td>{request.form['reason']}</td></tr>
                  </table>
                  <!--<p style="margin-top:12px">
                    Review &amp; take action: <a href="{url_for('admin_pending', _external=True)}">Pending Approvals</a>
                  </p>-->
                </div>
                """
                send_email(recipients, subject, html)
        except Exception as e:
            # Don’t block the user’s flow if email fails; log/flash if you like
            # print(f"Email notify failed: {e}")
            pass


        flash("Leave request submitted.", "success")
        return redirect(url_for("index"))
    return render_template("apply.html")

@app.route("/decision/<int:row>", methods=["POST"])
@login_required
def decision(row):
    if current_user.role != "admin":
        return "Unauthorized", 403

    comment = request.form.get("comment", "").strip()
    decision_val = request.form.get("decision", "Pending")

    sheet = get_sheet()
    # row is 0-based data index; +2 points to sheet row (header + 1)
    sheet.update_cell(row + 2, COLUMNS.index("Apprved/Rejected") + 1, decision_val)
    sheet.update_cell(row + 2, COLUMNS.index("comment") + 1, comment)

# Get current data so we can email the correct employee
    records = sheet.get_all_records()  # list of dicts, 0-based
    if row < 0 or row >= len(records):
        flash("Invalid row index.", "danger")
        return redirect(url_for("admin_pending"))

    rec = records[row]  
    emp_email = (rec.get("Email Address") or "").strip()
    emp_name  = (rec.get("Name") or "").strip()
    from_date = rec.get("Leave From Date", "")
    to_date   = rec.get("Leave To Date", "")
    session   = rec.get("Half Day / Full Day", "")
    ltype     = rec.get("Type of leave", "")
    reason    = rec.get("Reason for leave", "")
    applied   = rec.get("Timestamp", "")

    # Update status & comment in the sheet
    sheet.update_cell(row + 2, COLUMNS.index("Apprved/Rejected") + 1, decision_val)
    sheet.update_cell(row + 2, COLUMNS.index("comment") + 1, comment)

    # Email the employee about the decision (non-blocking try/except)
    try:
        if emp_email:
            subject = f"Your Leave Request ({from_date} → {to_date}) is {decision_val}"
            html = f"""
            <div style="font-family:Arial,sans-serif">
              <p>Hi {emp_name},</p>
              <p>Your leave request has been <b>{decision_val}</b>.</p>
              <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse">
                <tr><th align="left">Applied On</th><td>{applied}</td></tr>
                <tr><th align="left">From</th><td>{from_date}</td></tr>
                <tr><th align="left">To</th><td>{to_date}</td></tr>
                <tr><th align="left">Session</th><td>{session}</td></tr>
                <tr><th align="left">Type</th><td>{ltype}</td></tr>
                <tr><th align="left">Reason</th><td>{reason}</td></tr>
                <tr><th align="left">Admin Comment</th><td>{comment or '-'}</td></tr>
              </table>
              <!--<p style="margin-top:12px">
                You can view your requests here: <a href="{url_for('index_gform', _external=True)}">My Leave Requests</a>
              </p>-->
            </div>
            """
            send_email([emp_email], subject, html)
    except Exception:
        # avoid breaking the flow if email fails
        pass

    flash(f"Leave {decision_val.lower()} with comment: {comment}", "info")
    
     # --- Email the employee about the decision ---
    try:
        if emp_email:
            subject = f"Your Leave Request ({from_date} → {to_date}) is {decision_val}"
            html = f"""
            <p>Hello {emp_name},</p>
            <p>Your leave request has been <b>{decision_val}</b>.</p>
            <p><b>Type:</b> {ltype}<br>
               <b>From:</b> {from_date} → {to_date}<br>
               <b>Session:</b> {session}<br>
               <b>Reason:</b> {reason}</p>
            <p>Regards,<br>{FROM_NAME}</p>
            """
            send_email([emp_email], subject, html)
            flash(f"Leave {decision_val.lower()}; email sent to {emp_email}.", "success")
        else:
            flash("Leave updated, but employee email is missing.", "warning")
    except Exception as e:
        flash(f"Leave {decision_val.lower()}, but email failed: {e}", "warning")

    return redirect(url_for("admin_pending"))



@app.route("/admin/history")
@login_required
def admin_history():
    if current_user.role != "admin":
        return "Unauthorized", 403

    sheet = get_sheet()
    records = sheet.get_all_records()

    # Optional filters
    year_filter = request.args.get("year", "").strip()
    month_filter = request.args.get("month", "").strip().lower()
    name_filter = request.args.get("name", "").strip().lower()
    session_filter = request.args.get("session", "").strip()

    def match_filters(row):
        from_date = row.get("Leave From Date", "").strip()
        name = row.get("Name", "").lower()
        session = row.get("Half Day / Full Day", "")

        try:
            #mm, dd, yyyy = from_date.split("/")
            date_obj = try_parse_date(row.get("Leave From Date"))
            if not date_obj:
                return False
            month_name_str = month_name[date_obj.month]
            yyyy = str(date_obj.year)

            #month_name_str = month_name[int(mm)]
        except Exception:
            return False  # skip invalid dates

        return (
            (not year_filter or yyyy == year_filter) and
            (not month_filter or month_filter == month_name_str.lower()) and
            (not name_filter or name_filter in name) and
            (not session_filter or session_filter == session)
        )

    filtered = list(filter(match_filters, records))

    def parse_timestamp(r):
        try:
            return datetime.strptime(r.get("Timestamp",""), TS_FORMAT)
        except Exception:
            return datetime.min

    filtered.sort(key=parse_timestamp, reverse=True)
    return render_template("admin_history.html", leaves=filtered)

@app.route("/admin/pending")
@login_required
def admin_pending():
    if current_user.role != "admin":
        return "Unauthorized", 403

    sheet = get_sheet()
    records = sheet.get_all_records()

    pending = []
    for idx, r in enumerate(records):
        if r.get("Name", "").strip() != "" and r.get("Apprved/Rejected","") == "":
            r["row_number"] = idx  # 0-based; decision() adds +2 when updating
            pending.append(r)

    return render_template("admin_pending.html", leaves=pending)

@app.route("/admin")
@login_required
def admin_dashboard():
    if current_user.role != "admin":
        return "Unauthorized", 403
    # keep admins landing on Pending Approvals
    return redirect(url_for("admin_pending"))

#if __name__ == "__main__":
#    app.run(debug=True)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)



