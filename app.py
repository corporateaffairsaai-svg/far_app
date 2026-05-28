from flask import Flask, render_template, request, redirect, url_for, flash
from googleapiclient.discovery import build
from google.oauth2 import service_account
import json, os, tempfile

app = Flask(__name__)
app.secret_key = "far_app_secret_123"

SPREADSHEET_ID = "1XhDK0DmYhfx8RSosUCwFFwsVAUujXI0IWIC3XAjjVUM"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds_json = os.environ.get("GOOGLE_CREDENTIALS")
if creds_json:
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
    tmp.write(creds_json)
    tmp.close()
    KEY_FILE = tmp.name
else:
    KEY_FILE = "key.json"

ENTITIES = ["YT", "YBS", "BRAT", "DTCI", "AMA", "LCPL", "GCAI", "LIS"]

CATEGORY_FIELDS = {
    "LPT": {"label": "Laptop", "icon": "💻", "fields": ["Asset Tag","Category","Make","Model","Serial Number","Company Tag","Current User","Current Department","Allotted Date","Billing Entity","Current Operating Entity","Invoice Number","Invoice Date","Purchase Amount","Processor","RAM","Storage","Purchase Vendor","Current Location","Status","Warranty Expiry","Asset Condition"]},
    "DSK": {"label": "Desktop", "icon": "🖥️", "fields": ["Asset Tag","Category","Make","Model","Serial Number","Company Tag","Current User","Current Department","Allotted Date","Billing Entity","Current Operating Entity","Invoice Number","Invoice Date","Purchase Amount","Processor","RAM","Storage","Purchase Vendor","Current Location","Status","Warranty Expiry","Asset Condition"]},
    "MOB": {"label": "Mobile", "icon": "📱", "fields": ["Asset Tag","Category","Make","Model","Serial Number","IMEI 1","IMEI 2","Company Tag","Current User","Current Department","Invoice Date","Invoice Number","Purchase Amount","Billing Entity","RAM","Storage","Allotted Date","Purchase Vendor","Current Location","Status","Warranty Expiry"]},
    "PTR": {"label": "Printer", "icon": "🖨️", "fields": ["Asset Tag","Category","Make","Model","Serial Number","New Tag Number","Old Tag Number","Invoice Date","Billing Entity","Invoice Number","Purchase Amount","Current User","Current Department","Current Location","Status"]},
    "PRJ": {"label": "Projector", "icon": "📽️", "fields": ["Asset Tag","Category","Make","Model","Serial Number","New Tag Number","Old Tag Number","Invoice Date","Billing Entity","Invoice Number","Purchase Amount","Current User","Current Location","Status"]},
    "UPS": {"label": "UPS", "icon": "🔋", "fields": ["Asset Tag","Category","Make","Model","Serial Number","Tag Number","Invoice Number","Invoice Date","Purchase Amount","Current User","Billing Entity","Purchase Vendor","Current Location","Status"]},
    "FW":  {"label": "Firewall", "icon": "🔥", "fields": ["Asset Tag","Category","Make","Model","Serial Number","New Tag Number","Old Tag Number","Invoice Date","Billing Entity","Invoice Number","Purchase Amount","Current User","Current Department","Current Location","Status"]},
    "SW":  {"label": "Switch & Rack", "icon": "🔌", "fields": ["Asset Tag","Category","Make","Model","Serial Number","New Tag Number","Old Tag Number","Invoice Date","Billing Entity","Invoice Number","Purchase Amount","Current User","Current Department","Current Location","Status"]},
}

SHEETS = {
    "assets":     "Asset Master",
    "allocation": "Allocation",
    "transfer":   "Transfer Log",
    "maintenance":"Maintenance",
    "purchase":   "Purchase Register",
    "disposal":   "Disposal",
    "software":   "Software Tracking",
    "employees":  "Employee Master",
    "vendors":    "Vendor Master",
    "entities":   "Entity Master",
    "locations":  "Location Master",
    "categories": "Category Master",
    "departments":"Department Master",
}

def get_service():
    creds = service_account.Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)

def read_sheet(sheet_name, range_end="AO500"):
    try:
        service = get_service()
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1:{range_end}"
        ).execute()
        values = result.get("values", [])
        if not values or len(values) < 1:
            return [], []
        headers = values[0]
        rows = []
        for row in values[1:]:
            while len(row) < len(headers):
                row.append("")
            rows.append(dict(zip(headers, row)))
        return headers, rows
    except Exception as e:
        print(f"Error reading {sheet_name}: {e}")
        return [], []

def parse_amount(val):
    try:
        clean = str(val).replace("₹","").replace(",","").replace(" ","").replace("\xa0","").strip()
        return float(clean) if clean else 0
    except:
        return 0

def append_row(sheet_name, row_data):
    service = get_service()
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": [row_data]}
    ).execute()

def update_row(sheet_name, row_index, row_data):
    service = get_service()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A{row_index + 2}",
        valueInputOption="USER_ENTERED",
        body={"values": [row_data]}
    ).execute()

def delete_row_from_sheet(sheet_name, row_index):
    service = get_service()
    sheet_meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = None
    for s in sheet_meta["sheets"]:
        if s["properties"]["title"] == sheet_name:
            sheet_id = s["properties"]["sheetId"]
            break
    if sheet_id is None:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"deleteDimension": {"range": {
            "sheetId": sheet_id, "dimension": "ROWS",
            "startIndex": row_index + 1, "endIndex": row_index + 2
        }}}]}
    ).execute()

# ── DASHBOARD ──────────────────────────────────────────
@app.route("/")
def dashboard():
    _, all_assets   = read_sheet("Asset Master")
    _, all_employees = read_sheet("Employee Master")

    entity_stats = {}
    grand_total_cost = 0

    for entity in ENTITIES:
        assets = [a for a in all_assets if
                  a.get("Billing Entity","").strip() == entity or
                  a.get("Current Operating Entity","").strip() == entity]
        by_cat = {}
        total_cost = 0
        for a in assets:
            cat = a.get("Category","?").strip()
            by_cat[cat] = by_cat.get(cat, 0) + 1
            total_cost += parse_amount(a.get("Purchase Amount","0"))
        grand_total_cost += total_cost
        entity_stats[entity] = {
            "total": len(assets),
            "by_cat": by_cat,
            "total_cost": total_cost,
            "employees": len([e for e in all_employees if e.get("Operating Entity","").strip() == entity])
        }

    return render_template("dashboard.html",
        entity_stats=entity_stats,
        entities=ENTITIES,
        total_assets=len(all_assets),
        grand_total_cost=grand_total_cost,
        category_fields=CATEGORY_FIELDS
    )

# ── ENTITY DETAIL ──────────────────────────────────────
@app.route("/entity/<entity_code>")
def entity_detail(entity_code):
    _, all_assets = read_sheet("Asset Master")
    assets = [a for a in all_assets if
              a.get("Billing Entity","").strip() == entity_code or
              a.get("Current Operating Entity","").strip() == entity_code]

    by_category = {}
    total_cost = 0
    for a in assets:
        cat = a.get("Category","?").strip()
        if cat not in by_category:
            by_category[cat] = {"count":0,"cost":0,"assets":[]}
        by_category[cat]["count"] += 1
        by_category[cat]["assets"].append(a)
        cost = parse_amount(a.get("Purchase Amount","0"))
        by_category[cat]["cost"] += cost
        total_cost += cost

    return render_template("entity_detail.html",
        entity_code=entity_code,
        assets=assets,
        by_category=by_category,
        total_cost=total_cost,
        category_fields=CATEGORY_FIELDS
    )

# ── CATEGORY ASSETS ────────────────────────────────────
@app.route("/entity/<entity_code>/category/<cat_code>")
def category_assets(entity_code, cat_code):
    _, all_assets = read_sheet("Asset Master")
    assets = [a for a in all_assets if
        (a.get("Billing Entity","").strip() == entity_code or
         a.get("Current Operating Entity","").strip() == entity_code) and
        a.get("Category","").strip() == cat_code]
    cat_info = CATEGORY_FIELDS.get(cat_code, {"label":cat_code,"icon":"📦","fields":[]})
    search = request.args.get("q","").lower()
    if search:
        assets = [a for a in assets if any(search in str(v).lower() for v in a.values())]
    total_cost = sum(parse_amount(a.get("Purchase Amount","0")) for a in assets)
    return render_template("category_assets.html",
        entity_code=entity_code, cat_code=cat_code,
        cat_info=cat_info, assets=assets,
        search=search, total_cost=total_cost
    )

# ── ADD ASSET (smart category form) ───────────────────
@app.route("/add-asset/<cat_code>", methods=["GET","POST"])
def add_asset(cat_code):
    cat_info = CATEGORY_FIELDS.get(cat_code, {"label":cat_code,"icon":"📦","fields":[]})
    if request.method == "POST":
        row_data = [request.form.get(f,"") for f in cat_info["fields"]]
        append_row("Asset Master", row_data)
        flash(f"{cat_info['label']} added to Google Sheets!", "success")
        return redirect(url_for("dashboard"))
    return render_template("add_asset.html",
        cat_code=cat_code, cat_info=cat_info,
        entities=ENTITIES, category_fields=CATEGORY_FIELDS
    )

# ── GENERIC TABLE VIEW ─────────────────────────────────
@app.route("/table/<key>")
def table_view(key):
    if key not in SHEETS:
        return redirect(url_for("dashboard"))
    sheet_name = SHEETS[key]
    search = request.args.get("q","").lower()
    headers, rows = read_sheet(sheet_name)
    if search:
        rows = [r for r in rows if any(search in str(v).lower() for v in r.values())]
    return render_template("table.html",
        key=key, sheet_name=sheet_name,
        headers=headers, rows=rows,
        search=search, enumerate=enumerate
    )

# ── ADD ROW (generic for all tables) ──────────────────
@app.route("/add/<key>", methods=["GET","POST"])
def add_row(key):
    if key not in SHEETS:
        return redirect(url_for("dashboard"))
    sheet_name = SHEETS[key]
    headers, rows = read_sheet(sheet_name)
    if request.method == "POST":
        row_data = [request.form.get(h,"") for h in headers]
        append_row(sheet_name, row_data)
        flash(f"Row added to {sheet_name}!", "success")
        return redirect(url_for("table_view", key=key))
    return render_template("form.html",
        key=key, sheet_name=sheet_name,
        headers=headers, row=None, mode="add"
    )

# ── EDIT ROW ───────────────────────────────────────────
@app.route("/edit/<key>/<int:row_index>", methods=["GET","POST"])
def edit_row(key, row_index):
    if key not in SHEETS:
        return redirect(url_for("dashboard"))
    sheet_name = SHEETS[key]
    headers, rows = read_sheet(sheet_name)
    if row_index >= len(rows):
        return redirect(url_for("table_view", key=key))
    row = rows[row_index]
    if request.method == "POST":
        row_data = [request.form.get(h,"") for h in headers]
        update_row(sheet_name, row_index, row_data)
        flash("Updated successfully!", "success")
        return redirect(url_for("table_view", key=key))
    return render_template("form.html",
        key=key, sheet_name=sheet_name,
        headers=headers, row=row, mode="edit", row_index=row_index
    )

# ── DELETE ROW ─────────────────────────────────────────
@app.route("/delete/<key>/<int:row_index>", methods=["POST"])
def delete_row_route(key, row_index):
    if key not in SHEETS:
        return redirect(url_for("dashboard"))
    delete_row_from_sheet(SHEETS[key], row_index)
    flash("Deleted!", "success")
    return redirect(url_for("table_view", key=key))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
