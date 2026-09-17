from flask import Flask, request, jsonify, send_file
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from datetime import datetime
import tempfile
import os
import re
import uuid
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote
import json

app = Flask(__name__)


# ============================================================
# WEBSITE DESIGN SETTINGS
# ============================================================

SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "settings.json"
)

DEFAULT_SETTINGS = {
    "theme": "light",
    "primary_color": "#1F5C9F",
    "primary_hover": "#174A82",
    "accent_color": "#1F5C9F",
    "background_color": "#F5F7F9",
    "panel_color": "#FFFFFF",
    "text_color": "#202833",
    "muted_text_color": "#667085",
    "border_color": "#D7DCE2",
    "logo_width": 180,
    "border_radius": 8,
    "button_radius": 8
}


def load_settings():
    """Load saved website design settings, using defaults for missing values."""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)

            settings = DEFAULT_SETTINGS.copy()
            if isinstance(saved, dict):
                settings.update(saved)
            return settings

    except (OSError, json.JSONDecodeError) as e:
        print(f"Settings load error: {e}")

    return DEFAULT_SETTINGS.copy()


def save_settings(settings):
    """Save website design settings permanently to settings.json."""
    clean = DEFAULT_SETTINGS.copy()

    if isinstance(settings, dict):
        clean.update(settings)

    # Keep numeric design values within sensible ranges.
    try:
        clean["logo_width"] = max(80, min(500, int(clean["logo_width"])))
    except (TypeError, ValueError):
        clean["logo_width"] = DEFAULT_SETTINGS["logo_width"]

    try:
        clean["border_radius"] = max(0, min(30, int(clean["border_radius"])))
    except (TypeError, ValueError):
        clean["border_radius"] = DEFAULT_SETTINGS["border_radius"]

    try:
        clean["button_radius"] = max(0, min(30, int(clean["button_radius"])))
    except (TypeError, ValueError):
        clean["button_radius"] = DEFAULT_SETTINGS["button_radius"]

    # Basic color validation. Invalid values fall back to defaults.
    color_keys = [
        "primary_color",
        "primary_hover",
        "accent_color",
        "background_color",
        "panel_color",
        "text_color",
        "muted_text_color",
        "border_color"
    ]

    hex_color = re.compile(r"^#[0-9A-Fa-f]{6}$")

    for key in color_keys:
        value = str(clean.get(key, "")).strip()
        if not hex_color.fullmatch(value):
            clean[key] = DEFAULT_SETTINGS[key]

    temp_file = SETTINGS_FILE + ".tmp"

    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=4)

    os.replace(temp_file, SETTINGS_FILE)

    return clean


def ensure_settings_file():
    """Create settings.json automatically the first time the app starts."""
    if not os.path.exists(SETTINGS_FILE):
        save_settings(DEFAULT_SETTINGS.copy())


ensure_settings_file()


@app.context_processor
def inject_design_settings():
    """Make saved design settings available to every HTML template."""
    return {"design_settings": load_settings()}


@app.route("/settings", methods=["GET", "POST"])
def settings():
    """In-site design settings editor."""
    if request.method == "POST":
        submitted = {
            "theme": request.form.get("theme", "light"),
            "primary_color": request.form.get("primary_color", DEFAULT_SETTINGS["primary_color"]),
            "primary_hover": request.form.get("primary_hover", DEFAULT_SETTINGS["primary_hover"]),
            "accent_color": request.form.get("accent_color", DEFAULT_SETTINGS["accent_color"]),
            "background_color": request.form.get("background_color", DEFAULT_SETTINGS["background_color"]),
            "panel_color": request.form.get("panel_color", DEFAULT_SETTINGS["panel_color"]),
            "text_color": request.form.get("text_color", DEFAULT_SETTINGS["text_color"]),
            "muted_text_color": request.form.get("muted_text_color", DEFAULT_SETTINGS["muted_text_color"]),
            "border_color": request.form.get("border_color", DEFAULT_SETTINGS["border_color"]),
            "logo_width": request.form.get("logo_width", DEFAULT_SETTINGS["logo_width"]),
            "border_radius": request.form.get("border_radius", DEFAULT_SETTINGS["border_radius"]),
            "button_radius": request.form.get("button_radius", DEFAULT_SETTINGS["button_radius"])
        }

        saved = save_settings(submitted)

        # Return the settings page with a success message.
        return render_settings_page(saved, True)

    return render_settings_page(load_settings(), False)


def render_settings_page(settings_data, saved=False):
    """Render the settings page without requiring a separate template file."""
    success_message = (
        '<div class="success">✓ Settings saved successfully.</div>'
        if saved else ""
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>C&C Website Settings</title>
    <style>
        * {{
            box-sizing: border-box;
        }}

        body {{
            margin: 0;
            font-family: Arial, sans-serif;
            background: {settings_data["background_color"]};
            color: {settings_data["text_color"]};
        }}

        .settings-page {{
            max-width: 900px;
            margin: 50px auto;
            padding: 0 24px;
        }}

        .header {{
            background: {settings_data["panel_color"]};
            border: 1px solid {settings_data["border_color"]};
            border-radius: {settings_data["border_radius"]}px;
            padding: 24px 28px;
            margin-bottom: 20px;
        }}

        h1 {{
            margin: 0 0 8px;
            font-size: 28px;
        }}

        .subtitle {{
            color: {settings_data["muted_text_color"]};
        }}

        .panel {{
            background: {settings_data["panel_color"]};
            border: 1px solid {settings_data["border_color"]};
            border-radius: {settings_data["border_radius"]}px;
            padding: 28px;
        }}

        .grid {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 20px;
        }}

        .field {{
            display: flex;
            flex-direction: column;
            gap: 7px;
        }}

        label {{
            font-weight: 600;
            font-size: 14px;
        }}

        input, select {{
            width: 100%;
            min-height: 44px;
            padding: 9px 12px;
            border: 1px solid {settings_data["border_color"]};
            border-radius: {settings_data["button_radius"]}px;
            background: {settings_data["panel_color"]};
            color: {settings_data["text_color"]};
            font: inherit;
        }}

        input[type="color"] {{
            padding: 4px;
            height: 44px;
        }}

        .actions {{
            margin-top: 26px;
            display: flex;
            gap: 10px;
        }}

        button, .back {{
            border: 0;
            border-radius: {settings_data["button_radius"]}px;
            padding: 12px 18px;
            font-weight: 700;
            font: inherit;
            cursor: pointer;
            text-decoration: none;
        }}

        button {{
            background: {settings_data["primary_color"]};
            color: white;
        }}

        button:hover {{
            background: {settings_data["primary_hover"]};
        }}

        .back {{
            background: {settings_data["background_color"]};
            border: 1px solid {settings_data["border_color"]};
            color: {settings_data["text_color"]};
        }}

        .success {{
            margin-bottom: 18px;
            padding: 12px 14px;
            border-radius: {settings_data["button_radius"]}px;
            background: #EAF7EE;
            border: 1px solid #B9DFC2;
            color: #216E39;
            font-weight: 600;
        }}

        @media (max-width: 700px) {{
            .grid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>
    <main class="settings-page">
        <div class="header">
            <h1>C&C Website Settings</h1>
            <div class="subtitle">
                Change the website design without editing code.
            </div>
        </div>

        {success_message}

        <form class="panel" method="POST">
            <div class="grid">
                <div class="field">
                    <label for="theme">Theme</label>
                    <select id="theme" name="theme">
                        <option value="light" {"selected" if settings_data["theme"] == "light" else ""}>Light</option>
                        <option value="dark" {"selected" if settings_data["theme"] == "dark" else ""}>Dark</option>
                    </select>
                </div>

                <div class="field">
                    <label for="logo_width">Logo Width (px)</label>
                    <input id="logo_width" name="logo_width" type="number"
                           min="80" max="500" value="{settings_data["logo_width"]}">
                </div>

                <div class="field">
                    <label for="primary_color">Primary Button Color</label>
                    <input id="primary_color" name="primary_color" type="color"
                           value="{settings_data["primary_color"]}">
                </div>

                <div class="field">
                    <label for="primary_hover">Button Hover Color</label>
                    <input id="primary_hover" name="primary_hover" type="color"
                           value="{settings_data["primary_hover"]}">
                </div>

                <div class="field">
                    <label for="accent_color">Accent Color</label>
                    <input id="accent_color" name="accent_color" type="color"
                           value="{settings_data["accent_color"]}">
                </div>

                <div class="field">
                    <label for="background_color">Page Background</label>
                    <input id="background_color" name="background_color" type="color"
                           value="{settings_data["background_color"]}">
                </div>

                <div class="field">
                    <label for="panel_color">Panel Background</label>
                    <input id="panel_color" name="panel_color" type="color"
                           value="{settings_data["panel_color"]}">
                </div>

                <div class="field">
                    <label for="text_color">Main Text Color</label>
                    <input id="text_color" name="text_color" type="color"
                           value="{settings_data["text_color"]}">
                </div>

                <div class="field">
                    <label for="muted_text_color">Secondary Text Color</label>
                    <input id="muted_text_color" name="muted_text_color" type="color"
                           value="{settings_data["muted_text_color"]}">
                </div>

                <div class="field">
                    <label for="border_color">Border Color</label>
                    <input id="border_color" name="border_color" type="color"
                           value="{settings_data["border_color"]}">
                </div>

                <div class="field">
                    <label for="border_radius">Panel Radius (px)</label>
                    <input id="border_radius" name="border_radius" type="number"
                           min="0" max="30" value="{settings_data["border_radius"]}">
                </div>

                <div class="field">
                    <label for="button_radius">Button Radius (px)</label>
                    <input id="button_radius" name="button_radius" type="number"
                           min="0" max="30" value="{settings_data["button_radius"]}">
                </div>
            </div>

            <div class="actions">
                <button type="submit">Save Settings</button>
                <a class="back" href="/">Back to Fetcher</a>
            </div>
        </form>
    </main>
</body>
</html>"""



# ============================================================
# CONTROLS & COMPONENTS
# ============================================================

WEBSITE = "https://controlsandcomponents.com"

SEARCH_URL = (
    "https://controlsandcomponents.com/api/algolia/search"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": WEBSITE,
    "Referer": WEBSITE + "/"
}

# Maximum number of SKUs accepted by the web app.
MAX_SKUS = 500

# A small worker pool speeds up large jobs without opening hundreds of
# simultaneous connections to the C&C website.
MAX_WORKERS = 5

# Each running job is kept here until its download is retrieved.
jobs = {}
jobs_lock = threading.Lock()


# ============================================================
# HOME PAGE
# ============================================================

@app.route("/")
def home():
    # Saved design settings are injected into the page.
    with open("index.html", "r", encoding="utf-8") as f:
        html = f.read()

    settings_data = load_settings()

    # Apply the settings directly to the existing HTML without
    # requiring the user to edit CSS in VS Code.
    css = f"""
    <style id="saved-design-settings">
        :root {{
            --cc-primary: {settings_data["primary_color"]};
            --cc-primary-hover: {settings_data["primary_hover"]};
            --cc-accent: {settings_data["accent_color"]};
            --cc-background: {settings_data["background_color"]};
            --cc-panel: {settings_data["panel_color"]};
            --cc-text: {settings_data["text_color"]};
            --cc-muted: {settings_data["muted_text_color"]};
            --cc-border: {settings_data["border_color"]};
            --cc-radius: {settings_data["border_radius"]}px;
            --cc-button-radius: {settings_data["button_radius"]}px;
        }}

        body {{
            background-color: var(--cc-background);
            color: var(--cc-text);
        }}

        {"body { background-color: #12161C !important; color: #F3F4F6 !important; } .panel { background-color: #1B2129 !important; border-color: #303846 !important; } .muted, .subtitle, .tagline { color: #AAB4C0 !important; }" if settings_data["theme"] == "dark" else ""}

        .company-logo {{
            width: {settings_data["logo_width"]}px !important;
            height: auto !important;
        }}

        .fetch-button,
        button.fetch-button {{
            background-color: var(--cc-primary) !important;
            border-radius: var(--cc-button-radius) !important;
        }}

        .fetch-button:hover,
        button.fetch-button:hover {{
            background-color: var(--cc-primary-hover) !important;
        }}

        .download {{
            background-color: var(--cc-primary) !important;
            border-radius: var(--cc-button-radius) !important;
        }}

        .download:hover {{
            background-color: var(--cc-primary-hover) !important;
        }}

        .panel {{
            background-color: var(--cc-panel);
            border-color: var(--cc-border);
            border-radius: var(--cc-radius);
        }}

        .clear-parts-button {{
            border-color: var(--cc-border) !important;
            border-radius: var(--cc-button-radius) !important;
        }}

        .clear-parts-button:hover {{
            border-color: var(--cc-primary) !important;
        }}
    </style>
    """

    if "</head>" in html:
        html = html.replace("</head>", css + "\n</head>", 1)
    else:
        html = css + html

    return html


# ============================================================
# START FETCH JOB
# ============================================================

@app.route("/start_fetch", methods=["POST"])
def start_fetch():
    part_numbers_text = request.form.get("part_numbers", "")
    excel_file = request.files.get("excel")

    if not part_numbers_text.strip():
        return jsonify({
            "error": "Please enter at least one part number."
        }), 400

    if not excel_file:
        return jsonify({
            "error": "Please upload an Excel file."
        }), 400

    # Read and de-duplicate part numbers, preserving order.
    part_numbers = []

    for line in part_numbers_text.splitlines():
        part = line.strip()

        if not part:
            continue

        if not any(
            part.upper() == x.upper()
            for x in part_numbers
        ):
            part_numbers.append(part)

    if not part_numbers:
        return jsonify({
            "error": "No valid part numbers found."
        }), 400

    if len(part_numbers) > MAX_SKUS:
        return jsonify({
            "error": f"Please enter {MAX_SKUS} or fewer part numbers."
        }), 400

    original_temp = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    )

    excel_file.save(original_temp.name)
    original_temp.close()

    job_id = uuid.uuid4().hex

    with jobs_lock:
        jobs[job_id] = {
            "status": "starting",
            "total": len(part_numbers),
            "processed": 0,
            "updated": 0,
            "not_found": 0,
            "review": 0,
            "results_by_index": {},
            "download_id": None,
            "error": None,
            "created": datetime.now().isoformat()
        }

    thread = threading.Thread(
        target=run_fetch_job,
        args=(job_id, part_numbers, original_temp.name),
        daemon=True
    )
    thread.start()

    return jsonify({
        "job_id": job_id,
        "total": len(part_numbers)
    })


# ============================================================
# JOB PROGRESS
# ============================================================

@app.route("/progress/<job_id>")
def get_progress(job_id):
    with jobs_lock:
        job = jobs.get(job_id)

        if not job:
            return jsonify({"error": "Job not found."}), 404

        # Copy only what the browser needs.
        partial_results = [
            job["results_by_index"][i]
            for i in sorted(job["results_by_index"])
        ]

        response = {
            "status": job["status"],
            "total": job["total"],
            "processed": job["processed"],
            "updated": job["updated"],
            "not_found": job["not_found"],
            "review": job["review"],
            "results": partial_results,
            "download_id": job["download_id"],
            "error": job["error"]
        }

    return jsonify(response)


# ============================================================
# BACKGROUND FETCH JOB
# ============================================================

def run_fetch_job(job_id, part_numbers, original_temp_path):
    try:
        with jobs_lock:
            jobs[job_id]["status"] = "fetching"

        results_by_index = {}

        # Each worker gets its own requests.Session. This keeps the
        # existing search logic intact while allowing several SKUs
        # to be processed at once.
        def process_one(index, part):
            session = requests.Session()
            session.headers.update(HEADERS)

            print(
                f"[{index + 1}/{len(part_numbers)}] "
                f"Searching C&C: {part}"
            )

            return index, part, search_product(session, part)

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = [
                executor.submit(process_one, index, part)
                for index, part in enumerate(part_numbers)
            ]

            for future in as_completed(futures):
                index, part, result = future.result()

                if result["status"] == "Updated":
                    status = "Updated"
                    description = result["description"]
                    product_url = result["url"]
                    updated = True
                else:
                    status = "Not Found"
                    description = ""
                    product_url = ""
                    updated = False

                results_by_index[index] = {
                    "part_number": part,
                    "description": description,
                    "status": status,
                    "website": product_url
                }

                with jobs_lock:
                    jobs[job_id]["results_by_index"] = results_by_index.copy()
                    jobs[job_id]["processed"] += 1

                    if updated:
                        jobs[job_id]["updated"] += 1
                    else:
                        jobs[job_id]["not_found"] += 1

        # --------------------------------------------------------
        # Build the updated workbook after all searches finish.
        # --------------------------------------------------------

        with jobs_lock:
            jobs[job_id]["status"] = "saving"

        workbook = load_workbook(original_temp_path)

        if "Products" in workbook.sheetnames:
            sheet = workbook["Products"]
        else:
            sheet = workbook.create_sheet("Products")

        workbook.active = workbook.sheetnames.index("Products")

        # Exact headers used by the original fetcher.
        sheet.cell(row=1, column=1).value = "Part Number"
        sheet.cell(row=1, column=2).value = "Description"
        sheet.cell(row=1, column=3).value = "Status"
        sheet.cell(row=1, column=4).value = "Website Match"
        sheet.cell(row=1, column=5).value = "Last Updated"

        # Clear previous A:E results.
        if sheet.max_row >= 2:
            for row in range(2, sheet.max_row + 1):
                for column in range(1, 6):
                    sheet.cell(
                        row=row,
                        column=column
                    ).value = None

                    sheet.cell(
                        row=row,
                        column=column
                    ).hyperlink = None

        # Write results in the original input order.
        for index, part in enumerate(part_numbers):
            row = index + 2
            result = results_by_index[index]

            sheet.cell(row=row, column=1).value = part

            if result["status"] == "Updated":
                sheet.cell(
                    row=row,
                    column=2
                ).value = result["description"]

                sheet.cell(
                    row=row,
                    column=3
                ).value = "Updated"

                website_cell = sheet.cell(
                    row=row,
                    column=4
                )
                website_cell.value = "View Product"
                website_cell.hyperlink = result["website"]

            else:
                sheet.cell(row=row, column=2).value = ""
                sheet.cell(row=row, column=3).value = "Not Found"
                sheet.cell(row=row, column=4).value = ""

            sheet.cell(
                row=row,
                column=5
            ).value = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )

        sheet.column_dimensions["A"].width = 20
        sheet.column_dimensions["B"].width = 55
        sheet.column_dimensions["C"].width = 18
        sheet.column_dimensions["D"].width = 20
        sheet.column_dimensions["E"].width = 22

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:E{len(part_numbers) + 1}"

        output_file = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )
        output_file.close()

        workbook.save(output_file.name)

        download_id = os.path.basename(output_file.name)

        with jobs_lock:
            jobs[job_id]["download_id"] = download_id
            jobs[job_id]["status"] = "complete"

        download_storage[download_id] = output_file.name

        print("")
        print("=" * 70)
        print("FETCH COMPLETE")
        print("=" * 70)
        print(f"Results saved: {len(part_numbers)}")
        print("")

    except Exception as e:
        print("")
        print("=" * 70)
        print("ERROR")
        print("=" * 70)
        print(str(e))
        print("")

        with jobs_lock:
            jobs[job_id]["status"] = "error"
            jobs[job_id]["error"] = str(e)

    finally:
        try:
            os.remove(original_temp_path)
        except Exception:
            pass


# ============================================================
# DOWNLOAD EXCEL
# ============================================================

download_storage = {}


@app.route("/download")
def download_file():
    file_id = request.args.get("file")

    if (
        not file_id
        or file_id not in download_storage
    ):
        return "File not found.", 404

    file_path = download_storage[file_id]

    return send_file(
        file_path,
        as_attachment=True,
        download_name="CC_Updated_Descriptions.xlsx"
    )




# ============================================================
# SEARCH C&C
# ============================================================

def search_product(
    session,
    part_number
):

    part_number = part_number.strip()

    if not part_number:

        return {
            "status": "Not Found",
            "description": "",
            "url": ""
        }

    # --------------------------------------------------------
    # Exact C&C API request
    # --------------------------------------------------------

    payload = {

        "query": part_number,

        "hitsPerPage": 8,

        "includeFacets": False

    }

    try:

        response = session.post(

            SEARCH_URL,

            json=payload,

            timeout=20

        )

        print(
            f"    Search API: "
            f"{response.status_code}"
        )

        if response.status_code != 200:

            return {
                "status": "Not Found",
                "description": "",
                "url": ""
            }

        data = response.json()

    except Exception as e:

        print(
            f"    Search error: {e}"
        )

        return {
            "status": "Not Found",
            "description": "",
            "url": ""
        }

    # --------------------------------------------------------
    # Get results
    # --------------------------------------------------------

    hits = data.get(
        "results",
        []
    )

    if not hits:

        return {
            "status": "Not Found",
            "description": "",
            "url": ""
        }

    # --------------------------------------------------------
    # Find exact SKU
    # --------------------------------------------------------

    exact_match = None

    for hit in hits:

        sku_data = hit.get(
            "sku",
            ""
        )

        if isinstance(
            sku_data,
            dict
        ):

            sku_value = str(
                sku_data.get(
                    "value",
                    ""
                )
            ).strip()

        else:

            sku_value = str(
                sku_data
            ).strip()

        clean_sku = clean_sku_value(
            sku_value
        )

        print(
            f"    SKU: "
            f"{sku_value} -> {clean_sku}"
        )

        if (
            clean_sku.upper()
            == part_number.upper()
        ):

            exact_match = hit

            break

    # --------------------------------------------------------
    # No exact match
    # --------------------------------------------------------

    if exact_match is None:

        print(
            f"    No exact SKU match: "
            f"{part_number}"
        )

        return {
            "status": "Not Found",
            "description": "",
            "url": ""
        }

    # --------------------------------------------------------
    # Get product description
    # --------------------------------------------------------

    description = str(
        exact_match.get(
            "name",
            ""
        )
    ).strip()

    if not description:

        return {
            "status": "Not Found",
            "description": "",
            "url": ""
        }

    # --------------------------------------------------------
    # Get category
    # --------------------------------------------------------

    category = get_value(
        exact_match.get(
            "category",
            ""
        )
    )

    # --------------------------------------------------------
    # Find product URL
    # --------------------------------------------------------

    product_url = find_product_url(
        session,
        part_number,
        category
    )

    # --------------------------------------------------------
    # If URL cannot be found, use the
    # C&C search URL as a fallback.
    #
    # The description itself came directly
    # from the exact C&C search result.
    # --------------------------------------------------------

    if not product_url:

        product_url = (
            f"{WEBSITE}/?s="
            f"{quote(part_number)}"
        )

    return {

        "status": "Updated",

        "description": description,

        "url": product_url

    }


# ============================================================
# CLEAN SKU
# ============================================================

def clean_sku_value(
    sku
):

    sku = str(
        sku
    ).strip()

    # Remove HTML tags
    sku = re.sub(
        r"<[^>]+>",
        "",
        sku
    )

    # C&C may return:
    #
    # cmASQF-6F/cm
    #
    # We need:
    #
    # ASQF-6F

    if sku.startswith(
        "cm"
    ):

        sku = sku[2:]

    if sku.endswith(
        "/cm"
    ):

        sku = sku[:-3]

    return sku.strip()


# ============================================================
# GET VALUE
# ============================================================

def get_value(
    value
):

    if isinstance(
        value,
        dict
    ):

        return str(
            value.get(
                "value",
                ""
            )
        ).strip()

    return str(
        value
    ).strip()


# ============================================================
# FIND PRODUCT URL
# ============================================================

def find_product_url(
    session,
    part_number,
    category
):

    if not category:

        return None

    category_slug = slugify(
        category
    )

    category_url = (
        f"{WEBSITE}/product-categories/"
        f"{category_slug}"
    )

    try:

        response = session.get(

            category_url,

            timeout=20

        )

        if response.status_code != 200:

            return None

        soup = BeautifulSoup(

            response.text,

            "html.parser"
        )

        for link in soup.find_all(
            "a",
            href=True
        ):

            href = link.get(
                "href",
                ""
            )

            text = link.get_text(
                " ",
                strip=True
            )

            combined = (
                href
                + " "
                + text
            )

            if (
                part_number.lower()
                in combined.lower()
            ):

                if href.startswith("/"):

                    href = (
                        WEBSITE
                        + href
                    )

                if href.startswith(
                    "http"
                ):

                    print(
                        f"    Product URL: "
                        f"{href}"
                    )

                    return href

    except Exception as e:

        print(
            f"    URL lookup error: {e}"
        )

    return None


# ============================================================
# SLUGIFY
# ============================================================

def slugify(
    text
):

    text = str(
        text
    ).lower()

    text = text.replace(
        "&",
        "and"
    )

    text = text.replace(
        "'",
        ""
    )

    text = re.sub(
        r"[^a-z0-9]+",
        "-",
        text
    )

    text = re.sub(
        r"-+",
        "-",
        text
    )

    return text.strip("-")

# ============================================================
# START SERVER
# ============================================================

if __name__ == "__main__":
    print("")
    print("=" * 70)
    print("C&C PRODUCT DESCRIPTION FETCHER")
    print("=" * 70)
    print("")
    print("Website running at:")
    print("http://127.0.0.1:5000")
    print("Settings page: http://127.0.0.1:5000/settings")
    print("")
    print(f"Maximum SKUs per job: {MAX_SKUS}")
    print(f"Concurrent workers: {MAX_WORKERS}")
    print("")

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )
