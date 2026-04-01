from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
import sqlite3
from datetime import datetime, date, timedelta
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = "supervisor-waste-hours-secret-key"

DB_NAME = "waste_hours.db"


def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_name TEXT NOT NULL UNIQUE,
        is_active INTEGER NOT NULL DEFAULT 1
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS structures (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        structure_name TEXT NOT NULL,
        is_active INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY (project_id) REFERENCES projects(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS waste_hour_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        badge_no TEXT NOT NULL,
        project_id INTEGER NOT NULL,
        structure_id INTEGER NOT NULL,
        people_count INTEGER NOT NULL,
        time_from TEXT NOT NULL,
        time_to TEXT NOT NULL,
        total_hours REAL NOT NULL,
        reason_disruption TEXT NOT NULL,
        responsibility_tag TEXT NOT NULL,
        impact_type TEXT NOT NULL,
        remarks TEXT,
        entry_date TEXT NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY (project_id) REFERENCES projects(id),
        FOREIGN KEY (structure_id) REFERENCES structures(id)
    )
    """)

    # Seed sample projects
    cur.execute("SELECT COUNT(*) as cnt FROM projects")
    if cur.fetchone()["cnt"] == 0:
        cur.executemany(
            "INSERT INTO projects (project_name) VALUES (?)",
            [
                ("Project 2820",),
                ("Project 2883",),
                ("Project 2745",),
            ]
        )

    # Seed sample structures
    cur.execute("SELECT COUNT(*) as cnt FROM structures")
    if cur.fetchone()["cnt"] == 0:
        cur.executemany(
            "INSERT INTO structures (project_id, structure_name) VALUES (?, ?)",
            [
                (1, "GPF"),
                (1, "PDP"),
                (2, "WHT-2"),
                (2, "WHT-3"),
                (3, "Deck Module"),
                (3, "Pipe Rack"),
            ]
        )

    conn.commit()
    conn.close()


def calculate_hours(time_from_str, time_to_str):
    fmt = "%H:%M"
    start = datetime.strptime(time_from_str, fmt)
    end = datetime.strptime(time_to_str, fmt)
    diff = end - start
    return diff.total_seconds() / 3600


@app.route("/")
def index():
    conn = get_db_connection()
    projects = conn.execute(
        "SELECT id, project_name FROM projects WHERE is_active = 1 ORDER BY project_name"
    ).fetchall()
    conn.close()

    today_str = datetime.now().strftime("%d-%b-%Y")

    return render_template("index.html", projects=projects, today_str=today_str)


@app.route("/get_structures/<int:project_id>")
def get_structures(project_id):
    conn = get_db_connection()
    structures = conn.execute(
        "SELECT id, structure_name FROM structures WHERE project_id = ? AND is_active = 1 ORDER BY structure_name",
        (project_id,)
    ).fetchall()
    conn.close()

    data = [{"id": s["id"], "name": s["structure_name"]} for s in structures]
    return jsonify(data)


@app.route("/submit", methods=["POST"])
def submit():
    badge_no = request.form.get("badge_no", "").strip()
    project_id = request.form.get("project_id", "").strip()
    structure_id = request.form.get("structure_id", "").strip()
    people_count = request.form.get("people_count", "").strip()
    time_from = request.form.get("time_from", "").strip()
    time_to = request.form.get("time_to", "").strip()
    reason_disruption = request.form.get("reason_disruption", "").strip()
    responsibility_tag = request.form.get("responsibility_tag", "").strip()
    impact_type = request.form.get("impact_type", "").strip()
    remarks = request.form.get("remarks", "").strip()

    entry_date = date.today().isoformat()
    created_at = datetime.now().isoformat(timespec="seconds")

    # Basic validations
    if not badge_no:
        flash("Supervisor Badge Number is mandatory.", "danger")
        return redirect(url_for("index"))

    if not badge_no.isdigit():
        flash("Supervisor Badge Number must be numeric only.", "danger")
        return redirect(url_for("index"))

    if not project_id:
        flash("Project is mandatory.", "danger")
        return redirect(url_for("index"))

    if not structure_id:
        flash("Structure is mandatory.", "danger")
        return redirect(url_for("index"))

    if not people_count:
        flash("Total number of people is mandatory.", "danger")
        return redirect(url_for("index"))

    if not people_count.isdigit():
        flash("Total number of people must be numeric only.", "danger")
        return redirect(url_for("index"))

    if not time_from or not time_to:
        flash("Both Timing From and Timing To are mandatory.", "danger")
        return redirect(url_for("index"))
    
    if not reason_disruption:
        flash("Reason for Disruption is mandatory.", "danger")
        return redirect(url_for("index"))

    if not responsibility_tag:
        flash("Responsibility Tag is mandatory.", "danger")
        return redirect(url_for("index"))

    if not impact_type:
        flash("Impact Type is mandatory.", "danger")
        return redirect(url_for("index"))

    try:
        total_hours = calculate_hours(time_from, time_to)
    except ValueError:
        flash("Invalid time format.", "danger")
        return redirect(url_for("index"))

    if total_hours <= 0:
        flash("Timing To must be greater than Timing From.", "danger")
        return redirect(url_for("index"))

    conn = get_db_connection()

    # Prevent exact same timing again for same supervisor on same day
    duplicate = conn.execute("""
        SELECT COUNT(*) as cnt
        FROM waste_hour_entries
        WHERE badge_no = ?
          AND entry_date = ?
          AND time_from = ?
          AND time_to = ?
    """, (badge_no, entry_date, time_from, time_to)).fetchone()["cnt"]

    if duplicate > 0:
        conn.close()
        flash("This same timing is already recorded for this supervisor today.", "danger")
        return redirect(url_for("index"))

    # Optional overlap prevention too
    existing_entries = conn.execute("""
        SELECT time_from, time_to
        FROM waste_hour_entries
        WHERE badge_no = ?
          AND entry_date = ?
    """, (badge_no, entry_date)).fetchall()

    new_start = datetime.strptime(time_from, "%H:%M")
    new_end = datetime.strptime(time_to, "%H:%M")

    for row in existing_entries:
        old_start = datetime.strptime(row["time_from"], "%H:%M")
        old_end = datetime.strptime(row["time_to"], "%H:%M")

        overlap = new_start < old_end and new_end > old_start
        if overlap:
            conn.close()
            flash("This timing overlaps with an existing record for the same supervisor today.", "danger")
            return redirect(url_for("index"))

    # Check daily total not more than 11 hours
    day_total = conn.execute("""
        SELECT COALESCE(SUM(total_hours), 0) as total
        FROM waste_hour_entries
        WHERE badge_no = ?
          AND entry_date = ?
    """, (badge_no, entry_date)).fetchone()["total"]

    if day_total + total_hours > 11:
        conn.close()
        flash("Daily total waste hours cannot exceed 11 hours for the same supervisor.", "danger")
        return redirect(url_for("index"))

    conn.execute("""
        INSERT INTO waste_hour_entries
        (badge_no, project_id, structure_id, people_count, time_from, time_to, total_hours,
         reason_disruption, responsibility_tag, impact_type, remarks, entry_date, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        badge_no,
        int(project_id),
        int(structure_id),
        int(people_count),
        time_from,
        time_to,
        round(total_hours, 2),
        reason_disruption,
        responsibility_tag,
        impact_type,
        remarks,
        entry_date,
        created_at
    ))

    conn.commit()
    conn.close()

    flash("Waste hours record saved successfully.", "success")
    return redirect(url_for("records"))



@app.route("/records")
def records():
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT w.id,
               w.badge_no,
               p.project_name,
               s.structure_name,
               w.people_count,
               w.time_from,
               w.time_to,
               w.total_hours,
               w.reason_disruption,
               w.responsibility_tag,
               w.impact_type,
               w.entry_date,
               w.created_at
        FROM waste_hour_entries w
        JOIN projects p ON w.project_id = p.id
        JOIN structures s ON w.structure_id = s.id
        ORDER BY w.entry_date DESC, w.created_at DESC
    """).fetchall()
    conn.close()
    return render_template("records.html", rows=rows)

@app.route("/download_excel")
def download_excel():
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT w.id,
               w.badge_no,
               p.project_name,
               s.structure_name,
               w.people_count,
               w.time_from,
               w.time_to,
               w.total_hours,
               w.reason_disruption,
               w.responsibility_tag,
               w.impact_type,
               w.entry_date,
               w.created_at
        FROM waste_hour_entries w
        JOIN projects p ON w.project_id = p.id
        JOIN structures s ON w.structure_id = s.id
        ORDER BY w.entry_date DESC, w.created_at DESC
    """).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Waste Hours Records"

    headers = [
        "ID",
        "Badge No",
        "Project",
        "Structure",
        "People Count",
        "Time From",
        "Time To",
        "Total Hours",
        "Reason for Disruption",
        "Responsibility Tag",
        "Impact Type",
        "Entry Date",
        "Created At"
    ]
    ws.append(headers)

    for row in rows:
        ws.append([
            row["id"],
            row["badge_no"],
            row["project_name"],
            row["structure_name"],
            row["people_count"],
            row["time_from"],
            row["time_to"],
            row["total_hours"],
            row["reason_disruption"],
            row["responsibility_tag"],
            row["impact_type"],
            row["entry_date"],
            row["created_at"]
        ])

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            cell_value = str(cell.value) if cell.value is not None else ""
            if len(cell_value) > max_length:
                max_length = len(cell_value)
        ws.column_dimensions[column_letter].width = max_length + 2

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Supervisor_Waste_Hours_{date.today().strftime('%Y%m%d')}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    init_db()
    app.run(debug=True)