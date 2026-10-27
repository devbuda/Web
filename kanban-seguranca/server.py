#!/usr/bin/env python3
import json
import os
import re
import sqlite3
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "kanban.db")
SEED_FILE = os.path.join(BASE_DIR, "initial_tasks.json")
PORT = int(os.environ.get("PORT", "8080"))

db_lock = threading.Lock()
state_cond = threading.Condition()
state_version = int(time.time() * 1000)

MONTH_MAP = {
    "janeiro": "01", "fevereiro": "02", "março": "03", "marco": "03",
    "abril": "04", "maio": "05", "junho": "06", "julho": "07",
    "agosto": "08", "setembro": "09", "outubro": "10", "novembro": "11", "dezembro": "12"
}
WEEK_DAY = {"1": "07", "2": "14", "3": "21", "4": "28", "5": "31"}


def normalize_schedule_ddmmyyyy(sched):
    if not sched or sched == "Sem data definida":
        return ""
    s = str(sched).strip()
    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        return s
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        y, m, d = s.split("-")
        return f"{d}/{m}/{y}"
    parts = [p.strip() for p in s.split("•")]
    last = parts[-1]
    m_match = re.match(r"^([A-Za-zçÇ]+)\s+de\s+(\d{4})", last)
    if not m_match:
        return ""
    m_name = m_match.group(1).lower()
    year = m_match.group(2)
    month = MONTH_MAP.get(m_name, "10")
    weeks = re.findall(r"(\d)ª\s*Semana", last)
    day = WEEK_DAY.get(weeks[-1], "15") if weeks else "15"
    return f"{day}/{month}/{year}"


def init_db():
    with db_lock:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY,
                solution TEXT NOT NULL,
                pillar TEXT NOT NULL,
                activity TEXT NOT NULL,
                month TEXT,
                all_months TEXT,
                schedule TEXT,
                status TEXT NOT NULL,
                assignee TEXT DEFAULT '',
                description TEXT DEFAULT '',
                subtasks TEXT DEFAULT '[]',
                comments TEXT DEFAULT '[]',
                updated_at REAL
            )
        """)
        # Migrate existing table if new columns do not exist yet
        c.execute("PRAGMA table_info(tasks)")
        cols = [row[1] for row in c.fetchall()]
        if "assignee" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN assignee TEXT DEFAULT ''")
        if "description" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN description TEXT DEFAULT ''")
        if "subtasks" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN subtasks TEXT DEFAULT '[]'")
        if "comments" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN comments TEXT DEFAULT '[]'")

        c.execute("SELECT COUNT(*) FROM tasks")
        count = c.fetchone()[0]
        if count == 0 and os.path.exists(SEED_FILE):
            with open(SEED_FILE, "r", encoding="utf-8") as f:
                seed_tasks = json.load(f)
            now = time.time()
            for t in seed_tasks:
                c.execute("""
                    INSERT INTO tasks (id, solution, pillar, activity, month, all_months, schedule, status, assignee, description, subtasks, comments, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    int(t["id"]),
                    t.get("solution", ""),
                    t.get("pillar", ""),
                    t.get("activity", ""),
                    t.get("month", ""),
                    json.dumps(t.get("all_months", []), ensure_ascii=False),
                    normalize_schedule_ddmmyyyy(t.get("schedule", "")),
                    t.get("status", "todo"),
                    t.get("assignee", ""),
                    t.get("description", ""),
                    json.dumps(t.get("subtasks", []), ensure_ascii=False),
                    json.dumps(t.get("comments", []), ensure_ascii=False),
                    now
                ))
        else:
            # Migrate any existing free-text schedules in DB to DD/MM/AAAA
            c.execute("SELECT id, schedule FROM tasks")
            for row_id, row_sched in c.fetchall():
                new_sched = normalize_schedule_ddmmyyyy(row_sched or "")
                if new_sched != (row_sched or ""):
                    c.execute("UPDATE tasks SET schedule = ? WHERE id = ?", (new_sched, row_id))
        conn.commit()
        conn.close()


def get_all_tasks():
    with db_lock:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT id, solution, pillar, activity, month, all_months, schedule, status, assignee, description, subtasks, comments FROM tasks ORDER BY id ASC")
        rows = c.fetchall()
        conn.close()
        result = []
        for r in rows:
            try:
                all_m = json.loads(r["all_months"]) if r["all_months"] else []
            except Exception:
                all_m = []
            try:
                subtasks = json.loads(r["subtasks"]) if r["subtasks"] else []
            except Exception:
                subtasks = []
            try:
                comments = json.loads(r["comments"]) if r["comments"] else []
            except Exception:
                comments = []
            result.append({
                "id": r["id"],
                "solution": r["solution"],
                "pillar": r["pillar"],
                "activity": r["activity"],
                "month": r["month"] or "",
                "all_months": all_m,
                "schedule": r["schedule"] or "",
                "status": r["status"] or "todo",
                "assignee": r["assignee"] or "",
                "description": r["description"] or "",
                "subtasks": subtasks,
                "comments": comments
            })
        return result


def save_all_tasks(tasks_list):
    global state_version
    now = time.time()
    with db_lock:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        keep_ids = [int(t["id"]) for t in tasks_list if "id" in t]
        if keep_ids:
            placeholders = ",".join("?" for _ in keep_ids)
            c.execute(f"DELETE FROM tasks WHERE id NOT IN ({placeholders})", keep_ids)
        else:
            c.execute("DELETE FROM tasks")
        for t in tasks_list:
            c.execute("""
                INSERT INTO tasks (id, solution, pillar, activity, month, all_months, schedule, status, assignee, description, subtasks, comments, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    solution=excluded.solution,
                    pillar=excluded.pillar,
                    activity=excluded.activity,
                    month=excluded.month,
                    all_months=excluded.all_months,
                    schedule=excluded.schedule,
                    status=excluded.status,
                    assignee=excluded.assignee,
                    description=excluded.description,
                    subtasks=excluded.subtasks,
                    comments=excluded.comments,
                    updated_at=excluded.updated_at
            """, (
                int(t["id"]),
                t.get("solution", ""),
                t.get("pillar", ""),
                t.get("activity", ""),
                t.get("month", ""),
                json.dumps(t.get("all_months", []), ensure_ascii=False),
                t.get("schedule", ""),
                t.get("status", "todo"),
                t.get("assignee", ""),
                t.get("description", ""),
                json.dumps(t.get("subtasks", []), ensure_ascii=False),
                json.dumps(t.get("comments", []), ensure_ascii=False),
                now
            ))
        conn.commit()
        conn.close()

    with state_cond:
        state_version = int(time.time() * 1000)
        state_cond.notify_all()


class KanbanHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def log_message(self, format, *args):
        pass

    def do_GET(self):
        if self.path.startswith("/api/events"):
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream; charset=utf-8")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Connection", "keep-alive")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            try:
                last_seen = state_version
                self.wfile.write(f"data: {json.dumps({'version': last_seen})}\n\n".encode("utf-8"))
                self.wfile.flush()
                while True:
                    with state_cond:
                        state_cond.wait(timeout=15.0)
                        current_ver = state_version
                    if current_ver != last_seen:
                        last_seen = current_ver
                        payload = json.dumps({"version": last_seen, "tasks": get_all_tasks()}, ensure_ascii=False)
                        self.wfile.write(f"data: {payload}\n\n".encode("utf-8"))
                        self.wfile.flush()
                    else:
                        self.wfile.write(b": keepalive\n\n")
                        self.wfile.flush()
            except Exception:
                return

        if self.path.startswith("/api/tasks"):
            try:
                tasks = get_all_tasks()
                payload = json.dumps({"version": state_version, "tasks": tasks}, ensure_ascii=False)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(payload.encode("utf-8"))
                return
            except Exception as e:
                self.send_error(500, str(e))
                return

        return super().do_GET()

    def do_POST(self):
        if self.path.startswith("/api/tasks"):
            try:
                length = int(self.headers.get("Content-Length", 0))
                raw_body = self.rfile.read(length).decode("utf-8")
                parsed = json.loads(raw_body)
                if isinstance(parsed, dict) and "tasks" in parsed:
                    tasks_list = parsed["tasks"]
                elif isinstance(parsed, list):
                    tasks_list = parsed
                else:
                    self.send_error(400, "Expected JSON array or {tasks: [...]}")
                    return
                save_all_tasks(tasks_list)
                resp = json.dumps({"ok": True, "version": state_version})
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(resp.encode("utf-8"))
            except Exception as e:
                self.send_error(500, str(e))
            return
        self.send_error(404, "Not found")


if __name__ == "__main__":
    init_db()
    server = ThreadingHTTPServer(("0.0.0.0", PORT), KanbanHandler)
    print(f"Kanban Web Server + SQLite + Real-time SSE rodando em http://0.0.0.0:{PORT}")
    server.serve_forever()
