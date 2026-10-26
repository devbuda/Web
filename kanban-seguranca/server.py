#!/usr/bin/env python3
import json
import os
import sqlite3
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "kanban.db")
SEED_FILE = os.path.join(BASE_DIR, "initial_tasks.json")
PORT = int(os.environ.get("PORT", "80"))

db_lock = threading.Lock()
state_cond = threading.Condition()
state_version = int(time.time() * 1000)


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
                updated_at REAL
            )
        """)
        c.execute("SELECT COUNT(*) FROM tasks")
        count = c.fetchone()[0]
        if count == 0 and os.path.exists(SEED_FILE):
            with open(SEED_FILE, "r", encoding="utf-8") as f:
                seed_tasks = json.load(f)
            now = time.time()
            for t in seed_tasks:
                c.execute("""
                    INSERT INTO tasks (id, solution, pillar, activity, month, all_months, schedule, status, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    int(t["id"]),
                    t.get("solution", ""),
                    t.get("pillar", ""),
                    t.get("activity", ""),
                    t.get("month", ""),
                    json.dumps(t.get("all_months", []), ensure_ascii=False),
                    t.get("schedule", ""),
                    t.get("status", "todo"),
                    now
                ))
            conn.commit()
        conn.close()


def get_all_tasks():
    with db_lock:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT id, solution, pillar, activity, month, all_months, schedule, status FROM tasks ORDER BY id ASC")
        rows = c.fetchall()
        conn.close()
        result = []
        for r in rows:
            try:
                all_m = json.loads(r["all_months"]) if r["all_months"] else []
            except Exception:
                all_m = []
            result.append({
                "id": r["id"],
                "solution": r["solution"],
                "pillar": r["pillar"],
                "activity": r["activity"],
                "month": r["month"] or "",
                "all_months": all_m,
                "schedule": r["schedule"] or "",
                "status": r["status"] or "todo"
            })
        return result


def save_all_tasks(tasks_list):
    global state_version
    now = time.time()
    with db_lock:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        for t in tasks_list:
            c.execute("""
                INSERT INTO tasks (id, solution, pillar, activity, month, all_months, schedule, status, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    solution=excluded.solution,
                    pillar=excluded.pillar,
                    activity=excluded.activity,
                    month=excluded.month,
                    all_months=excluded.all_months,
                    schedule=excluded.schedule,
                    status=excluded.status,
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
