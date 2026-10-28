#!/usr/bin/env python3
import json
import os
import sqlite3
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "kanban_pessoal.db")
PORT = int(os.environ.get("PORT", "8001"))

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
                title TEXT NOT NULL,
                created_date TEXT DEFAULT '27/10/2026',
                date_label TEXT,
                description TEXT,
                subtasks TEXT,
                comments TEXT,
                status TEXT NOT NULL,
                assignee TEXT DEFAULT 'Victor Freire',
                updated_at REAL
            )
        """)
        c.execute("PRAGMA table_info(tasks)")
        cols = [row[1] for row in c.fetchall()]
        if "assignee" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN assignee TEXT DEFAULT 'Victor Freire'")
        if "created_date" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN created_date TEXT DEFAULT '27/10/2026'")
        c.execute("UPDATE tasks SET assignee = 'Victor Freire' WHERE assignee IS NULL OR TRIM(assignee) = ''")
        c.execute("UPDATE tasks SET created_date = '27/10/2026' WHERE created_date IS NULL OR TRIM(created_date) = ''")
        # Set deadline of current demands to last day of next month (30/11/2026) if still on default seed date or empty
        c.execute("UPDATE tasks SET date_label = '30/11/2026' WHERE date_label IS NULL OR TRIM(date_label) = '' OR date_label = '2026-10-27' OR date_label = '27/10/2026'")
        # Update Serlyo demand title to the improved corporate wording and ensure status is 'todo'
        improved_title = "Alinhar com Serlyo a emissão de procurações eletrônicas para centralizar e controlar o uso do Certificado Digital"
        c.execute("SELECT id FROM tasks WHERE LOWER(title) LIKE '%serlyo%'")
        serlyo_rows = c.fetchall()
        if serlyo_rows:
            c.execute("UPDATE tasks SET title = ?, status = 'todo', date_label = '30/11/2026' WHERE id = ?", (improved_title, serlyo_rows[0][0]))
        conn.commit()
        conn.close()


def get_all_tasks():
    with db_lock:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT id, title, created_date, date_label, description, subtasks, comments, status, assignee FROM tasks ORDER BY id ASC")
        rows = c.fetchall()
        conn.close()
        result = []
        for r in rows:
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
                "title": r["title"],
                "created_date": (r["created_date"] or "27/10/2026").strip() or "27/10/2026",
                "date": r["date_label"] or "",
                "description": r["description"] or "",
                "subtasks": subtasks,
                "comments": comments,
                "status": r["status"] or "todo",
                "assignee": (r["assignee"] or "Victor Freire").strip() or "Victor Freire"
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
                INSERT INTO tasks (id, title, created_date, date_label, description, subtasks, comments, status, assignee, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    title=excluded.title,
                    created_date=excluded.created_date,
                    date_label=excluded.date_label,
                    description=excluded.description,
                    subtasks=excluded.subtasks,
                    comments=excluded.comments,
                    status=excluded.status,
                    assignee=excluded.assignee,
                    updated_at=excluded.updated_at
            """, (
                int(t["id"]),
                t.get("title", ""),
                (t.get("created_date") or "27/10/2026").strip() or "27/10/2026",
                t.get("date", ""),
                t.get("description", ""),
                json.dumps(t.get("subtasks", []), ensure_ascii=False),
                json.dumps(t.get("comments", []), ensure_ascii=False),
                t.get("status", "todo"),
                (t.get("assignee") or "Victor Freire").strip() or "Victor Freire",
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
    print(f"Amazoncopy - Painel de demandas rodando em http://0.0.0.0:{PORT}")
    server.serve_forever()
