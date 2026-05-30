# -*- coding: utf-8 -*-
import os, re, sys, json, time, base64, zipfile, datetime
import threading, openpyxl
import urllib.request, urllib.error
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkinter.scrolledtext import ScrolledText
from tkinter.font import Font as TkFont, BOLD

if getattr(sys, 'frozen', False):
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(_base_dir, ".dog_config.json")

GITHUB_USER  = ""
GITHUB_REPO  = ""
GITHUB_TOKEN = ""
NUM_WORKERS  = 3


def load_config():
    global GITHUB_USER, GITHUB_REPO, GITHUB_TOKEN, NUM_WORKERS
    if os.path.exists(CONFIG_FILE):
        try:
            c = json.load(open(CONFIG_FILE))
            GITHUB_USER  = c.get("user", "")
            GITHUB_REPO  = c.get("repo", "")
            GITHUB_TOKEN = c.get("token", "")
            NUM_WORKERS  = int(c.get("workers", 3))
        except Exception:
            pass


def save_config():
    try:
        json.dump({
            "user": GITHUB_USER, "repo": GITHUB_REPO,
            "token": GITHUB_TOKEN, "workers": NUM_WORKERS
        }, open(CONFIG_FILE, "w"))
    except Exception:
        pass


def gh_api(method, path, body=None):
    url = f"https://api.github.com{path}"
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"token {GITHUB_TOKEN}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("Content-Type", "application/json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            content = resp.read()
            return json.loads(content) if content else {}
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"GitHub API {e.code}: {e.read().decode()}")


def excel_serial_to_date(serial):
    base = datetime.date(1899, 12, 30)
    return base + datetime.timedelta(days=int(serial))


def excel_to_json(path):
    rows = []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        date_val = row[0]
        dog_name = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if not dog_name:
            continue
        if isinstance(date_val, datetime.datetime):
            d = date_val.date()
        elif isinstance(date_val, datetime.date):
            d = date_val
        elif isinstance(date_val, (int, float)):
            d = excel_serial_to_date(date_val)
        else:
            continue
        rows.append({"date": d.strftime("%Y-%m-%d"), "dog": dog_name})
    wb.close()
    return rows


def upload_input_json(rows, log):
    log(f"  Uploading {len(rows)} rows to GitHub...")
    content = base64.b64encode(json.dumps(rows, indent=2).encode()).decode()
    url = f"/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/input.json"
    sha = None
    try:
        data = gh_api("GET", url)
        sha = data.get("sha")
    except Exception:
        pass
    body = {"message": "update input", "content": content, "branch": "main"}
    if sha:
        body["sha"] = sha
    gh_api("PUT", url, body)
    log("  Input uploaded.")


def get_workflow_id(log):
    """Find workflow by name — avoids hardcoded filename issues."""
    data = gh_api("GET", f"/repos/{GITHUB_USER}/{GITHUB_REPO}/actions/workflows")
    workflows = data.get("workflows", [])
    log(f"  Found {len(workflows)} workflow(s) in repo:")
    for wf in workflows:
        log(f"    - {wf.get('name')} ({wf.get('path')}) state={wf.get('state')}")
        name = wf.get("name", "").lower()
        path = wf.get("path", "").lower()
        state = wf.get("state", "")
        if state == "active" and ("dog" in name or "scrape" in name or "scrape.yml" in path):
            log(f"  Using workflow: {wf.get('name')} (id={wf['id']})")
            return wf["id"]
    # fallback to first active
    for wf in workflows:
        if wf.get("state") == "active":
            log(f"  Fallback workflow: {wf.get('name')} (id={wf['id']})")
            return wf["id"]
    raise RuntimeError(
        f"No active workflow_dispatch workflow found.\n"
        f"Make sure scrape.yml has 'on: workflow_dispatch' and is committed to main branch.\n"
        f"Workflows found: {[w.get('name') for w in workflows]}"
    )


def trigger_workflow(log):
    wf_id = get_workflow_id(log)
    gh_api("POST",
           f"/repos/{GITHUB_USER}/{GITHUB_REPO}/actions/workflows/{wf_id}/dispatches",
           {"ref": "main", "inputs": {
               "num_workers": str(NUM_WORKERS),
               "gist_id": "",
           }})
    log(f"  Workflow triggered (id={wf_id}, workers={NUM_WORKERS})")


def get_latest_run_id(log):
    log("  Waiting for run to register...")
    wf_id = get_workflow_id(log)
    for _ in range(25):
        time.sleep(4)
        try:
            data = gh_api("GET",
                f"/repos/{GITHUB_USER}/{GITHUB_REPO}/actions/workflows/{wf_id}/runs?per_page=1")
            runs = data.get("workflow_runs", [])
            if runs and runs[0].get("status") in ("queued", "in_progress", "completed"):
                return runs[0]["id"]
        except Exception:
            pass
    raise RuntimeError("Could not find workflow run — check GitHub Actions tab manually")


def wait_for_run(run_id, log, stop_event):
    url = f"/repos/{GITHUB_USER}/{GITHUB_REPO}/actions/runs/{run_id}"
    job_id = None
    last_step = ""
    start = time.time()

    def poll_steps():
        nonlocal job_id, last_step
        while not stop_event.is_set():
            try:
                if job_id is None:
                    data = gh_api("GET",
                        f"/repos/{GITHUB_USER}/{GITHUB_REPO}/actions/runs/{run_id}/jobs")
                    jobs = data.get("jobs", [])
                    if jobs:
                        job_id = jobs[0]["id"]
                if job_id:
                    jdata = gh_api("GET",
                        f"/repos/{GITHUB_USER}/{GITHUB_REPO}/actions/jobs/{job_id}")
                    for step in jdata.get("steps", []):
                        if step.get("status") == "in_progress":
                            name = step.get("name", "")
                            if name != last_step:
                                last_step = name
                                log(f"  ▶ {name}")
                            break
            except Exception:
                pass
            time.sleep(8)

    threading.Thread(target=poll_steps, daemon=True).start()

    while True:
        try:
            data = gh_api("GET", url)
            status     = data.get("status", "")
            conclusion = data.get("conclusion", "")
            if status == "completed":
                stop_event.set()
                return conclusion
            elapsed = int(time.time() - start)
            mins, secs = divmod(elapsed, 60)
            log(f"  ⏱ Running... {mins}m {secs:02d}s")
        except Exception as e:
            log(f"  Poll error: {e}")
        time.sleep(20)


def download_artifact(run_id, out_dir, log):
    data = gh_api("GET",
        f"/repos/{GITHUB_USER}/{GITHUB_REPO}/actions/runs/{run_id}/artifacts")
    artifacts = data.get("artifacts", [])
    if not artifacts:
        raise RuntimeError("No artifacts found")

    artifact_id = artifacts[0]["id"]
    size_kb = artifacts[0].get("size_in_bytes", 0) // 1024
    log(f"  Downloading output ({size_kb} KB)...")

    zip_url = (f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}"
               f"/actions/artifacts/{artifact_id}/zip")

    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, headers, newurl):
            return None

    req = urllib.request.Request(zip_url)
    req.add_header("Authorization", f"token {GITHUB_TOKEN}")
    req.add_header("Accept", "application/vnd.github+json")

    download_url = None
    try:
        opener = urllib.request.build_opener(NoRedirect())
        opener.open(req, timeout=30)
    except urllib.error.HTTPError as e:
        if e.code in (301, 302, 307, 308):
            download_url = e.headers.get("Location")

    zip_path = os.path.join(out_dir, "_dl.zip")
    if download_url:
        with urllib.request.urlopen(download_url, timeout=120) as resp:
            with open(zip_path, "wb") as f:
                f.write(resp.read())
    else:
        with urllib.request.urlopen(req, timeout=120) as resp:
            with open(zip_path, "wb") as f:
                f.write(resp.read())

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(out_dir)
        for n in z.namelist():
            if n.endswith(".xlsx") or n.endswith(".json"):
                log(f"  → {n}")
    os.remove(zip_path)


class App:
    def __init__(self, root):
        load_config()
        self.root    = root
        self.running = False

        root.title("Dog Race Scraper (Cloud)")
        root.geometry("640x620")
        root.configure(bg="#ffffff")
        root.resizable(False, False)

        green  = "#055030"
        bold14 = TkFont(family="Arial", size=14, weight=BOLD)
        bold11 = TkFont(family="Arial", size=11, weight=BOLD)
        small9 = TkFont(family="Arial", size=9)

        tk.Label(root, text="Dog Race Scraper  ☁  Cloud Edition",
                 font=bold14, bg="#ffffff", fg=green).place(x=12, y=8)

        tk.Label(root, text="User",    font=small9, bg="#ffffff", fg="#555").place(x=12,  y=42)
        tk.Label(root, text="Repo",    font=small9, bg="#ffffff", fg="#555").place(x=160, y=42)
        tk.Label(root, text="Token",   font=small9, bg="#ffffff", fg="#555").place(x=308, y=42)
        tk.Label(root, text="Workers", font=small9, bg="#ffffff", fg="#555").place(x=516, y=42)

        self.e_user = tk.Entry(root, font=bold11, width=13, bg="#f5f5f5", relief="sunken", bd=2)
        self.e_user.insert(0, GITHUB_USER)
        self.e_user.place(x=12, y=58)

        self.e_repo = tk.Entry(root, font=bold11, width=13, bg="#f5f5f5", relief="sunken", bd=2)
        self.e_repo.insert(0, GITHUB_REPO)
        self.e_repo.place(x=160, y=58)

        self.e_token = tk.Entry(root, font=bold11, width=17, show="*", bg="#f5f5f5", relief="sunken", bd=2)
        self.e_token.insert(0, GITHUB_TOKEN)
        self.e_token.place(x=308, y=58)

        self.e_workers = tk.Entry(root, font=bold11, width=4, bg="#f5f5f5", relief="sunken", bd=2)
        self.e_workers.insert(0, str(NUM_WORKERS))
        self.e_workers.place(x=516, y=58)

        tk.Button(root, text="Save Config", font=small9, command=self._save_config,
                  bg=green, fg="#fff").place(x=572, y=58)

        tk.Label(root, text="Input Excel:", font=bold11,
                 bg="#ffffff", fg=green).place(x=12, y=95)
        self.input_var = tk.StringVar()
        tk.Entry(root, textvariable=self.input_var, font=small9, width=56,
                 bg="#f5f5f5", relief="sunken", bd=2).place(x=110, y=97)
        tk.Button(root, text="Browse", font=small9, command=self._browse_input,
                  bg=green, fg="#fff", width=7).place(x=580, y=95)

        tk.Label(root, text="Save To:", font=bold11,
                 bg="#ffffff", fg=green).place(x=12, y=129)
        self.output_var = tk.StringVar()
        tk.Entry(root, textvariable=self.output_var, font=small9, width=56,
                 bg="#f5f5f5", relief="sunken", bd=2).place(x=110, y=131)
        tk.Button(root, text="Browse", font=small9, command=self._browse_output,
                  bg=green, fg="#fff", width=7).place(x=580, y=129)

        tk.Label(root,
                 text="☁  Runs on GitHub servers. 3-5 workers recommended, up to 10 supported.",
                 font=small9, bg="#ffffff", fg="#555").place(x=12, y=162)

        tk.Button(root, text="▶ Start Scraping", font=bold11, width=18,
                  command=self._start, bg=green, fg="#fff").place(x=220, y=185)

        tk.Label(root, text="Live Progress:", font=bold11,
                 bg="#ffffff", fg=green).place(x=12, y=222)

        self.log_box = ScrolledText(root, state="disabled",
                                    font=TkFont(family="Courier", size=8),
                                    bg="#0d1117", fg="#58a6ff")
        self.log_box.place(x=10, y=242, width=620, height=365)
        self.log_box.tag_configure("green",  foreground="#3fb950")
        self.log_box.tag_configure("red",    foreground="#f85149")
        self.log_box.tag_configure("yellow", foreground="#d29922")
        self.log_box.tag_configure("dim",    foreground="#8b949e")

    def _log(self, msg):
        self.log_box.configure(state="normal")
        if any(x in msg for x in ["✓", "Done", "written", "Merged", "uploaded", "triggered"]):
            tag = "green"
        elif any(x in msg for x in ["✗", "ERROR", "error", "failed", "Failed", "not found"]):
            tag = "red"
        elif any(x in msg for x in ["▶", "⏱", "===", "Trigger", "Upload", "Download", "Using workflow"]):
            tag = "yellow"
        else:
            tag = "dim"
        self.log_box.insert(tk.END, msg + "\n", tag)
        self.log_box.see(tk.END)
        self.log_box.configure(state="disabled")
        self.root.update_idletasks()

    def _save_config(self):
        global GITHUB_USER, GITHUB_REPO, GITHUB_TOKEN, NUM_WORKERS
        GITHUB_USER  = self.e_user.get().strip()
        GITHUB_REPO  = self.e_repo.get().strip()
        GITHUB_TOKEN = self.e_token.get().strip()
        try:
            NUM_WORKERS = int(self.e_workers.get().strip())
        except Exception:
            NUM_WORKERS = 3
        save_config()
        self._log("✓ Config saved.")

    def _browse_input(self):
        p = fd.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if p:
            self.input_var.set(p)

    def _browse_output(self):
        p = fd.askdirectory(title="Choose folder to save results")
        if p:
            self.output_var.set(p)

    def _validate(self):
        self._save_config()
        if not all([GITHUB_USER, GITHUB_REPO, GITHUB_TOKEN]):
            mb.showerror("Config Missing", "Please fill in GitHub User, Repo and Token.")
            return False
        if not self.input_var.get().strip():
            mb.showerror("No Input", "Please select an input Excel file.")
            return False
        if not self.output_var.get().strip():
            mb.showerror("No Output", "Please select an output folder.")
            return False
        return True

    def _start(self):
        if self.running or not self._validate():
            return
        self.running = True
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", tk.END)
        self.log_box.configure(state="disabled")
        threading.Thread(target=self._run_job, daemon=True).start()

    def _run_job(self):
        stop_event = threading.Event()
        try:
            inp     = self.input_var.get().strip()
            out_dir = self.output_var.get().strip()

            self._log(f"\n{'='*52}")
            self._log("  Converting Excel to JSON...")
            rows = excel_to_json(inp)
            self._log(f"  {len(rows)} rows loaded.")

            upload_input_json(rows, self._log)

            self._log(f"  Triggering {NUM_WORKERS} parallel workers on GitHub Actions...")
            trigger_workflow(self._log)

            run_id = get_latest_run_id(self._log)
            self._log(f"  Run ID: {run_id}")
            self._log(f"  Watch: https://github.com/{GITHUB_USER}/{GITHUB_REPO}/actions/runs/{run_id}")
            self._log(f"{'='*52}")

            conclusion = wait_for_run(run_id, self._log, stop_event)

            self._log(f"\n{'='*52}")
            if conclusion == "success":
                self._log("✓ Scraping complete! Downloading Excel...")
                download_artifact(run_id, out_dir, self._log)
                self._log(f"\n✓ Done! Results saved to:")
                self._log(f"  {out_dir}")
                self._log(f"{'='*52}")
                mb.showinfo("Done ✓", f"Results saved to:\n{out_dir}")
            else:
                self._log(f"✗ Run ended: {conclusion}")
                self._log(f"  Logs: https://github.com/{GITHUB_USER}/{GITHUB_REPO}/actions/runs/{run_id}")
                mb.showerror("Failed", f"Scraper run ended with: {conclusion}\nCheck logs link above.")

        except Exception as e:
            self._log(f"\n✗ ERROR: {e}")
            mb.showerror("Error", str(e))
        finally:
            stop_event.set()
            self.running = False


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
