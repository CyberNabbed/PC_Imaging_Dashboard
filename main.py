import tkinter as tk
from tkinter import ttk, messagebox
import os
import datetime
import time
import re
import subprocess
import threading

# change these for your environment
LOG_DIRECTORY = r'\\SERVER\SHARE\EnrollmentLogs'
ADMIN_SERVICE_URL = "https://SERVER/AdminService/wmi"
EMAIL_TO = "recipient@example.com"


class EnrollmentDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Device Enrollment Dashboard")
        self.root.geometry("1100x650")

        self.log_directory = LOG_DIRECTORY
        self.tracked_hosts = set()
        self.all_hostnames = []

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=28)
        style.configure("Treeview.Heading", font=("Helvetica", 11, "bold"))

        top = tk.Frame(root)
        top.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(top, text="Hostname").pack(side=tk.LEFT)
        self.hostname_entry = tk.Entry(top, width=20)
        self.hostname_entry.pack(side=tk.LEFT, padx=5)

        self.hostname_entry.bind("<KeyRelease>", self.on_key_release)
        self.hostname_entry.bind("<Return>", self.add_device_from_entry)
        self.hostname_entry.bind("<Down>", self.on_arrow_down)

        tk.Button(top, text="Add", command=self.add_device_from_entry).pack(side=tk.LEFT, padx=5)
        tk.Button(top, text="Refresh", command=self.manual_refresh).pack(side=tk.LEFT, padx=5)
        tk.Button(top, text="Report", command=self.show_report_window).pack(side=tk.LEFT, padx=20)
        tk.Button(top, text="Remove", command=self.remove_selected).pack(side=tk.RIGHT)

        self.suggestion_box = tk.Listbox(root, height=8)
        self.suggestion_box.bind("<ButtonRelease-1>", self.on_listbox_select)

        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10)

        cols = ("hostname", "serial", "status", "start", "details")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings")

        for c in cols:
            self.tree.heading(c, text=c.title())

        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.tag_configure("success", background="#dff0d8")
        self.tree.tag_configure("failed", background="#f2dede")
        self.tree.tag_configure("working", background="#fcf8e3")

        self.status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.root.after(200, self.load_hostnames)
        self.start_auto_refresh()

    def add_device_from_entry(self, event=None):
        name = self.hostname_entry.get().strip()
        if not name:
            return

        self.hostname_entry.delete(0, tk.END)
        self.suggestion_box.place_forget()
        self.register_host(name)

    def register_host(self, hostname):
        key = hostname.lower()
        if key in self.tracked_hosts:
            return

        self.tracked_hosts.add(key)
        self.tree.insert("", tk.END, iid=key,
                         values=(hostname, "Loading", "Checking", "-", "-"))

        self.check_single_host(hostname)

        threading.Thread(
            target=self.fetch_serial_background,
            args=(hostname,),
            daemon=True
        ).start()

    def remove_selected(self):
        for item in self.tree.selection():
            self.tree.delete(item)
            self.tracked_hosts.discard(item)

    def start_auto_refresh(self):
        self.manual_refresh()
        self.root.after(300000, self.start_auto_refresh)

    def manual_refresh(self):
        for h in list(self.tracked_hosts):
            self.check_single_host(h)
        self.status_bar.config(
            text=f"Last updated {datetime.datetime.now().strftime('%H:%M:%S')}"
        )

    def fetch_serial_background(self, hostname):
        # stubbed out for public version
        try:
            ps = f'''
            $base = "{ADMIN_SERVICE_URL}"
            # real lookup removed
            '''
            subprocess.Popen(["powershell", "-Command", ps],
                             stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE)
            serial = "Unknown"
        except Exception:
            serial = "Error"

        self.root.after(0, lambda: self.update_serial(hostname, serial))

    def update_serial(self, hostname, serial):
        key = hostname.lower()
        if self.tree.exists(key):
            self.tree.set(key, "serial", serial)

    def check_single_host(self, hostname):
        key = hostname.lower()
        path = os.path.join(self.log_directory, f"{hostname}.log")

        if not os.path.exists(path):
            self.tree.item(
                key,
                values=(hostname, "-", "Missing", "-", "Log not found"),
                tags=("failed",)
            )
            return

        lines = self.read_log(path)
        self.parse_log(key, hostname, lines)

    def read_log(self, path):
        try:
            with open(path, encoding="utf-8", errors="ignore") as f:
                return f.readlines()
        except Exception:
            return []

    def parse_log(self, key, hostname, lines):
        if not lines:
            self.tree.item(
                key,
                values=(hostname, "-", "Empty", "-", "No data"),
                tags=("working",)
            )
            return

        last = ""
        for line in reversed(lines):
            if line.strip():
                last = line.strip()
                break

        self.tree.item(
            key,
            values=(hostname, "-", "Working", "-", last[:80]),
            tags=("working",)
        )

    def show_report_window(self):
        messagebox.showinfo(
            "Report",
            "Email generation removed for public example."
        )

    def load_hostnames(self):
        if not os.path.exists(self.log_directory):
            return

        self.all_hostnames = [
            f[:-4] for f in os.listdir(self.log_directory)
            if f.lower().endswith(".log")
        ]

    def on_key_release(self, event):
        text = self.hostname_entry.get().lower()
        if not text:
            self.suggestion_box.place_forget()
            return

        matches = [h for h in self.all_hostnames if text in h.lower()]
        if not matches:
            self.suggestion_box.place_forget()
            return

        self.suggestion_box.delete(0, tk.END)
        for m in matches[:8]:
            self.suggestion_box.insert(tk.END, m)

        self.suggestion_box.place(
            x=self.hostname_entry.winfo_x(),
            y=self.hostname_entry.winfo_y() + 25
        )

    def on_arrow_down(self, event):
        if self.suggestion_box.winfo_ismapped():
            self.suggestion_box.focus_set()
            self.suggestion_box.selection_set(0)

    def on_listbox_select(self, event):
        if not self.suggestion_box.curselection():
            return

        value = self.suggestion_box.get(self.suggestion_box.curselection())
        self.hostname_entry.delete(0, tk.END)
        self.hostname_entry.insert(0, value)
        self.suggestion_box.place_forget()


if __name__ == "__main__":
    root = tk.Tk()
    EnrollmentDashboard(root)
    root.mainloop()
