from __future__ import annotations

import json
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext
from typing import Any

from config import CONFIG_FILE, load_config, save_config
from scheduler import MailScheduler
from sender import send_mail
from state import load_state


class AutoMailUI:
    def __init__(self, root: tk.Tk, scheduler: MailScheduler) -> None:
        self.root = root
        self.scheduler = scheduler
        self.config = load_config()
        root.title("AutoMail - cấu hình gửi mail tự động")
        root.geometry("820x720")
        self._build()
        self._load_to_form()

    def _build(self) -> None:
        frm = tk.Frame(self.root, padx=10, pady=10)
        frm.pack(fill="both", expand=True)
        self.entries: dict[str, tk.Entry] = {}
        fields = [("to", "Người nhận (To)"), ("cc", "CC"), ("bcc", "BCC"), ("subject", "Tiêu đề"), ("font_family", "Font"), ("font_size", "Cỡ chữ")]
        for i, (key, label) in enumerate(fields):
            tk.Label(frm, text=label).grid(row=i, column=0, sticky="w")
            ent = tk.Entry(frm, width=90)
            ent.grid(row=i, column=1, sticky="ew", pady=2)
            self.entries[key] = ent
        tk.Label(frm, text="Nội dung HTML/Text").grid(row=6, column=0, sticky="nw")
        self.body = scrolledtext.ScrolledText(frm, height=12)
        self.body.grid(row=6, column=1, sticky="nsew")
        self.format_var = tk.StringVar(value="html")
        tk.Checkbutton(frm, text="HTML", variable=self.format_var, onvalue="html", offvalue="plain").grid(row=7, column=1, sticky="w")
        self.schedule_enabled = tk.BooleanVar()
        self.schedule_type = tk.StringVar(value="daily")
        self.schedule_time = tk.StringVar(value="08:00")
        self.interval = tk.StringVar(value="60")
        sch = tk.LabelFrame(frm, text="Lịch gửi")
        sch.grid(row=8, column=1, sticky="ew", pady=8)
        tk.Checkbutton(sch, text="Bật schedule", variable=self.schedule_enabled).pack(side="left")
        tk.OptionMenu(sch, self.schedule_type, "daily", "weekly", "interval").pack(side="left")
        tk.Label(sch, text="Giờ HH:MM").pack(side="left")
        tk.Entry(sch, textvariable=self.schedule_time, width=8).pack(side="left")
        tk.Label(sch, text="Interval phút").pack(side="left")
        tk.Entry(sch, textvariable=self.interval, width=6).pack(side="left")
        btns = tk.Frame(frm)
        btns.grid(row=9, column=1, sticky="ew", pady=8)
        tk.Button(btns, text="Chọn .eml", command=self._pick_eml).pack(side="left")
        tk.Button(btns, text="Lưu config", command=self.save).pack(side="left", padx=5)
        tk.Button(btns, text="Gửi thử", command=self.send_test).pack(side="left")
        tk.Button(btns, text="Xem trạng thái", command=self.show_state).pack(side="left", padx=5)
        self.status = tk.Label(frm, text=f"Config: {CONFIG_FILE}", anchor="w")
        self.status.grid(row=10, column=0, columnspan=2, sticky="ew")
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(6, weight=1)

    def _load_to_form(self) -> None:
        mail = self.config["mail"]
        for key, ent in self.entries.items():
            value = mail.get(key, "")
            ent.delete(0, "end")
            ent.insert(0, ";".join(value) if isinstance(value, list) else str(value))
        self.body.delete("1.0", "end")
        self.body.insert("1.0", mail.get("body", ""))
        self.format_var.set(mail.get("body_format", "html"))
        sch = self.config["schedule"]
        self.schedule_enabled.set(bool(sch.get("enabled")))
        self.schedule_type.set(sch.get("type", "daily"))
        self.schedule_time.set(sch.get("time", "08:00"))
        self.interval.set(str(sch.get("interval_minutes", 60)))

    def _form_config(self) -> dict[str, Any]:
        cfg = load_config()
        mail = cfg["mail"]
        for key, ent in self.entries.items():
            val = ent.get().strip()
            mail[key] = [x.strip() for x in val.replace(",", ";").split(";") if x.strip()] if key in {"to","cc","bcc"} else val
        mail["font_size"] = int(mail.get("font_size") or 11)
        mail["body"] = self.body.get("1.0", "end").strip()
        mail["body_format"] = self.format_var.get()
        cfg["schedule"] = {"enabled": self.schedule_enabled.get(), "type": self.schedule_type.get(), "time": self.schedule_time.get(), "weekdays": [0,1,2,3,4], "interval_minutes": int(self.interval.get() or 60)}
        return cfg

    def save(self) -> None:
        self.config = self._form_config()
        save_config(self.config)
        self.status.config(text="Đã lưu config.json")

    def send_test(self) -> None:
        self.save()
        threading.Thread(target=lambda: messagebox.showinfo("AutoMail", send_mail(self.config["mail"])["message"]), daemon=True).start()

    def show_state(self) -> None:
        messagebox.showinfo("State", json.dumps(load_state(), ensure_ascii=False, indent=2))

    def _pick_eml(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Email template", "*.eml")])
        if path:
            self.config["mail"]["template"] = str(Path(path))
            save_config(self.config)
            self.status.config(text=f"Template: {path}")
