import json
import traceback
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

import LDefin2D

from cad_ai.config import (
    APP_TITLE,
    APP_GEOMETRY,
    APP_MINSIZE,
    LLM_ENABLED,
    LLM_MODEL_PATH,
)
from cad_ai.kompas.connect import connect_kompas, new_document_part
from cad_ai.kompas.builder import Kompas3DBuilder
from cad_ai.templates import TEMPLATES
from cad_ai.llm.engine import LocalLLMEngine
from cad_ai.llm.errors import LLMJSONError


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(*APP_MINSIZE)

        self.ks_const = None
        self.ks_const_3d = None
        self.api5 = None
        self.api7 = None
        self.kompas_object = None
        self.application = None

        self.builder = None
        self.iPart = None

        self.llm_json = None
        self.llm_raw = None
        self.llm_extracted = None
        self.llm_prompt = None
        self._llm_engine = None
        self._llm_busy = False

        self._build_ui()

    # -----------------
    # UI
    # -----------------
    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        conn = ttk.LabelFrame(root, text="KOMPAS connection", padding=10)
        conn.pack(fill="x")

        self.status_var = tk.StringVar(value="Not connected")
        ttk.Label(conn, textvariable=self.status_var).pack(side="left")
        ttk.Button(conn, text="Connect", command=self.on_connect).pack(side="right")

        llm_box = ttk.LabelFrame(
            root, text="Text request (local LLM -> JSON)", padding=10
        )
        llm_box.pack(fill="x", pady=(12, 0))

        self.prompt_var = tk.StringVar(
            value="пластина 120 на 80 толщиной 8, 4 отверстия 10, отступ 15"
        )
        ttk.Entry(llm_box, textvariable=self.prompt_var).pack(fill="x")

        row = ttk.Frame(llm_box)
        row.pack(fill="x", pady=(8, 0))

        self.btn_llm_generate = ttk.Button(
            row, text="Generate JSON (LLM)", command=self.on_generate_llm
        )
        self.btn_llm_generate.pack(side="left")

        self.btn_llm_build = ttk.Button(
            row, text="Build LLM result", command=self.on_build_llm
        )
        self.btn_llm_build.pack(side="left", padx=8)

        self.btn_llm_show_json = ttk.Button(
            row, text="Show LLM JSON", command=self.on_show_llm_json
        )
        self.btn_llm_show_json.pack(side="left", padx=8)

        self.btn_llm_show_raw = ttk.Button(
            row, text="Show LLM raw", command=self.on_show_llm_raw
        )
        self.btn_llm_show_raw.pack(side="left", padx=8)

        self.llm_status_var = tk.StringVar(value="LLM: idle")
        ttk.Label(llm_box, textvariable=self.llm_status_var).pack(
            anchor="w", pady=(8, 0)
        )

        self.llm_progress = ttk.Progressbar(llm_box, mode="indeterminate")
        self.llm_progress.pack(fill="x", pady=(4, 0))

        box = ttk.LabelFrame(root, text="AI examples (templates)", padding=10)
        box.pack(fill="x", pady=(12, 0))

        self.template_var = tk.StringVar(value=list(TEMPLATES.keys())[0])
        self.template_cb = ttk.Combobox(
            box,
            state="readonly",
            textvariable=self.template_var,
            values=list(TEMPLATES.keys()),
        )
        self.template_cb.pack(fill="x")
        self.template_cb.bind("<<ComboboxSelected>>", lambda e: self.render_params())

        self.params_frame = ttk.LabelFrame(
            root, text="Parameters (templates)", padding=10
        )
        self.params_frame.pack(fill="x", pady=(12, 0))
        self.param_entries = {}
        self.render_params()

        opts = ttk.LabelFrame(root, text="Options", padding=10)
        opts.pack(fill="x", pady=(12, 0))

        self.new_doc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts, text="Create new document each build", variable=self.new_doc_var
        ).pack(anchor="w")

        btns = ttk.Frame(root)
        btns.pack(fill="x", pady=(12, 0))

        ttk.Button(
            btns, text="Build template in KOMPAS", command=self.on_build_template
        ).pack(side="left")
        ttk.Button(
            btns, text="Show template JSON", command=self.on_show_template_json
        ).pack(side="left", padx=8)
        ttk.Button(btns, text="Exit", command=self.destroy).pack(side="right")

        logbox = ttk.LabelFrame(root, text="Log", padding=10)
        logbox.pack(fill="both", expand=True, pady=(12, 0))

        self.log = tk.Text(logbox, height=10, wrap="word")
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")

    def log_write(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # -----------------
    # Template UI helpers
    # -----------------
    def render_params(self):
        for w in self.params_frame.winfo_children():
            w.destroy()
        self.param_entries.clear()

        tpl = TEMPLATES[self.template_var.get()]
        for key, label, default in tpl["params"]:
            row = ttk.Frame(self.params_frame)
            row.pack(fill="x", pady=3)

            ttk.Label(row, text=label, width=26).pack(side="left")
            ent = ttk.Entry(row)
            ent.pack(side="left", fill="x", expand=True)
            ent.insert(0, str(default))
            self.param_entries[key] = ent

    def show_text_window(self, title: str, text: str):
        top = tk.Toplevel(self)
        top.title(title)
        top.geometry("820x520")
        t = tk.Text(top, wrap="word")
        t.pack(fill="both", expand=True)
        t.insert("1.0", text or "")
        t.focus_set()

    def set_llm_busy(self, busy: bool, *, status: str | None = None):
        self._llm_busy = bool(busy)
        state = "disabled" if busy else "normal"

        for btn in (
            "btn_llm_generate",
            "btn_llm_build",
            "btn_llm_show_json",
            "btn_llm_show_raw",
        ):
            b = getattr(self, btn, None)
            if b is not None:
                try:
                    b.configure(state=state)
                except Exception:
                    pass

        if status is not None:
            self.llm_status_var.set(status)

        try:
            if busy:
                self.llm_progress.start(12)
            else:
                self.llm_progress.stop()
        except Exception:
            pass

    def on_show_llm_raw(self):
        if not self.llm_raw and not self.llm_extracted:
            messagebox.showinfo("No data", "Сначала сгенерируй JSON (LLM).")
            return

        parts = []
        if self.llm_prompt:
            parts.append("=== PROMPT SENT TO LLM ===\n" + self.llm_prompt.strip())
        if self.llm_raw:
            parts.append("\n\n=== RAW LLM OUTPUT ===\n" + self.llm_raw.strip())
        if self.llm_extracted:
            parts.append(
                "\n\n=== EXTRACTED JSON (what we tried to parse) ===\n"
                + self.llm_extracted.strip()
            )

        self.show_text_window("LLM debug output", "\n".join(parts))

    def read_params(self):
        p = {}
        for k, ent in self.param_entries.items():
            v = ent.get().strip().replace(",", ".")
            p[k] = float(v)
        return p

    def build_template_json(self):
        tpl = TEMPLATES[self.template_var.get()]
        params = self.read_params()
        return tpl["build"](params)

    # -----------------
    # KOMPAS helpers
    # -----------------
    def ensure_connected(self):
        if self.application is None or self.ks_const is None:
            raise RuntimeError("Not connected. Click 'Connect' first.")

    def ensure_builder(self):
        self.ensure_connected()
        if self.new_doc_var.get() or self.iPart is None:
            self.log_write("Creating new 3D Part document...")
            _, _, _, iPart = new_document_part(
                self.ks_const,
                self.ks_const_3d,
                self.api5,
                self.api7,
                self.kompas_object,
                self.application,
            )
            self.iPart = iPart
            self.builder = Kompas3DBuilder(self.ks_const, self.ks_const_3d, self.iPart)

    def on_connect(self):
        try:
            self.log_write("Connecting to KOMPAS...")
            ks_const, ks_const_3d, api5, api7, kompas_object, application = (
                connect_kompas()
            )
            self.ks_const = ks_const
            self.ks_const_3d = ks_const_3d
            self.api5 = api5
            self.api7 = api7
            self.kompas_object = kompas_object
            self.application = application

            self.status_var.set("Connected ✅")
            self.log_write("Connected OK. Constants and APIs loaded.")
        except Exception as e:
            self.status_var.set("Not connected")
            self.log_write("ERROR while connecting:\n" + traceback.format_exc())
            messagebox.showerror("Connect error", str(e))

    # -----------------
    # LLM
    # -----------------
    def get_llm_engine(self) -> LocalLLMEngine:
        if not LLM_ENABLED:
            raise RuntimeError("LLM is disabled (LLM_ENABLED=False).")
        if self._llm_engine is None:
            model_path = Path(LLM_MODEL_PATH)
            self.log_write(f"Loading local LLM: {model_path} ...")
            self._llm_engine = LocalLLMEngine(
                model_path=str(model_path),
                n_ctx=2048,
                n_threads=8,
                n_gpu_layers=0,
            )
            self.log_write("LLM loaded ✅")
        return self._llm_engine

    def generate_llm_json(self, user_text: str) -> dict:
        eng = self.get_llm_engine()
        data = eng.generate_json(user_text)

        self.llm_raw = eng.last_raw
        self.llm_extracted = eng.last_extracted
        self.llm_prompt = eng.last_prompt

        return data

    def on_generate_llm(self):
        text = self.prompt_var.get().strip()
        if not text:
            messagebox.showwarning("Empty", "Введите текстовый запрос")
            return
        if self._llm_busy:
            return

        self.set_llm_busy(True, status="LLM: generating...")

        def worker():
            try:
                self.log_write("LLM: generating JSON...")
                data = self.generate_llm_json(text)

                def ok():
                    self.llm_json = data
                    self.log_write("LLM JSON ready ✅")
                    self.set_llm_busy(False, status="LLM: idle")
                    self.show_json_window("LLM JSON output", data)

                self.after(0, ok)

            except LLMJSONError as e:

                def err():
                    self.log_write(
                        "LLM ERROR (parse/validate):\n" + traceback.format_exc()
                    )
                    self.llm_raw = getattr(e, "raw", "") or self.llm_raw
                    self.llm_extracted = (
                        getattr(e, "extracted", "") or self.llm_extracted
                    )
                    self.llm_prompt = getattr(e, "prompt", "") or self.llm_prompt
                    self.set_llm_busy(False, status="LLM: error (see raw)")
                    self.on_show_llm_raw()
                    messagebox.showerror("LLM error", str(e))

                self.after(0, err)

            except Exception as e:

                def err2():
                    self.log_write("LLM ERROR:\n" + traceback.format_exc())
                    self.set_llm_busy(False, status="LLM: error")
                    messagebox.showerror("LLM error", str(e))

                self.after(0, err2)

        threading.Thread(target=worker, daemon=True).start()

    def on_build_llm(self):
        try:
            self.ensure_builder()
            if not self.llm_json:
                raise RuntimeError("Сначала нажми Generate JSON (LLM).")

            self.log_write(
                f"Building LLM model: {self.llm_json.get('name','(no name)')}"
            )
            self.builder.process_json(self.llm_json)
            self.log_write("Build complete ✅")
            messagebox.showinfo("Done", "Модель (LLM) построена в KOMPAS.")
        except Exception as e:
            self.log_write("BUILD LLM ERROR:\n" + traceback.format_exc())
            messagebox.showerror("Build error", str(e))

    def on_show_llm_json(self):
        if not self.llm_json:
            messagebox.showinfo("No data", "LLM JSON еще не сгенерирован.")
            return
        self.show_json_window("LLM JSON output", self.llm_json)

    # -----------------
    # Templates actions
    # -----------------
    def on_show_template_json(self):
        try:
            data = self.build_template_json()
            self.show_json_window("Template JSON output", data)
        except Exception as e:
            messagebox.showerror("JSON error", str(e))

    def on_build_template(self):
        try:
            self.ensure_builder()
            data = self.build_template_json()

            self.log_write(
                f"Building template: {data.get('name', self.template_var.get())}"
            )
            self.builder.process_json(data)
            self.log_write("Build complete ✅")
            messagebox.showinfo("Done", "Model построена в KOMPAS.")
        except Exception as e:
            self.log_write("ERROR while building:\n" + traceback.format_exc())
            messagebox.showerror("Build error", str(e))

    # -----------------
    # JSON viewer
    # -----------------
    def show_json_window(self, title: str, data: dict):
        txt = json.dumps(data, ensure_ascii=False, indent=2)
        top = tk.Toplevel(self)
        top.title(title)
        top.geometry("760x560")
        t = tk.Text(top, wrap="none")
        t.pack(fill="both", expand=True)
        t.insert("1.0", txt)
