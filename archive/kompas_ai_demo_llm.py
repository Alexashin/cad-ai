# -*- coding: utf-8 -*-
"""
KOMPAS-3D AI Demo (LLM -> JSON -> KOMPAS) — single file
------------------------------------------------------
What this demo does:
- Connects to KOMPAS-3D via COM (API5 + API7) using pywin32.
- Lets you:
  1) Build pre-made "AI examples" (templates).
  2) Type a natural-language request and generate JSON steps locally using a small LLM (GGUF via llama-cpp-python).
  3) Build the generated JSON in KOMPAS.

Why JSON (DSL) instead of Python-macros:
- Much more stable for a demo: the LLM outputs a constrained set of actions,
  and your interpreter executes them.

Requirements:
- Windows + installed KOMPAS-3D with COM registered
- Python 3.10+ recommended
- pip install pywin32 llama-cpp-python

Local files required in the same folder:
- MiscellaneousHelpers.py
- LDefin2D.py (usually shipped with KOMPAS examples)

Model:
- Put a GGUF model file somewhere (default: ./models/qwen2.5-1.5b-instruct-q4_k_m.gguf)
- Recommended lightweight model for JSON:
  Qwen2.5-1.5B-Instruct (GGUF)

Notes:
- This file is intentionally "presentation-stable": it validates LLM output and limits actions.
- If LLM is missing/unavailable, templates still work.

Author: You + ChatGPT
"""

import json
import traceback
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

import pythoncom
from win32com.client import Dispatch, gencache

# Local helper modules from your project/examples
import MiscellaneousHelpers as MH
import LDefin2D  # noqa: F401  (imported for compatibility with KOMPAS examples)


# -----------------------------
# LLM config (GGUF via llama-cpp-python)
# -----------------------------

LLM_MODEL_PATH = str(Path("models") / "qwen2.5-1.5b-instruct-q4_k_m.gguf")

# If you want to disable LLM without removing code:
LLM_ENABLED = True


def make_llm_prompt(user_text: str) -> str:
    """
    Prompt to force a strict JSON output following our mini-DSL.
    Keep it short for small models.
    """
    return f"""
Ты — генератор JSON-инструкций для построения 3D модели в КОМПАС-3D.
Верни ТОЛЬКО валидный JSON-объект без пояснений, без markdown, без текста.

Формат:
{{
  "name": "string",
  "steps": [ ... ]
}}

Разрешенные action и поля:

1) sketch:
{{
  "action": "sketch",
  "plane": "XOY" | "XOZ" | "YOZ",
  "entities": [
    {{ "type":"line", "start":[x,y], "end":[x,y] }},
    {{ "type":"circle", "center":[x,y], "radius": r }}
  ]
}}

2) extrude:
{{ "action":"extrude", "height": number, "direction":"normal"|"reverse"|"both" }}

3) cut:
{{ "action":"cut", "through_all": true, "direction":"normal"|"reverse"|"both" }}
или
{{ "action":"cut", "through_all": false, "depth": number, "direction":"normal"|"reverse"|"both" }}

4) workplane_offset:
{{ "action":"workplane_offset", "base_plane":"XOY"|"XOZ"|"YOZ", "offset": number, "name":"string" }}

5) sketch_on_plane:
{{
  "action":"sketch_on_plane",
  "plane":"XOY"|"XOZ"|"YOZ"|"<name from workplane_offset>",
  "name":"optional_step_name",
  "entities":[ ...как в sketch... ]
}}

Ограничения:
- Единицы: миллиметры.
- Все числа — только числа (не строки).
- Используй только разрешенные action.
- Если параметров не хватает, выбери разумные значения по умолчанию.
- Для отверстия "насквозь" используй cut with through_all=true.
- Обычно direction для extrude: "both" (стабильно для демо), для cut: "both" (если не уверен).

Запрос пользователя: {user_text}
""".strip()


# -----------------------------
# COM / KOMPAS bootstrap
# -----------------------------


class LLMJSONError(Exception):
    """Raised when the LLM output cannot be parsed/validated as JSON."""

    def __init__(
        self, message: str, *, raw: str = "", extracted: str = "", prompt: str = ""
    ):
        super().__init__(message)
        self.raw = raw or ""
        self.extracted = extracted or ""
        self.prompt = prompt or ""


GUID_KS_CONST = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"
GUID_KS_CONST_3D = "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}"
GUID_API5 = "{0422828C-F174-495E-AC5D-D31014DBBE87}"
GUID_API7 = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"


def connect_kompas():
    """
    Connect to KOMPAS (API5 and API7), load constants, and return:
    (kompas6_constants, kompas6_constants_3d, api5, api7, kompas_object_api5, application_api7)
    """
    ks_const = gencache.EnsureModule(GUID_KS_CONST, 0, 1, 0).constants
    ks_const_3d = gencache.EnsureModule(GUID_KS_CONST_3D, 0, 1, 0).constants

    api5 = gencache.EnsureModule(GUID_API5, 0, 1, 0)
    api7 = gencache.EnsureModule(GUID_API7, 0, 1, 0)

    kompas5_disp = Dispatch("Kompas.Application.5")
    kompas7_disp = Dispatch("Kompas.Application.7")

    kompas_object = api5.KompasObject(
        kompas5_disp._oleobj_.QueryInterface(
            api5.KompasObject.CLSID, pythoncom.IID_IDispatch
        )
    )
    application = api7.IApplication(
        kompas7_disp._oleobj_.QueryInterface(
            api7.IApplication.CLSID, pythoncom.IID_IDispatch
        )
    )

    MH.iKompasObject = kompas_object
    MH.iApplication = application

    try:
        MH.iKompasObject.Visible = True
    except Exception:
        pass

    return ks_const, ks_const_3d, api5, api7, kompas_object, application


def new_document_part(ks_const, ks_const_3d, api5, api7, kompas_object, application):
    """
    Create a new 3D part document and return:
    (kompas_document, kompas_document_3d_api7, iDocument3D_api5, iPart_api5)
    """
    documents = application.Documents
    kompas_document = documents.AddWithDefaultSettings(ks_const.ksDocumentPart, True)
    kompas_document_3d = api7.IKompasDocument3D(kompas_document)

    iDocument3D = kompas_object.ActiveDocument3D()
    iPart = iDocument3D.GetPart(ks_const_3d.pTop_Part)

    return kompas_document, kompas_document_3d, iDocument3D, iPart


# -----------------------------
# Builder (JSON -> KOMPAS)
# -----------------------------


class Kompas3DBuilder:
    def __init__(self, ks_const, ks_const_3d, iPart):
        self.ks_const = ks_const
        self.ks_const_3d = ks_const_3d
        self.iPart = iPart
        self.last_sketch = None
        self.named_planes = {}

    def start_sketch(self, plane_name: str):
        plane_name = (plane_name or "XOY").upper().strip()
        plane_attr = f"o3d_plane{plane_name}"
        if not hasattr(self.ks_const_3d, plane_attr):
            raise ValueError(f"Unknown plane '{plane_name}'. Use one of: XOY, XOZ, YOZ")

        sketch = self.iPart.NewEntity(self.ks_const_3d.o3d_sketch)
        definition = sketch.GetDefinition()

        plane_obj = self.iPart.GetDefaultEntity(getattr(self.ks_const_3d, plane_attr))
        definition.SetPlane(plane_obj)

        sketch.Create()
        self.last_sketch = sketch
        return definition.BeginEdit()

    def finish_sketch(self):
        if not self.last_sketch:
            raise RuntimeError("No active sketch to finish.")
        self.last_sketch.GetDefinition().EndEdit()

    # --- 2D primitives ---
    @staticmethod
    def add_line(doc2d, x1, y1, x2, y2):
        doc2d.ksLineSeg(float(x1), float(y1), float(x2), float(y2), 1)

    @staticmethod
    def add_circle(doc2d, x, y, radius):
        doc2d.ksCircle(float(x), float(y), float(radius), 1)

    # --- Workplanes ---
    def start_sketch_on_plane_any(self, plane_name: str):
        plane_name = (plane_name or "XOY").strip()
        up = plane_name.upper()

        if up in ("XOY", "XOZ", "YOZ"):
            return self.start_sketch(up)

        # иначе считаем, что это имя созданной плоскости
        return self.start_sketch_on_named_plane(plane_name)

    def create_offset_plane(self, base_plane: str, offset: float, name: str):
        base_plane = base_plane.upper().strip()
        attr = f"o3d_plane{base_plane}"
        if not hasattr(self.ks_const_3d, attr):
            raise ValueError("base_plane must be XOY / XOZ / YOZ")

        plane_entity = self.iPart.NewEntity(self.ks_const_3d.o3d_planeOffset)
        plane_def = plane_entity.GetDefinition()

        base = self.iPart.GetDefaultEntity(getattr(self.ks_const_3d, attr))
        plane_def.SetPlane(base)
        plane_def.offset = float(offset)

        plane_entity.Create()
        self.named_planes[name] = plane_entity

    def start_sketch_on_named_plane(self, plane_name: str):
        if plane_name not in self.named_planes:
            raise ValueError(f"Plane '{plane_name}' not found")

        sketch = self.iPart.NewEntity(self.ks_const_3d.o3d_sketch)
        definition = sketch.GetDefinition()

        definition.SetPlane(self.named_planes[plane_name])
        sketch.Create()
        self.last_sketch = sketch
        return definition.BeginEdit()

    # --- Features ---
    def extrude_boss(self, height: float, direction: str = "both"):
        if not self.last_sketch:
            raise RuntimeError("Extrude requires a sketch first.")
        height = float(height)

        extrusion = self.iPart.NewEntity(self.ks_const_3d.o3d_bossExtrusion)
        definition = extrusion.GetDefinition()
        definition.SetSketch(self.last_sketch)

        p = definition.ExtrusionParam()
        d = (direction or "both").lower().strip()

        if d in ("reverse", "rev", "back"):
            p.direction = self.ks_const_3d.dtReverse
            p.typeReverse = self.ks_const_3d.etBlind
            p.depthReverse = height
        elif d in ("normal", "norm", "forward"):
            p.direction = self.ks_const_3d.dtNormal
            p.typeNormal = self.ks_const_3d.etBlind
            p.depthNormal = height
        else:
            p.direction = self.ks_const_3d.dtBoth
            half = height / 2.0
            p.typeNormal = self.ks_const_3d.etBlind
            p.depthNormal = half
            p.typeReverse = self.ks_const_3d.etBlind
            p.depthReverse = half

        extrusion.Create()

    def cut_extrusion(
        self,
        *,
        through_all: bool = True,
        depth: float | None = None,
        direction: str = "normal",
    ):
        if not self.last_sketch:
            raise RuntimeError("Cut requires a sketch first.")

        cut_feature = self.iPart.NewEntity(self.ks_const_3d.o3d_cutExtrusion)
        definition = cut_feature.GetDefinition()
        definition.SetSketch(self.last_sketch)

        p = definition.ExtrusionParam()

        dir_norm = (direction or "normal").lower().strip()
        is_reverse = dir_norm in ("reverse", "rev", "back")
        is_both = dir_norm in ("both", "two", "2", "both_sides")

        if is_both:
            p.direction = self.ks_const_3d.dtBoth
            if through_all:
                p.typeNormal = self.ks_const_3d.etThroughAll
                p.typeReverse = self.ks_const_3d.etThroughAll
            else:
                if depth is None:
                    raise ValueError("cut: depth is required when through_all=False")
                half = float(depth) / 2.0
                p.typeNormal = self.ks_const_3d.etBlind
                p.depthNormal = half
                p.typeReverse = self.ks_const_3d.etBlind
                p.depthReverse = half
        elif is_reverse:
            p.direction = self.ks_const_3d.dtReverse
            if through_all:
                p.typeReverse = self.ks_const_3d.etThroughAll
            else:
                if depth is None:
                    raise ValueError("cut: depth is required when through_all=False")
                p.typeReverse = self.ks_const_3d.etBlind
                p.depthReverse = float(depth)
        else:
            p.direction = self.ks_const_3d.dtNormal
            if through_all:
                p.typeNormal = self.ks_const_3d.etThroughAll
            else:
                if depth is None:
                    raise ValueError("cut: depth is required when through_all=False")
                p.typeNormal = self.ks_const_3d.etBlind
                p.depthNormal = float(depth)

        cut_feature.Create()

    def process_json(self, data: dict):
        steps = data.get("steps", [])
        for step in steps:
            action = (step.get("action") or "").lower().strip()

            if action == "sketch":
                plane = step.get("plane", "XOY")
                doc2d = self.start_sketch(plane)

                for ent in step.get("entities", []):
                    et = (ent.get("type") or "").lower().strip()
                    if et == "line":
                        x1, y1 = ent["start"]
                        x2, y2 = ent["end"]
                        self.add_line(doc2d, x1, y1, x2, y2)
                    elif et == "circle":
                        x, y = ent["center"]
                        r = ent["radius"]
                        self.add_circle(doc2d, x, y, r)
                    else:
                        raise ValueError(f"Unknown entity type: {ent.get('type')}")
                self.finish_sketch()

            elif action == "extrude":
                self.extrude_boss(
                    step.get("height", 10), direction=step.get("direction", "both")
                )

            elif action == "cut":
                through_all = bool(step.get("through_all", False))
                depth = step.get("depth", None)
                direction = step.get("direction", "normal")
                if through_all:
                    self.cut_extrusion(through_all=True, direction=direction)
                else:
                    if depth is None:
                        self.cut_extrusion(through_all=True, direction=direction)
                    else:
                        self.cut_extrusion(
                            through_all=False, depth=float(depth), direction=direction
                        )

            elif action == "workplane_offset":
                self.create_offset_plane(
                    base_plane=step["base_plane"],
                    offset=step["offset"],
                    name=step["name"],
                )

            elif action == "sketch_on_plane":
                # plane обязателен, но если LLM забыл — по умолчанию XOY
                plane_name = (
                    step.get("plane")
                    or step.get("plane_name")
                    or step.get("on_plane")
                    or "XOY"
                )

                doc2d = self.start_sketch_on_plane_any(plane_name)

                for ent in step.get("entities", []):
                    et = (ent.get("type") or "").lower().strip()
                    if et == "circle":
                        x, y = ent["center"]
                        r = ent["radius"]
                        self.add_circle(doc2d, x, y, r)
                    elif et == "line":
                        x1, y1 = ent["start"]
                        x2, y2 = ent["end"]
                        self.add_line(doc2d, x1, y1, x2, y2)
                    else:
                        raise ValueError(f"Unknown entity type: {ent.get('type')}")
                self.finish_sketch()

            else:
                raise ValueError(f"Unknown action: {step.get('action')}")


# -----------------------------
# Pre-made "AI" templates (stable)
# -----------------------------


def tpl_cube(size=100.0, plane="XOY"):
    s = float(size)
    return {
        "name": "Cube",
        "steps": [
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "line", "start": [0, 0], "end": [s, 0]},
                    {"type": "line", "start": [s, 0], "end": [s, s]},
                    {"type": "line", "start": [s, s], "end": [0, s]},
                    {"type": "line", "start": [0, s], "end": [0, 0]},
                ],
            },
            {"action": "extrude", "height": s, "direction": "both"},
        ],
    }


def tpl_cube_with_through_hole(size=100.0, hole_d=20.0, plane="XOY"):
    s = float(size)
    d = float(hole_d)
    return {
        "name": "Cube + Through hole",
        "steps": [
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "line", "start": [0, 0], "end": [s, 0]},
                    {"type": "line", "start": [s, 0], "end": [s, s]},
                    {"type": "line", "start": [s, s], "end": [0, s]},
                    {"type": "line", "start": [0, s], "end": [0, 0]},
                ],
            },
            {"action": "extrude", "height": s, "direction": "both"},
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "circle", "center": [s / 2, s / 2], "radius": d / 2},
                ],
            },
            {"action": "cut", "through_all": True, "direction": "both"},
        ],
    }


def tpl_plate_with_holes(
    w=120.0, h=80.0, thickness=8.0, hole_d=10.0, margin=15.0, plane="XOY"
):
    w = float(w)
    h = float(h)
    thickness = float(thickness)
    hole_d = float(hole_d)
    margin = float(margin)

    holes = [
        {"type": "circle", "center": [margin, margin], "radius": hole_d / 2},
        {"type": "circle", "center": [w - margin, margin], "radius": hole_d / 2},
        {"type": "circle", "center": [w - margin, h - margin], "radius": hole_d / 2},
        {"type": "circle", "center": [margin, h - margin], "radius": hole_d / 2},
    ]

    return {
        "name": "Plate (4 holes)",
        "steps": [
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "line", "start": [0, 0], "end": [w, 0]},
                    {"type": "line", "start": [w, 0], "end": [w, h]},
                    {"type": "line", "start": [w, h], "end": [0, h]},
                    {"type": "line", "start": [0, h], "end": [0, 0]},
                ],
            },
            {"action": "extrude", "height": thickness, "direction": "both"},
            {"action": "sketch", "plane": plane, "entities": holes},
            {"action": "cut", "through_all": True, "direction": "both"},
        ],
    }


TEMPLATES = {
    "Куб (AI)": {
        "params": [("size", "Размер куба", 100.0)],
        "build": lambda p: tpl_cube(size=p["size"], plane="XOY"),
    },
    "Куб с отверстием насквозь (AI)": {
        "params": [
            ("size", "Размер куба", 100.0),
            ("hole_d", "Диаметр отверстия", 20.0),
        ],
        "build": lambda p: tpl_cube_with_through_hole(
            size=p["size"], hole_d=p["hole_d"], plane="XOY"
        ),
    },
    "Пластина 4 отверстия (AI)": {
        "params": [
            ("w", "Ширина", 120.0),
            ("h", "Высота", 80.0),
            ("thickness", "Толщина", 8.0),
            ("hole_d", "Диаметр отверстия", 10.0),
            ("margin", "Отступ от края", 15.0),
        ],
        "build": lambda p: tpl_plate_with_holes(
            w=p["w"],
            h=p["h"],
            thickness=p["thickness"],
            hole_d=p["hole_d"],
            margin=p["margin"],
            plane="XOY",
        ),
    },
}


# -----------------------------
# GUI
# -----------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KOMPAS-3D AI Demo (LLM -> JSON)")
        self.geometry("720x640")
        self.minsize(720, 640)

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
        self._llm = None
        self._llm_busy = False

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        # Connection panel
        conn = ttk.LabelFrame(root, text="KOMPAS connection", padding=10)
        conn.pack(fill="x")

        self.status_var = tk.StringVar(value="Not connected")
        ttk.Label(conn, textvariable=self.status_var).pack(side="left")
        ttk.Button(conn, text="Connect", command=self.on_connect).pack(side="right")

        # LLM panel
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

        # busy indicator (prevents "app freeze" feeling)
        self.llm_status_var = tk.StringVar(value="LLM: idle")
        ttk.Label(llm_box, textvariable=self.llm_status_var).pack(
            anchor="w", pady=(8, 0)
        )

        self.llm_progress = ttk.Progressbar(llm_box, mode="indeterminate")
        self.llm_progress.pack(fill="x", pady=(4, 0))

        # Template panel
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

        # Params panel
        self.params_frame = ttk.LabelFrame(
            root, text="Parameters (templates)", padding=10
        )
        self.params_frame.pack(fill="x", pady=(12, 0))
        self.param_entries = {}
        self.render_params()

        # Options panel
        opts = ttk.LabelFrame(root, text="Options", padding=10)
        opts.pack(fill="x", pady=(12, 0))

        self.new_doc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts, text="Create new document each build", variable=self.new_doc_var
        ).pack(anchor="w")

        # Buttons
        btns = ttk.Frame(root)
        btns.pack(fill="x", pady=(12, 0))

        ttk.Button(
            btns, text="Build template in KOMPAS", command=self.on_build_template
        ).pack(side="left")
        ttk.Button(
            btns, text="Show template JSON", command=self.on_show_template_json
        ).pack(side="left", padx=8)
        ttk.Button(btns, text="Exit", command=self.destroy).pack(side="right")

        # Log panel
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

        # buttons can be absent if UI not built yet
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

        if hasattr(self, "llm_status_var") and status is not None:
            self.llm_status_var.set(status)

        if hasattr(self, "llm_progress"):
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
        data = tpl["build"](params)
        return data

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
    # LLM functions
    # -----------------
    def get_llm(self):
        if not LLM_ENABLED:
            raise RuntimeError("LLM is disabled (LLM_ENABLED=False).")

        try:
            from llama_cpp import (
                Llama,
            )  # imported here to keep templates usable without llama-cpp-python
        except Exception as e:
            raise RuntimeError(
                "llama-cpp-python is not installed. Run: pip install llama-cpp-python"
            ) from e

        if self._llm is None:
            model_path = Path(LLM_MODEL_PATH)
            if not model_path.exists():
                raise RuntimeError(
                    f"GGUF model not found: {model_path}\n"
                    f"Put a GGUF model there or change LLM_MODEL_PATH in the script."
                )

            self.log_write(f"Loading local LLM: {model_path} ...")
            self._llm = Llama(
                model_path=str(model_path),
                n_ctx=2048,
                n_threads=8,
                n_gpu_layers=0,
                verbose=False,
            )
            self.log_write("LLM loaded ✅")

        return self._llm

    @staticmethod
    def _extract_json_object(text: str) -> str:
        # crude but effective: take the first {...} block
        first = text.find("{")
        last = text.rfind("}")
        if first == -1 or last == -1 or last <= first:
            raise ValueError("LLM did not return a JSON object.")
        return text[first : last + 1]

    def validate_generated_json(self, data: dict):
        if not isinstance(data, dict):
            raise ValueError("JSON root must be an object.")
        if "steps" not in data or not isinstance(data["steps"], list):
            raise ValueError("JSON must contain 'steps' as a list.")

        allowed_actions = {
            "sketch",
            "extrude",
            "cut",
            "workplane_offset",
            "sketch_on_plane",
        }
        allowed_planes = {"XOY", "XOZ", "YOZ"}
        allowed_dirs = {"normal", "reverse", "both"}

        for i, step in enumerate(data["steps"]):
            if not isinstance(step, dict):
                raise ValueError(f"Step #{i} must be an object.")
            act = (step.get("action") or "").lower().strip()
            if act not in allowed_actions:
                raise ValueError(f"Unknown action '{act}' in step #{i}.")

            if act == "sketch":
                plane = (step.get("plane") or "XOY").upper()
                if plane not in allowed_planes:
                    raise ValueError(f"Bad plane '{plane}' in step #{i}.")
                ents = step.get("entities", [])
                if not isinstance(ents, list) or len(ents) == 0:
                    raise ValueError(f"Sketch must have entities in step #{i}.")
                for e in ents:
                    et = (e.get("type") or "").lower()
                    if et not in ("line", "circle"):
                        raise ValueError(f"Bad entity type '{et}' in step #{i}.")
                    # minimal field checks:
                    if et == "line":
                        if "start" not in e or "end" not in e:
                            raise ValueError(f"Line must have start/end in step #{i}.")
                    if et == "circle":
                        if "center" not in e or "radius" not in e:
                            raise ValueError(
                                f"Circle must have center/radius in step #{i}."
                            )

            if act == "extrude":
                if "height" not in step:
                    raise ValueError(f"Extrude missing 'height' in step #{i}.")
                d = (step.get("direction") or "both").lower()
                if d not in allowed_dirs:
                    raise ValueError(f"Bad direction '{d}' in step #{i}.")

            if act == "cut":
                d = (step.get("direction") or "both").lower()
                if d not in allowed_dirs:
                    raise ValueError(f"Bad direction '{d}' in step #{i}.")
                through_all = bool(step.get("through_all", False))
                if not through_all and step.get("depth", None) is None:
                    # allowed: builder defaults to through_all if depth missing
                    pass

            if act == "workplane_offset":
                bp = (step.get("base_plane") or "").upper()
                if bp not in allowed_planes:
                    raise ValueError(f"Bad base_plane '{bp}' in step #{i}.")
                if "offset" not in step or "name" not in step:
                    raise ValueError(
                        f"workplane_offset must have offset and name in step #{i}."
                    )

            if act == "sketch_on_plane":
                # plane может отсутствовать — тогда по умолчанию XOY
                plane = (step.get("plane") or "XOY").upper()
                # plane может быть и именем смещенной плоскости, тогда не валидируем по allowed_planes
                ents = step.get("entities", [])
                if not isinstance(ents, list) or len(ents) == 0:
                    raise ValueError(
                        f"sketch_on_plane must have entities in step #{i}."
                    )

    def generate_llm_json(self, user_text: str) -> dict:
        llm = self.get_llm()
        prompt = make_llm_prompt(user_text)

        out = llm.create_chat_completion(
            messages=[
                {"role": "system", "content": "You output ONLY JSON. No extra text."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.0,
            max_tokens=900,
        )

        raw = out["choices"][0]["message"]["content"] or ""
        content = raw.strip()
        extracted = self._extract_json_object(content)

        try:
            data = json.loads(extracted)
        except Exception as e:
            raise LLMJSONError(str(e), raw=raw, extracted=extracted, prompt=prompt)

        # --- repair / normalize after parsing ---
        for st in data.get("steps", []):
            act = (st.get("action") or "").lower().strip()
            if act == "sketch_on_plane" and not (
                st.get("plane") or st.get("plane_name") or st.get("on_plane")
            ):
                st["plane"] = "XOY"

        try:
            self.validate_generated_json(data)
        except Exception as e:
            raise LLMJSONError(str(e), raw=raw, extracted=extracted, prompt=prompt)

        # store for debugging UI
        self.llm_raw = raw
        self.llm_extracted = extracted
        self.llm_prompt = prompt

        return data

    def on_generate_llm(self):
        text = self.prompt_var.get().strip()
        if not text:
            messagebox.showwarning("Empty", "Введите текстовый запрос")
            return

        if getattr(self, "_llm_busy", False):
            return

        self.set_llm_busy(True, status="LLM: generating...")

        def worker():
            try:
                self.log_write("LLM: generating JSON...")
                data = self.generate_llm_json(text)

                # success
                def ok():
                    self.llm_json = data
                    self.log_write("LLM JSON ready ✅")
                    self.set_llm_busy(False, status="LLM: idle")
                    self.show_json_window("LLM JSON output", data)

                self.after(0, ok)

            except LLMJSONError as e:

                def err(e=e):
                    self.log_write(
                        "LLM ERROR (parse/validate):\n" + traceback.format_exc()
                    )
                    # store debug fields even on failure
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

                def err2(e=e):
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
    # Template actions
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
    # UI helper
    # -----------------
    def show_json_window(self, title: str, data: dict):
        txt = json.dumps(data, ensure_ascii=False, indent=2)
        top = tk.Toplevel(self)
        top.title(title)
        top.geometry("760x560")
        t = tk.Text(top, wrap="none")
        t.pack(fill="both", expand=True)
        t.insert("1.0", txt)


def main():
    pythoncom.CoInitialize()
    try:
        app = App()
        app.mainloop()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
