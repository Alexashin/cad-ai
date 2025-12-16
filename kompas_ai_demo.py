# -*- coding: utf-8 -*-
"""
KOMPAS-3D AI Demo (stable for presentation)
------------------------------------------
Single-file demo with a simple GUI (Tkinter):
- Connects to KOMPAS via COM (API5 + API7) using pywin32.
- Lets you pick a "neural-generated" example template and tweak parameters.
- Generates JSON-instructions and builds a 3D model in KOMPAS.
- Fixes the "through all hole" behavior (cut uses etThroughAll correctly).

How to run:
1) Put this file into the same folder as:
   - MiscellaneousHelpers.py
   - LDefin2D.py (optional, but typically shipped with examples)
   - (and any other helper files from your project)
2) Ensure KOMPAS-3D is installed and its COM servers are registered.
3) Install deps: pip install pywin32
4) Run: python kompas_ai_demo.py
"""

import json
import sys
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

import pythoncom
from win32com.client import Dispatch, gencache

# Local helper modules from your project/examples
import MiscellaneousHelpers as MH
import LDefin2D  # noqa: F401  (imported for compatibility with KOMPAS examples)


# -----------------------------
# COM / KOMPAS bootstrap
# -----------------------------

GUID_KS_CONST = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"
GUID_KS_CONST_3D = "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}"
GUID_API5 = "{0422828C-F174-495E-AC5D-D31014DBBE87}"
GUID_API7 = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"


def connect_kompas():
    """
    Connect to KOMPAS (API5 and API7), load constants, and return:
    (kompas6_constants, kompas6_constants_3d, kompas_object_api5, application_api7)
    """
    # Generate / load type libraries wrappers + constants
    ks_const = gencache.EnsureModule(GUID_KS_CONST, 0, 1, 0).constants
    ks_const_3d = gencache.EnsureModule(GUID_KS_CONST_3D, 0, 1, 0).constants

    api5 = gencache.EnsureModule(GUID_API5, 0, 1, 0)
    api7 = gencache.EnsureModule(GUID_API7, 0, 1, 0)

    # Connect to running KOMPAS instance or create it
    kompas5_disp = Dispatch("Kompas.Application.5")
    kompas7_disp = Dispatch("Kompas.Application.7")

    kompas_object = api5.KompasObject(
        kompas5_disp._oleobj_.QueryInterface(api5.KompasObject.CLSID, pythoncom.IID_IDispatch)
    )
    application = api7.IApplication(
        kompas7_disp._oleobj_.QueryInterface(api7.IApplication.CLSID, pythoncom.IID_IDispatch)
    )

    # Put into helper globals (used by some example helper scripts)
    MH.iKompasObject = kompas_object
    MH.iApplication = application

    # Make visible
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

    # ActiveDocument3D from API5 points to the newly created doc
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
        """
        Start a sketch on one of default planes: XOY / XOZ / YOZ
        """
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
    

    # --- Features ---
    def create_offset_plane(self, base_plane: str, offset: float, name: str):
        """
        Создаёт плоскость, параллельную базовой, со смещением offset (мм).
        base_plane: "XOY" / "XOZ" / "YOZ"
        name: ключ, по которому потом будем ссылаться
        """
        base_plane = base_plane.upper().strip()
        attr = f"o3d_plane{base_plane}"
        if not hasattr(self.ks_const_3d, attr):
            raise ValueError("base_plane must be XOY / XOZ / YOZ")

        # entity of plane
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


    
    def extrude_boss(self, height: float, direction: str = "both"):
        """
        Boss extrusion from last sketch.

        direction:
        - "normal"  -> dtNormal + typeNormal/depthNormal
        - "reverse" -> dtReverse + typeReverse/depthReverse
        - "both"    -> dtBoth with half depth to each side (stable demo behavior)
        """
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
            # safest: always visible regardless of sketch normal orientation
            p.direction = self.ks_const_3d.dtBoth
            half = height / 2.0
            p.typeNormal = self.ks_const_3d.etBlind
            p.depthNormal = half
            p.typeReverse = self.ks_const_3d.etBlind
            p.depthReverse = half

        extrusion.Create()


    def cut_extrusion(self, *, through_all: bool = True, depth: float | None = None, direction: str = "normal"):
        """
        Cut by last sketch.

        IMPORTANT (KOMPAS API nuance):
        - For dtNormal you must set typeNormal/depthNormal
        - For dtReverse you must set typeReverse/depthReverse
        Otherwise KOMPAS may keep default (often 10mm) and you get a shallow cut.

        Params:
        - through_all=True  -> etThroughAll (depth ignored)
        - through_all=False -> etBlind + depth (required)
        - direction: 'normal' or 'reverse'
        """

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
                # режем на половину в каждую сторону
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
        """
        Supported actions:
        - sketch: {plane: "XOY", entities: [...]}
        - extrude: {height: 100}
        - cut: {through_all: true} OR {depth: 20, through_all:false}
        Color is intentionally WIP for demo stability.
        """
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
                    step.get("height", 10),
                    direction=step.get("direction", "both")  # <-- по умолчанию both
                )

            elif action == "cut":
                through_all = bool(step.get("through_all", False))
                depth = step.get("depth", None)
                direction = step.get("direction", "normal")
                # If user didn't specify through_all but specified huge depth, still do ThroughAll
                if through_all:
                    self.cut_extrusion(through_all=True, direction=direction)
                else:
                    # Default for demo: if depth is missing -> ThroughAll
                    if depth is None:
                        self.cut_extrusion(through_all=True, direction=direction)
                    else:
                        self.cut_extrusion(through_all=False, depth=float(depth), direction=direction)

            elif action == "color":
                # Color is intentionally disabled for stability
                # (Material/appearance API integration is WIP)
                pass
            
            elif action == "workplane_offset":
                self.create_offset_plane(
                    base_plane=step["base_plane"],
                    offset=step["offset"],
                    name=step["name"],
                )

            elif action == "sketch_on_plane":
                doc2d = self.start_sketch_on_named_plane(step["name"])
                for ent in step.get("entities", []):
                    et = ent["type"]
                    if et == "circle":
                        x, y = ent["center"]
                        r = ent["radius"]
                        self.add_circle(doc2d, x, y, r)
                    elif et == "line":
                        x1, y1 = ent["start"]
                        x2, y2 = ent["end"]
                        self.add_line(doc2d, x1, y1, x2, y2)
                self.finish_sketch()

            else:
                raise ValueError(f"Unknown action: {step.get('action')}")


# -----------------------------
# "AI" templates (pre-generated)
# -----------------------------

def tpl_cube(size=100.0, plane="XOY"):
    s = float(size)
    return {
        "name": "Cube",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [s, 0]},
                {"type": "line", "start": [s, 0], "end": [s, s]},
                {"type": "line", "start": [s, s], "end": [0, s]},
                {"type": "line", "start": [0, s], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": s},
        ]
    }


def tpl_cube_with_through_hole(size=100.0, hole_d=20.0, plane="XOY"):
    s = float(size)
    d = float(hole_d)
    return {
        "name": "Cube + Through hole",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [s, 0]},
                {"type": "line", "start": [s, 0], "end": [s, s]},
                {"type": "line", "start": [s, s], "end": [0, s]},
                {"type": "line", "start": [0, s], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": s},
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [s / 2, s / 2], "radius": d / 2},
            ]},

            {"action": "cut", "through_all": True, "direction": "both"},
        ]
    }


def tpl_angle_perforated(
    a=100.0, b=60.0, t=6.0, length=80.0,
    big_d=30.0, small_d=8.0, edge_x=15.0, edge_z=15.0
):
    a = float(a); b = float(b); t = float(t); length = float(length)
    big_r = float(big_d) / 2.0
    small_r = float(small_d) / 2.0
    edge_x = float(edge_x)
    edge_z = float(edge_z)

    # ---------- HOLES ON HORIZONTAL LEG (XOZ): centers are [x, z] ----------
    x1 = max(t + edge_x, t + small_r + 0.1)
    x2 = min(a - edge_x, a - small_r - 0.1)

    z1 = max(edge_z, small_r + 0.1)
    z2 = min(length - edge_z, length - small_r - 0.1)

    cx = max(t + edge_x, min(a - edge_x, a * 0.6))
    cz = max(edge_z, min(length - edge_z, length * 0.5))

    holes_xoz = [
        {"type": "circle", "center": [cx, cz], "radius": big_r},
        {"type": "circle", "center": [x1, z1], "radius": small_r},
        {"type": "circle", "center": [x2, z1], "radius": small_r},
        {"type": "circle", "center": [x1, z2], "radius": small_r},
        {"type": "circle", "center": [x2, z2], "radius": small_r},
    ]

    # XOZ у тебя уже “лечится” инверсией Z
    holes_xoz = [{"type": "circle", "center": [c["center"][0], -c["center"][1]], "radius": c["radius"]} for c in holes_xoz]

    # ---------- HOLES ON VERTICAL LEG (YOZ): base centers are [y, z] ----------
    y1 = max(t + edge_x, t + small_r + 0.1)
    y2 = min(b - edge_x, b - small_r - 0.1)
    cy = max(t + edge_x, min(b - edge_x, b * 0.6))

    holes_yoz_raw = [
        {"type": "circle", "center": [cy, cz], "radius": big_r},
        {"type": "circle", "center": [y1, z1], "radius": small_r},
        {"type": "circle", "center": [y2, z1], "radius": small_r},
        {"type": "circle", "center": [y1, z2], "radius": small_r},
        {"type": "circle", "center": [y2, z2], "radius": small_r},
    ]

    # --- auto-fit YOZ orientation ---
    # Мы хотим, чтобы центры после трансформации попали примерно в диапазоны:
    # y ~ [t..b], z ~ [0..length]
    def penalty(points):
        # штраф за выход из диапазонов (чем больше - тем хуже)
        py0, py1 = t - 1.0, b + 1.0
        pz0, pz1 = -1.0, length + 1.0
        s = 0.0
        for (yy, zz) in points:
            if yy < py0: s += (py0 - yy) * (py0 - yy)
            if yy > py1: s += (yy - py1) * (yy - py1)
            if zz < pz0: s += (pz0 - zz) * (pz0 - zz)
            if zz > pz1: s += (zz - pz1) * (zz - pz1)
        return s

    # 8 вариантов: без swap и со swap, плюс флипы
    def transform(points, swap=False, fy=False, fz=False):
        out = []
        for (u, v) in points:
            yy, zz = (v, u) if swap else (u, v)
            if fy: yy = -yy
            if fz: zz = -zz
            out.append((yy, zz))
        return out

    raw_pts = [tuple(h["center"]) for h in holes_yoz_raw]

    candidates = []
    for swap in (False, True):
        for fy in (False, True):
            for fz in (False, True):
                pts = transform(raw_pts, swap=swap, fy=fy, fz=fz)
                candidates.append((penalty(pts), swap, fy, fz, pts))

    # выбираем ориентацию с минимальным штрафом
    candidates.sort(key=lambda x: x[0])
    _, best_swap, best_fy, best_fz, best_pts = candidates[0]

    holes_yoz = []
    for h, (yy, zz) in zip(holes_yoz_raw, best_pts):
        holes_yoz.append({"type": "circle", "center": [yy, zz], "radius": h["radius"]})

    return {
        "name": "Angle perforated",
        "steps": [
            {"action": "sketch", "plane": "XOY", "entities": [
                {"type": "line", "start": [0, 0], "end": [a, 0]},
                {"type": "line", "start": [a, 0], "end": [a, t]},
                {"type": "line", "start": [a, t], "end": [t, t]},
                {"type": "line", "start": [t, t], "end": [t, b]},
                {"type": "line", "start": [t, b], "end": [0, b]},
                {"type": "line", "start": [0, b], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": length, "direction": "normal"},

            {"action": "sketch", "plane": "XOZ", "entities": holes_xoz},
            {"action": "cut", "through_all": True, "direction": "both"},

            {"action": "sketch", "plane": "YOZ", "entities": holes_yoz},
            {"action": "cut", "through_all": True, "direction": "both"},
        ]
    }





def tpl_plate_with_holes(w=120.0, h=80.0, thickness=8.0, hole_d=10.0, margin=15.0, plane="XOY"):
    """
    Rectangular plate extruded along +Z, with 4 through holes.
    """
    w = float(w); h = float(h); thickness = float(thickness)
    hole_d = float(hole_d); margin = float(margin)
    if w <= 0 or h <= 0 or thickness <= 0:
        raise ValueError("Plate dims must be > 0")
    if margin*2 >= min(w, h):
        raise ValueError("Margin too large for given plate size")

    holes = [
        {"type": "circle", "center": [margin, margin], "radius": hole_d/2},
        {"type": "circle", "center": [w - margin, margin], "radius": hole_d/2},
        {"type": "circle", "center": [w - margin, h - margin], "radius": hole_d/2},
        {"type": "circle", "center": [margin, h - margin], "radius": hole_d/2},
    ]

    return {
        "name": "Plate (4 holes)",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [w, 0]},
                {"type": "line", "start": [w, 0], "end": [w, h]},
                {"type": "line", "start": [w, h], "end": [0, h]},
                {"type": "line", "start": [0, h], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": thickness},
            {"action": "sketch", "plane": plane, "entities": holes},
            {"action": "cut", "through_all": True, "direction": "both"},
        ]
    }


def tpl_stepped_block(w=120.0, h=80.0, base_z=20.0, step_w=60.0, step_h=40.0, step_z=20.0, pocket_w=30.0, pocket_h=20.0, pocket_depth=10.0, plane="XOY"):
    """
    Stepped block:
    - Base rectangle (w x h) extruded base_z
    - Step rectangle (step_w x step_h) extruded step_z (on same base plane for demo simplicity)
    - Pocket (pocket_w x pocket_h) cut blind pocket_depth
    Note: Without face-sketching, the 'step' is placed from the same base plane (XOY),
          so it becomes a second boss extrusion starting at Z=0 (will fuse). It's still a valid "more complex" demo shape.
    """
    w=float(w); h=float(h); base_z=float(base_z)
    step_w=float(step_w); step_h=float(step_h); step_z=float(step_z)
    pocket_w=float(pocket_w); pocket_h=float(pocket_h); pocket_depth=float(pocket_depth)

    if step_w > w or step_h > h:
        raise ValueError("Step must fit inside base footprint")
    if pocket_w > step_w or pocket_h > step_h:
        raise ValueError("Pocket must fit inside step footprint")

    # Step placed in top-right corner of base footprint
    sx0 = w - step_w
    sy0 = h - step_h

    # Pocket centered in the step
    px0 = sx0 + (step_w - pocket_w) / 2
    py0 = sy0 + (step_h - pocket_h) / 2

    return {
        "name": "Stepped block + pocket",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [w, 0]},
                {"type": "line", "start": [w, 0], "end": [w, h]},
                {"type": "line", "start": [w, h], "end": [0, h]},
                {"type": "line", "start": [0, h], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": base_z},

            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [sx0, sy0], "end": [w, sy0]},
                {"type": "line", "start": [w, sy0], "end": [w, h]},
                {"type": "line", "start": [w, h], "end": [sx0, h]},
                {"type": "line", "start": [sx0, h], "end": [sx0, sy0]},
            ]},
            {"action": "extrude", "height": step_z},

            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [px0, py0], "end": [px0 + pocket_w, py0]},
                {"type": "line", "start": [px0 + pocket_w, py0], "end": [px0 + pocket_w, py0 + pocket_h]},
                {"type": "line", "start": [px0 + pocket_w, py0 + pocket_h], "end": [px0, py0 + pocket_h]},
                {"type": "line", "start": [px0, py0 + pocket_h], "end": [px0, py0]},
            ]},
            {"action": "cut", "through_all": False, "depth": pocket_depth, "direction": "normal"},
        ]
    }
TEMPLATES = {
    "Куб (AI)": {
        "params": [("size", "Размер куба", 100.0)],
        "build": lambda p: tpl_cube(size=p["size"], plane="XOY"),
    },
    "Куб с отверстием насквозь (AI)": {
        "params": [("size", "Размер куба", 100.0), ("hole_d", "Диаметр отверстия", 20.0)],
        "build": lambda p: tpl_cube_with_through_hole(size=p["size"], hole_d=p["hole_d"], plane="XOY"),
    },
    "Уголок перфорированный (AI)": {
    "params": [
        ("a", "Полка A", 100.0),
        ("b", "Полка B", 60.0),
        ("t", "Толщина t", 6.0),
        ("length", "Длина", 80.0),
        ("big_d", "Большое отверстие Ø", 30.0),
        ("small_d", "Малые отверстия Ø", 8.0),
        ("edge_x", "Отступ от края (X/Y)", 15.0),
        ("edge_z", "Отступ по длине (Z)", 15.0),
    ],
    "build": lambda p: tpl_angle_perforated(
        a=p["a"], b=p["b"], t=p["t"], length=p["length"],
        big_d=p["big_d"], small_d=p["small_d"],
        edge_x=p["edge_x"], edge_z=p["edge_z"],
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
            w=p["w"], h=p["h"], thickness=p["thickness"],
            hole_d=p["hole_d"], margin=p["margin"], plane="XOY"
        ),
    },
    "Ступенчатый блок + карман (AI)": {
        "params": [
            ("w", "Ширина основания", 120.0),
            ("h", "Высота основания", 80.0),
            ("base_z", "Толщина основания", 20.0),
            ("step_w", "Ширина ступени", 60.0),
            ("step_h", "Высота ступени", 40.0),
            ("step_z", "Высота ступени (выдавл.)", 20.0),
            ("pocket_w", "Ширина кармана", 30.0),
            ("pocket_h", "Высота кармана", 20.0),
            ("pocket_depth", "Глубина кармана", 10.0),
        ],
        "build": lambda p: tpl_stepped_block(
            w=p["w"], h=p["h"], base_z=p["base_z"],
            step_w=p["step_w"], step_h=p["step_h"], step_z=p["step_z"],
            pocket_w=p["pocket_w"], pocket_h=p["pocket_h"], pocket_depth=p["pocket_depth"],
            plane="XOY"
        ),
    }
}


# -----------------------------
# GUI
# -----------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KOMPAS-3D AI Demo (stable)")
        self.geometry("640x520")
        self.minsize(640, 520)

        self.ks_const = None
        self.ks_const_3d = None
        self.api5 = None
        self.api7 = None
        self.kompas_object = None
        self.application = None

        self.builder = None
        self.iPart = None

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
        self.params_frame = ttk.LabelFrame(root, text="Parameters", padding=10)
        self.params_frame.pack(fill="x", pady=(12, 0))
        self.param_entries = {}
        self.render_params()

        # Options panel
        opts = ttk.LabelFrame(root, text="Options", padding=10)
        opts.pack(fill="x", pady=(12, 0))

        self.new_doc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Create new document each build", variable=self.new_doc_var).pack(anchor="w")

        # Buttons
        btns = ttk.Frame(root)
        btns.pack(fill="x", pady=(12, 0))

        ttk.Button(btns, text="Build in KOMPAS", command=self.on_build).pack(side="left")
        ttk.Button(btns, text="Show JSON", command=self.on_show_json).pack(side="left", padx=8)
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

    def read_params(self):
        p = {}
        for k, ent in self.param_entries.items():
            v = ent.get().strip().replace(",", ".")
            p[k] = float(v)
        return p

    def ensure_connected(self):
        if self.application is None or self.ks_const is None:
            raise RuntimeError("Not connected. Click 'Connect' first.")

    def on_connect(self):
        try:
            self.log_write("Connecting to KOMPAS...")
            ks_const, ks_const_3d, api5, api7, kompas_object, application = connect_kompas()
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

    def build_json(self):
        tpl = TEMPLATES[self.template_var.get()]
        params = self.read_params()
        data = tpl["build"](params)
        return data

    def on_show_json(self):
        try:
            data = self.build_json()
            txt = json.dumps(data, ensure_ascii=False, indent=2)
            top = tk.Toplevel(self)
            top.title("Generated JSON (AI output)")
            top.geometry("720x520")
            t = tk.Text(top, wrap="none")
            t.pack(fill="both", expand=True)
            t.insert("1.0", txt)
        except Exception as e:
            messagebox.showerror("JSON error", str(e))

    def on_build(self):
        try:
            self.ensure_connected()

            # Create new document if requested (recommended for demo)
            if self.new_doc_var.get() or self.iPart is None:
                self.log_write("Creating new 3D Part document...")
                _, _, _, iPart = new_document_part(
                    self.ks_const, self.ks_const_3d, self.api5, self.api7, self.kompas_object, self.application
                )
                self.iPart = iPart
                self.builder = Kompas3DBuilder(self.ks_const, self.ks_const_3d, self.iPart)

            data = self.build_json()
            self.log_write(f"Building template: {data.get('name', self.template_var.get())}")
            self.builder.process_json(data)
            self.log_write("Build complete ✅")

            messagebox.showinfo("Done", "Model построена в KOMPAS.")
        except Exception as e:
            self.log_write("ERROR while building:\n" + traceback.format_exc())
            messagebox.showerror("Build error", str(e))


def main():
    # Tkinter + COM sometimes benefit from initializing COM in STA
    pythoncom.CoInitialize()
    try:
        app = App()
        app.mainloop()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
