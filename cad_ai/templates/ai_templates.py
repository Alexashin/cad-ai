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
            {"action": "extrude", "height": s},
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
            {"action": "extrude", "height": s},
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

    if w <= 0 or h <= 0 or thickness <= 0:
        raise ValueError("Plate dims must be > 0")
    if margin * 2 >= min(w, h):
        raise ValueError("Margin too large for given plate size")

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
            {"action": "extrude", "height": thickness},
            {"action": "sketch", "plane": plane, "entities": holes},
            {"action": "cut", "through_all": True, "direction": "both"},
        ],
    }


def tpl_stepped_block(
    w=120.0,
    h=80.0,
    base_z=20.0,
    step_w=60.0,
    step_h=40.0,
    step_z=20.0,
    pocket_w=30.0,
    pocket_h=20.0,
    pocket_depth=10.0,
    plane="XOY",
):
    w = float(w)
    h = float(h)
    base_z = float(base_z)
    step_w = float(step_w)
    step_h = float(step_h)
    step_z = float(step_z)
    pocket_w = float(pocket_w)
    pocket_h = float(pocket_h)
    pocket_depth = float(pocket_depth)

    if step_w > w or step_h > h:
        raise ValueError("Step must fit inside base footprint")
    if pocket_w > step_w or pocket_h > step_h:
        raise ValueError("Pocket must fit inside step footprint")

    sx0 = w - step_w
    sy0 = h - step_h

    px0 = sx0 + (step_w - pocket_w) / 2
    py0 = sy0 + (step_h - pocket_h) / 2

    return {
        "name": "Stepped block + pocket",
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
            {"action": "extrude", "height": base_z},
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "line", "start": [sx0, sy0], "end": [w, sy0]},
                    {"type": "line", "start": [w, sy0], "end": [w, h]},
                    {"type": "line", "start": [w, h], "end": [sx0, h]},
                    {"type": "line", "start": [sx0, h], "end": [sx0, sy0]},
                ],
            },
            {"action": "extrude", "height": step_z},
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "line", "start": [px0, py0], "end": [px0 + pocket_w, py0]},
                    {
                        "type": "line",
                        "start": [px0 + pocket_w, py0],
                        "end": [px0 + pocket_w, py0 + pocket_h],
                    },
                    {
                        "type": "line",
                        "start": [px0 + pocket_w, py0 + pocket_h],
                        "end": [px0, py0 + pocket_h],
                    },
                    {"type": "line", "start": [px0, py0 + pocket_h], "end": [px0, py0]},
                ],
            },
            {
                "action": "cut",
                "through_all": False,
                "depth": pocket_depth,
                "direction": "normal",
            },
        ],
    }


def tpl_angle_perforated(
    a=100.0,
    b=60.0,
    t=6.0,
    length=80.0,
    big_d=30.0,
    small_d=8.0,
    edge_x=15.0,
    edge_z=15.0,
):
    a = float(a)
    b = float(b)
    t = float(t)
    length = float(length)
    big_r = float(big_d) / 2.0
    small_r = float(small_d) / 2.0
    edge_x = float(edge_x)
    edge_z = float(edge_z)

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

    # XOZ inversion fix (as in your old code)
    holes_xoz = [
        {
            "type": "circle",
            "center": [c["center"][0], -c["center"][1]],
            "radius": c["radius"],
        }
        for c in holes_xoz
    ]

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

    def penalty(points):
        py0, py1_ = t - 1.0, b + 1.0
        pz0, pz1_ = -1.0, length + 1.0
        s = 0.0
        for yy, zz in points:
            if yy < py0:
                s += (py0 - yy) * (py0 - yy)
            if yy > py1_:
                s += (yy - py1_) * (yy - py1_)
            if zz < pz0:
                s += (pz0 - zz) * (pz0 - zz)
            if zz > pz1_:
                s += (zz - pz1_) * (zz - pz1_)
        return s

    def transform(points, swap=False, fy=False, fz=False):
        out = []
        for u, v in points:
            yy, zz = (v, u) if swap else (u, v)
            if fy:
                yy = -yy
            if fz:
                zz = -zz
            out.append((yy, zz))
        return out

    raw_pts = [tuple(h["center"]) for h in holes_yoz_raw]

    candidates = []
    for swap in (False, True):
        for fy in (False, True):
            for fz in (False, True):
                pts = transform(raw_pts, swap=swap, fy=fy, fz=fz)
                candidates.append((penalty(pts), swap, fy, fz, pts))

    candidates.sort(key=lambda x: x[0])
    _, best_swap, best_fy, best_fz, best_pts = candidates[0]

    holes_yoz = []
    for h, (yy, zz) in zip(holes_yoz_raw, best_pts):
        holes_yoz.append({"type": "circle", "center": [yy, zz], "radius": h["radius"]})

    return {
        "name": "Angle perforated",
        "steps": [
            {
                "action": "sketch",
                "plane": "XOY",
                "entities": [
                    {"type": "line", "start": [0, 0], "end": [a, 0]},
                    {"type": "line", "start": [a, 0], "end": [a, t]},
                    {"type": "line", "start": [a, t], "end": [t, t]},
                    {"type": "line", "start": [t, t], "end": [t, b]},
                    {"type": "line", "start": [t, b], "end": [0, b]},
                    {"type": "line", "start": [0, b], "end": [0, 0]},
                ],
            },
            {"action": "extrude", "height": length, "direction": "normal"},
            {"action": "sketch", "plane": "XOZ", "entities": holes_xoz},
            {"action": "cut", "through_all": True, "direction": "both"},
            {"action": "sketch", "plane": "YOZ", "entities": holes_yoz},
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
            a=p["a"],
            b=p["b"],
            t=p["t"],
            length=p["length"],
            big_d=p["big_d"],
            small_d=p["small_d"],
            edge_x=p["edge_x"],
            edge_z=p["edge_z"],
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
            w=p["w"],
            h=p["h"],
            base_z=p["base_z"],
            step_w=p["step_w"],
            step_h=p["step_h"],
            step_z=p["step_z"],
            pocket_w=p["pocket_w"],
            pocket_h=p["pocket_h"],
            pocket_depth=p["pocket_depth"],
            plane="XOY",
        ),
    },
}
