import math


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

# НУЖНО ДОБАВИТЬ ВАРИАНТ ВЫДАВЛИВАНИЯ/ВЫРЕЗАНИЯ ПОД УГЛОМ ДЛЯ ШТУК ПО ТИПУ ЗЕНКОВКИ И Т.П.
def tpl_cylinder_countersink(d=40.0, h=30.0, hole_d=10.0, cs_d=10.0, cs_depth=5.0, plane="XOY"):
    return {
        "name": "Cylinder + countersink",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [0, 0], "radius": d / 2}
            ]},
            {"action": "extrude", "height": h},

            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [0, 0], "radius": hole_d / 2}
            ]},
            {"action": "cut", "through_all": True},

            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [0, 0], "radius": cs_d / 2}
            ]},
            {"action": "cut", "through_all": False, "depth": cs_depth},
        ]
    }

def tpl_plate_with_slot(w=120.0, h=60.0, t=10.0, slot_w=12.0, slot_l=80.0, plane="XOY"):
    sx = (w - slot_l) / 2
    sy = (h - slot_w) / 2

    return {
        "name": "Plate with slot",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [w, 0]},
                {"type": "line", "start": [w, 0], "end": [w, h]},
                {"type": "line", "start": [w, h], "end": [0, h]},
                {"type": "line", "start": [0, h], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": t},

            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [sx, sy], "end": [sx + slot_l, sy]},
                {"type": "line", "start": [sx + slot_l, sy], "end": [sx + slot_l, sy + slot_w]},
                {"type": "line", "start": [sx + slot_l, sy + slot_w], "end": [sx, sy + slot_w]},
                {"type": "line", "start": [sx, sy + slot_w], "end": [sx, sy]},
            ]},
            {"action": "cut", "through_all": True},
        ]
    }

def tpl_flange_with_holes(d_outer=120.0, d_inner=40.0, h=15.0, holes_count=6, bolt_d=10.0, pcd=90.0, plane="XOY"):
    ents = [
        {"type": "circle", "center": [0, 0], "radius": d_outer / 2},
        {"type": "circle", "center": [0, 0], "radius": d_inner / 2},
    ]

    holes = []
    for i in range(int(holes_count)):
        a = 2 * math.pi * i / holes_count
        holes.append({
            "type": "circle",
            "center": [pcd / 2 * math.cos(a), pcd / 2 * math.sin(a)],
            "radius": bolt_d / 2
        })

    return {
        "name": "Flange",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": ents},
            {"action": "extrude", "height": h, "direction": "reverse"},
            {"action": "sketch", "plane": plane, "entities": holes},
            {"action": "cut", "through_all": True},
        ]
    }

def tpl_stepped_shaft(d1=40.0, l1=50.0, d2=30.0, l2=40.0, plane="XOY"):
    return {
        "name": "Stepped shaft",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [0, 0], "radius": d1 / 2}
            ]},
            {"action": "extrude", "height": l1},

            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [0, 0], "radius": d2 / 2}
            ]},
            {"action": "extrude", "height": l2},
        ]
    }



def tpl_perforated_plate(w=120.0, h=80.0, t=8.0, holes_x=5, holes_y=3, d=6.0, pitch=20.0, plane="XOY"):
    holes = []
    ox = (w - (holes_x - 1) * pitch) / 2
    oy = (h - (holes_y - 1) * pitch) / 2

    for iy in range(holes_y):
        for ix in range(holes_x):
            holes.append({
                "type": "circle",
                "center": [ox + ix * pitch, oy + iy * pitch],
                "radius": d / 2
            })

    return {
        "name": "Perforated plate",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [w, 0]},
                {"type": "line", "start": [w, 0], "end": [w, h]},
                {"type": "line", "start": [w, h], "end": [0, h]},
                {"type": "line", "start": [0, h], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": t},
            {"action": "sketch", "plane": plane, "entities": holes},
            {"action": "cut", "through_all": True},
        ]
    }


def tpl_box_with_pocket(w=100.0, h=60.0, base_z=30.0, pocket_w=60.0, pocket_h=30.0, pocket_depth=10.0, plane="XOY"):
    w = float(w); h = float(h); base_z = float(base_z)
    pocket_w = float(pocket_w); pocket_h = float(pocket_h); pocket_depth = float(pocket_depth)
    px0 = (w - pocket_w) / 2
    py0 = (h - pocket_h) / 2
    return {
        "name": "Box with pocket",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [w, 0]},
                {"type": "line", "start": [w, 0], "end": [w, h]},
                {"type": "line", "start": [w, h], "end": [0, h]},
                {"type": "line", "start": [0, h], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": base_z},
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [px0, py0], "end": [px0 + pocket_w, py0]},
                {"type": "line", "start": [px0 + pocket_w, py0], "end": [px0 + pocket_w, py0 + pocket_h]},
                {"type": "line", "start": [px0 + pocket_w, py0 + pocket_h], "end": [px0, py0 + pocket_h]},
                {"type": "line", "start": [px0, py0 + pocket_h], "end": [px0, py0]},
            ]},
            {"action": "cut", "through_all": False, "depth": pocket_depth}
        ]
    }

def tpl_bracket_hole_offset_plane(a=80.0, b=40.0, t=6.0, length=60.0, hole_d=10.0, offset_z=20.0, plane="XOZ"):
    # Простой уголок с отверстием на смещённой плоскости
    a = float(a); b = float(b); t = float(t); length = float(length)
    hole_r = float(hole_d) / 2.0
    offset_z = float(offset_z)
    cx = a / 2; cz = 0; 
    return {
        "name": "Bracket hole on offset plane",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "line", "start": [0, 0], "end": [a, 0]},
                {"type": "line", "start": [a, 0], "end": [a, t]},
                {"type": "line", "start": [a, t], "end": [t, t]},
                {"type": "line", "start": [t, t], "end": [t, b]},
                {"type": "line", "start": [t, b], "end": [0, b]},
                {"type": "line", "start": [0, b], "end": [0, 0]},
            ]},
            {"action": "extrude", "height": length},
            {"action": "workplane_offset", "base_plane": "XOY", "offset": offset_z, "name": "hole_plane"},
            {"action": "sketch_on_plane", "name": "hole_plane", "entities": [
                {"type": "circle", "center": [cx, cz], "radius": hole_r}
            ]},
            {"action": "cut", "through_all": True}
        ]
    }

def tpl_plate_multi_blind_holes(w=100.0, h=60.0, thickness=10.0, holes=None, plane="XOY"):
    w = float(w); h = float(h); thickness = float(thickness)

    if holes is None:
        holes = []

    outline = [
        {"type": "line", "start": [0, 0], "end": [w, 0]},
        {"type": "line", "start": [w, 0], "end": [w, h]},
        {"type": "line", "start": [w, h], "end": [0, h]},
        {"type": "line", "start": [0, h], "end": [0, 0]},
    ]

    steps = [
        {"action": "sketch", "plane": plane, "entities": outline},
        {"action": "extrude", "height": thickness},
    ]

    for hinfo in holes:
        steps.append({
            "action": "sketch",
            "plane": plane,
            "entities": [
                {
                    "type": "circle",
                    "center": [hinfo["x"], hinfo["y"]],
                    "radius": float(hinfo["d"]) / 2.0
                }
            ]
        })
        steps.append({
            "action": "cut",
            "through_all": False,
            "depth": float(hinfo["depth"]),
            "direction": "normal"
        })

    return {"name": "Plate multi blind holes", "steps": steps}

def tpl_ribbed_beam(w=120.0, h=30.0, thickness=10.0, rib_w=4.0, rib_h=20.0, rib_count=4, spacing=20.0, plane="XOY"):
    w = float(w); h = float(h); thickness = float(thickness)
    rib_w = float(rib_w); rib_h = float(rib_h); rib_count = int(rib_count); spacing = float(spacing)
    # base rectangle
    outline = [
        {"type": "line", "start": [0, 0], "end": [w, 0]},
        {"type": "line", "start": [w, 0], "end": [w, h]},
        {"type": "line", "start": [w, h], "end": [0, h]},
        {"type": "line", "start": [0, h], "end": [0, 0]},
    ]
    steps = [
        {"action": "sketch", "plane": plane, "entities": outline},
        {"action": "extrude", "height": thickness},
    ]
    # ribs
    for i in range(rib_count):
        x0 = spacing + i * spacing
        if x0 + rib_w > w:
            break
        steps.append({"action": "sketch", "plane": plane, "entities": [
            {"type": "line", "start": [x0, 0], "end": [x0 + rib_w, 0]},
            {"type": "line", "start": [x0 + rib_w, 0], "end": [x0 + rib_w, h]},
            {"type": "line", "start": [x0 + rib_w, h], "end": [x0, h]},
            {"type": "line", "start": [x0, h], "end": [x0, 0]},
        ]})
        steps.append({"action": "extrude", "height": rib_h})
    return {"name": "Ribbed beam", "steps": steps}

def tpl_cover_with_slot(w=100.0, h=60.0, thickness=8.0, slot_len=60.0, slot_w=10.0, slot_pos=30.0, plane="XOY"):
    w = float(w); h = float(h); thickness = float(thickness)
    slot_len = float(slot_len); slot_w = float(slot_w); slot_pos = float(slot_pos)
    # outline
    outline = [
        {"type": "line", "start": [0, 0], "end": [w, 0]},
        {"type": "line", "start": [w, 0], "end": [w, h]},
        {"type": "line", "start": [w, h], "end": [0, h]},
        {"type": "line", "start": [0, h], "end": [0, 0]},
    ]
    # slot (прямоугольник + 2 полуокружности)
    sx = (w - slot_len) / 2
    sy = slot_pos
    slot = [
        {"type": "line", "start": [sx, sy], "end": [sx + slot_len, sy]},
        {"type": "line", "start": [sx + slot_len, sy], "end": [sx + slot_len, sy + slot_w]},
        {"type": "line", "start": [sx + slot_len, sy + slot_w], "end": [sx, sy + slot_w]},
        {"type": "line", "start": [sx, sy + slot_w], "end": [sx, sy]},
        # полуокружности можно добавить отдельными sketch/cut, если нужно
    ]
    return {
        "name": "Cover with slot",
        "steps": [
            {"action": "sketch", "plane": plane, "entities": outline},
            {"action": "extrude", "height": thickness},
            {"action": "sketch", "plane": plane, "entities": slot},
            {"action": "cut", "through_all": True}
        ]
    }

def tpl_bolt(
    thread_d=10.0,      # Диаметр резьбы (M10)
    shank_length=90.0,  # Длина стержня
    head_d=20.0,        # Диаметр головки
    head_height=10.0,   # Высота головки
    plane="XOY"
    ):
    """
    Болт с головкой и стержнем.
    """
    thread_d = float(thread_d)
    shank_length = float(shank_length)
    head_d = float(head_d)
    head_height = float(head_height)
    
    # Радиусы
    shank_radius = thread_d / 2.0
    head_radius = head_d / 2.0
    
    return {
        "name": "Болт",
        "steps": [
            # Стержень болта
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [0, 0], "radius": shank_radius}
            ]},
            {"action": "extrude", "height": shank_length, "direction": "normal"},
            
            # Головка болта
            {"action": "sketch", "plane": plane, "entities": [
                {"type": "circle", "center": [0, 0], "radius": head_radius}
            ]},
            {"action": "extrude", "height": head_height, "direction": "normal"},
            
            # Шестигранный вырез в головке (опционально, если хотите добавить)
            # Для простоты демо оставим без шестигранника
        ]
    }

def tpl_radial_bearing(
    d_inner=20.0,      # Внутренний диаметр
    d_outer=47.0,      # Наружный диаметр  
    width=14.0,        # Ширина подшипника
    chamfer=1.0,       # Фаска на кольцах
    plane="XOY"
):
    """Радиальный шарикоподшипник (упрощённая модель)"""
    d_inner = float(d_inner)
    d_outer = float(d_outer)
    width = float(width)
    
    # Проверка параметров
    if d_outer <= d_inner:
        raise ValueError("Наружный диаметр должен быть больше внутреннего")
    if width <= 0:
        raise ValueError("Ширина подшипника должна быть > 0")
    
    # Вычисление средней линии для дорожки качения
    d_mid = (d_inner + d_outer) / 2
    ball_d = (d_outer - d_inner - 4) / 4  # Упрощённый диаметр шарика
    
    return {
        "name": "Радиальный подшипник",
        "steps": [
            # Наружное кольцо
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "circle", "center": [0, 0], "radius": d_outer/2},
                    {"type": "circle", "center": [0, 0], "radius": d_outer/2 - width/2},
                ]
            },
            {"action": "extrude", "height": width, "direction": "normal"},
            
            # Внутреннее кольцо  
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {"type": "circle", "center": [0, 0], "radius": d_inner/2 + width/2},
                    {"type": "circle", "center": [0, 0], "radius": d_inner/2},
                ]
            },
            {"action": "extrude", "height": width, "direction": "normal"},
            
 
            # Отверстия в сепараторе для шариков (упрощённо - 8 отверстий)
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    *[
                        {
                            "type": "circle",
                            "center": [
                                d_mid/2 * math.cos(2*math.pi*i/8),
                                d_mid/2 * math.sin(2*math.pi*i/8)
                            ],
                            "radius": ball_d/2 * 0.8
                        }
                        for i in range(8)
                    ]
                ]
            },
            {"action": "cut", "through_all": True, "direction": "reverse"},
        ]
    }

def tpl_nut(
    thread_d=10.0,        # Диаметр резьбы (M10)
    flat_width=16.0,      # Ширина между параллельными гранями
    thickness=8.0,        # Толщина гайки
    plane="XOY"
):
    """
    Гайка шестигранная
    """
    import math
    
    thread_d = float(thread_d)
    flat_width = float(flat_width)
    thickness = float(thickness)
    
    # Радиус описанной окружности шестигранника
    circumradius = flat_width / math.sqrt(3)
    
    # Создаем точки шестигранника
    hex_points = []
    for i in range(6):
        angle = math.radians(30) + i * math.radians(60)  # 30°, 90°, 150°, 210°, 270°, 330°
        x = circumradius * math.cos(angle)
        y = circumradius * math.sin(angle)
        hex_points.append([x, y])
    
    # Сущности шестигранника
    hex_entities = []
    for i in range(6):
        hex_entities.append({
            "type": "line",
            "start": hex_points[i],
            "end": hex_points[(i + 1) % 6]
        })
    
    return {
        "name": "Гайка шестигранная",
        "steps": [
            # Шестигранник - основание гайки
            {
                "action": "sketch",
                "plane": plane,
                "entities": hex_entities
            },
            {
                "action": "extrude",
                "height": thickness,
                "direction": "normal"
            },
            
            # Резьбовое отверстие
            {
                "action": "sketch",
                "plane": plane,
                "entities": [
                    {
                        "type": "circle",
                        "center": [0, 0],
                        "radius": thread_d / 2.0
                    }
                ]
            },
            {
                "action": "cut",
                "through_all": True,
                "direction": "reverse"
            },
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
    },
    "Крышка со шпоночным пазом (AI)": {
        "params": [
            ("w", "Ширина", 100.0),
            ("h", "Высота", 60.0),
            ("thickness", "Толщина", 8.0),
            ("slot_len", "Длина паза", 60.0),
            ("slot_w", "Ширина паза", 10.0),
            ("slot_pos", "Положение паза по Y", 30.0),
        ],
        "build": lambda p: tpl_cover_with_slot(
            w=p["w"], h=p["h"], thickness=p["thickness"], slot_len=p["slot_len"], slot_w=p["slot_w"], slot_pos=p["slot_pos"], plane="XOY"
        ),
    },
    "Коробка с карманом (AI)": {
        "params": [
            ("w", "Ширина", 100.0),
            ("h", "Высота", 60.0),
            ("base_z", "Толщина основания", 30.0),
            ("pocket_w", "Ширина кармана", 60.0),
            ("pocket_h", "Высота кармана", 30.0),
            ("pocket_depth", "Глубина кармана", 10.0),
        ],
        "build": lambda p: tpl_box_with_pocket(
            w=p["w"], h=p["h"], base_z=p["base_z"],
            pocket_w=p["pocket_w"], pocket_h=p["pocket_h"], pocket_depth=p["pocket_depth"], plane="XOY"
        ),
    },
    "Уголок с отверстием на смещённой плоскости (AI)": {
        "params": [
            ("a", "Полка A", 80.0),
            ("b", "Полка B", 40.0),
            ("t", "Толщина t", 6.0),
            ("length", "Длина", 60.0),
            ("hole_d", "Диаметр отверстия", 10.0),
            ("offset_z", "Смещение плоскости", 20.0),
        ],
        "build": lambda p: tpl_bracket_hole_offset_plane(
            a=p["a"], b=p["b"], t=p["t"], length=p["length"], hole_d=p["hole_d"], offset_z=p["offset_z"], plane="XOZ"
        ),
    },
    "Пластина с несколькими слепыми отверстиями (AI)": {
        "params": [
            ("w", "Ширина", 100.0),
            ("h", "Высота", 60.0),
            ("thickness", "Толщина", 10.0),
            ("holes_count", "Число отверстий", 2),
        ],
        "build": lambda p, holes=None: tpl_plate_multi_blind_holes(
            w=p["w"],
            h=p["h"],
            thickness=p["thickness"],
            holes=holes,
            plane="XOY"
        ),
    },
    "Балка с рёбрами (AI)": {
        "params": [
            ("w", "Ширина", 120.0),
            ("h", "Высота", 30.0),
            ("thickness", "Толщина", 10.0),
            ("rib_w", "Ширина ребра", 4.0),
            ("rib_h", "Высота ребра", 20.0),
            ("rib_count", "Число рёбер", 4),
            ("spacing", "Шаг рёбер", 20.0),
        ],
        "build": lambda p: tpl_ribbed_beam(
            w=p["w"], h=p["h"], thickness=p["thickness"], rib_w=p["rib_w"], rib_h=p["rib_h"], rib_count=p["rib_count"], spacing=p["spacing"], plane="XOY"
        ),
    },
    "Цилиндр с зенковкой (AI)": {
        "params": [
            ("d", "Диаметр", 40.0),
            ("h", "Высота", 30.0),
            ("hole_d", "Диаметр отверстия", 10.0),
            ("cs_d", "Диаметр зенковки", 10.0),
            ("cs_depth", "Глубина зенковки", 5.0),
        ],
        "build": lambda p: tpl_cylinder_countersink(
            d=p["d"],
            h=p["h"],
            hole_d=p["hole_d"],
            cs_d=p["cs_d"],
            cs_depth=p["cs_depth"],
        ),
    },
    "Пластина с пазом (AI)": {
        "params": [
            ("w", "Ширина", 120.0),
            ("h", "Высота", 60.0),
            ("t", "Толщина", 10.0),
            ("slot_w", "Ширина паза", 12.0),
            ("slot_l", "Длина паза", 80.0),
        ],
        "build": lambda p: tpl_plate_with_slot(
            w=p["w"],
            h=p["h"],
            t=p["t"],
            slot_w=p["slot_w"],
            slot_l=p["slot_l"],
        ),
    },
    "Фланец с отверстиями (AI)": {
        "params": [
            ("d_outer", "Наружный диаметр", 120.0),
            ("d_inner", "Внутренний диаметр", 40.0),
            ("h", "Толщина", 15.0),
            ("holes_count", "Число отверстий", 6),
            ("bolt_d", "Диаметр отверстий", 10.0),
            ("pcd", "Диаметр окружности", 90.0),
        ],
        "build": lambda p: tpl_flange_with_holes(
            d_outer=p["d_outer"],
            d_inner=p["d_inner"],
            h=p["h"],
            holes_count=int(p["holes_count"]),
            bolt_d=p["bolt_d"],
            pcd=p["pcd"],
        ),
    },
    "Ступенчатый вал (AI)": {
        "params": [
            ("d1", "Диаметр 1", 40.0),
            ("l1", "Длина 1", 50.0),
            ("d2", "Диаметр 2", 30.0),
            ("l2", "Длина 2", 40.0),
        ],
        "build": lambda p: tpl_stepped_shaft(
            d1=p["d1"],
            l1=p["l1"],
            d2=p["d2"],
            l2=p["l2"],
        ),
    },
    "Перфорированная пластина (AI)": {
        "params": [
            ("w", "Ширина", 120.0),
            ("h", "Высота", 80.0),
            ("t", "Толщина", 8.0),
            ("holes_x", "Отверстий по X", 5),
            ("holes_y", "Отверстий по Y", 3),
            ("d", "Диаметр отверстий", 6.0),
            ("pitch", "Шаг", 20.0),
        ],
        "build": lambda p: tpl_perforated_plate(
            w=p["w"],
            h=p["h"],
            t=p["t"],
            holes_x=int(p["holes_x"]),
            holes_y=int(p["holes_y"]),
            d=p["d"],
            pitch=p["pitch"],
        ),
    },
    "Болт M10 (AI)": {
        "params": [
            ("thread_d", "Диаметр резьбы", 10.0),
            ("shank_length", "Длина стержня", 90.0),
            ("head_d", "Диаметр головки", 20.0),
            ("head_height", "Высота головки", 10.0),
        ],
        "build": lambda p: tpl_bolt(
            thread_d=p["thread_d"],
            shank_length=p["shank_length"],
            head_d=p["head_d"],
            head_height=p["head_height"],
            plane="XOY"
        ),
    },
    "Радиальный подшипник (AI)": {
        "params": [
            ("d_inner", "Внутренний диаметр, мм", 20.0),
            ("d_outer", "Наружный диаметр, мм", 47.0),
            ("width", "Ширина, мм", 14.0),
        ],
        "build": lambda p: tpl_radial_bearing(
            d_inner=p["d_inner"],
            d_outer=p["d_outer"],
            width=p["width"],
            plane="XOY"
        ),
    },
     "Гайка шестигранная M10 (AI)": {
        "params": [
            ("thread_d", "Диаметр резьбы, мм", 10.0),
            ("flat_width", "Ширина между гранями, мм", 16.0),
            ("thickness", "Толщина гайки, мм", 8.0),
        ],
        "build": lambda p: tpl_nut(
            thread_d=p["thread_d"],
            flat_width=p["flat_width"],
            thickness=p["thickness"],
            plane="XOY"
        ),
    },
}
