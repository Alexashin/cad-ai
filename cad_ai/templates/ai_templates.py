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
