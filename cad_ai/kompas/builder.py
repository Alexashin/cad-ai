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
