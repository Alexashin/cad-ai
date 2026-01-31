import json
import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH

# Константы Kompas API (API5) для типов объектов и операций
PLANE_XOY = 1
PLANE_XOZ = 2
PLANE_YOZ = 3
OBJ_SKETCH = 5
OBJ_BOSS_EXTRUSION = 25
OBJ_CUT_EXTRUSION = 26
DIR_NORMAL = 0  # Выдавливание в прямом направлении (по нормали эскиза)
DIR_REVERSE = 1  # Выдавливание в обратном направлении
DIR_BOTH = 2  # Выдавливание в обе стороны
EXTR_TYPE_BLIND = 0  # Выдавливание на заданную глубину (Blind)
# (При необходимости можно добавить EXTR_TYPE_THROUGH_ALL и др.)


# Подключаем константы API Компас
kompas6_constants = gencache.EnsureModule(
    "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
).constants
kompas6_constants_3d = gencache.EnsureModule(
    "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0
).constants

# Подключаем API Kompas 5 и 7 через MiscellaneousHelpers
kompas6_api5_module = gencache.EnsureModule(
    "{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0
)
kompas_object = kompas6_api5_module.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(
        kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch
    )
)
MH.iKompasObject = kompas_object

kompas_api7_module = gencache.EnsureModule(
    "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0
)
application = kompas_api7_module.IApplication(
    Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
        kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch
    )
)
MH.iApplication = application

# Создаем новый документ
Documents = application.Documents
kompas_document = Documents.AddWithDefaultSettings(
    kompas6_constants.ksDocumentPart, True
)

kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
iDocument3D = kompas_object.ActiveDocument3D()
iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

MH.iKompasObject.Visible = True

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

# Пример JSON для построения куба с отверстием
json_instruction = """
{
  "name": "Cube + Through hole",
  "steps": [
    {
      "action": "sketch",
      "plane": "XOY",
      "entities": [
        {
          "type": "line",
          "start": [0.0, 0.0],
          "end": [60.0, 0.0]
        },
        {
          "type": "line",
          "start": [60.0, 0.0],
          "end": [60.0, 60.0]
        },
        {
          "type": "line",
          "start": [60.0, 60.0],
          "end": [0.0, 60.0]
        },
        {
          "type": "line",
          "start": [0.0, 60.0],
          "end": [0.0, 0.0]
        }
      ]
    },
    {
      "action": "extrude",
      "height": 60.0
    },
    {
      "action": "sketch",
      "plane": "XOY",
      "entities": [
        {
          "type": "circle",
          "center": [30.0, 30.0],
          "radius": 7.5
        }
      ]
    },
    {
      "action": "cut",
      "through_all": 1.0,
      "direction": "both"
    }
  ]
}
"""

# Запуск построения модели
builder = Kompas3DBuilder(kompas6_constants, kompas6_constants_3d, iPart)
builder.process_json(json.loads(json_instruction))

# Сохранение файла
# kompas_document.SaveAs(r"C:\\Users\\artem\\Desktop\\Деталь.m3d")
