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
    def __init__(self):
        self.iPart = iPart
        self.last_sketch = None

    def start_sketch(self, plane):
        """Создание нового эскиза на заданной плоскости."""
        sketch = self.iPart.NewEntity(kompas6_constants_3d.o3d_sketch)
        definition = sketch.GetDefinition()
        plane_obj = self.iPart.GetDefaultEntity(
            getattr(kompas6_constants_3d, f"o3d_plane{plane}")
        )
        definition.SetPlane(plane_obj)
        sketch.Create()
        self.last_sketch = sketch
        return definition.BeginEdit()

    def add_line(self, doc2d, x1, y1, x2, y2):
        """Добавление линии в эскиз."""
        doc2d.ksLineSeg(x1, y1, x2, y2, 1)

    def add_circle(self, doc2d, x, y, radius):
        """Добавление окружности в эскиз."""
        doc2d.ksCircle(x, y, radius, 1)

    def finish_sketch(self):
        """Завершение эскиза."""
        if not self.last_sketch:
            raise RuntimeError("Эскиз не найден, невозможно завершить редактирование.")
        self.last_sketch.GetDefinition().EndEdit()

    def extrude(self, height):
        """Выдавливание последнего эскиза."""
        extrusion = self.iPart.NewEntity(kompas6_constants_3d.o3d_bossExtrusion)
        definition = extrusion.GetDefinition()
        definition.SetSketch(self.last_sketch)
        extrude_params = definition.ExtrusionParam()
        extrude_params.direction = kompas6_constants_3d.dtNormal
        extrude_params.depthNormal = height
        extrude_params.typeNormal = kompas6_constants_3d.etBlind
        extrusion.Create()

    def cut(self, depth):
        """Вырезание материала по последнему эскизу."""
        cut_feature = self.iPart.NewEntity(kompas6_constants_3d.o3d_cutExtrusion)
        definition = cut_feature.GetDefinition()
        definition.SetSketch(self.last_sketch)
        extrude_params = definition.ExtrusionParam()
        extrude_params.direction = kompas6_constants_3d.dtReverse
        extrude_params.depthNormal = depth
        extrude_params.typeNormal = kompas6_constants_3d.etThroughAll
        cut_feature.Create()

    def apply_color(self, color_value):
        """Применение цвета к объекту."""
        color_param = self.last_sketch.ColorParam()
        r, g, b = color_value
        color_int = (b << 16) | (g << 8) | r
        color_param.color = color_int
        color_param.transparency = 1

    def process_json(self, json_data):
        """Обработка JSON-инструкций для создания модели."""
        for step in json_data.get("steps", []):
            match step.get("action"):
                case "sketch":
                    plane = step.get("plane")
                    doc2d = self.start_sketch(plane)
                    for entity in step.get("entities", []):
                        match entity.get("type"):
                            case "line":
                                self.add_line(doc2d, *entity["start"], *entity["end"])
                            case "circle":
                                self.add_circle(
                                    doc2d, *entity["center"], entity["radius"]
                                )
                    self.finish_sketch()
                case "extrude":
                    self.extrude(step.get("height"))
                case "cut":
                    self.cut(step.get("depth"))
                case "color":
                    self.apply_color(step.get("value"))
                case _:
                    raise ValueError(f"Неизвестная команда: {step.get('action')}")


# Пример JSON для построения куба с отверстием
json_instruction = """
{
  "name": "Болт M10x100",
  "steps": [
    {
      "action": "sketch",
      "plane": "XOY",
      "entities": [
        {"type": "circle", "center": [0, 0], "radius": 5}
      ]
    },
    {
      "action": "extrude",
      "height": 90,
      "direction": "normal"
    },
    {
      "action": "sketch",
      "plane": "XOY",
      "entities": [
        {"type": "circle", "center": [0, 0], "radius": 10}
      ]
    },
    {
      "action": "extrude",
      "height": 10,
      "direction": "normal"
    }
  ]
}
"""

# Запуск построения модели
builder = Kompas3DBuilder()
builder.process_json(json.loads(json_instruction))

# Сохранение файла
# kompas_document.SaveAs(r"C:\\Users\\artem\\Desktop\\Деталь.m3d")
