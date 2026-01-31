# cad_ai/llm/prompt.py
import json

from cad_ai.templates.ai_templates import (
    tpl_cube,
    tpl_cube_with_through_hole,
    tpl_plate_with_holes,
    tpl_stepped_block,
)


def _compact_json(obj: dict) -> str:
    # компактно, но читаемо
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def make_llm_prompt(user_text: str) -> str:
    # few-shot examples (stable “golden” outputs)
    ex1 = _compact_json(tpl_cube(60.0, "XOY"))
    ex2 = _compact_json(tpl_cube_with_through_hole(60.0, 12.0, "XOY"))
    ex3 = _compact_json(tpl_plate_with_holes(120.0, 80.0, 8.0, 10.0, 15.0, "XOY"))
    ex4 = _compact_json(tpl_stepped_block(120, 80, 20, 60, 40, 20, 30, 20, 10, "XOY"))

    return f"""
Ты генерируешь ТОЛЬКО валидный JSON-объект для построения модели в КОМПАС-3D.
Никакого текста, объяснений, markdown, комментариев — только JSON.

Схема результата:
{{
  "name": "string",
  "steps": [ ... ]
}}

Разрешённые action:
- sketch (plane: XOY|XOZ|YOZ; entities: line/circle)
- extrude (height:number; direction: normal|reverse|both)  // direction можно не указывать, по умолчанию both
- cut (through_all:true|false; depth:number если through_all=false; direction: normal|reverse|both)
- workplane_offset (base_plane: XOY|XOZ|YOZ; offset:number; name:string)
- sketch_on_plane (plane: XOY|XOZ|YOZ|<name>; entities: line/circle)

Жёсткие правила:
- Все числа — числа (НЕ строки).
- Не добавляй поля, которых нет в описании.
- Не используй другие action.
- Если не уверен — direction="both" для extrude и cut.
- Если нужен “насквозь” — cut with through_all=true.

Примеры корректного JSON (учись формату!):
1) {ex1}
2) {ex2}
3) {ex3}
4) {ex4}

Запрос пользователя: {user_text}
""".strip()
