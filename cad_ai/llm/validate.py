def extract_json_object(text: str) -> str:
    first = text.find("{")
    last = text.rfind("}")
    if first == -1 or last == -1 or last <= first:
        raise ValueError("LLM did not return a JSON object.")
    return text[first : last + 1]


def validate_generated_json(data: dict):
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
                if et == "line" and ("start" not in e or "end" not in e):
                    raise ValueError(f"Line must have start/end in step #{i}.")
                if et == "circle" and ("center" not in e or "radius" not in e):
                    raise ValueError(f"Circle must have center/radius in step #{i}.")

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
            # depth can be missing; builder will default to through_all

        if act == "workplane_offset":
            bp = (step.get("base_plane") or "").upper()
            if bp not in allowed_planes:
                raise ValueError(f"Bad base_plane '{bp}' in step #{i}.")
            if "offset" not in step or "name" not in step:
                raise ValueError(
                    f"workplane_offset must have offset and name in step #{i}."
                )

        if act == "sketch_on_plane":
            ents = step.get("entities", [])
            if not isinstance(ents, list) or len(ents) == 0:
                raise ValueError(f"sketch_on_plane must have entities in step #{i}.")
