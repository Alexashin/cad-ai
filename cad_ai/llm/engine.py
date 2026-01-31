import json
from pathlib import Path

from .errors import LLMJSONError
from .prompt import make_llm_prompt
from .validate import extract_json_object, validate_generated_json


class LocalLLMEngine:
    """
    Thin wrapper around llama-cpp-python.
    Keeps last_raw/last_extracted/last_prompt for UI debug windows.
    """

    def __init__(
        self,
        model_path: str,
        *,
        n_ctx: int = 2048,
        n_threads: int = 8,
        n_gpu_layers: int = 0,
    ):
        self.model_path = model_path
        self.n_ctx = n_ctx
        self.n_threads = n_threads
        self.n_gpu_layers = n_gpu_layers

        self._llm = None

        self.last_raw = ""
        self.last_extracted = ""
        self.last_prompt = ""

    def _get_llm(self):
        try:
            from llama_cpp import Llama
        except Exception as e:
            raise RuntimeError(
                "llama-cpp-python is not installed. Run: pip install llama-cpp-python"
            ) from e

        if self._llm is None:
            p = Path(self.model_path)
            if not p.exists():
                raise RuntimeError(
                    f"GGUF model not found: {p}\n"
                    f"Put a GGUF model there or change LLM_MODEL_PATH."
                )
            self._llm = Llama(
                model_path=str(p),
                n_ctx=self.n_ctx,
                n_threads=self.n_threads,
                n_gpu_layers=self.n_gpu_layers,
                verbose=False,
            )
        return self._llm

    def generate_json(self, user_text: str) -> dict:
        llm = self._get_llm()
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
        extracted = extract_json_object(raw.strip())

        try:
            data = json.loads(extracted)
        except Exception as e:
            self.last_raw, self.last_extracted, self.last_prompt = (
                raw,
                extracted,
                prompt,
            )
            raise LLMJSONError(str(e), raw=raw, extracted=extracted, prompt=prompt)

        # normalize (same behavior as your file)
        for st in data.get("steps", []):
            act = (st.get("action") or "").lower().strip()
            if act == "sketch_on_plane" and not (
                st.get("plane") or st.get("plane_name") or st.get("on_plane")
            ):
                st["plane"] = "XOY"

        try:
            validate_generated_json(data)
        except Exception as e:
            self.last_raw, self.last_extracted, self.last_prompt = (
                raw,
                extracted,
                prompt,
            )
            raise LLMJSONError(str(e), raw=raw, extracted=extracted, prompt=prompt)

        self.last_raw, self.last_extracted, self.last_prompt = raw, extracted, prompt
        return data
