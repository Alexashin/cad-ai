class LLMJSONError(Exception):
    """Raised when the LLM output cannot be parsed/validated as JSON."""

    def __init__(
        self, message: str, *, raw: str = "", extracted: str = "", prompt: str = ""
    ):
        super().__init__(message)
        self.raw = raw or ""
        self.extracted = extracted or ""
        self.prompt = prompt or ""
