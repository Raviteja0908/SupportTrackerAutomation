from dataclasses import dataclass
from pathlib import Path


class RunLogger:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def log(self, message: str):
        line = message.rstrip()
        with self.path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
        print(line, flush=True)


@dataclass(frozen=True)
class MarkingReason:
    maintenance: str = "Maintenance - left unfilled"
    unknown: str = "Unknown - missing required data"
    blue: str = "Created time fallback = ack time"
