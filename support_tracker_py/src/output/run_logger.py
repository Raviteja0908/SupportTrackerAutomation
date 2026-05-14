from dataclasses import dataclass
from pathlib import Path


class RunLogger:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def log(self, message: str):
        line = self._clean_prefix(message.rstrip())
        with self.path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
        print(line, flush=True)

    @staticmethod
    def _clean_prefix(line: str) -> str:
        replacements = {
            "[INFO]": "[info]",
            "[WARNING]": "[warn]",
            "[ERROR]": "[error]",
        }
        for old, new in replacements.items():
            if line.startswith(old):
                return new + line[len(old):]
        return line


@dataclass(frozen=True)
class MarkingReason:
    maintenance: str = "Maintenance - left unfilled"
    unknown: str = "Unknown - missing required data"
    blue: str = "Created time fallback = ack time"
