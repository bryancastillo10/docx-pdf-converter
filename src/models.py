from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ConversionResult:
    source: Path
    output: Path | None = None
    error: str | None = None

    @property
    def succeeded(self) -> bool:
        return self.error is None and self.output is not None
