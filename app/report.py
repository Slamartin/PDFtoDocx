"""Result shape shared by both converters so the UI can render them identically."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field


@dataclass
class ConversionReport:
    pages: int = 0
    images: int = 0
    tables: int = 0
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def warn(self, text: str) -> None:
        if text not in self.warnings:
            self.warnings.append(text)

    def note(self, text: str) -> None:
        if text not in self.notes:
            self.notes.append(text)

    def to_dict(self) -> dict:
        return asdict(self)


class PasswordError(Exception):
    """The PDF is encrypted and the given password (if any) did not open it."""


class CorruptFileError(Exception):
    """The input could not be parsed as the claimed format."""
