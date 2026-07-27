from dataclasses import dataclass
from pptx.dml.color import RGBColor

@dataclass(frozen=True)
class Color:
    """Normalized RGB color with optional opacity."""
    r: int
    g: int
    b: int
    opacity: float = 1.0

    @classmethod
    def from_hex(cls, value: str, opacity: float = 1.0) -> "Color":
        v = value.strip().lstrip("#")
        if len(v) == 3:
            v = "".join(c * 2 for c in v)
        if len(v) == 8:
            opacity = int(v[6:8], 16) / 255.0
            v = v[:6]
        return cls(int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16), opacity)

    def to_rgb(self) -> RGBColor:
        return RGBColor(self.r, self.g, self.b)
