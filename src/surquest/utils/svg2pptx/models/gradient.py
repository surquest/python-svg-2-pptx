from dataclasses import dataclass, field
from typing import List
from .color import Color


@dataclass(frozen=True)
class GradientStop:
    """A single color stop in a gradient."""
    color: Color
    position: float  # 0.0 to 1.0


@dataclass(frozen=True)
class GradientFill:
    """Linear gradient fill."""
    stops: tuple  # tuple of GradientStop
    angle: float = 0.0  # Degrees, 0 = left-to-right
    gradient_units: str = "objectBoundingBox"
