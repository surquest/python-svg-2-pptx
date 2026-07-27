from dataclasses import dataclass
from typing import List, Optional, Union
from .base import IRNode
from .canvas import Point
from .color import Color
from .gradient import GradientFill

@dataclass
class IRPolygon(IRNode):
    """Handles both polygon and polyline."""
    waypoints: List[Point]
    is_closed: bool
    fill: Optional[Union[Color, GradientFill]]
    stroke: Optional[Color]
    stroke_width_emu: int
    dashed: bool = False
    shape_id: Optional[str] = None
