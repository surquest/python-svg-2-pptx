from dataclasses import dataclass
from typing import Optional, Union
from .base import IRNode
from .canvas import Geometry
from .color import Color
from .gradient import GradientFill

@dataclass
class IRRectangle(IRNode):
    geometry: Geometry
    fill: Optional[Union[Color, GradientFill]]
    stroke: Optional[Color]
    stroke_width_emu: int
    corner_radius_emu: int
    dashed: bool = False
    shape_id: Optional[str] = None
