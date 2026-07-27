from dataclasses import dataclass
from typing import Optional
from .base import IRNode
from .canvas import Geometry


@dataclass
class IRImage(IRNode):
    """Image element parsed from SVG <image> tag."""
    geometry: Geometry
    href: str
    image_bytes: bytes
    content_type: str
    preserve_aspect_ratio: str = "xMidYMid meet"
    shape_id: Optional[str] = None
