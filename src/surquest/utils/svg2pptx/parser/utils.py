import re
from typing import Tuple, List, Optional, Union
from xml.etree import ElementTree as ET
from pptx.util import Emu
from pptx.enum.text import PP_ALIGN

from ..models import (
    IRSlide, IRNode, IRRectangle, IRIcon, IRText, IRLine, IRConnector, IRInfoBox, IRGroup,
    IREllipse, IRPolygon,
    Geometry, Color, GradientStop, GradientFill, FontStyle, TextRun, TextBlock, Point,
    ArrowType, ConnectorType
)

class CoordinateNormalizer:
    """Converts SVG viewbox units to EMU."""

    def __init__(self, viewbox: Tuple[float, float, float, float],
                 target_width: int, target_height: int):
        _, _, vb_w, vb_h = viewbox
        self.scale_x = target_width / vb_w
        self.scale_y = target_height / vb_h
        self.offset_x = -viewbox[0] * self.scale_x
        self.offset_y = -viewbox[1] * self.scale_y

    def x(self, v: float) -> Emu:
        return Emu(int(v * self.scale_x + self.offset_x))

    def y(self, v: float) -> Emu:
        return Emu(int(v * self.scale_y + self.offset_y))

    def w(self, v: float) -> Emu:
        return Emu(int(v * self.scale_x))

    def h(self, v: float) -> Emu:
        return Emu(int(v * self.scale_y))

    def pt(self, v: float) -> float:
        """Convert SVG font-size (px) to PowerPoint points (approx scale)."""
        avg = (self.scale_x + self.scale_y) / 2
        emu_per_pt = 12700
        return (v * avg) / emu_per_pt

def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag

CSS_COLORS = {
    "black": "#000000", "silver": "#c0c0c0", "gray": "#808080", "white": "#ffffff",
    "maroon": "#800000", "red": "#ff0000", "purple": "#800080", "fuchsia": "#ff00ff",
    "green": "#008000", "lime": "#00ff00", "olive": "#808000", "yellow": "#ffff00",
    "navy": "#000080", "blue": "#0000ff", "teal": "#008080", "aqua": "#00ffff",
    "orange": "#ffa500", "brown": "#a52a2a", "transparent": "none", "none": "none"
}

FONT_FAMILY_MAP = {
    "monospace": "Consolas",
    "sans-serif": "Arial",
    "serif": "Times New Roman",
    "cursive": "Brush Script MT",
    "fantasy": "Papyrus",
}

def _parse_color(value: Optional[str], opacity: float = 1.0) -> Optional[Color]:
    if not value:
        return None
    v = value.strip().lower()
    if v in ("none", "transparent"):
        return None
    if v in CSS_COLORS:
        v = CSS_COLORS[v]
        if v == "none":
            return None
    if v.startswith("#"):
        return Color.from_hex(v, opacity)
    if v.startswith("rgb("):
        m = re.match(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", v)
        if m:
            return Color(int(m.group(1)), int(m.group(2)), int(m.group(3)), opacity)
    return None

def _parse_font_family(value: Optional[str]) -> str:
    if not value:
        return "Arial"
    family = value.strip().lower().strip("'\"")
    return FONT_FAMILY_MAP.get(family, value.strip().strip("'\""))

def _parse_float(value: Optional[str], default: float = 0.0) -> float:
    if value is None:
        return default
    m = re.match(r"-?\d*\.?\d+", value.strip())
    return float(m.group()) if m else default

def _parse_dash(value: Optional[str]) -> bool:
    return value is not None and value.strip() not in ("", "none")


def _parse_gradient_stops(grad_el: ET.Element) -> tuple:
    """Parse <stop> elements inside a gradient, return tuple of GradientStop."""
    ns = "{http://www.w3.org/2000/svg}"
    stops = []
    for stop_el in grad_el.findall(f"{ns}stop") or grad_el.findall("stop"):
        offset_str = stop_el.get("offset", "0%")
        if offset_str.endswith("%"):
            pos = float(offset_str.rstrip("%")) / 100.0
        else:
            pos = float(offset_str)
        stop_color_str = stop_el.get("stop-color", "#000000")
        stop_opacity = _parse_float(stop_el.get("stop-opacity"), 1.0)
        color = _parse_color(stop_color_str, stop_opacity)
        if color:
            stops.append(GradientStop(color=color, position=pos))
    return tuple(stops)


def _parse_gradient(grad_el: ET.Element) -> Optional[GradientFill]:
    """Parse a <linearGradient> or <radialGradient> element."""
    stops = _parse_gradient_stops(grad_el)
    if not stops:
        return None
    gradient_units = grad_el.get("gradientUnits", "objectBoundingBox")
    x1 = _parse_float(grad_el.get("x1"), 0.0)
    y1 = _parse_float(grad_el.get("y1"), 0.0)
    x2 = _parse_float(grad_el.get("x2"), 1.0)
    y2 = _parse_float(grad_el.get("y2"), 0.0)
    import math
    angle = math.degrees(math.atan2(y2 - y1, x2 - x1))
    return GradientFill(
        stops=stops,
        angle=angle,
        gradient_units=gradient_units,
    )


def _parse_fill(value: Optional[str], opacity: float = 1.0,
                 gradients: Optional[dict] = None) -> Optional[Union[Color, GradientFill]]:
    """Parse fill attribute, supporting both solid colors and url(#id) gradient refs."""
    if not value:
        return None
    v = value.strip()
    if v.startswith("url("):
        m = re.match(r"url\(#([^)]+)\)", v)
        if m and gradients:
            return gradients.get(m.group(1))
        return None
    return _parse_color(v, opacity)


def _resolve_image_path(href: str, svg_path: Optional[str] = None) -> Optional[str]:
    """Resolve image href to absolute file path.

    Handles:
    - data: URIs (returned as-is, caller handles decoding)
    - Absolute file paths (returned as-is)
    - Relative paths (resolved against svg_path's directory)
    """
    from pathlib import Path

    if not href:
        return None

    href = href.strip()

    # data: URI - return as-is
    if href.startswith("data:"):
        return href

    # Absolute path
    if Path(href).is_absolute():
        return href if Path(href).exists() else None

    # Relative path - resolve against SVG file location
    if svg_path:
        svg_dir = Path(svg_path).parent
        resolved = (svg_dir / href).resolve()
        if resolved.exists():
            return str(resolved)

    return None


def _load_image_bytes(href: str, svg_path: Optional[str] = None) -> Optional[Tuple[bytes, str]]:
    """Load image bytes and detect content type.

    Returns (image_bytes, content_type) or None if not found.
    """
    import mimetypes
    from pathlib import Path

    # data: URI
    if href.startswith("data:"):
        import base64
        # format: data:image/png;base64,XXXX
        m = re.match(r"data:([^;]+);base64,(.+)", href, re.DOTALL)
        if m:
            content_type = m.group(1)
            image_bytes = base64.b64decode(m.group(2))
            return image_bytes, content_type
        return None

    # File path
    resolved = _resolve_image_path(href, svg_path)
    if resolved:
        content_type, _ = mimetypes.guess_type(resolved)
        if content_type is None:
            content_type = "image/png"
        with open(resolved, "rb") as f:
            return f.read(), content_type

    return None

