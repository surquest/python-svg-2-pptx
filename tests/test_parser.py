import pytest
from xml.etree import ElementTree as ET
from pptx.util import Emu

from surquest.utils.svg2pptx.parser import CoordinateNormalizer, _parse_color, _parse_float, SVGParser, _parse_font_family
from surquest.utils.svg2pptx.models import Color

class TestCoordinateNormalizer:
    def test_coordinate_normalizer(self):
        norm = CoordinateNormalizer(viewbox=(0, 0, 1000, 500), target_width=10000, target_height=5000)
        assert norm.x(100) == Emu(1000)
        assert norm.y(50) == Emu(500)
        assert norm.w(200) == Emu(2000)
        assert norm.h(100) == Emu(1000)

class TestParseColor:
    def test_parse_color(self):
        assert _parse_color("none") is None
        assert _parse_color("transparent") is None
        
        red = _parse_color("red")
        if hasattr(red, 'r'):
            assert red.r == 255
            assert red.g == 0
            assert red.b == 0
        else:
            assert isinstance(red, Color)
        
        blue = _parse_color("#0000ff")
        if hasattr(blue, 'b'):
            assert blue.b == 255
            
        rgb = _parse_color("rgb(10, 20, 30)")
        assert rgb.r == 10
        assert rgb.g == 20
        assert rgb.b == 30

    def test_parse_color_with_opacity(self):
        red = _parse_color("red", opacity=0.5)
        assert red.opacity == 0.5
        assert red.r == 255

        hex_color = _parse_color("#00ff00", opacity=0.3)
        assert hex_color.opacity == 0.3

class TestParseFontFamily:
    def test_parse_font_family(self):
        assert _parse_font_family("monospace") == "Consolas"
        assert _parse_font_family("sans-serif") == "Arial"
        assert _parse_font_family("serif") == "Times New Roman"
        assert _parse_font_family("Arial") == "Arial"
        assert _parse_font_family("'Courier New'") == "Courier New"

class TestParseFloat:
    def test_parse_float(self):
        assert _parse_float("12.5") == 12.5
        assert _parse_float("-3.5") == -3.5
        assert _parse_float(None, 4.0) == 4.0

class TestSVGParser:
    def test_svg_parser_basic(self):
        svg_source = """
        <svg viewBox="0 0 100 100">
            <rect x="10" y="10" width="80" height="80" fill="red" stroke="blue" />
            <circle cx="50" cy="50" r="40" fill="#00FF00" />
        </svg>
        """
        parser = SVGParser(svg_source, slide_width=9144000, slide_height=5143500, svg_ns="")
        slide = parser.parse()
        
        assert slide.width_emu == 9144000
        assert slide.height_emu == 5143500
        assert len(slide.nodes) == 2
        
        rect_node = slide.nodes[0]
        assert type(rect_node).__name__ == "IRRectangle"
        assert rect_node.geometry.width == Emu(7315200)
        
        circle_node = slide.nodes[1]
        assert type(circle_node).__name__ == "IREllipse"
        assert circle_node.geometry.width == Emu(7315200)

    def test_svg_parser_opacity(self):
        svg_source = """
        <svg viewBox="0 0 100 100">
            <rect x="10" y="10" width="80" height="80" fill="#7E45AF" fill-opacity="0.3" />
        </svg>
        """
        parser = SVGParser(svg_source, slide_width=9144000, slide_height=5143500, svg_ns="")
        slide = parser.parse()
        rect = slide.nodes[0]
        assert rect.fill.opacity == 0.3

    def test_svg_parser_monospace_font(self):
        svg_source = """
        <svg viewBox="0 0 100 100">
            <text x="10" y="50" font-family="monospace" font-size="12">Hello</text>
        </svg>
        """
        parser = SVGParser(svg_source, slide_width=9144000, slide_height=5143500, svg_ns="")
        slide = parser.parse()
        text = slide.nodes[0]
        assert text.block.runs[0].font.family == "Consolas"
