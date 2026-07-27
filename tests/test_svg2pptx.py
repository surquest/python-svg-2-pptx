import os
import pytest
from pathlib import Path
from surquest.utils.svg2pptx import SVG2Pptx
from surquest.utils.svg2pptx.parser import SVGParser
from surquest.utils.svg2pptx.svg2pptx import _ir_object_hook
import json

DATA_DIR = Path(__file__).parent / "data"
SVG_FILES = list(DATA_DIR.glob("*.svg"))

class TestSVG2Pptx:
    def test_svg2pptx_string_input(self, tmp_path):
        svg_source = """
        <svg viewBox="0 0 100 100">
            <rect x="10" y="10" width="80" height="80" fill="red" />
        </svg>
        """
        converter = SVG2Pptx(slide_width_emu=9144000, slide_height_emu=9144000)
        
        out_pptx = tmp_path / "output.pptx"
        converter.convert(svg_input=svg_source, output_path=out_pptx, export_as_json=True)
        
        assert out_pptx.exists()
        out_json = tmp_path / "output.json"
        assert out_json.exists()

    def test_svg2pptx_file_input(self, tmp_path):
        svg_file = tmp_path / "input.svg"
        svg_file.write_text("""
        <svg viewBox="0 0 100 100">
            <circle cx="50" cy="50" r="40" fill="blue" />
        </svg>
        """)
        
        converter = SVG2Pptx()
        out_pptx = tmp_path / "output_file.pptx"
        
        converter.convert(svg_input=svg_file, output_path=out_pptx, export_as_json=True)
        
        assert out_pptx.exists()
        out_json = tmp_path / "output_file.json"
        assert out_json.exists()

    def test_svg2pptx_multiple_files(self, tmp_path):
        svg_file1 = tmp_path / "input1.svg"
        svg_file1.write_text('<svg viewBox="0 0 100 100"><rect x="10" y="10" width="80" height="80" fill="red" /></svg>')
        svg_file2 = tmp_path / "input2.svg"
        svg_file2.write_text('<svg viewBox="0 0 100 100"><circle cx="50" cy="50" r="40" fill="blue" /></svg>')
        
        converter = SVG2Pptx()
        out_pptx = tmp_path / "output_multi.pptx"
        
        # Test convert list of files
        converter.convert(svg_input=[svg_file1, svg_file2], output_path=out_pptx, export_as_json=True)
        
        assert out_pptx.exists()
        out_json = tmp_path / "output_multi.json"
        assert out_json.exists()
        
        # Verify JSON has array of IR slides
        with open(out_json, "r") as f:
            data = json.load(f)
        assert isinstance(data, list)
        assert len(data) == 2
        
        # Test to_json with multiple files
        ir_json = converter.to_json([svg_file1, svg_file2])
        parsed_ir = json.loads(ir_json)
        assert isinstance(parsed_ir, list)
        assert len(parsed_ir) == 2

    @pytest.mark.parametrize("svg_file", SVG_FILES, ids=lambda path: path.name)
    def test_svg_to_ir_matches_json_ir(self, svg_file):
        """Assure the SVG parser produces identical IR as stored in the reference JSON files."""
        converter = SVG2Pptx()
        svg_source, svg_path = converter._get_svg_source(svg_file)
        
        ir_from_svg = SVGParser(
            svg_source,
            slide_width=converter.slide_width_emu,
            slide_height=converter.slide_height_emu,
            svg_ns=converter.svg_ns,
            svg_path=svg_path,
        ).parse()
        
        json_file = svg_file.with_suffix(".json")
        if not json_file.exists():
            pytest.skip(f"No corresponding JSON file found for {svg_file.name}")
            
        with open(json_file, "r", encoding="utf-8") as f:
            ir_from_json = json.load(f, object_hook=_ir_object_hook)
            
        assert ir_from_svg == ir_from_json
