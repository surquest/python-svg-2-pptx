import os
import json
import logging
import dataclasses
from enum import Enum
from pathlib import Path
from typing import Union, List

from .parser import SVGParser
from .generator import PPTXBackend

logger = logging.getLogger(__name__)

class _IRJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if dataclasses.is_dataclass(obj):
            d = dict(obj.__dict__)
            d["__type__"] = obj.__class__.__name__
            return d
        if isinstance(obj, Enum):
            return obj.value
        if isinstance(obj, bytes):
            return obj.decode("utf-8", errors="replace")
        return super().default(obj)


def _ir_object_hook(d):
    """Deserialize JSON dictionary back into IRNode dataclasses based on __type__."""
    if "__type__" not in d:
        return d
        
    from . import models
    from pptx.enum.text import PP_ALIGN
    
    d_copy = d.copy()
    cls_name = d_copy.pop("__type__")
    
    if hasattr(models, cls_name):
        node_cls = getattr(models, cls_name)
    else:
        return d
        
    try:
        fields = dataclasses.fields(node_cls)
        import typing
        hints = typing.get_type_hints(node_cls)
    except TypeError:
        return d
        
    kwargs = {}
    for f in fields:
        if f.name in d_copy:
            val = d_copy[f.name]
            
            # Reconstruct enums
            if f.name == "alignment" and val is not None:
                try:
                    val = PP_ALIGN(val)
                except ValueError:
                    pass
            elif f.name == "connector_type" and val is not None:
                val = models.ConnectorType(val)
            elif f.name in ("start_arrow", "end_arrow") and val is not None:
                val = models.ArrowType(val)
                
            # Cast list to tuple if field type is Tuple
            if isinstance(val, list):
                f_type = hints.get(f.name)
                if typing.get_origin(f_type) is tuple:
                    val = tuple(val)
            
            # Cast string to bytes if field type is bytes
            if isinstance(val, str) and hints.get(f.name) is bytes:
                val = val.encode("utf-8")
                
            kwargs[f.name] = val
            
    return node_cls(**kwargs)

class SVG2Pptx:

    """Main class for converting SVG to PowerPoint."""
    
    # Centralized configuration and constants
    SVG_NS: str = "{http://www.w3.org/2000/svg}"
    
    # Standard 16:9 slide dimensions in EMU (13.333" x 7.5")
    SLIDE_WIDTH_EMU: int = 12192000
    SLIDE_HEIGHT_EMU: int = 6858000

    def __init__(self, slide_width_emu: int = None, slide_height_emu: int = None, svg_ns: str = None):
        """
        Initialize the SVG to PPTX converter with optional configuration.
        """
        self.slide_width_emu = slide_width_emu or self.SLIDE_WIDTH_EMU
        self.slide_height_emu = slide_height_emu or self.SLIDE_HEIGHT_EMU
        self.svg_ns = svg_ns or self.SVG_NS
    
    def _get_svg_source(self, svg_input: Union[str, Path]) -> tuple:
        """Return (svg_source, svg_path) tuple."""
        if isinstance(svg_input, (str, Path)):
            try:
                if os.path.isfile(svg_input):
                    with open(svg_input, "r", encoding="utf-8") as f:
                        return f.read(), str(svg_input)
            except Exception:
                pass
            return str(svg_input), None
        raise TypeError("svg_input must be a string containing SVG content or a file path.")

    def to_json(self, svg_input: Union[str, Path, List[Union[str, Path]]], indent: int = 2) -> str:
        """
        Convert an SVG string or file (or list of them) to its JSON IR representation string.
        """
        if not isinstance(svg_input, list):
            svg_input = [svg_input]
            
        irs = []
        for inp in svg_input:
            svg_source, svg_path = self._get_svg_source(inp)
            ir = SVGParser(
                svg_source,
                slide_width=self.slide_width_emu,
                slide_height=self.slide_height_emu,
                svg_ns=self.svg_ns,
                svg_path=svg_path,
            ).parse()
            irs.append(ir)
            
        # Return a single object if only one was passed, else return list
        return json.dumps(irs[0] if len(irs) == 1 else irs, cls=_IRJSONEncoder, indent=indent)

    def convert(self, svg_input: Union[str, Path, List[Union[str, Path]]], output_path: Union[str, Path], export_as_json: bool = False) -> None:
        """
        Convert an SVG string or file (or list of them) to a PPTX presentation and JSON output.
        
        Args:
            svg_input: Either an SVG string format or a path to an SVG file, or a list of such.
            output_path: Path where the output .pptx should be saved.
                         A corresponding .json file will also be created.
            export_as_json: If True, only export the JSON IR without creating a PPTX file.
        """
        if not isinstance(svg_input, list):
            svg_input = [svg_input]

        irs = []
        for inp in svg_input:
            svg_source, svg_path = self._get_svg_source(inp)
            ir = SVGParser(
                svg_source,
                slide_width=self.slide_width_emu,
                slide_height=self.slide_height_emu,
                svg_ns=self.svg_ns,
                svg_path=svg_path,
            ).parse()
            irs.append(ir)
        
        if export_as_json:
            self.export_ir_to_json(output_path, irs)
            
        prs = PPTXBackend(irs).render()
        prs.save(str(output_path))
        logger.info(f"Wrote PPTX to %s", output_path)

        return str(output_path)

    def export_ir_to_json(self, output_path, ir):
        output_path_str = str(output_path)
        json_path = output_path_str.replace(".pptx", ".json")
        if json_path == output_path_str:
            json_path += ".json"
            
        with open(json_path, "w", encoding="utf-8") as f:
            # Drop the list wrapping if it's a single item for backwards compat.
            output_data = ir[0] if isinstance(ir, list) and len(ir) == 1 else ir
            json.dump(output_data, f, cls=_IRJSONEncoder, indent=2)
        logger.info(f"Wrote JSON IR to %s", json_path)

    def from_json(self, json_input: Union[str, Path], output_path: Union[str, Path]) -> str:
        """
        Import an IRSlide (or list of IRSlides) from a JSON string or file path and convert it to PPTX.
        
        Args:
            json_input: Either a JSON string containing IR representation or a path to a JSON file.
            output_path: Path where the output .pptx should be saved.
        """
        if isinstance(json_input, (str, Path)) and os.path.isfile(str(json_input)):
            with open(json_input, "r", encoding="utf-8") as f:
                ir = json.load(f, object_hook=_ir_object_hook)
        else:
            ir = json.loads(str(json_input), object_hook=_ir_object_hook)
            
        prs = PPTXBackend(ir).render()
        prs.save(str(output_path))
        logger.info(f"Wrote PPTX to %s", output_path)
        return str(output_path)
