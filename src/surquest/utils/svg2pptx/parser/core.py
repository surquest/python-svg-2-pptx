import re
from typing import Tuple, List, Optional
from xml.etree import ElementTree as ET
from pptx.util import Emu
from pptx.enum.text import PP_ALIGN

from ..models import (
    IRSlide, IRNode, IRRectangle, IRIcon, IRText, IRLine, IRConnector, IRInfoBox, IRGroup,
    IREllipse, IRPolygon, IRImage,
    Geometry, Color, GradientFill, FontStyle, TextRun, TextBlock, Point, ArrowType, ConnectorType
)
from .utils import CoordinateNormalizer, _strip_ns, _parse_color, _parse_float, _parse_dash, _parse_font_family, _parse_gradient, _parse_fill, _load_image_bytes

class SVGParser:
    """Frontend: SVG -> IR."""

    def __init__(self, svg_source: str, slide_width: int, slide_height: int, svg_ns: str,
                 svg_path: str = None):
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.svg_ns = svg_ns
        self.svg_path = svg_path
        self.root = ET.fromstring(svg_source)
        vb = self.root.get("viewBox", "0 0 960 540").split()
        viewbox = tuple(float(v) for v in vb)  # type: ignore
        self.norm = CoordinateNormalizer(viewbox, self.slide_width, self.slide_height)
        self.gradients = self._parse_defs()

    # ---------- Public API ----------

    def parse(self) -> IRSlide:
        slide = IRSlide(width_emu=self.slide_width, height_emu=self.slide_height)
        self._traverse(self.root, slide.nodes)
        return slide

    def _parse_defs(self) -> dict:
        """Extract gradient definitions from <defs> elements."""
        gradients = {}
        for defs_el in self.root.iter():
            if _strip_ns(defs_el.tag) == "defs":
                for child in defs_el:
                    tag = _strip_ns(child.tag)
                    if tag in ("linearGradient", "radialGradient"):
                        grad_id = child.get("id")
                        if grad_id:
                            grad = _parse_gradient(child)
                            if grad:
                                gradients[grad_id] = grad
        return gradients

    def _parse_transform(self, transform_str: str) -> Tuple[float, float]:
        dx, dy = 0.0, 0.0
        if not transform_str:
            return dx, dy
        import re
        for m in re.finditer(r'translate\(\s*(-?\d*\.?\d+)(?:[\s,]+(-?\d*\.?\d+))?\s*\)', transform_str):
            dx += float(m.group(1))
            dy += float(m.group(2) if m.group(2) is not None else 0.0)
        return dx, dy

    def _traverse(self, parent: ET.Element, nodes: List[IRNode], dx: float = 0.0, dy: float = 0.0) -> None:
        for child in parent:
            tag = _strip_ns(child.tag)
            etype = child.get("data-element-type", "")

            transform = child.get("transform", "")
            cdx, cdy = self._parse_transform(transform)
            ndx = dx + cdx
            ndy = dy + cdy

            if tag == "defs":
                continue
            elif tag == "g" and etype == "infoBox":
                nodes.append(self._parse_infobox(child, ndx, ndy))
            elif tag in ("line", "polyline") and etype == "connector":
                nodes.append(self._parse_connector(child, ndx, ndy))
            elif tag == "text":
                nodes.extend(self._parse_text(child, dx=ndx, dy=ndy))
            elif tag == "rect":
                nodes.append(self._parse_rect(child, ndx, ndy))
            elif tag in ("circle", "ellipse"):
                nodes.append(self._parse_ellipse(child, ndx, ndy))
            elif tag in ("polygon", "polyline") and etype != "connector":
                nodes.append(self._parse_poly(child, ndx, ndy))
            elif tag == "path":
                nodes.append(self._parse_path_as_svg(child, ndx, ndy))
            elif tag == "image":
                img = self._parse_image(child, ndx, ndy)
                if img is not None:
                    nodes.append(img)
            elif tag == "g":
                group_nodes = []
                self._traverse(child, group_nodes, ndx, ndy)
                if group_nodes:
                    nodes.append(IRGroup(children=group_nodes, shape_id=child.get("id")))
            elif tag == "svg" and (etype == "icon" or parent.get("data-element-type") == "icon"):
                nodes.append(self._parse_icon(child, ndx, ndy))
            elif tag == "line" and etype != "connector":
                nodes.append(self._parse_decorative_line(child, ndx, ndy))

    # ---------- Element parsers ----------

    def _parse_infobox(self, g: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IRInfoBox:
        rect_el = None
        children: List[IRNode] = []
        shape_id = g.get("id")

        def _traverse_infobox(parent: ET.Element, pdx: float, pdy: float):
            nonlocal rect_el
            for child in parent:
                tag = _strip_ns(child.tag)
                etype = child.get("data-element-type", "")

                transform = child.get("transform", "")
                cdx, cdy = self._parse_transform(transform)
                ndx = pdx + cdx
                ndy = pdy + cdy

                if tag == "rect":
                    if rect_el is None:
                        rect_el = self._parse_rect(child, ndx, ndy)
                    else:
                        children.append(self._parse_rect(child, ndx, ndy))
                elif tag == "svg":
                    children.append(self._parse_icon(child, ndx, ndy))
                elif tag == "text":
                    children.extend(self._parse_text(child, container=rect_el, dx=ndx, dy=ndy))
                elif tag == "line":
                    children.append(self._parse_decorative_line(child, ndx, ndy))
                elif tag in ("circle", "ellipse"):
                    children.append(self._parse_ellipse(child, ndx, ndy))
                elif tag in ("polygon", "polyline"):
                    children.append(self._parse_poly(child, ndx, ndy))
                elif tag == "path":
                    children.append(self._parse_path_as_svg(child, ndx, ndy))
                elif tag == "g":
                    group_nodes = []
                    self._traverse(child, group_nodes, ndx, ndy)
                    if group_nodes:
                        children.append(IRGroup(children=group_nodes, shape_id=child.get("id")))

        _traverse_infobox(g, dx, dy)

        if rect_el is None:
            # Fallback: empty container
            rect_el = IRRectangle(
                geometry=Geometry(Emu(0), Emu(0), Emu(0), Emu(0)),
                fill=None, stroke=None,
                stroke_width_emu=0, corner_radius_emu=0,
                shape_id=shape_id
            )
        else:
            if shape_id and not rect_el.shape_id:
                rect_el.shape_id = shape_id
        return IRInfoBox(rectangle=rect_el, children=children)

    def _parse_rect(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IRRectangle:
        x = _parse_float(el.get("x")) + dx
        y = _parse_float(el.get("y")) + dy
        w = _parse_float(el.get("width"))
        h = _parse_float(el.get("height"))
        rx = _parse_float(el.get("rx"))
        fill_opacity = _parse_float(el.get("fill-opacity"), 1.0)
        stroke_opacity = _parse_float(el.get("stroke-opacity"), 1.0)
        return IRRectangle(
            geometry=Geometry(
                self.norm.x(x), self.norm.y(y),
                self.norm.w(w), self.norm.h(h)
            ),
            fill=_parse_fill(el.get("fill"), fill_opacity, self.gradients),
            stroke=_parse_color(el.get("stroke"), stroke_opacity),
            stroke_width_emu=int(self.norm.w(_parse_float(el.get("stroke-width"), 1))),
            corner_radius_emu=int(self.norm.w(rx)),
            dashed=_parse_dash(el.get("stroke-dasharray")),
            shape_id=el.get("id"),
        )
        
    def _parse_ellipse(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IREllipse:
        tag = _strip_ns(el.tag)
        if tag == "circle":
            cx = _parse_float(el.get("cx")) + dx
            cy = _parse_float(el.get("cy")) + dy
            r = _parse_float(el.get("r"))
            rx = ry = r
        else:
            cx = _parse_float(el.get("cx")) + dx
            cy = _parse_float(el.get("cy")) + dy
            rx = _parse_float(el.get("rx"))
            ry = _parse_float(el.get("ry"))
            
        fill_opacity = _parse_float(el.get("fill-opacity"), 1.0)
        stroke_opacity = _parse_float(el.get("stroke-opacity"), 1.0)
        return IREllipse(
            geometry=Geometry(
                self.norm.x(cx - rx), self.norm.y(cy - ry),
                self.norm.w(rx * 2), self.norm.h(ry * 2)
            ),
            fill=_parse_fill(el.get("fill"), fill_opacity, self.gradients),
            stroke=_parse_color(el.get("stroke"), stroke_opacity),
            stroke_width_emu=int(self.norm.w(_parse_float(el.get("stroke-width"), 1))),
            dashed=_parse_dash(el.get("stroke-dasharray")),
            shape_id=el.get("id"),
        )
        
    def _parse_poly(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IRPolygon:
        is_closed = _strip_ns(el.tag) == "polygon"
        pts_raw = el.get("points", "").replace(",", " ").split()
        coords = [float(v) for v in pts_raw if v.strip()]
        waypoints = []
        for i in range(0, len(coords)-1, 2):
            waypoints.append(Point(
                self.norm.x(coords[i] + dx),
                self.norm.y(coords[i + 1] + dy)
            ))
             
        return IRPolygon(
             waypoints=waypoints,
             is_closed=is_closed,
             fill=_parse_fill(el.get("fill"), gradients=self.gradients) if is_closed else None,
             stroke=_parse_color(el.get("stroke")),
             stroke_width_emu=int(self.norm.w(_parse_float(el.get("stroke-width"), 1))),
             dashed=_parse_dash(el.get("stroke-dasharray")),
             shape_id=el.get("id")
        )
        
    def _parse_path_as_svg(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IRIcon:
        svg_copy = ET.Element("svg", {
            "xmlns": "http://www.w3.org/2000/svg",
            "viewBox": self.root.get("viewBox", "0 0 960 540"),
            "width": "100%",
            "height": "100%"
        })
        
        path_copy = ET.fromstring(ET.tostring(el))
        
        if dx != 0.0 or dy != 0.0:
            g = ET.SubElement(svg_copy, "g", {"transform": f"translate({dx}, {dy})"})
            g.append(path_copy)
        else:
            svg_copy.append(path_copy)
            
        svg_str = ET.tostring(svg_copy, encoding="utf-8")
        # Ensure it has basic width height in geometry equivalent to slide (rendered whole-slide overlay)
        return IRIcon(
            geometry=Geometry(Emu(0), Emu(0), Emu(self.slide_width), Emu(self.slide_height)),
            svg_bytes=svg_str,
        )

    def _parse_icon(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IRIcon:
        x = _parse_float(el.get("x")) + dx
        y = _parse_float(el.get("y")) + dy
        w = _parse_float(el.get("width"))
        h = _parse_float(el.get("height"))

        svg_copy = ET.fromstring(ET.tostring(el))
        if not svg_copy.tag.startswith("{"):
            svg_copy.set("xmlns", "http://www.w3.org/2000/svg")
        for attr in ("x", "y"):
            if attr in svg_copy.attrib:
                del svg_copy.attrib[attr]
        svg_copy.set("width", str(w))
        svg_copy.set("height", str(h))
        svg_str = ET.tostring(svg_copy, encoding="utf-8")
        if b"xmlns" not in svg_str.split(b">", 1)[0]:
            svg_str = svg_str.replace(
                b"<svg", b'<svg xmlns="http://www.w3.org/2000/svg"', 1
            )

        return IRIcon(
            geometry=Geometry(
                self.norm.x(x), self.norm.y(y),
                self.norm.w(w), self.norm.h(h)
            ),
            svg_bytes=svg_str,
        )

    def _parse_image(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> Optional[IRImage]:
        x = _parse_float(el.get("x")) + dx
        y = _parse_float(el.get("y")) + dy
        w = _parse_float(el.get("width"))
        h = _parse_float(el.get("height"))
        href = el.get("href") or el.get("{http://www.w3.org/1999/xlink}href") or ""
        preserve_aspect = el.get("preserveAspectRatio", "xMidYMid meet")

        result = _load_image_bytes(href, self.svg_path)
        if result is None:
            return None

        image_bytes, content_type = result
        return IRImage(
            geometry=Geometry(
                self.norm.x(x), self.norm.y(y),
                self.norm.w(w), self.norm.h(h)
            ),
            href=href,
            image_bytes=image_bytes,
            content_type=content_type,
            preserve_aspect_ratio=preserve_aspect,
            shape_id=el.get("id"),
        )

    def _parse_text(self, el: ET.Element,
                    container: Optional[IRRectangle] = None, dx: float = 0.0, dy: float = 0.0) -> List[IRText]:
        base_font = self._extract_font(el)
        anchor = el.get("text-anchor", "start")
        align_map = {
            "start": PP_ALIGN.LEFT,
            "middle": PP_ALIGN.CENTER,
            "end": PP_ALIGN.RIGHT,
        }
        alignment = align_map.get(anchor, PP_ALIGN.LEFT)

        base_x = _parse_float(el.get("x")) + dx
        base_y = _parse_float(el.get("y")) + dy

        tspans = list(el.findall(f"{self.svg_ns}tspan")) or list(el.findall("tspan"))

        if not tspans:
            txt = (el.text or "").strip()
            if txt:
                block = TextBlock(
                    runs=(TextRun(text=txt, font=base_font),),
                    anchor_x=self.norm.x(base_x),
                    anchor_y=self.norm.y(base_y),
                    alignment=alignment,
                )
                return [IRText(block=block, container_geometry=container.geometry if container else None)]
            return []

        blocks_info = []
        current_runs = []
        current_x = base_x
        current_y = base_y
        current_block_x = base_x
        current_block_y = base_y

        active_font = base_font

        for i, ts in enumerate(tspans):
            ts_x_str = ts.get("x")
            ts_y_str = ts.get("y")
            ts_dy_str = ts.get("dy")
            ts_dx_str = ts.get("dx")

            if ts_x_str is not None:
                ts_x = _parse_float(ts_x_str) + dx
            else:
                ts_x = current_x + _parse_float(ts_dx_str)
                
            if ts_y_str is not None:
                ts_y = _parse_float(ts_y_str) + dy
            else:
                ts_y = current_y + _parse_float(ts_dy_str)

            is_newline = False
            is_new_block = False

            if i == 0:
                current_block_x = ts_x
                current_block_y = ts_y
            else:
                if ts_y_str is not None and abs(ts_y - current_y) > 0.1:
                    is_new_block = True
                elif ts_dy_str is not None and abs(_parse_float(ts_dy_str)) > 0.1:
                    is_newline = True
                elif ts_x_str is not None and abs(ts_x - current_x) > 0.1:
                    if abs(ts_x - current_block_x) < 0.1:
                        # Back to same x as block start -> probably a newline
                        is_newline = True
                    else:
                        is_new_block = True

            if is_new_block:
                if current_runs:
                    blocks_info.append((current_block_x, current_block_y, current_runs))
                current_runs = []
                current_block_x = ts_x
                current_block_y = ts_y

            ts_font = self._extract_font(ts, fallback=base_font)
            txt = (ts.text or "").strip()
            
            # If same font and no newline and no new block, we could merge text, but we can just use TextRun
            # with line_break_before = is_newline
            current_runs.append(TextRun(
                text=txt,
                font=ts_font,
                line_break_before=is_newline,
            ))

            current_x = ts_x
            current_y = ts_y

        if current_runs:
            blocks_info.append((current_block_x, current_block_y, current_runs))

        result = []
        for bx, by, bruns in blocks_info:
            block = TextBlock(
                runs=tuple(bruns),
                anchor_x=self.norm.x(bx),
                anchor_y=self.norm.y(by),
                alignment=alignment,
            )
            result.append(IRText(block=block, container_geometry=container.geometry if container else None))
            
        return result

    def _extract_font(self, el: ET.Element,
                      fallback: Optional[FontStyle] = None) -> FontStyle:
        fb = fallback or FontStyle()
        family = _parse_font_family(el.get("font-family", fb.family))
        size_attr = el.get("font-size")
        size = self.norm.pt(_parse_float(size_attr, 16)) if size_attr else fb.size_pt
        weight = el.get("font-weight", "bold" if fb.bold else "normal")
        bold = weight in ("bold", "bolder") or (
            weight.isdigit() and int(weight) >= 600
        )
        fill = _parse_color(el.get("fill"))
        return FontStyle(
            family=family,
            size_pt=size,
            bold=bold,
            italic=fb.italic,
            color=fill or fb.color,
        )

    def _parse_decorative_line(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IRLine:
        x1 = _parse_float(el.get("x1")) + dx
        y1 = _parse_float(el.get("y1")) + dy
        x2 = _parse_float(el.get("x2")) + dx
        y2 = _parse_float(el.get("y2")) + dy
        return IRLine(
            start=Point(self.norm.x(x1), self.norm.y(y1)),
            end=Point(self.norm.x(x2), self.norm.y(y2)),
            stroke=_parse_color(el.get("stroke")) or Color(0, 0, 0),
            stroke_width_emu=int(self.norm.w(_parse_float(el.get("stroke-width"), 1))),
            dashed=_parse_dash(el.get("stroke-dasharray")),
        )

    def _parse_connector(self, el: ET.Element, dx: float = 0.0, dy: float = 0.0) -> IRConnector:
        tag = _strip_ns(el.tag)
        ctype = ConnectorType(el.get("data-connector-type", "straight"))

        waypoints: List[Point] = []
        if tag == "line":
            waypoints = [
                Point(self.norm.x(_parse_float(el.get("x1")) + dx),
                      self.norm.y(_parse_float(el.get("y1")) + dy)),
                Point(self.norm.x(_parse_float(el.get("x2")) + dx),
                      self.norm.y(_parse_float(el.get("y2")) + dy)),
            ]
        elif tag == "polyline":
            pts_raw = el.get("points", "").replace(",", " ").split()
            coords = [float(v) for v in pts_raw]
            for i in range(0, len(coords), 2):
                waypoints.append(Point(
                    self.norm.x(coords[i] + dx),
                    self.norm.y(coords[i + 1] + dy)
                ))

        start_arrow = ArrowType.TRIANGLE if el.get("marker-start") else ArrowType.NONE
        end_arrow = ArrowType.TRIANGLE if el.get("marker-end") else ArrowType.NONE

        return IRConnector(
            connector_type=ctype,
            waypoints=tuple(waypoints),
            stroke=_parse_color(el.get("stroke")) or Color(0, 0, 0),
            stroke_width_emu=int(self.norm.w(_parse_float(el.get("stroke-width"), 1))),
            dashed=_parse_dash(el.get("stroke-dasharray")),
            start_arrow=start_arrow,
            end_arrow=end_arrow,
            start_ref=el.get("data-start"),
            end_ref=el.get("data-end"),
        )
