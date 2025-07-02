"""Objects related to shapes.

A shape is a visual object that appears on the drawing layer of a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.shape import WD_INLINE_SHAPE
from docx.oxml.ns import nsmap
from docx.shared import Parented

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body
    from docx.oxml.shape import CT_Inline, CT_Pict
    from docx.parts.story import StoryPart
    from docx.shared import Length


class InlineShapes(Parented):
    """Sequence of |InlineShape| instances, supporting len(), iteration, and indexed access."""

    def __init__(self, body_elm: CT_Body, parent: StoryPart):
        super(InlineShapes, self).__init__(parent)
        self._body = body_elm

    def __getitem__(self, idx: int):
        """Provide indexed access, e.g. 'inline_shapes[idx]'."""
        try:
            inline = self._inline_lst[idx]
        except IndexError:
            msg = "inline shape index [%d] out of range" % idx
            raise IndexError(msg)

        return InlineShape(inline)

    def __iter__(self):
        return (InlineShape(inline) for inline in self._inline_lst)

    def __len__(self):
        return len(self._inline_lst)

    @property
    def _inline_lst(self):
        body = self._body
        xpath = "//w:p/w:r/w:drawing/wp:inline"
        return body.xpath(xpath)


class InlineShape:
    """Proxy for an ``<wp:inline>`` element, representing the container for an inline
    graphical object."""

    def __init__(self, inline: CT_Inline):
        super(InlineShape, self).__init__()
        self._inline = inline

    @property
    def height(self) -> Length:
        """Read/write.

        The display height of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cy

    @height.setter
    def height(self, cy: Length):
        self._inline.extent.cy = cy
        self._inline.graphic.graphicData.pic.spPr.cy = cy

    @property
    def type(self):
        """The type of this inline shape as a member of
        ``docx.enum.shape.WD_INLINE_SHAPE``, e.g. ``LINKED_PICTURE``.

        Read-only.
        """
        graphicData = self._inline.graphic.graphicData
        uri = graphicData.uri
        if uri == nsmap["pic"]:
            blip = graphicData.pic.blipFill.blip
            if blip.link is not None:
                return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap["c"]:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap["dgm"]:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED

    @property
    def width(self):
        """Read/write.

        The display width of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cx

    @width.setter
    def width(self, cx: Length):
        self._inline.extent.cx = cx
        self._inline.graphic.graphicData.pic.spPr.cx = cx


class Textbox:
    """Proxy for a VML textbox element, providing access to textbox properties and content."""

    def __init__(self, pict: CT_Pict):
        super(Textbox, self).__init__()
        self._pict = pict
        self._shape = pict.shape
        self._textbox = self._shape.textbox

    @property
    def text(self) -> str:
        """Read/write.

        The text content of the textbox. For now, this is a placeholder
        that returns empty string. Text content management will be enhanced
        in future versions.
        """
        # Simplified implementation for now
        return ""

    @text.setter
    def text(self, value: str):
        """Set the text content of the textbox.

        This is a placeholder implementation. Text content management
        will be enhanced in future versions.
        """
        # Simplified implementation for now
        pass

    @property
    def left(self) -> float:
        """Read/write.

        The left position of the textbox in points.
        """
        style = self._shape.get("style") or ""
        if "margin-left:" in style:
            left_part = style.split("margin-left:")[1].split(";")[0]
            return float(left_part.replace("pt", ""))
        return 0.0

    @left.setter
    def left(self, value: float):
        """Set the left position of the textbox in points."""
        self._update_style_property("margin-left", f"{value}pt")

    @property
    def top(self) -> float:
        """Read/write.

        The top position of the textbox in points.
        """
        style = self._shape.get("style") or ""
        if "margin-top:" in style:
            top_part = style.split("margin-top:")[1].split(";")[0]
            return float(top_part.replace("pt", ""))
        return 0.0

    @top.setter
    def top(self, value: float):
        """Set the top position of the textbox in points."""
        self._update_style_property("margin-top", f"{value}pt")

    @property
    def width(self) -> float:
        """Read/write.

        The width of the textbox in points.
        """
        style = self._shape.get("style") or ""
        if "width:" in style:
            width_part = style.split("width:")[1].split(";")[0]
            return float(width_part.replace("pt", ""))
        return 0.0

    @width.setter
    def width(self, value: float):
        """Set the width of the textbox in points."""
        self._update_style_property("width", f"{value}pt")

    @property
    def height(self) -> float:
        """Read/write.

        The height of the textbox in points.
        """
        style = self._shape.get("style") or ""
        if "height:" in style:
            height_part = style.split("height:")[1].split(";")[0]
            return float(height_part.replace("pt", ""))
        return 0.0

    @height.setter
    def height(self, value: float):
        """Set the height of the textbox in points."""
        self._update_style_property("height", f"{value}pt")

    def _update_style_property(self, property_name: str, value: str):
        """Update a specific property in the style attribute."""
        style = self._shape.get("style") or ""

        # Remove existing property if present
        style_parts = [part.strip() for part in style.split(";") if part.strip()]
        style_parts = [part for part in style_parts if not part.startswith(f"{property_name}:")]

        # Add new property
        style_parts.append(f"{property_name}:{value}")

        # Update the style using lxml set method
        self._shape.set("style", ";".join(style_parts) + ";")
