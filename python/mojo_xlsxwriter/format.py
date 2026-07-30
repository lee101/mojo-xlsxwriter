"""Cell formatting for the covered XlsxWriter API."""

from __future__ import annotations


class Format:
    def __init__(self, properties=None, workbook=None):
        self.workbook = workbook
        self.properties = dict(properties or {})
        self._style_index: int | None = None

    def _set(self, name, value=1):
        self.properties[name] = value
        return self

    def set_bold(self, bold=True):
        return self._set("bold", bold)

    def set_italic(self, italic=True):
        return self._set("italic", italic)

    def set_underline(self, underline=1):
        return self._set("underline", underline)

    def set_font_name(self, font_name):
        return self._set("font_name", font_name)

    def set_font_size(self, font_size):
        return self._set("font_size", font_size)

    def set_font_color(self, font_color):
        return self._set("font_color", font_color)

    set_font_colour = set_font_color

    def set_bg_color(self, bg_color):
        return self._set("bg_color", bg_color)

    set_bg_colour = set_bg_color

    def set_fg_color(self, fg_color):
        return self._set("fg_color", fg_color)

    set_fg_colour = set_fg_color

    def set_pattern(self, pattern=1):
        return self._set("pattern", pattern)

    def set_num_format(self, num_format):
        return self._set("num_format", num_format)

    def set_align(self, align):
        if align in {"top", "vcenter", "bottom", "vjustify", "vdistributed"}:
            return self._set("valign", align)
        return self._set("align", align)

    def set_text_wrap(self, text_wrap=True):
        return self._set("text_wrap", text_wrap)

    def set_rotation(self, rotation):
        return self._set("rotation", rotation)

    def set_indent(self, indent=1):
        return self._set("indent", indent)

    def set_shrink(self, shrink=True):
        return self._set("shrink", shrink)

    def set_border(self, border=1):
        return self._set("border", border)

    def set_border_color(self, color):
        return self._set("border_color", color)

    set_border_colour = set_border_color

    def set_left(self, style=1):
        return self._set("left", style)

    def set_right(self, style=1):
        return self._set("right", style)

    def set_top(self, style=1):
        return self._set("top", style)

    def set_bottom(self, style=1):
        return self._set("bottom", style)

    def set_locked(self, locked=True):
        return self._set("locked", locked)

    def set_hidden(self, hidden=True):
        return self._set("hidden", hidden)
