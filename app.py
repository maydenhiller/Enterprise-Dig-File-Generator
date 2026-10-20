"""Enterprise Dig File Generator - single file Streamlit app.

Upload a staking report template, a dig stake cheat sheet template, one or more
Enterprise dig packages (.xlsm) and, optionally, the pipeline KMZ. Get back one
filled staking report per dig sheet in the package, plus a filled cheat sheet
for each package.

How the Enterprise template works: its StakingReport tab is nothing but
formulas that read the template's own "Dig Sheet" tab. So filling a report means
copying the dig package's "Dig (NN)" sheet into the template's "Dig Sheet" tab,
then adding the things a dig sheet cannot know - driving directions, the aerial
image, the surveyor - and leaving every formula alone.

Everything the uploads can establish is filled in. Everything they cannot is
left blank and flagged - a wrong value in a signed report is worse than a
missing one.

Deliberately kept as one file with no local package imports: Streamlit Cloud
deployments break silently when a subpackage does not make it into the repo.

Latitude, longitude, elevation, EDOC, survey date and photos are field
measurements and stay manual.
"""

from __future__ import annotations

import datetime as _dt
import io
import math
import os
import re
import tempfile
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from copy import copy
from dataclasses import dataclass, field
from datetime import date
from typing import Iterable, Optional
from xml.etree import ElementTree
from xml.sax.saxutils import escape

import requests
from PIL import Image, ImageDraw, ImageFont


# ==========================================================================
# MODELS
# ==========================================================================

UNKNOWN = "Unknown"

# "AGM 090 - Sta. 3248+74 - 37.9' ENE of C/L R Ave" and "MLV  1135  9329+66  Check
# Valve" are the two shapes a reference cell takes. Both start with a type and a
# number; that is all the report and the cheat sheet want.
REFERENCE_NAME_RE = re.compile(r"^\s*(AGM|MLV|MLB|BV)\s*0*(\d+)", re.I)


def reference_name(reference) -> Optional[str]:
    """'AGM 090 - Sta. 3248+74 - ...' -> 'AGM 090'. 'MLV  1135  9329+66 ...' -> 'MLV 1135'.

    The number keeps the zero padding the package prints (AGM 090, not AGM 90) so
    the name matches the marker as it is labelled in the field.
    """
    if reference is None:
        return None
    text = re.sub(r"\s+", " ", str(reference)).strip()
    if not text:
        return None
    match = re.match(r"^(AGM|MLV|MLB|BV)\s+(\d+)", text, re.I)
    if match:
        return f"{match.group(1).upper()} {match.group(2)}"
    head = re.split(r"\s+-\s+|,", text)[0].strip()
    return head or None


def station_label(feet) -> str:
    """332977 -> '3329+77', the way the dig package prints it."""
    if feet is None:
        return ""
    feet = int(round(float(feet)))
    return f"{feet // 100}+{feet % 100:02d}"


def _number(value) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


@dataclass
class Dig:
    """One dig sheet from a package, plus everything derived for it."""

    # --- identity --------------------------------------------------------
    package_id: str = ""          # the assessment ID, e.g. "115"
    package_file: str = ""
    sheet_name: str = ""          # e.g. "Dig (02)"
    name: str = ""                # as the sheet prints it: "Dig # 02B"
    number: str = ""              # "02B"
    feature_id: str = ""

    # --- straight off the dig sheet --------------------------------------
    odometer: Optional[float] = None
    station_ft: Optional[int] = None
    us_ref: str = ""
    us_ref_odometer: Optional[float] = None
    us_ref_station_ft: Optional[int] = None
    us_ref_distance: Optional[float] = None       # D11, feet from the U/S reference
    us_ref_lat: Optional[float] = None
    us_ref_lon: Optional[float] = None
    ds_ref: str = ""
    ds_ref_odometer: Optional[float] = None
    ds_ref_station_ft: Optional[int] = None
    ds_ref_distance: Optional[float] = None       # J11, feet to the D/S reference
    ds_ref_lat: Optional[float] = None
    ds_ref_lon: Optional[float] = None
    us_weld_distance: Optional[float] = None      # G33
    us_weld_id: str = ""                          # Site Ref: U/S GW ID_No.
    us_weld_station_ft: Optional[int] = None
    ds_weld_id: str = ""
    ds_weld_station_ft: Optional[int] = None
    us_weld_lat: Optional[float] = None           # Site Ref LAT/LON of those welds
    us_weld_lon: Optional[float] = None
    ds_weld_lat: Optional[float] = None
    ds_weld_lon: Optional[float] = None
    ds_weld_distance: Optional[float] = None      # I33
    joint_length: Optional[float] = None
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    hca: str = ""                                 # "YES" / "NO" as printed
    county_state: str = ""
    legal_description: str = ""
    tract_number: str = ""
    alignment_sheet: str = ""
    line_segment: str = ""
    header: str = ""                              # A5, the line / assessment banner

    # Every non-empty cell of the source sheet, {"G8": "Dig # 02B", ...}. This is
    # what gets copied into the template's "Dig Sheet" tab.
    cells: dict = field(default_factory=dict)

    # True when AS stationing rises with absolute odometer on this line.
    station_ascends: bool = True

    # --- filled in by the app / by hand ----------------------------------
    directions: str = ""
    aerial_image: Optional[bytes] = None   # JPEG bytes
    notes: str = ""
    warnings: list = field(default_factory=list)

    # --- convenience -----------------------------------------------------
    @property
    def flow_bearing(self) -> Optional[float]:
        """Compass bearing (degrees clockwise from north) the product flows, U/S -> D/S.

        From the two welds either side of the dig; the U/S and D/S reference
        points are the fallback when the Site Ref tab has no coordinates.
        """
        for a_lat, a_lon, b_lat, b_lon in (
            (self.us_weld_lat, self.us_weld_lon, self.ds_weld_lat, self.ds_weld_lon),
            (self.us_ref_lat, self.us_ref_lon, self.ds_ref_lat, self.ds_ref_lon),
        ):
            if None in (a_lat, a_lon, b_lat, b_lon):
                continue
            if (a_lat, a_lon) == (b_lat, b_lon):
                continue
            p1, p2 = math.radians(a_lat), math.radians(b_lat)
            dl = math.radians(b_lon - a_lon)
            x = math.sin(dl) * math.cos(p2)
            y = math.cos(p1) * math.sin(p2) - math.sin(p1) * math.cos(p2) * math.cos(dl)
            return math.degrees(math.atan2(x, y)) % 360.0
        return None

    @property
    def upstream_reference(self) -> str:
        return reference_name(self.us_ref) or ""

    @property
    def downstream_reference(self) -> str:
        return reference_name(self.ds_ref) or ""

    @property
    def nearest_reference_label(self) -> str:
        """The cheat sheet's 'Feet from AGM U/S-D/S' cell, e.g. "AGM 090 1717.92'"."""
        candidates = []
        if self.us_ref_distance is not None and self.upstream_reference:
            candidates.append((self.us_ref_distance, self.upstream_reference))
        if self.ds_ref_distance is not None and self.downstream_reference:
            candidates.append((self.ds_ref_distance, self.downstream_reference))
        if not candidates:
            return ""
        distance, label = min(candidates, key=lambda pair: pair[0])
        return f"{label} {distance:.2f}'"

    @property
    def hca_label(self) -> str:
        """'YES' -> 'Yes', the way the cheat sheet writes it."""
        text = (self.hca or "").strip()
        return text.title() if text else ""

    @property
    def output_basename(self) -> str:
        return f"AID {self.package_id} - {self.name.replace('#', '').strip()}_Staking Report".replace("  ", " ")

    @property
    def key(self) -> tuple:
        return (self.package_id, self.sheet_name)



# ==========================================================================
# XLSMPATCH
# ==========================================================================
# Edit an .xlsm in place at the package level.
#
# openpyxl cannot round-trip this template: it drops DrawingML shapes, which
# would take the macro-linked buttons and the logos with them. So the workbook is
# treated as what it is - a zip of XML parts - and only the bytes that need to
# change are changed. Styles, macros, buttons, print settings, comments and the
# logos all survive untouched.

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

EXCEL_EPOCH = _dt.datetime(1899, 12, 30)


@dataclass
class ImageSlot:
    """A picture already anchored in the template, ready to be swapped out.

    ``drawing_path``, ``rel_id`` and the ``pic_*`` offsets are what let the
    aerial be given its own media part instead of overwriting the one already
    in the package. In the Enterprise template the StakingReport image slot and
    the Pre_Dig_Photos placeholders are the *same* media part (image6.jpeg), so
    overwriting its bytes would replace the pre-dig photo placeholders too.
    """

    media_path: str
    extension: str
    width_emu: int
    height_emu: int
    drawing_path: str = ""
    rel_id: str = ""
    pic_start: int = 0
    pic_end: int = 0

    @property
    def aspect(self) -> float:
        if not self.height_emu:
            return 2.0
        return self.width_emu / self.height_emu


class XlsmPatcher:
    def __init__(self, data: bytes):
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            self.names = list(archive.namelist())
            self.parts = {name: archive.read(name) for name in self.names}
            self.infos = {info.filename: info for info in archive.infolist()}
        self._sheet_paths = self._map_sheets()

    # -- structure ------------------------------------------------------
    def _text(self, path: str) -> str:
        return self.parts[path].decode("utf-8")

    def _map_sheets(self) -> dict:
        workbook = self._text("xl/workbook.xml")
        rels = self._text("xl/_rels/workbook.xml.rels")

        targets = {
            match.group(1): match.group(2)
            for match in re.finditer(
                r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels
            )
        }
        mapping = {}
        for match in re.finditer(r"<sheet\b[^>]*/?>", workbook):
            tag = match.group(0)
            name = re.search(r'name="([^"]*)"', tag)
            rid = re.search(r'r:id="([^"]*)"', tag)
            if not name or not rid:
                continue
            target = targets.get(rid.group(1), "")
            if not target:
                continue
            path = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
            path = path.replace("/./", "/")
            mapping[name.group(1)] = path
        return mapping

    def sheet_names(self) -> list:
        return list(self._sheet_paths)

    # -- cell values ----------------------------------------------------
    def set_values(self, sheet_name: str, values: dict, literal: bool = False) -> list:
        """Write {'C19': value} into a sheet. Returns cells it could not place.

        ``literal=True`` writes text exactly as given. Without it, text that
        starts with "=" becomes a formula - right for the few cells the app
        sets itself, wrong for data copied out of a dig sheet.
        """
        path = self._sheet_paths.get(sheet_name)
        if path is None:
            raise KeyError(f"No sheet named {sheet_name!r} in this workbook")

        xml = self._text(path)
        missing = []
        for ref, value in values.items():
            if value is None:
                continue
            xml, ok = _write_cell(xml, ref, value, literal)
            if not ok:
                missing.append(ref)
        self.parts[path] = xml.encode("utf-8")
        return missing

    def restyle(self, sheet_name: str, refs, size: Optional[float] = None,
                shrink: bool = False) -> None:
        """Give cells a copy of their own style with a smaller font and/or
        shrink-to-fit. The original style is left alone, so other cells that
        share it are unaffected."""
        path = self._sheet_paths.get(sheet_name)
        if path is None:
            return
        xml = self._text(path)
        styles = self._text("xl/styles.xml")

        fonts_match = re.search(r"(<fonts\b[^>]*>)(.*?)(</fonts>)", styles, re.S)
        xfs_match = re.search(r"(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)", styles, re.S)
        if not fonts_match or not xfs_match:
            return
        font_pattern = r"<font\b[^>]*?/>|<font\b[^>]*?>.*?</font>"
        xf_pattern = r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>"
        fonts = re.findall(font_pattern, fonts_match.group(2), re.S)
        xfs = re.findall(xf_pattern, xfs_match.group(2), re.S)

        made = {}

        def variant(index: int) -> int:
            if index in made:
                return made[index]
            xf = xfs[index]
            if size is not None:
                font_match = re.search(r'fontId="(\d+)"', xf)
                font_id = int(font_match.group(1)) if font_match else 0
                font = fonts[font_id]
                if "<sz " in font:
                    font = re.sub(r'<sz val="[^"]*"\s*/>', f'<sz val="{size:g}"/>', font, 1)
                else:
                    font = font.replace("<font>", f'<font><sz val="{size:g}"/>', 1)
                fonts.append(font)
                new_font = len(fonts) - 1
                xf = (re.sub(r'fontId="\d+"', f'fontId="{new_font}"', xf, 1)
                      if font_match else xf.replace("<xf ", f'<xf fontId="{new_font}" ', 1))
                if "applyFont=" in xf:
                    xf = re.sub(r'applyFont="\d"', 'applyFont="1"', xf, 1)
                else:
                    xf = xf.replace("<xf ", '<xf applyFont="1" ', 1)
            if shrink:
                if xf.endswith("/>") and "</xf>" not in xf:
                    xf = xf[:-2] + "></xf>"
                if "<alignment" in xf:
                    xf = re.sub(r'\swrapText="\d"', "", xf, 1)
                    if "shrinkToFit=" not in xf:
                        xf = xf.replace("<alignment", '<alignment shrinkToFit="1"', 1)
                else:
                    xf = re.sub(r"(<xf\b[^>]*?>)", r'\1<alignment shrinkToFit="1"/>', xf, 1)
                if "applyAlignment=" in xf:
                    xf = re.sub(r'applyAlignment="\d"', 'applyAlignment="1"', xf, 1)
                else:
                    xf = xf.replace("<xf ", '<xf applyAlignment="1" ', 1)
            xfs.append(xf)
            made[index] = len(xfs) - 1
            return made[index]

        for ref in refs:
            match = re.search(rf'<c r="{ref}"([^>]*?)(/?)>', xml)
            if not match:
                continue
            style = re.search(r'\ss="(\d+)"', match.group(1))
            index = int(style.group(1)) if style else 0
            if index >= len(xfs):
                continue
            attrs = match.group(1)
            new_index = variant(index)
            attrs = (re.sub(r'\ss="\d+"', f' s="{new_index}"', attrs, 1)
                     if style else attrs + f' s="{new_index}"')
            xml = xml[:match.start()] + f'<c r="{ref}"{attrs}{match.group(2)}>' + xml[match.end():]

        if not made:
            return
        fonts_head = re.sub(r'count="\d+"', f'count="{len(fonts)}"', fonts_match.group(1), 1)
        xfs_head = re.sub(r'count="\d+"', f'count="{len(xfs)}"', xfs_match.group(1), 1)
        # Replace the later section first so the earlier match offsets stay valid.
        first, second = sorted(
            [(fonts_match, fonts_head, "".join(fonts)), (xfs_match, xfs_head, "".join(xfs))],
            key=lambda item: item[0].start(), reverse=True)
        for match, head, body in (first, second):
            styles = styles[:match.start()] + head + body + match.group(3) + styles[match.end():]
        self.parts["xl/styles.xml"] = styles.encode("utf-8")
        self.parts[path] = xml.encode("utf-8")

    def cell_text(self, sheet_name: str, ref: str) -> str:
        """The text a cell shows: an inline or shared string, or a formula's last result."""
        path = self._sheet_paths.get(sheet_name)
        if path is None:
            return ""
        match = _cell_pattern(ref).search(self._text(path))
        if not match or not match.group("body"):
            return ""
        attrs, body = match.group("attrs"), match.group("body")
        kind = re.search(r'\st="(\w+)"', attrs)
        kind = kind.group(1) if kind else ""
        if kind == "inlineStr":
            return _unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S)))
        value = re.search(r"<v>(.*?)</v>", body, re.S)
        if value is None:
            return ""
        if kind == "s" and "xl/sharedStrings.xml" in self.parts:
            items = re.findall(r"<si>(.*?)</si>", self._text("xl/sharedStrings.xml"), re.S)
            index = int(value.group(1))
            if index < len(items):
                return _unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", items[index], re.S)))
            return ""
        return _unescape(value.group(1))

    def box_size(self, sheet_name: str, ref: str) -> Optional[tuple]:
        """(width in pixels, height in points) of the merged box a cell starts."""
        path = self._sheet_paths.get(sheet_name)
        if path is None:
            return None
        xml = self._text(path)
        merge = re.search(rf'<mergeCell ref="{ref}:([A-Z]+)(\d+)"', xml)
        start_row = int(re.search(r"\d+", ref).group(0))
        first_col = _column_index(re.match(r"[A-Z]+", ref).group(0))
        if merge:
            last_col, last_row = _column_index(merge.group(1)), int(merge.group(2))
        else:
            last_col, last_row = first_col, start_row

        default_width = re.search(r'<sheetFormatPr[^>]*defaultColWidth="([\d.]+)"', xml)
        default_width = float(default_width.group(1)) if default_width else 8.43
        default_height = re.search(r'<sheetFormatPr[^>]*defaultRowHeight="([\d.]+)"', xml)
        default_height = float(default_height.group(1)) if default_height else 15.0
        widths = {}
        for col in re.finditer(r"<col\b[^>]*/>", xml):
            tag = col.group(0)
            low, high = re.search(r'min="(\d+)"', tag), re.search(r'max="(\d+)"', tag)
            width = re.search(r'width="([\d.]+)"', tag)
            if low and high and width:
                for index in range(int(low.group(1)), min(int(high.group(1)), 200) + 1):
                    widths[index] = float(width.group(1))
        heights = {}
        for row in re.finditer(r"<row\b[^>]*>", xml):
            tag = row.group(0)
            number, height = re.search(r' r="(\d+)"', tag), re.search(r' ht="([\d.]+)"', tag)
            if number and height:
                heights[int(number.group(1))] = float(height.group(1))
        width_px = sum(widths.get(i, default_width) * 7 for i in range(first_col, last_col + 1))
        height_pt = sum(heights.get(r, default_height) for r in range(start_row, last_row + 1))
        return width_px, height_pt

    def set_cached(self, sheet_name: str, values: dict) -> None:
        """Refresh the value a formula cell last computed, leaving the formula.

        The template's formulas carry whatever Excel last computed (zeros, in a
        blank template). Excel recalculates on open, but anything that previews
        the file without calculating - a phone, a mail client, SharePoint's
        viewer - would show those stale zeros. Writing the right cached value
        beside each formula makes the file correct everywhere.
        """
        path = self._sheet_paths.get(sheet_name)
        if path is None:
            return
        xml = self._text(path)
        for ref, value in values.items():
            xml = _set_cached_value(xml, ref, value)
        self.parts[path] = xml.encode("utf-8")

    # -- images ---------------------------------------------------------
    def find_image_slot(self, sheet_name: str) -> Optional[ImageSlot]:
        """The largest picture anchored on a sheet - the aerial image slot."""
        drawing = self._drawing_for(sheet_name)
        if drawing is None:
            return None
        drawing_xml = self._text(drawing)
        rels_path = _rels_path(drawing)
        if rels_path not in self.parts:
            return None
        rels = self._text(rels_path)
        targets = {
            match.group(1): match.group(2)
            for match in re.finditer(
                r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels
            )
        }

        best = None
        for block in re.finditer(r"<xdr:pic\b.*?</xdr:pic>", drawing_xml, re.S):
            chunk = block.group(0)
            embed = re.search(r'r:embed="([^"]+)"', chunk)
            extent = re.search(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"', chunk)
            if not embed or not extent:
                continue
            width, height = int(extent.group(1)), int(extent.group(2))
            target = targets.get(embed.group(1))
            if not target:
                continue
            media = _resolve(drawing, target)
            if media not in self.parts:
                continue
            if best is None or width * height > best.width_emu * best.height_emu:
                best = ImageSlot(
                    media_path=media,
                    extension=media.rsplit(".", 1)[-1].lower(),
                    width_emu=width,
                    height_emu=height,
                    drawing_path=drawing,
                    rel_id=embed.group(1),
                    pic_start=block.start(),
                    pic_end=block.end(),
                )
        return best

    def replace_image(self, slot: ImageSlot, image_bytes: bytes,
                      extension: str = "png") -> None:
        """Point this one picture at a new media part.

        The obvious implementation - overwrite ``slot.media_path`` - is wrong
        for this template. Its StakingReport image slot and the Pre_Dig_Photos
        placeholders share a single media part, so overwriting the bytes put
        the aerial in place of the pre-dig photo placeholders as well. Those
        are filled in by hand after the survey, so they have to be left alone.

        Instead the aerial is added as a new part and only the StakingReport
        picture's own relationship is repointed at it. Every other picture,
        on this sheet or any other, keeps the media it already had.
        """
        if not slot.drawing_path or not slot.rel_id:
            # Nothing to repoint - fall back to the in-place swap.
            self.parts[slot.media_path] = image_bytes
            return

        extension = extension.lower().lstrip(".")
        media_path = self._new_media_path(extension)
        self.parts[media_path] = image_bytes
        self._ensure_content_type(extension)

        rels_path = _rels_path(slot.drawing_path)
        rels = self._text(rels_path)
        new_id = self._new_rel_id(rels)
        target = _relative_part(media_path, slot.drawing_path)
        entry = (
            f'<Relationship Id="{new_id}" Type="http://schemas.openxmlformats.org/'
            f'officeDocument/2006/relationships/image" Target="{target}"/>'
        )
        self.parts[rels_path] = rels.replace(
            "</Relationships>", entry + "</Relationships>", 1
        ).encode("utf-8")

        drawing_xml = self._text(slot.drawing_path)
        chunk = drawing_xml[slot.pic_start:slot.pic_end]
        chunk = chunk.replace(f'r:embed="{slot.rel_id}"', f'r:embed="{new_id}"', 1)
        self.parts[slot.drawing_path] = (
            drawing_xml[:slot.pic_start] + chunk + drawing_xml[slot.pic_end:]
        ).encode("utf-8")

    def _new_media_path(self, extension: str) -> str:
        index = 1
        while f"xl/media/image{index}.{extension}" in self.parts:
            index += 1
        while any(
            name.startswith(f"xl/media/image{index}.") for name in self.parts
        ):
            index += 1
        return f"xl/media/image{index}.{extension}"

    @staticmethod
    def _new_rel_id(rels: str) -> str:
        used = {int(n) for n in re.findall(r'Id="rId(\d+)"', rels)}
        candidate = max(used) + 1 if used else 1
        return f"rId{candidate}"

    def _ensure_content_type(self, extension: str) -> None:
        path = "[Content_Types].xml"
        if path not in self.parts:
            return
        xml = self._text(path)
        if f'Extension="{extension}"' in xml:
            return
        kind = "image/jpeg" if extension in ("jpg", "jpeg") else f"image/{extension}"
        entry = f'<Default Extension="{extension}" ContentType="{kind}"/>'
        self.parts[path] = xml.replace("<Types", "<Types", 1).replace(
            "</Types>", entry + "</Types>", 1
        ).encode("utf-8")

    def _drawing_for(self, sheet_name: str) -> Optional[str]:
        path = self._sheet_paths.get(sheet_name)
        if path is None:
            return None
        rels_path = _rels_path(path)
        if rels_path not in self.parts:
            return None
        rels = self._text(rels_path)
        match = re.search(
            r'<Relationship[^>]*Type="[^"]*/drawing"[^>]*Target="([^"]+)"', rels
        )
        if not match:
            match = re.search(
                r'<Relationship[^>]*Target="([^"]+)"[^>]*Type="[^"]*/drawing"', rels
            )
        if not match:
            return None
        return _resolve(path, match.group(1))

    def force_full_recalc(self) -> None:
        """Make Excel recalculate every formula when the file is opened.

        Formula cells carry the value Excel last computed. The template's
        phone number is an XLOOKUP against the surveyor name with a stale
        cached result, so without this the report opens showing the old value
        until something nudges the cell.
        """
        path = "xl/workbook.xml"
        if path not in self.parts:
            return
        xml = self._text(path)
        match = re.search(r"<calcPr\b[^>]*/?>", xml)
        if match:
            tag = match.group(0)
            if "fullCalcOnLoad" in tag:
                return
            updated = tag.rstrip("/>").rstrip() + ' fullCalcOnLoad="1"/>'
            xml = xml[: match.start()] + updated + xml[match.end():]
        else:
            xml = xml.replace(
                "</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>'
            )
        self.parts[path] = xml.encode("utf-8")

    def keep_only_sheet(self, keep: str, rename_to: Optional[str] = None) -> None:
        """Drop every other sheet from the workbook, optionally renaming the one kept.

        Everything that belongs only to a dropped sheet (its XML, drawing,
        relationships, defined names, calculation chain) goes with it.
        """
        workbook = self._text("xl/workbook.xml")
        rels = self._text("xl/_rels/workbook.xml.rels")
        sheets = list(re.finditer(r"<sheet\b[^>]*/>", workbook))
        order = [re.search(r'name="([^"]*)"', m.group(0)).group(1) for m in sheets]
        if keep not in order:
            raise KeyError(f"No sheet named {keep!r} in this workbook")
        keep_index = order.index(keep)

        dropped_paths = []
        for match, name in zip(sheets, order):
            if name == keep:
                continue
            rid = re.search(r'r:id="([^"]*)"', match.group(0)).group(1)
            rel = re.search(rf'<Relationship\b[^>]*Id="{rid}"[^>]*/>', rels)
            if rel:
                rels = rels.replace(rel.group(0), "")
            path = self._sheet_paths.get(name)
            if path:
                dropped_paths.append(path)
            workbook = workbook.replace(match.group(0), "")

        # Defined names are tied to a sheet by position.
        def fix_name(match):
            tag = match.group(0)
            local = re.search(r'localSheetId="(\d+)"', tag)
            if local:
                if int(local.group(1)) != keep_index:
                    return ""
                tag = tag.replace(local.group(0), 'localSheetId="0"')
            return tag
        workbook = re.sub(r"<definedName\b[^>]*>.*?</definedName>", fix_name, workbook, flags=re.S)
        workbook = re.sub(r"<definedNames>\s*</definedNames>", "", workbook)

        if rename_to and rename_to != keep:
            workbook = workbook.replace(f'name="{keep}"', f'name="{escape(rename_to)}"', 1)
            workbook = re.sub(rf"(?<![\w.]){re.escape(keep)}!", f"{rename_to}!", workbook)
        self.parts["xl/workbook.xml"] = workbook.encode("utf-8")

        # Excel rebuilds the calculation chain; a stale one names cells on dropped sheets.
        chain = "xl/calcChain.xml"
        if chain in self.parts:
            rels = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', "", rels)
            dropped_paths.append(chain)
        self.parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

        dropped = set()
        for path in dropped_paths:
            dropped.add(path)
            sheet_rels = _rels_path(path)
            if sheet_rels in self.parts:
                dropped.add(sheet_rels)
                for target in re.findall(r'Target="([^"]+)"', self._text(sheet_rels)):
                    resolved = _resolve(path, target)
                    if resolved.startswith("xl/printerSettings/") and resolved in self.parts:
                        dropped.add(resolved)
                    if resolved.startswith("xl/drawings/") and resolved in self.parts:
                        dropped.add(resolved)
                        drawing_rels = _rels_path(resolved)
                        if drawing_rels in self.parts:
                            dropped.add(drawing_rels)

        content_types = self._text("[Content_Types].xml")
        for path in dropped:
            content_types = re.sub(
                rf'<Override\b[^>]*PartName="/{re.escape(path)}"[^>]*/>', "", content_types)
        self.parts["[Content_Types].xml"] = content_types.encode("utf-8")

        for path in dropped:
            self.parts.pop(path, None)
        # Media that only a dropped drawing used.
        referenced = set()
        for name, data in self.parts.items():
            if name.endswith(".rels"):
                base = name.replace("_rels/", "").rsplit(".rels", 1)[0]
                for target in re.findall(r'Target="([^"]+)"', data.decode("utf-8")):
                    referenced.add(_resolve(base, target))
        for name in [n for n in self.parts if n.startswith("xl/media/")]:
            if name not in referenced:
                dropped.add(name)
                self.parts.pop(name, None)
        self.names = [n for n in self.names if n not in dropped]
        self._sheet_paths = self._map_sheets()

    # -- output ---------------------------------------------------------
    def to_bytes(self) -> bytes:
        buffer = io.BytesIO()
        # Parts added since the template was opened (the aerial's own media
        # part) are not in self.names, so they are appended at the end.
        added = [name for name in self.parts if name not in self.infos]
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name in self.names + added:
                info = self.infos.get(name)
                if info is None:
                    new = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                else:
                    new = zipfile.ZipInfo(name, date_time=info.date_time)
                    new.external_attr = info.external_attr
                new.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(new, self.parts[name])
        return buffer.getvalue()


# ---------------------------------------------------------------------------
# Cell writing
# ---------------------------------------------------------------------------

def _cell_pattern(ref: str) -> re.Pattern:
    return re.compile(
        rf'<c r="{ref}"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</c>)', re.S
    )


def _style_of(attrs: str) -> str:
    match = re.search(r'\ss="(\d+)"', attrs or "")
    return f' s="{match.group(1)}"' if match else ""


_ERROR_VALUES = {"#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"}


def _excel_time(value: _dt.time) -> float:
    return (value.hour * 3600 + value.minute * 60 + value.second
            + value.microsecond / 1e6) / 86400.0


def _render(ref: str, style: str, value, literal: bool = False) -> str:
    if isinstance(value, bool):
        return f'<c r="{ref}"{style} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{style}><v>{value!r}</v></c>'
    if isinstance(value, _dt.datetime):
        serial = (value - EXCEL_EPOCH).total_seconds() / 86400.0
        return f'<c r="{ref}"{style}><v>{serial!r}</v></c>'
    if isinstance(value, _dt.date):
        serial = (_dt.datetime(value.year, value.month, value.day) - EXCEL_EPOCH).days
        return f'<c r="{ref}"{style}><v>{serial}</v></c>'
    if isinstance(value, _dt.time):
        return f'<c r="{ref}"{style}><v>{_excel_time(value)!r}</v></c>'
    if isinstance(value, _dt.timedelta):
        return f'<c r="{ref}"{style}><v>{value.total_seconds() / 86400.0!r}</v></c>'

    text = str(value)
    if literal and text in _ERROR_VALUES:
        return f'<c r="{ref}"{style} t="e"><v>{text}</v></c>'
    if text.startswith("=") and not literal:
        return f'<c r="{ref}"{style}><f>{escape(text[1:])}</f></c>'
    return (
        f'<c r="{ref}"{style} t="inlineStr">'
        f'<is><t xml:space="preserve">{escape(text)}</t></is></c>'
    )


def _set_cached_value(xml: str, ref: str, value) -> str:
    """Replace the <v> of a formula cell, keeping its <f> and style."""
    match = _cell_pattern(ref).search(xml)
    if not match or not match.group("body"):
        return xml
    body = match.group("body")
    formula = re.search(r"<f\b[^>]*?(?:/>|>.*?</f>)", body, re.S)
    if not formula:
        return xml
    attrs = re.sub(r'\st="[^"]*"', "", match.group("attrs") or "")

    if value is None or value == "":
        kind, text = ' t="str"', ""
    elif isinstance(value, bool):
        kind, text = ' t="b"', "1" if value else "0"
    elif isinstance(value, (int, float)):
        kind, text = "", repr(value)
    else:
        kind, text = ' t="str"', escape(str(value))
    cell = f'<c r="{ref}"{attrs}{kind}>{formula.group(0)}<v>{text}</v></c>'
    return xml[: match.start()] + cell + xml[match.end():]


def _write_cell(xml: str, ref: str, value, literal: bool = False) -> tuple:
    match = _cell_pattern(ref).search(xml)
    if match:
        replacement = _render(ref, _style_of(match.group("attrs")), value, literal)
        return xml[: match.start()] + replacement + xml[match.end():], True

    inserted = _insert_cell(xml, ref, value, literal)
    return (inserted, True) if inserted is not None else (xml, False)


_COLUMN_RE = re.compile(r"([A-Z]+)(\d+)")


def _column_index(letters: str) -> int:
    index = 0
    for character in letters:
        index = index * 26 + (ord(character) - 64)
    return index


def _insert_cell(xml: str, ref: str, value, literal: bool = False) -> Optional[str]:
    """Add a cell to an existing row, keeping columns in order."""
    parsed = _COLUMN_RE.fullmatch(ref)
    if not parsed:
        return None
    letters, number = parsed.group(1), parsed.group(2)
    target = _column_index(letters)

    row = re.search(rf'<row[^>]*\br="{number}"[^>]*>(.*?)</row>', xml, re.S)
    if not row:
        return None

    body = row.group(1)
    style = ""
    insert_at = len(body)
    for cell in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>.*?</c>)', body, re.S):
        index = _column_index(cell.group(1))
        if index < target:
            style = _style_of(cell.group(2)) or style
        else:
            insert_at = cell.start()
            break

    new_body = body[:insert_at] + _render(ref, style, value, literal) + body[insert_at:]
    return xml[: row.start(1)] + new_body + xml[row.end(1):]


# ---------------------------------------------------------------------------
# Package path helpers
# ---------------------------------------------------------------------------

def _rels_path(part: str) -> str:
    folder, _, name = part.rpartition("/")
    return f"{folder}/_rels/{name}.rels"


def _relative_part(part: str, base_part: str) -> str:
    """``part`` expressed relative to the folder holding ``base_part``.

    The inverse of ``_resolve``, used when adding a relationship.
    """
    target = part.split("/")
    folder = base_part.rpartition("/")[0].split("/") if "/" in base_part else []
    shared = 0
    while (shared < len(folder) and shared < len(target) - 1
           and folder[shared] == target[shared]):
        shared += 1
    return "/".join([".."] * (len(folder) - shared) + target[shared:])


def _resolve(base_part: str, target: str) -> str:
    folder = base_part.rpartition("/")[0]
    parts = folder.split("/") if folder else []
    for chunk in target.split("/"):
        if chunk in ("", "."):
            continue
        if chunk == "..":
            if parts:
                parts.pop()
        else:
            parts.append(chunk)
    return "/".join(parts)



# ==========================================================================
# DIGPACKAGE
# ==========================================================================
# Read an Enterprise dig package (.xlsm).
#
# A package is one workbook with a "Dig (NN)" sheet per dig, each laid out
# identically, plus Dashboard / Dig List / Site Ref tabs that are not needed here.
# The values the staking report shows are the cached results of formulas on those
# sheets, so the package is read with data_only=True - what Excel last computed.
#
# Fixed cell positions on a Dig sheet (verified on both example packages):
#
#   G8   dig name ("Dig # 02B")        Q15  dig number ("02B")     Q22  Feature ID
#   G9   anomaly odometer              I9   anomaly station (feet)
#   B10  U/S reference text            N10  D/S reference text
#   B11  U/S ref odometer              N11  D/S ref odometer
#   B12  U/S ref station               N12  D/S ref station
#   D11  feet from U/S ref             J11  feet to D/S ref
#   G33  anomaly -> U/S weld (ft)      I33  anomaly -> D/S weld (ft)
#   AR13/AS13  U/S ref lat/lon         AR16/AS16  D/S ref lat/lon
#   AR15/AS15  dig lat/lon
#   C36 county, state  C37 legal  C38 tract  C41 alignment sheet  C42 line
#   D49  HCA ("YES"/"NO")              A5   line / assessment banner

DIG_SHEET_RE = re.compile(r"^Dig \((\d+)\)$")
MAX_DIG_ROW = 57
MAX_DIG_COLUMN = 73  # BU


def _column_letters(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _package_id(workbook, filename: str) -> str:
    """The assessment ID: Dashboard's Asmt ID, else the AID_nnn in the filename."""
    match = re.search(r"AID[_\s-]*(\d+)", filename or "", re.I)
    try:
        sheet = workbook["Dashboard"]
        row = next(sheet.iter_rows(min_row=4, max_row=4, max_col=12, values_only=True), ())
        # C4 = Line ID, E4 = Asmt ID on both example packages.
        if len(row) >= 5 and row[4] not in (None, ""):
            value = row[4]
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            return str(value).strip()
    except Exception:  # noqa: BLE001 - the filename is a fine fallback
        pass
    return match.group(1) if match else (filename or "package").rsplit(".", 1)[0]


def _read_dig_sheet(sheet) -> dict:
    cells = {}
    for row_index, row in enumerate(
        sheet.iter_rows(min_row=1, max_row=MAX_DIG_ROW, max_col=MAX_DIG_COLUMN,
                        values_only=True),
        start=1,
    ):
        for column_index, value in enumerate(row, start=1):
            if value is None or value == "":
                continue
            cells[f"{_column_letters(column_index)}{row_index}"] = value
    return cells


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def _station(value) -> Optional[int]:
    number = _number(value)
    return int(round(number)) if number is not None else None


def dig_from_cells(cells: dict, package_id: str, package_file: str,
                   sheet_name: str) -> Dig:
    """Build a Dig from one sheet's cell dictionary."""
    number = _text(cells.get("Q15"))
    name = _text(cells.get("G8")) or (f"Dig # {number}" if number else sheet_name)

    dig = Dig(
        package_id=package_id,
        package_file=package_file,
        sheet_name=sheet_name,
        name=name,
        number=number,
        feature_id=_text(cells.get("Q22") or cells.get("T15")),
        odometer=_number(cells.get("G9")) if cells.get("G9") is not None
        else _number(cells.get("W15")),
        station_ft=_station(cells.get("I9")),
        us_ref=_text(cells.get("B10")),
        us_ref_odometer=_number(cells.get("B11")),
        us_ref_station_ft=_station(cells.get("B12")),
        us_ref_distance=_number(cells.get("D11")),
        us_ref_lat=_number(cells.get("AR13")),
        us_ref_lon=_number(cells.get("AS13")),
        ds_ref=_text(cells.get("N10")),
        ds_ref_odometer=_number(cells.get("N11")),
        ds_ref_station_ft=_station(cells.get("N12")),
        ds_ref_distance=_number(cells.get("J11")),
        ds_ref_lat=_number(cells.get("AR16")),
        ds_ref_lon=_number(cells.get("AS16")),
        us_weld_distance=_number(cells.get("G33")),
        ds_weld_distance=_number(cells.get("I33")),
        joint_length=_number(cells.get("H24")),
        latitude=_number(cells.get("AR15")),
        longitude=_number(cells.get("AS15")),
        hca=_text(cells.get("D49")),
        county_state=_text(cells.get("C36")),
        legal_description=_text(cells.get("C37")),
        tract_number=_text(cells.get("C38")),
        alignment_sheet=_text(cells.get("C41")),
        line_segment=_text(cells.get("C42")),
        header=_text(cells.get("A5")),
        cells=cells,
    )

    # A weld distance the sheet does not print can be recovered from the others.
    if dig.us_weld_distance is None and cells.get("Q24") is not None:
        dig.us_weld_distance = _number(cells.get("Q24"))
    if (dig.ds_weld_distance is None and dig.joint_length is not None
            and dig.us_weld_distance is not None):
        dig.ds_weld_distance = round(dig.joint_length - dig.us_weld_distance, 2)

    dig.station_ascends = detect_station_ascends(dig)
    dig.warnings.extend(check_dig(dig))
    return dig


def detect_station_ascends(dig: Dig) -> bool:
    """True when stationing rises with odometer, read from the dig's own references.

    Both example packages ascend (station 332977 at odometer 57975.75), but the
    cheat sheet's formulas flip sign on a descending line, so it is detected
    per dig rather than assumed.
    """
    pairs = [
        (dig.us_ref_station_ft, dig.us_ref_odometer),
        (dig.station_ft, dig.odometer),
        (dig.ds_ref_station_ft, dig.ds_ref_odometer),
    ]
    pairs = [(s, o) for s, o in pairs if s is not None and o is not None]
    if len(pairs) < 2:
        return True
    d_station = pairs[-1][0] - pairs[0][0]
    d_odometer = pairs[-1][1] - pairs[0][1]
    if d_station == 0 or d_odometer == 0:
        return True
    return (d_station > 0) == (d_odometer > 0)


def check_dig(dig: Dig) -> list:
    """Things worth a second look. Nothing here blocks a report being written."""
    problems = []
    if not dig.number:
        problems.append("No dig number found in Q15.")
    if dig.odometer is None or dig.station_ft is None:
        problems.append("Odometer or station is missing.")
    if not dig.upstream_reference or not dig.downstream_reference:
        problems.append("A U/S or D/S reference could not be read.")
    if dig.latitude is None or dig.longitude is None:
        problems.append("No latitude/longitude - no aerial image or directions.")
    if not dig.county_state:
        problems.append("County/State is blank on the dig sheet.")
    if not dig.legal_description or not dig.tract_number:
        problems.append("Legal description or tract number is blank on the dig sheet.")
    if not dig.alignment_sheet:
        problems.append("Alignment sheet is blank on the dig sheet.")
    if not dig.hca:
        problems.append("HCA is blank on the dig sheet.")

    if (dig.us_weld_distance is not None and dig.ds_weld_distance is not None
            and dig.joint_length is not None):
        gap = abs(dig.us_weld_distance + dig.ds_weld_distance - dig.joint_length)
        if gap > 0.06:
            problems.append(
                f"Weld distances ({dig.us_weld_distance} + {dig.ds_weld_distance}) "
                f"do not add up to the joint length ({dig.joint_length})."
            )
    for label, distance in (("U/S", dig.us_ref_distance), ("D/S", dig.ds_ref_distance)):
        if distance is not None and distance < 0:
            problems.append(f"{label} reference distance is negative ({distance:.2f}).")
    return problems


SITE_REF_RE = re.compile(r"^Site Ref \(", re.I)
WELD_MATCH_FEET = 1.5


def read_site_ref_welds(workbook) -> list:
    """Every WELD row on the Site Ref tabs as (odometer, gw_id, station)."""
    welds = []
    for name in workbook.sheetnames:
        if not SITE_REF_RE.match(name) or "Template" in name:
            continue
        for row in workbook[name].iter_rows(min_row=2, max_col=29, values_only=True):
            if len(row) < 24 or str(row[4] or "").strip().upper() != "WELD":
                continue
            odometer, station = _number(row[6]), _station(row[7])
            gw = row[23]
            if odometer is None or gw in (None, ""):
                continue
            lat = _number(row[27]) if len(row) > 28 else None
            lon = _number(row[28]) if len(row) > 28 else None
            welds.append((odometer, _text(gw), station, lat, lon))
    return welds


def _neighbour_weld_id(dig: Dig, known_id: str, step: int) -> str:
    """The weld number next to ``known_id`` in the dig sheet's own weld table (S34:S45)."""
    numbers = []
    for row in range(34, 46):
        value = _number(dig.cells.get(f"S{row}"))
        if value is not None:
            numbers.append(_text(int(value)))
    if known_id in numbers:
        index = numbers.index(known_id) + step
        if 0 <= index < len(numbers):
            return numbers[index]
    return ""


def assign_welds(dig: Dig, welds: list) -> None:
    """Fill the dig's U/S and D/S weld IDs and stationing from the Site Ref welds.

    A weld the Site Ref tabs do not list is worked out from the other one: its
    stationing is the known weld's plus (or minus) the joint length, and its
    number is the neighbour in the dig sheet's weld table.
    """
    if dig.odometer is None:
        return
    missing = []
    for side, distance, sign in (("us", dig.us_weld_distance, -1),
                                 ("ds", dig.ds_weld_distance, 1)):
        if distance is None:
            missing.append(side)
            continue
        target = dig.odometer + sign * distance
        near = [w for w in welds if abs(w[0] - target) <= WELD_MATCH_FEET]
        if not near:
            missing.append(side)
            continue
        odometer, gw, station, lat, lon = min(near, key=lambda w: abs(w[0] - target))
        setattr(dig, f"{side}_weld_id", gw)
        setattr(dig, f"{side}_weld_station_ft", station)
        setattr(dig, f"{side}_weld_lat", lat)
        setattr(dig, f"{side}_weld_lon", lon)

    for side in missing:
        other = "us" if side == "ds" else "ds"
        other_station = getattr(dig, f"{other}_weld_station_ft")
        other_id = getattr(dig, f"{other}_weld_id")
        if other_station is None or dig.joint_length is None:
            dig.warnings.append(f"No {side.upper()} weld found on the Site Ref tabs.")
            continue
        # Downstream is up the odometer; stationing follows it or runs against it.
        along = 1 if side == "ds" else -1
        direction = 1 if dig.station_ascends else -1
        station = int(round(other_station + along * direction * dig.joint_length))
        setattr(dig, f"{side}_weld_station_ft", station)
        setattr(dig, f"{side}_weld_id", _neighbour_weld_id(dig, other_id, along))
        dig.warnings.append(
            f"{side.upper()} weld is not on the Site Ref tabs; worked out from the "
            f"{other.upper()} weld and the {dig.joint_length} ft joint length "
            f"(station {station_label(station)}).")


def parse_dig_package(data: bytes, filename: str) -> list:
    """Every Dig sheet in a package, as Dig records in odometer order."""
    import openpyxl

    workbook = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        package_id = _package_id(workbook, filename)
        digs = []
        for sheet_name in workbook.sheetnames:
            if not DIG_SHEET_RE.match(sheet_name):
                continue
            # The hidden "Template- Dig Site" tab does not match the pattern, so
            # only real dig sheets reach this point.
            cells = _read_dig_sheet(workbook[sheet_name])
            if not cells.get("Q15") and not cells.get("G8"):
                continue
            digs.append(dig_from_cells(cells, package_id, filename, sheet_name))
        welds = read_site_ref_welds(workbook)
        for dig in digs:
            assign_welds(dig, welds)
    finally:
        workbook.close()

    digs.sort(key=lambda d: (d.odometer is None, d.odometer or 0.0))
    return digs



# ==========================================================================
# KMZ
# ==========================================================================
# Read the pipeline centerline and the dig placemarks out of a KMZ/KML export.
#
# The Enterprise KMZ carries two kinds of placemark:
#
#   * the pipeline itself - one LineString (a Point rides along in its
#     MultiGeometry, which is not a dig), named with a chunk of HTML;
#   * one Point per dig, with the dig's whole table in ExtendedData. The two dig
#     packages share dig numbers (02A, 03A ... exist in both), so a placemark is
#     matched to a dig by assessment ID *and* Feature ID, never by name alone.
#     The assessment ID is only in the schema name ("S_AID_115_CSV_...").
#
# The dig packages already carry each dig's latitude and longitude, and on the
# example files they agree with the KMZ to the foot, so the KMZ's job is the
# centerline for the aerial plus a cross-check on the coordinates.

@dataclass
class DigPlacemark:
    name: str
    package_id: str
    feature_id: str
    longitude: float
    latitude: float


@dataclass
class PipelineData:
    """Centerlines and dig placemarks, all in (longitude, latitude) degrees."""

    lines: list = field(default_factory=list)       # list[list[tuple[float, float]]]
    line_names: list = field(default_factory=list)  # parallel to lines
    line_colours: list = field(default_factory=list)  # parallel to lines: (r, g, b) or None
    line_ids: list = field(default_factory=list)      # parallel to lines: PipelineNumber text
    digs: list = field(default_factory=list)        # list[DigPlacemark]
    source_name: str = ""

    @property
    def is_empty(self) -> bool:
        return not self.lines and not self.digs

    @property
    def pipeline_label(self) -> str:
        for name in self.line_names:
            if name:
                return name
        name = (self.source_name or "Pipeline").rsplit("/", 1)[-1]
        return name.rsplit(".", 1)[0] or "Pipeline"


def _strip_namespace(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _parse_coordinates(text: str):
    points = []
    for chunk in re.split(r"\s+", (text or "").strip()):
        if not chunk:
            continue
        parts = chunk.split(",")
        if len(parts) < 2:
            continue
        try:
            longitude = float(parts[0])
            latitude = float(parts[1])
        except ValueError:
            continue
        points.append((longitude, latitude))
    return points


def _plain_text(markup: str) -> str:
    """'<html>...<p>East Leg Mainline</p>...</html>' -> 'East Leg Mainline'."""
    text = re.sub(r"<[^>]+>", " ", markup or "")
    return re.sub(r"\s+", " ", text).strip()


def _kml_colour(text: str):
    """KML writes colour as aabbggrr hex. Returns (r, g, b), or None."""
    text = (text or "").strip().lstrip("#")
    if len(text) != 8:
        return None
    try:
        blue, green, red = int(text[2:4], 16), int(text[4:6], 16), int(text[6:8], 16)
    except ValueError:
        return None
    return (red, green, blue)


def _line_style_colours(root) -> dict:
    """Map '#styleId' -> the LineStyle colour it draws, following StyleMaps.

    A line placemark points at a StyleMap, whose 'normal' pair points at a Style.
    Google Earth draws the normal pair, so that is the colour the user sees.
    """
    direct, maps = {}, {}
    for element in root.iter():
        tag = _strip_namespace(element.tag)
        if tag == "Style" and element.get("id"):
            for child in element.iter():
                if _strip_namespace(child.tag) == "LineStyle":
                    for item in child:
                        if _strip_namespace(item.tag) == "color":
                            colour = _kml_colour(item.text)
                            if colour:
                                direct["#" + element.get("id")] = colour
        elif tag == "StyleMap" and element.get("id"):
            normal = None
            for pair in element:
                key = url = ""
                for item in pair:
                    item_tag = _strip_namespace(item.tag)
                    if item_tag == "key":
                        key = (item.text or "").strip()
                    elif item_tag == "styleUrl":
                        url = (item.text or "").strip()
                if key == "normal":
                    normal = url
            if normal:
                maps["#" + element.get("id")] = normal
    resolved = dict(direct)
    for name, target in maps.items():
        if target in direct:
            resolved[name] = direct[target]
    return resolved


def _description_field(description: str, field_name: str) -> str:
    match = re.search(rf"{field_name}\s*</t[dh]>\s*<td[^>]*>(.*?)</td>", description or "",
                      re.S | re.I)
    return re.sub(r"<[^>]+>|\s+", " ", match.group(1)).strip() if match else ""


def _pipeline_number(description: str, extended: dict) -> str:
    return (extended.get("PipelineNumber") or
            _description_field(description, "PipelineNumber")).strip()


def _pipeline_name(description: str) -> str:
    return _description_field(description, "PipelineName")


def parse_kml_bytes(data: bytes, source_name: str = "") -> PipelineData:
    result = PipelineData(source_name=source_name)
    root = ElementTree.fromstring(data)
    style_colours = _line_style_colours(root)

    for placemark in root.iter():
        if _strip_namespace(placemark.tag) != "Placemark":
            continue

        name = ""
        schema_url = ""
        extended = {}
        description = ""
        line_points = []
        point = None
        colour = None

        for element in placemark.iter():
            tag = _strip_namespace(element.tag)
            if tag == "name" and not name and element in list(placemark):
                name = _plain_text(element.text)
            elif tag == "styleUrl" and colour is None and element in list(placemark):
                colour = style_colours.get((element.text or "").strip())
            elif tag == "Style" and colour is None and element in list(placemark):
                # A style written inline on the placemark itself.
                for line_style in element:
                    if _strip_namespace(line_style.tag) == "LineStyle":
                        for item in line_style:
                            if _strip_namespace(item.tag) == "color":
                                colour = _kml_colour(item.text)
            elif tag == "description" and not description and element in list(placemark):
                description = element.text or ""
            elif tag == "SchemaData":
                schema_url = element.get("schemaUrl", "")
                for item in element:
                    if _strip_namespace(item.tag) == "SimpleData":
                        extended[item.get("name", "")] = (item.text or "").strip()
            elif tag in ("LineString", "LinearRing"):
                for child in element:
                    if _strip_namespace(child.tag) == "coordinates":
                        points = _parse_coordinates(child.text)
                        if len(points) >= 2:
                            line_points.append(points)
            elif tag == "Point" and point is None:
                for child in element:
                    if _strip_namespace(child.tag) == "coordinates":
                        points = _parse_coordinates(child.text)
                        if points:
                            point = points[0]

        for points in line_points:
            result.lines.append(points)
            result.line_names.append(name)
            result.line_colours.append(colour)
            result.line_ids.append(_pipeline_number(description, extended))
            if description and _pipeline_name(description) and not name.strip():
                result.line_names[-1] = _pipeline_name(description)

        # A dig placemark has a table of dig data; the Point inside the
        # pipeline's own MultiGeometry does not.
        if point is not None and not line_points and extended.get("ID_No"):
            aid = re.search(r"AID[_\s-]*(\d+)", schema_url, re.I)
            result.digs.append(DigPlacemark(
                name=name or extended.get("Dig__", ""),
                package_id=aid.group(1) if aid else "",
                feature_id=extended.get("ID_No", ""),
                longitude=point[0],
                latitude=point[1],
            ))

    return result


def parse_kmz(data: bytes, filename: str = "") -> PipelineData:
    """Accepts a .kmz (zip) or a bare .kml."""
    if filename.lower().endswith(".kml") or data[:5] == b"<?xml":
        return parse_kml_bytes(data, filename)

    combined = PipelineData(source_name=filename)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = [n for n in archive.namelist() if n.lower().endswith(".kml")]
        names.sort(key=lambda n: (n.lower() != "doc.kml", n))
        for name in names:
            try:
                part = parse_kml_bytes(archive.read(name), name)
            except ElementTree.ParseError:
                continue
            combined.lines.extend(part.lines)
            combined.line_names.extend(part.line_names)
            combined.line_colours.extend(part.line_colours)
            combined.line_ids.extend(part.line_ids)
            combined.digs.extend(part.digs)
    return combined


def find_dig_placemark(data: Optional[PipelineData], dig: Dig) -> Optional[DigPlacemark]:
    """The KMZ placemark for this dig: same assessment ID and Feature ID."""
    if not data or not dig.feature_id:
        return None
    for placemark in data.digs:
        if placemark.feature_id == dig.feature_id and placemark.package_id == dig.package_id:
            return placemark
    # A KMZ whose schema name carries no assessment ID can still match when the
    # Feature ID is unique across the file.
    candidates = [p for p in data.digs if p.feature_id == dig.feature_id]
    if len(candidates) == 1 and not candidates[0].package_id:
        return candidates[0]
    return None


def _miles_apart(lat_a, lon_a, lat_b, lon_b) -> float:
    dx = (lon_a - lon_b) * math.cos(math.radians((lat_a + lat_b) / 2)) * 69.172
    dy = (lat_a - lat_b) * 69.0
    return (dx * dx + dy * dy) ** 0.5


def reconcile_with_kmz(digs: Iterable[Dig], data: Optional[PipelineData],
                       tolerance_feet: float = 25.0) -> None:
    """Cross-check each dig's coordinates against its KMZ placemark.

    The dig sheet's coordinates are used as-is. When the KMZ disagrees by more
    than ``tolerance_feet`` - or has no placemark for the dig - that is flagged,
    not silently resolved: the two files are meant to be the same dig list.
    """
    if not data or not data.digs:
        return
    for dig in digs:
        placemark = find_dig_placemark(data, dig)
        if placemark is None:
            dig.warnings.append("Not found in the KMZ (matched on assessment ID and Feature ID).")
            continue
        if dig.latitude is None or dig.longitude is None:
            dig.latitude, dig.longitude = placemark.latitude, placemark.longitude
            dig.warnings.append("Coordinates taken from the KMZ; the dig sheet had none.")
            continue
        feet = _miles_apart(dig.latitude, dig.longitude,
                            placemark.latitude, placemark.longitude) * 5280
        if feet > tolerance_feet:
            dig.warnings.append(
                f"KMZ placemark is {feet:,.0f} ft from the dig sheet's coordinates."
            )


# ==========================================================================
# AERIAL
# ==========================================================================
# Render the aerial image that goes into the staking report.
#
# Satellite tiles are composited, then the KMZ centerline, the dig's U/S and D/S
# references (their coordinates are on the dig sheet) and a labelled dig pin are
# drawn on top. The output is produced at the exact aspect ratio of the
# template's image slot so it drops in without distortion.

# ==========================================================================
# AERIAL
# ==========================================================================
# Render the aerial image that goes into the staking report.
#
# Satellite tiles are composited, then the KMZ centerline, the nearest AGM and a
# labelled dig pin are drawn on top - the same information the Google Earth Pro
# screenshots carry today. The output is produced at the exact aspect ratio of
# the template's image slot so it drops in without distortion.

# The StakingReport image slot in the Enterprise template: 6115050 x 3190043 EMU.
# The real aspect is read from whatever template is uploaded; this is the fallback.
SLOT_ASPECT = 6115050 / 3190043

BASEMAPS = {
    # Mapbox uses the same token as the directions generator, and serves 512px
    # retina tiles, so it is the default where a token is configured.
    "Mapbox Satellite": {
        "url": "https://api.mapbox.com/v4/mapbox.satellite/{z}/{x}/{y}@2x.jpg90",
        "tile_size": 512,
        "needs_token": True,
        "max_zoom": 19,
    },
    "Esri World Imagery": {
        "url": (
            "https://server.arcgisonline.com/ArcGIS/rest/services/"
            "World_Imagery/MapServer/tile/{z}/{y}/{x}"
        ),
        "tile_size": 256,
        "needs_token": False,
        "max_zoom": 19,
    },
    "USDA NAIP (USA only)": {
        "url": (
            "https://gis.apfo.usda.gov/arcgis/rest/services/NAIP/"
            "USDA_CONUS_PRIME/ImageServer/tile/{z}/{y}/{x}"
        ),
        "tile_size": 256,
        "needs_token": False,
        "max_zoom": 18,
    },
}


def basemap_options(has_token: bool) -> list:
    return [
        name for name, spec in BASEMAPS.items()
        if has_token or not spec["needs_token"]
    ]

USER_AGENT = "Enterprise-Dig-File-Generator/1.0 (staking report automation)"

# Web Mercator maths is done in the tile set's own pixel units, so a retina
# (@2x, 512px) tile set simply doubles the resolution at a given zoom.
DEFAULT_TILE_SIZE = 256

PIPE_RED = (226, 32, 32)
KML_DEFAULT_LINE = (255, 255, 255)
LABEL_WHITE = (255, 255, 255)
SHADOW = (0, 0, 0)


@dataclass
class MapView:
    zoom: int
    centre_lat: float
    centre_lon: float
    width: int
    height: int
    tile_size: int = DEFAULT_TILE_SIZE

    def project(self, longitude: float, latitude: float) -> tuple[float, float]:
        """Longitude/latitude -> pixel coordinates in the rendered image."""
        cx, cy = _lonlat_to_world(self.centre_lon, self.centre_lat, self.zoom, self.tile_size)
        px, py = _lonlat_to_world(longitude, latitude, self.zoom, self.tile_size)
        return px - cx + self.width / 2.0, py - cy + self.height / 2.0


def _lonlat_to_world(longitude, latitude, zoom, tile_size=DEFAULT_TILE_SIZE):
    scale = tile_size * (2 ** zoom)
    x = (longitude + 180.0) / 360.0 * scale
    sin_lat = math.sin(math.radians(max(min(latitude, 85.05112878), -85.05112878)))
    y = (0.5 - math.log((1 + sin_lat) / (1 - sin_lat)) / (4 * math.pi)) * scale
    return x, y


def metres_per_pixel(latitude, zoom, tile_size=DEFAULT_TILE_SIZE) -> float:
    """Ground metres covered by one rendered pixel."""
    base = 156543.03392 * math.cos(math.radians(latitude)) / (2 ** zoom)
    return base * (DEFAULT_TILE_SIZE / tile_size)


def _fetch_tile(session, template: str, zoom: int, x: int, y: int, token: str = ""):
    url = template.format(z=zoom, x=x, y=y)
    if token:
        url += ("&" if "?" in url else "?") + f"access_token={token}"
    try:
        response = session.get(url, timeout=20, headers={"User-Agent": USER_AGENT})
        if response.status_code != 200 or not response.content:
            return None
        return Image.open(io.BytesIO(response.content)).convert("RGB")
    except Exception:
        return None


def _basemap_image(view: MapView, template: str, token: str = "") -> tuple:
    tile = view.tile_size
    cx, cy = _lonlat_to_world(view.centre_lon, view.centre_lat, view.zoom, tile)
    left = cx - view.width / 2.0
    top = cy - view.height / 2.0

    first_x = int(math.floor(left / tile))
    first_y = int(math.floor(top / tile))
    last_x = int(math.floor((left + view.width) / tile))
    last_y = int(math.floor((top + view.height) / tile))

    canvas = Image.new("RGB", (view.width, view.height), (32, 38, 32))
    jobs = [
        (x, y)
        for x in range(first_x, last_x + 1)
        for y in range(first_y, last_y + 1)
    ]
    if not jobs:
        return canvas, 0

    fetched = 0
    max_index = 2 ** view.zoom
    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=8) as pool:
            futures = {
                pool.submit(
                    _fetch_tile, session, template, view.zoom, x % max_index, y, token
                ): (x, y)
                for x, y in jobs
                if 0 <= y < max_index
            }
            for future, (x, y) in futures.items():
                image = future.result()
                if image is None:
                    continue
                if image.size != (tile, tile):
                    image = image.resize((tile, tile), Image.LANCZOS)
                canvas.paste(
                    image,
                    (int(round(x * tile - left)), int(round(y * tile - top))),
                )
                fetched += 1
    return canvas, fetched


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------

_FONT_CACHE: dict = {}
FONT_IS_SCALABLE = True

_FONT_PATHS = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/arial.ttf",
)

_FONT_NAMES = ("DejaVuSans-Bold.ttf", "DejaVuSans.ttf", "LiberationSans-Bold.ttf")


def _font(size: int):
    """A font that actually honours the requested size.

    ``ImageFont.load_default()`` returns a FIXED-SIZE bitmap font - it ignores
    the size argument completely. Falling back to it means every label renders
    at about 11 px however large it was asked to be, and once Excel shrinks the
    image into its slot that is illegible. Streamlit Cloud has no DejaVu
    installed, so this fallback was being hit in production even though it
    never was in development, which is why scaling the sizes changed nothing.

    Pillow 10.1+ can scale its own bundled font through
    ``load_default(size=...)``, needing no system fonts and no extra files.
    """
    global FONT_IS_SCALABLE

    size = max(6, int(round(size)))
    if size in _FONT_CACHE:
        return _FONT_CACHE[size]

    font = None
    for path in _FONT_PATHS:
        try:
            font = ImageFont.truetype(path, size)
            break
        except Exception:  # noqa: BLE001
            continue

    if font is None:
        for name in _FONT_NAMES:
            try:
                font = ImageFont.truetype(name, size)
                break
            except Exception:  # noqa: BLE001
                continue

    if font is None:
        try:
            font = ImageFont.load_default(size=size)
        except TypeError:
            FONT_IS_SCALABLE = False
            font = ImageFont.load_default()

    _FONT_CACHE[size] = font
    return font


def _outlined_text(draw, xy, text, font, fill=LABEL_WHITE, outline=SHADOW, weight=2):
    x, y = xy
    for dx in range(-weight, weight + 1):
        for dy in range(-weight, weight + 1):
            if dx or dy:
                draw.text((x + dx, y + dy), text, font=font, fill=outline)
    draw.text((x, y), text, font=font, fill=fill)


def _centred_text(draw, centre, text, font, **kwargs):
    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
    _outlined_text(
        draw,
        (centre[0] - (right - left) / 2.0, centre[1] - (bottom - top) / 2.0),
        text,
        font,
        **kwargs,
    )


def _draw_pin(draw, x, y, scale=1.0, colour=(255, 214, 0)):
    """A small Google Earth style pushpin."""
    height = 30 * scale
    radius = 9 * scale
    draw.line([(x, y), (x, y - height + radius)], fill=(60, 60, 60),
              width=max(2, int(3 * scale)))
    draw.ellipse(
        [x - radius, y - height, x + radius, y - height + 2 * radius],
        fill=colour,
        outline=(60, 60, 60),
        width=max(1, int(2 * scale)),
    )


def _draw_flag(draw, x, y, scale=1.0, colour=PIPE_RED):
    """A small flag marker, used for AGMs."""
    height = 32 * scale
    draw.line([(x, y), (x, y - height)], fill=(245, 245, 245),
              width=max(2, int(4 * scale)))
    draw.polygon(
        [(x + scale, y - height),
         (x + 21 * scale, y - height + 8 * scale),
         (x + scale, y - height + 16 * scale)],
        fill=colour,
    )


def _scale_bar(draw, mpp, width, height, scale=1.0):
    """``mpp`` is the ground metres covered by one pixel of the *output* image."""
    feet_per_pixel = mpp * 3.28084
    for target in (200, 400, 500, 800, 1000, 1500, 2000, 3000, 5000):
        pixels = target / feet_per_pixel
        if 90 <= pixels <= width * 0.22:
            break
    else:
        target, pixels = 400, 400 / feet_per_pixel

    right = width - int(70 * scale)
    left = right - pixels
    baseline = height - int(20 * scale)
    thickness = max(2, int(4 * scale))
    tick = 8 * scale
    draw.line([(left, baseline), (right, baseline)], fill=LABEL_WHITE, width=thickness)
    draw.line([(left, baseline - tick), (left, baseline + tick * 0.7)],
              fill=LABEL_WHITE, width=thickness)
    draw.line([(right, baseline - tick), (right, baseline + tick * 0.7)],
              fill=LABEL_WHITE, width=thickness)
    font = _font(int(22 * scale))
    _outlined_text(draw, (left, baseline - 34 * scale), f"{target:,} ft", font,
                   weight=max(1, int(1.5 * scale)))


def _north_arrow(draw, width, height, scale=1.0):
    font = _font(int(26 * scale))
    x = width - 32 * scale
    y = height - 76 * scale
    draw.polygon(
        [(x, y), (x - 10 * scale, y + 28 * scale),
         (x, y + 20 * scale), (x + 10 * scale, y + 28 * scale)],
        fill=LABEL_WHITE,
        outline=SHADOW,
    )
    _centred_text(draw, (x, y + 46 * scale), "N", font,
                  weight=max(1, int(1.5 * scale)))


def _title_card(draw, text, scale=1.0):
    if not text:
        return
    font = _font(int(20 * scale))
    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
    pad_x = 11 * scale
    pad_y = 8 * scale
    x0 = 12 * scale
    y0 = 12 * scale
    box_width = (right - left) + pad_x * 2
    box_height = (bottom - top) + pad_y * 2
    draw.rectangle([x0, y0, x0 + box_width, y0 + box_height],
                   fill=(255, 255, 255, 235), outline=(120, 120, 120),
                   width=max(1, int(scale)))
    draw.text((x0 + pad_x - left, y0 + pad_y - top), text, font=font,
              fill=(20, 20, 20))



# ---------------------------------------------------------------------------
# Legend
# ---------------------------------------------------------------------------

PIN_YELLOW = (255, 214, 0)
NEIGHBOUR_WHITE = (245, 245, 245)


def _layer_count(layer, box) -> int:
    """How many pixels of a transparent layer are drawn inside ``box``."""
    if layer is None:
        return 0
    left, top, right, bottom = (int(round(v)) for v in box)
    left, top = max(left, 0), max(top, 0)
    right, bottom = min(right, layer.width), min(bottom, layer.height)
    if right <= left or bottom <= top:
        return 0
    alpha = layer.crop((left, top, right, bottom)).getchannel("A")
    return alpha.point(lambda value: 255 if value else 0).histogram()[255]


def _legend(layer, entries, width, height, scale=1.0, avoid=None, soft=None):
    """Draw the legend on its own layer, in the corner that hides the least.

    ``entries`` is a list of (label, swatch) or, for a line, (label, 'line', colour);
    swatch is 'line', 'flag' or 'pin'. ``avoid`` is the layer holding the dig pin,
    flags and labels (covering those is what must not happen) and ``soft`` the one
    holding the pipeline lines (covering a stretch of line is a lesser evil).
    """
    if not entries:
        return
    draw = ImageDraw.Draw(layer, "RGBA")

    # Start at a readable size and shrink until the box is a modest part of the
    # picture - a long list of lines must not take over the aerial.
    size = 11.5
    while True:
        unit = size * scale / 17.0
        font = _font(max(8, int(17 * unit)))
        line_height = 20 * unit
        box_width = 120 * unit
        for entry in entries:
            left, top, right, bottom = draw.textbbox((0, 0), entry[0], font=font)
            box_width = max(box_width, (right - left) + 52 * unit)
        box_height = 30 * unit + line_height * len(entries)
        if (box_width <= width * 0.36 and box_height <= height * 0.42) or size <= 7:
            break
        size -= 0.5

    margin = 12 * scale
    candidates = [
        (width - box_width - margin, margin),                          # top right
        (margin, height - box_height - margin),                        # bottom left
        (margin, 62 * scale + margin),                                 # top left, under the title card
        (width - box_width - margin, height - box_height - 100 * scale),  # bottom right, above the scale bar
    ]
    best, best_score = candidates[0], None
    for order, (x, y) in enumerate(candidates):
        x = min(max(x, 0), max(width - box_width, 0))
        y = min(max(y, 0), max(height - box_height, 0))
        box = (x, y, x + box_width, y + box_height)
        score = _layer_count(avoid, box) * 10 + _layer_count(soft, box) + order
        if best_score is None or score < best_score:
            best, best_score = (x, y), score
    x0, y0 = best

    draw.rectangle(
        [x0, y0, x0 + box_width, y0 + box_height],
        fill=(255, 255, 255, 230),
        outline=(120, 120, 120),
        width=max(1, int(scale)),
    )
    draw.text((x0 + 9 * unit, y0 + 5 * unit), "Legend", font=font, fill=(20, 20, 20))
    for index, entry in enumerate(entries):
        label, swatch = entry[0], entry[1]
        y = y0 + 28 * unit + index * line_height
        centre_y = y + line_height * 0.42
        draw.text((x0 + 38 * unit, y), label, font=font, fill=(20, 20, 20))
        if swatch == "line":
            draw.line([(x0 + 9 * unit, centre_y), (x0 + 29 * unit, centre_y)],
                      fill=entry[2] if len(entry) > 2 else PIPE_RED,
                      width=max(2, int(4 * unit)))
        else:
            colour = PIN_YELLOW if swatch == "pin" else PIPE_RED
            radius = 5.5 * unit
            draw.ellipse([x0 + 19 * unit - radius, centre_y - radius,
                          x0 + 19 * unit + radius, centre_y + radius],
                         fill=colour, outline=(60, 60, 60))


# ---------------------------------------------------------------------------
# Geometry
# ---------------------------------------------------------------------------

def aerial_geometry(latitude: float, span_feet: float, width: int,
                    tile_size: int, max_zoom: int) -> tuple:
    """Pick the tile zoom and resampling factor that make the image exactly
    ``span_feet`` wide on the ground.

    Returns ``(tile_zoom, factor, metres_per_output_pixel)``. Tiles come in
    whole zoom levels, so the canvas is composited at the first level at least
    as detailed as asked for and then resampled by ``1/factor``. (Picking the
    nearest *coarser* level instead - the obvious loop - lands anywhere between
    one and two times wider than requested, so the slider would not mean what
    it says.)
    """
    target_mpp = (span_feet / 3.28084) / width
    mpp_zoom0 = 156543.03392 * math.cos(math.radians(latitude))
    exact = math.log2(mpp_zoom0 / target_mpp)        # zoom needed in a 256 px world
    bonus = math.log2(tile_size / 256.0)             # a 512 px tile is one level denser
    zoom = max(0, min(int(math.ceil(exact - bonus - 1e-9)), max_zoom))
    factor = 2.0 ** (zoom + bonus - exact)           # canvas pixels per output pixel
    return zoom, factor, target_mpp


def _label_line(draw, runs, text, dig_x, dig_y, width, height, scale):
    """Put the pipeline's name beside its line, clear of the dig's pin and label."""
    font = _font(int(22 * scale))
    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
    text_w, text_h = right - left, bottom - top
    best, best_score = None, None
    for points in runs:
        for (x1, y1), (x2, y2) in zip(points, points[1:]):
            for t in (0.25, 0.5, 0.75):
                x, y = x1 + (x2 - x1) * t, y1 + (y2 - y1) * t
                gap = math.hypot(x - dig_x, y - dig_y)
                if gap < 150 * scale:
                    continue
                box_x = x - text_w / 2
                box_y = y - 30 * scale - text_h
                if (box_x < 6 * scale or box_x + text_w > width - 6 * scale
                        or box_y < 70 * scale or box_y + text_h > height - 70 * scale):
                    continue
                score = abs(gap - 260 * scale)
                if best_score is None or score < best_score:
                    best, best_score = (x, y, box_x, box_y), score
    if best is None:
        return
    x, y, box_x, box_y = best
    draw.line([(x, y - 4 * scale), (x, box_y + text_h + 4 * scale)],
              fill=(255, 255, 255, 255), width=max(2, int(2 * scale)))
    _outlined_text(draw, (box_x, box_y), text, font, weight=max(1, int(2 * scale)))


def _line_colour(pipeline, index: int) -> tuple:
    colours = getattr(pipeline, "line_colours", None) or []
    colour = colours[index] if index < len(colours) else None
    # A LineStyle with no <color> is white in KML (and in Google Earth).
    return tuple(colour) if colour else KML_DEFAULT_LINE


def _norm_name(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip()


def _distance_to_line_feet(line, lat, lon) -> float:
    best = float("inf")
    cos = math.cos(math.radians(lat))
    for x, y in line:
        d = math.hypot((y - lat) * 364000.0, (x - lon) * 364000.0 * cos)
        best = min(best, d)
    return best


def dig_line_index(pipeline, dig) -> Optional[int]:
    """Which KMZ line is the dig's own pipeline.

    Matched by Line ID (PipelineNumber), then by the line name on the dig sheet,
    and last by being the nearest line to the dig (within 300 ft).
    """
    if pipeline is None or not pipeline.lines:
        return None
    info = asset_info(dig) if hasattr(dig, "header") else AssetInfo()
    ids = getattr(pipeline, "line_ids", None) or []
    wanted = _norm_name(info.line_id)
    if wanted:
        for index, value in enumerate(ids):
            if _norm_name(value) == wanted:
                return index
    names = [_norm_name(n) for n in pipeline.line_names]
    for target in (_norm_name(info.line_name), _norm_name(info.segment_name)):
        if target:
            for index, name in enumerate(names):
                if name and (name == target or name in target or target in name):
                    return index
    if dig.latitude is None or dig.longitude is None:
        return None
    nearest = min(range(len(pipeline.lines)),
                  key=lambda i: _distance_to_line_feet(pipeline.lines[i], dig.latitude, dig.longitude))
    if _distance_to_line_feet(pipeline.lines[nearest], dig.latitude, dig.longitude) <= 300:
        return nearest
    return None


def _line_runs(line, centre_lat, centre_lon, span_feet):
    """The stretches of a long centerline that matter to one image.

    A pipeline KMZ line has tens of thousands of vertices spanning hundreds of
    miles. Projecting them all at street-level zoom gives coordinates in the
    millions of pixels, which is slow to draw and pointless. Keep only vertices
    inside a generous window, plus the vertex on each side so a segment
    crossing the edge of the picture is still drawn to it.
    """
    half = span_feet * 1.5 / 364000.0                      # degrees of latitude
    half_lon = half / max(math.cos(math.radians(centre_lat)), 0.01)
    inside = [
        abs(lat - centre_lat) <= half and abs(lon - centre_lon) <= half_lon
        for lon, lat in line
    ]
    keep = [
        inside[i]
        or (i > 0 and inside[i - 1])
        or (i + 1 < len(line) and inside[i + 1])
        for i in range(len(line))
    ]
    runs, current = [], []
    for point, wanted in zip(line, keep):
        if wanted:
            current.append(point)
        elif current:
            runs.append(current)
            current = []
    if current:
        runs.append(current)
    return [run for run in runs if len(run) >= 2]


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

class AerialError(RuntimeError):
    pass


def render_aerial(
    dig,
    pipeline: Optional[PipelineData],
    width: int = 1600,
    basemap: str = "Esri World Imagery",
    span_feet: float = 2600.0,
    show_legend: bool = True,
    token: str = "",
    slot_aspect: Optional[float] = None,
    show_neighbours: bool = True,
) -> Optional[bytes]:
    """Return JPEG bytes for one dig, or None when there is no coordinate.

    ``slot_aspect`` is the aspect of the image placeholder in the staking
    report template. Rendering to it exactly means the image is never cropped
    to fit, which is what would otherwise clip the title card in the corner.
    """
    if dig.latitude is None or dig.longitude is None:
        return None

    aspect = slot_aspect or SLOT_ASPECT
    height = int(round(width / aspect))

    # Excel shrinks the image into a slot roughly 650 px wide, so everything
    # drawn on it has to be sized for that, not for the rendered pixels.
    scale = width / 640.0
    latitude = float(dig.latitude)
    longitude = float(dig.longitude)

    spec = BASEMAPS.get(basemap) or BASEMAPS["Esri World Imagery"]
    tile_size = spec["tile_size"]
    tile_token = token if spec["needs_token"] else ""
    if spec["needs_token"] and not tile_token:
        raise AerialError(f"{basemap} needs a Mapbox token.")

    zoom, factor, mpp = aerial_geometry(
        latitude, span_feet, width, tile_size, spec.get("max_zoom", 19)
    )
    canvas_width = max(1, int(round(width * factor)))
    canvas_height = max(1, int(round(height * factor)))
    view = MapView(zoom=zoom, centre_lat=latitude, centre_lon=longitude,
                   width=canvas_width, height=canvas_height, tile_size=tile_size)

    image, fetched = _basemap_image(view, spec["url"], tile_token)
    if fetched == 0:
        raise AerialError(
            f"No imagery tiles could be fetched from {basemap}. "
            "Check network access, or the Mapbox token if using Mapbox."
        )
    if image.size != (width, height):
        image = image.resize((width, height), Image.LANCZOS)

    def project(lon, lat):
        x, y = view.project(lon, lat)
        return x / factor, y / factor

    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay, "RGBA")
    legend_entries = []

    # --- pipeline centerline ---------------------------------------------
    # Each line keeps the colour it has in the KMZ (the loop and the mainline are
    # different colours there); a line with no colour of its own falls back to red.
    dig_line = dig_line_index(pipeline, dig)
    dig_label_spot = None
    if pipeline is not None and pipeline.lines:
        colours = [_line_colour(pipeline, i) for i in range(len(pipeline.lines))]   # exactly as the KMZ has them
        runs_by_line = {}
        for index, line in enumerate(pipeline.lines):
            projected = [[project(lon, lat) for lon, lat in run]
                         for run in _line_runs(line, latitude, longitude, span_feet)]
            if projected:
                runs_by_line[index] = projected

        def name_of(index):
            return (pipeline.line_names[index] if index < len(pipeline.line_names)
                    and pipeline.line_names[index] else pipeline.pipeline_label)

        # Other lines first and thinner; the dig's own line last, thick (no outline)
        # so it reads as the pipeline the report is about.
        thin = max(3, int(3 * scale))
        for index, projected in runs_by_line.items():
            if index == dig_line:
                continue
            for points in projected:
                draw.line(points, fill=colours[index] + (215,), width=thin, joint="curve")
        if dig_line in runs_by_line:
            heavy = max(6, int(8 * scale))
            for points in runs_by_line[dig_line]:
                draw.line(points, fill=colours[dig_line], width=heavy, joint="curve")
            dig_label_spot = runs_by_line[dig_line]

        drawn = []
        if dig_line in runs_by_line:
            drawn.append((name_of(dig_line) + " (this dig)", colours[dig_line], True))
        for index in runs_by_line:
            if index == dig_line:
                continue
            entry = (name_of(index), colours[index], False)
            if all(entry[:2] != d[:2] for d in drawn):
                drawn.append(entry)
        for label, colour, bold in drawn:
            legend_entries.append((label, "line", colour, bold))

    # Everything else goes on its own layer so the legend can be placed where
    # it covers the least (pin, flags and labels matter most, lines less).
    lines_layer = overlay
    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay, "RGBA")

    # --- other digs from the same package, when they fall in the picture ---
    dig_x, dig_y = project(longitude, latitude)
    if show_neighbours and pipeline is not None:
        for other in pipeline.digs:
            if other.package_id != dig.package_id or other.feature_id == dig.feature_id:
                continue
            ox, oy = project(other.longitude, other.latitude)
            if not (0 <= ox <= width and 0 <= oy <= height):
                continue
            # A neighbour a few feet from the dig (02A beside 02B) would sit
            # under its pin and label, so only mark ones clear of it.
            if math.hypot(ox - dig_x, oy - dig_y) < 70 * scale:
                continue
            radius = 6 * scale
            draw.ellipse([ox - radius, oy - radius, ox + radius, oy + radius],
                         fill=NEIGHBOUR_WHITE, outline=(60, 60, 60),
                         width=max(1, int(2 * scale)))
            _outlined_text(draw, (ox + 10 * scale, oy - 14 * scale),
                           other.name, _font(int(20 * scale)),
                           weight=max(1, int(1.5 * scale)))

    # --- the dig's own U/S and D/S references, from the dig sheet ----------
    for label, ref_lat, ref_lon in (
        (dig.upstream_reference, dig.us_ref_lat, dig.us_ref_lon),
        (dig.downstream_reference, dig.ds_ref_lat, dig.ds_ref_lon),
    ):
        if not label or ref_lat is None or ref_lon is None:
            continue
        ax, ay = project(ref_lon, ref_lat)
        if not (-50 <= ax <= width + 50 and -50 <= ay <= height + 50):
            continue
        _draw_flag(draw, ax, ay, scale)
        label_font = _font(int(24 * scale))
        left, top, right, bottom = draw.textbbox((0, 0), label, font=label_font)
        label_width = right - left
        # The legend sits in the top-right corner. Put the label on the other
        # side of the flag rather than let it run underneath.
        if ax + 26 * scale + label_width > width * 0.62:
            label_x = ax - 12 * scale - label_width
        else:
            label_x = ax + 26 * scale
        _outlined_text(draw, (label_x, ay - 48 * scale), label, label_font,
                       weight=max(1, int(2 * scale)))
        legend_entries.append((label, "flag"))

    # --- name the dig's own pipeline right on the picture -------------------
    if dig_line is not None and dig_label_spot:
        label_text = (pipeline.line_names[dig_line] if dig_line < len(pipeline.line_names)
                      and pipeline.line_names[dig_line] else pipeline.pipeline_label)
        _label_line(draw, dig_label_spot, label_text, dig_x, dig_y, width, height, scale)

    # --- the dig ----------------------------------------------------------
    _draw_pin(draw, dig_x, dig_y, scale)
    _centred_text(draw, (dig_x, dig_y + 34 * scale), dig.name,
                  _font(int(32 * scale)), weight=max(2, int(2.5 * scale)))
    if dig.name:
        legend_entries.append((dig.name, "pin"))

    _title_card(draw, dig.name, scale)
    _scale_bar(draw, mpp, width, height, scale)
    _north_arrow(draw, width, height, scale)

    legend_layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
    if show_legend:
        _legend(legend_layer, legend_entries, width, height, scale,
                avoid=overlay, soft=lines_layer)

    image = image.convert("RGBA")
    for layer in (lines_layer, overlay, legend_layer):
        image = Image.alpha_composite(image, layer)
    image = image.convert("RGB")

    # JPEG, not PNG: satellite imagery compresses far better as JPEG (roughly a
    # tenth of the size), which matters when a batch of 100+ reports is held in
    # memory and zipped. Chroma subsampling is off so the label text stays crisp.
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=90, subsampling=0, optimize=True)
    return buffer.getvalue()

# ==========================================================================
# DIRECTIONS
# ==========================================================================
# Driving directions, ported from Hayden's Dig Site Directions Generator.
#
# Source: github.com/maydenhiller/Dig-Site-Directions-Generator
#
# The logic is kept identical so the wording matches the reports already issued:
# reverse geocode the dig to its town, seed a route from the town centre to get a
# real road start point, name that start point by the two distinct roads meeting
# there, then phrase each step as "<instruction> and continue traveling
# <cardinal> for <n.nn> miles", closing with which side the dig sits on relative
# to the direction of travel.
#
# Needs a Mapbox token - the same one the existing app uses.

TIMEOUT = 25
CARDINALS = ["North", "Northeast", "East", "Southeast",
             "South", "Southwest", "West", "Northwest"]


class DirectionsError(RuntimeError):
    pass


@dataclass
class DirectionsResult:
    paragraph: str
    town: str = ""
    state: str = ""
    intersection: str = ""


# ---------------------------------------------------------------------------
# Web Mercator geometry - used to decide which side of the road the dig is on
# ---------------------------------------------------------------------------

def _mercator_xy(longitude: float, latitude: float) -> tuple:
    radius = 6378137.0
    x = math.radians(longitude) * radius
    y = math.log(math.tan(math.pi / 4 + math.radians(latitude) / 2)) * radius
    return x, y


def _segment_projection(a_lon, a_lat, b_lon, b_lat, p_lon, p_lat):
    ax, ay = _mercator_xy(a_lon, a_lat)
    bx, by = _mercator_xy(b_lon, b_lat)
    px, py = _mercator_xy(p_lon, p_lat)
    vx, vy = bx - ax, by - ay
    length_squared = vx * vx + vy * vy
    if length_squared == 0:
        return 0.0, ax, ay, ax, ay, bx, by
    t = ((px - ax) * vx + (py - ay) * vy) / length_squared
    t = max(0.0, min(1.0, t))
    return t, ax + t * vx, ay + t * vy, ax, ay, bx, by


def _nearest_segment(route, dig_lat, dig_lon):
    best, best_distance = None, float("inf")
    px, py = _mercator_xy(dig_lon, dig_lat)
    for index in range(len(route) - 1):
        a_lat, a_lon = route[index]
        b_lat, b_lon = route[index + 1]
        _, projx, projy, *_ = _segment_projection(
            a_lon, a_lat, b_lon, b_lat, dig_lon, dig_lat
        )
        dx, dy = px - projx, py - projy
        distance = dx * dx + dy * dy
        if distance < best_distance:
            best_distance = distance
            best = ((a_lon, a_lat), (b_lon, b_lat), (projx, projy))
    return best


def side_relative_to_route(route, dig_lat, dig_lon) -> str:
    nearest = _nearest_segment(route, dig_lat, dig_lon)
    if not nearest:
        return "right"
    (a_lon, a_lat), (b_lon, b_lat), (projx, projy) = nearest
    ax, ay = _mercator_xy(a_lon, a_lat)
    bx, by = _mercator_xy(b_lon, b_lat)
    px, py = _mercator_xy(dig_lon, dig_lat)
    cross = (bx - ax) * (py - projy) - (by - ay) * (px - projx)
    return "left" if cross > 0 else "right"


# ---------------------------------------------------------------------------
# Phrasing
# ---------------------------------------------------------------------------

def _bearing_to_cardinal(bearing: float) -> str:
    return CARDINALS[round((bearing % 360) / 45) % 8]


def _extract_town_state(feature) -> tuple:
    town, state = "", ""
    for context in feature.get("context", []):
        identifier = context.get("id", "")
        if identifier.startswith("place."):
            town = context.get("text", "")
        if identifier.startswith("region."):
            state = context.get("text", "")
    if not town:
        town = feature.get("text", "")
    return town, state


def _normalise_street_base(name: str) -> str:
    if not name:
        return ""
    text = re.sub(r"^(N|S|E|W)\s+", "", name.strip(), flags=re.I)
    text = re.sub(
        r"\b(Street|St\.?|Avenue|Ave\.?|Road|Rd\.?|Boulevard|Blvd\.?|Drive|Dr\.?"
        r"|Lane|Ln\.?|Terrace|Ter\.?|Court|Ct\.?)\b\.?",
        "",
        text,
        flags=re.I,
    ).strip()
    return re.sub(r"\s{2,}", " ", text)


def _intersection_label(longitude, latitude, token) -> str:
    url = (
        "https://api.mapbox.com/v4/mapbox.mapbox-streets-v8/tilequery/"
        f"{longitude},{latitude}.json?layers=road&radius=300&limit=50"
        f"&access_token={token}"
    )
    payload = requests.get(url, timeout=TIMEOUT).json()

    seen, roads = set(), []
    for feature in payload.get("features", []):
        properties = feature.get("properties", {})
        name = properties.get("name")
        if not name or name in seen:
            continue
        seen.add(name)
        roads.append((name, properties.get("class", ""), _normalise_street_base(name)))

    if not roads:
        return "Unknown Intersection"

    bases = [base for _, _, base in roads]
    if any("Washington" in b for b in bases) and any("Jefferson" in b for b in bases):
        washington = next(n for n, _, b in roads if "Washington" in b)
        jefferson = next(n for n, _, b in roads if "Jefferson" in b)
        return f"{washington} & {jefferson}"

    used, chosen = set(), []
    for name, _, base in roads:
        if base not in used:
            chosen.append(name)
            used.add(base)
        if len(chosen) == 2:
            break
    return " & ".join(chosen) if chosen else "Unknown Intersection"


def _format_step(step, miles: float) -> str:
    manoeuvre = step.get("maneuver", {})
    instruction = manoeuvre.get("instruction", "").rstrip(".")
    cardinal = _bearing_to_cardinal(manoeuvre.get("bearing_after", 0))
    return f"{instruction} and continue traveling {cardinal} for {miles:.2f} miles"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def generate_directions(latitude: float, longitude: float, token: str) -> DirectionsResult:
    if not token:
        raise DirectionsError("No Mapbox token configured.")

    import polyline

    town_url = (
        "https://api.mapbox.com/geocoding/v5/mapbox.places/"
        f"{longitude},{latitude}.json?types=place&language=en&access_token={token}"
    )
    town_payload = requests.get(town_url, timeout=TIMEOUT).json()
    features = town_payload.get("features") or []
    if not features:
        raise DirectionsError("Mapbox could not find a town near this coordinate.")
    town_feature = features[0]
    town, state = _extract_town_state(town_feature)
    town_centre = town_feature["center"]

    seed_url = (
        "https://api.mapbox.com/directions/v5/mapbox/driving/"
        f"{town_centre[0]},{town_centre[1]};{longitude},{latitude}"
        "?steps=true&geometries=polyline&overview=full&language=en"
        f"&access_token={token}"
    )
    seed = requests.get(seed_url, timeout=TIMEOUT).json()
    seed_routes = seed.get("routes") or []
    if not seed_routes:
        raise DirectionsError("Mapbox could not route to this coordinate.")
    start = seed_routes[0]["legs"][0]["steps"][0]["maneuver"]["location"]

    intersection = _intersection_label(start[0], start[1], token)

    route_url = (
        "https://api.mapbox.com/directions/v5/mapbox/driving/"
        f"{start[0]},{start[1]};{longitude},{latitude}"
        "?steps=true&geometries=polyline&overview=full&language=en"
        f"&access_token={token}"
    )
    payload = requests.get(route_url, timeout=TIMEOUT).json()
    routes = payload.get("routes") or []
    if not routes:
        raise DirectionsError("Mapbox returned no route from the start intersection.")
    route = routes[0]
    steps = route["legs"][0]["steps"]
    coordinates = polyline.decode(route["geometry"])

    narrative = [
        f"From the intersection of {intersection} in {town}, {state}, travel as follows"
    ]
    for index, step in enumerate(steps):
        if index == len(steps) - 1:
            side = side_relative_to_route(coordinates, latitude, longitude)
            narrative.append(f"The dig site will be located on your {side}.")
        else:
            narrative.append(_format_step(step, step["distance"] / 1609.34))

    paragraph = " ".join(line.strip().rstrip(".") + "." for line in narrative)
    return DirectionsResult(
        paragraph=paragraph, town=town, state=state, intersection=intersection
    )



# ==========================================================================
# CHEATSHEET
# ==========================================================================
# Fill the dig stake cheat sheet - three rows per dig: U/S weld, anomaly, D/S weld.
#
# The template ships in two flavours, ascending and descending station numbers.
# They differ only in the sign of the station formulas, so the right formulas are
# written per dig from the stationing direction detected on that dig's own sheet.
# Either template works as the starting point; when both are uploaded the one that
# matches the package is used, so the untouched cells (colours, notes, widths)
# are the ones the crew expects for that direction.

FIRST_ROW = 6
ROWS_PER_DIG = 3

COL_NAME = 1          # A
COL_ODO = 2           # B
COL_STATION = 3       # C
COL_FROM_ANOMALY = 4  # D
COL_FROM_REF = 5      # E
COL_NOTES = 6         # F
COL_HCA = 7           # G
COL_DEPTH = 8         # H
COL_LAT = 9           # I
COL_LON = 10          # J


def _round2(value):
    return None if value is None else round(float(value), 2)


def cheat_sheet_flavour(template_bytes: bytes) -> Optional[str]:
    """'ascending' or 'descending', read from the template's own U/S weld formula."""
    import openpyxl

    try:
        sheet = openpyxl.load_workbook(io.BytesIO(template_bytes)).worksheets[0]
    except Exception:  # noqa: BLE001
        return None
    formula = str(sheet.cell(FIRST_ROW, COL_STATION).value or "").replace(" ", "")
    # Ascending: U/S weld station = anomaly station - distance (=C7-D6).
    if re.fullmatch(r"=C\d+-D\d+", formula):
        return "ascending"
    if re.fullmatch(r"=C\d+\+D\d+", formula):
        return "descending"
    return None


def pick_cheat_template(templates: dict, digs: Iterable[Dig]) -> Optional[bytes]:
    """``templates`` maps flavour (or None) -> bytes. Prefer the package's own direction."""
    if not templates:
        return None
    digs = list(digs)
    ascending = sum(1 for d in digs if d.station_ascends)
    wanted = "ascending" if ascending * 2 >= len(digs) else "descending"
    if wanted in templates:
        return templates[wanted]
    return next(iter(templates.values()))


def build_cheat_sheet(
    template_bytes: bytes,
    digs: Iterable[Dig],
    auto_notes: bool = True,
    proximity_feet: float = 2000.0,
) -> bytes:
    import openpyxl

    workbook = openpyxl.load_workbook(io.BytesIO(template_bytes))
    sheet = workbook.worksheets[0]

    digs = sorted(digs, key=lambda d: (d.odometer is None, d.odometer or 0.0))
    _ensure_rows(sheet, len(digs))

    previous: Optional[Dig] = None
    for index, dig in enumerate(digs):
        top = FIRST_ROW + index * ROWS_PER_DIG
        anomaly = top + 1
        bottom = top + 2

        sheet.cell(top, COL_NAME).value = f"{dig.name} U/S Weld"
        sheet.cell(anomaly, COL_NAME).value = f"{dig.name} Anomaly"
        sheet.cell(bottom, COL_NAME).value = f"{dig.name} D/S Weld"

        sheet.cell(anomaly, COL_ODO).value = dig.odometer
        sheet.cell(anomaly, COL_STATION).value = dig.station_ft

        # Weld distances sit on the weld rows; the formulas above and below the
        # anomaly derive the weld odometer and station from them.
        sheet.cell(top, COL_FROM_ANOMALY).value = _round2(dig.us_weld_distance)
        sheet.cell(bottom, COL_FROM_ANOMALY).value = _round2(dig.ds_weld_distance)

        # The upstream weld is always at a lower odometer, so the ODO formulas
        # never change. Only the station column flips with the line's direction.
        sheet.cell(top, COL_ODO).value = f"=B{anomaly}-D{top}"
        sheet.cell(bottom, COL_ODO).value = f"=B{anomaly}+D{bottom}"
        up, down = ("-", "+") if dig.station_ascends else ("+", "-")
        sheet.cell(top, COL_STATION).value = f"=C{anomaly}{up}D{top}"
        sheet.cell(bottom, COL_STATION).value = f"=C{anomaly}{down}D{bottom}"

        sheet.cell(anomaly, COL_FROM_REF).value = dig.nearest_reference_label or None

        sheet.cell(anomaly, COL_NOTES).value = None
        if auto_notes and previous is not None and dig.odometer and previous.odometer:
            gap = abs(dig.odometer - previous.odometer)
            if gap <= proximity_feet:
                sheet.cell(anomaly, COL_NOTES).value = f"{gap:.2f}' past {previous.name}"
        if dig.notes:
            sheet.cell(anomaly, COL_NOTES).value = dig.notes

        sheet.cell(anomaly, COL_HCA).value = dig.hca_label or None

        # Latitude and longitude come straight from the dig sheet, on the anomaly
        # row. The weld rows are cleared so template leftovers cannot leak through.
        for row in (top, anomaly, bottom):
            sheet.cell(row, COL_LAT).value = None
            sheet.cell(row, COL_LON).value = None
        sheet.cell(anomaly, COL_LAT).value = dig.latitude
        sheet.cell(anomaly, COL_LON).value = dig.longitude

        previous = dig

    _clear_unused(sheet, len(digs))

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _ensure_rows(sheet, dig_count: int) -> None:
    """Extend the template past its pre-built dig blocks if needed."""
    needed = FIRST_ROW + dig_count * ROWS_PER_DIG - 1
    if needed <= sheet.max_row:
        return
    for row in range(sheet.max_row + 1, needed + 1):
        source_row = FIRST_ROW + ((row - FIRST_ROW) % ROWS_PER_DIG)
        for column in range(1, COL_LON + 1):
            source = sheet.cell(source_row, column)
            target = sheet.cell(row, column)
            if source.has_style:
                target._style = copy(source._style)
        height = sheet.row_dimensions[source_row].height
        if height:
            sheet.row_dimensions[row].height = height


def _clear_unused(sheet, dig_count: int) -> None:
    """Blank out the template's leftover 'Dig #n' placeholder blocks."""
    start = FIRST_ROW + dig_count * ROWS_PER_DIG
    for row in range(start, sheet.max_row + 1):
        for column in range(1, COL_LON + 1):
            sheet.cell(row, column).value = None



# ==========================================================================
# STAKING
# ==========================================================================
# Write one filled staking report per dig, from the uploaded .xlsm template.
#
# The template's StakingReport tab is almost entirely formulas into its "Dig
# Sheet" tab - ='Dig Sheet'!G8 and so on. So the report is filled by:
#
#   1. copying the package's Dig sheet, cell for cell, into the "Dig Sheet" tab
#      (values only: on the package those cells are formulas into its Dashboard
#      and References tabs, which the template does not have);
#   2. writing the few things a dig sheet cannot know - surveyor, directions,
#      staking notes - and swapping in the aerial image;
#   3. leaving every formula alone, but refreshing the value each one last
#      computed so the file reads correctly even where nothing recalculates it.

SHEET = "StakingReport"
DIG_SHEET = "Dig Sheet"
CONTACTS_SHEET = "Contacts"
CACHED_SHEETS = ("StakingReport", "Pre_Dig_Photos")

# Cells on the template's own sheets that the app writes.
CELL_SURVEYOR = "C11"
CELL_DIRECTIONS = "A28"
CELL_STAKING_NOTES = "J28"


@dataclass
class ReportSettings:
    surveyor_name: str = ""
    # None leaves whatever the template already says in the Staking Notes box.
    staking_notes: Optional[str] = None


@dataclass
class TemplateInfo:
    """What the staking report template says about itself."""

    stakers: list = field(default_factory=list)        # Contacts!A, in order
    phones: dict = field(default_factory=dict)         # name -> phone
    staking_notes: str = ""                            # the template's J28 text
    slot_aspect: Optional[float] = None
    problems: list = field(default_factory=list)


def read_template_info(template_bytes: bytes) -> TemplateInfo:
    """Names, phone numbers, default notes and the image slot of a template."""
    import openpyxl

    info = TemplateInfo()
    try:
        patcher = XlsmPatcher(template_bytes)
    except Exception as error:  # noqa: BLE001
        info.problems.append(f"Not a readable workbook: {error}")
        return info

    names = patcher.sheet_names()
    for required in (SHEET, DIG_SHEET):
        if required not in names:
            info.problems.append(
                f"The template has no '{required}' sheet - found {', '.join(names)}."
            )

    slot = patcher.find_image_slot(SHEET)
    if slot is not None:
        info.slot_aspect = slot.aspect
    elif SHEET in names:
        info.problems.append("No image placeholder found on the StakingReport sheet.")

    try:
        workbook = openpyxl.load_workbook(
            io.BytesIO(template_bytes), read_only=True, data_only=True
        )
        try:
            if CONTACTS_SHEET in workbook.sheetnames:
                for row in workbook[CONTACTS_SHEET].iter_rows(
                    min_row=2, max_row=50, max_col=2, values_only=True
                ):
                    name = _text(row[0]) if row else ""
                    if name:
                        info.stakers.append(name)
                        info.phones[name] = _text(row[1]) if len(row) > 1 else ""
            if SHEET in workbook.sheetnames:
                for row in workbook[SHEET].iter_rows(
                    min_row=28, max_row=28, min_col=10, max_col=10, values_only=True
                ):
                    info.staking_notes = _text(row[0]) if row else ""
        finally:
            workbook.close()
    except Exception as error:  # noqa: BLE001
        info.problems.append(f"Could not read the template's contacts: {error}")
    return info


# ---------------------------------------------------------------------------
# Cached values
# ---------------------------------------------------------------------------

_SIMPLE_FORMULA = re.compile(r"<c r=\"([A-Z]+\d+)\"[^>]*>\s*<f>([^<]*)</f>")
_DIG_SHEET_REF = re.compile(r"^'?Dig Sheet'?!\$?([A-Z]+)\$?(\d+)$")
_SHEET_REF = re.compile(r"^'?([A-Za-z_ ]+?)'?!\$?([A-Z]+)\$?(\d+)$")
_XLOOKUP_PHONE = re.compile(r"^_xlfn\.XLOOKUP\(([A-Z]+\d+),Contacts!", re.I)


def _unescape(text: str) -> str:
    return (text.replace("&apos;", "'").replace("&quot;", '"')
            .replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&"))


def computed_values(patcher: XlsmPatcher, dig: Dig, written: dict,
                    phones: dict) -> dict:
    """{sheet: {ref: value}} for the template's simple reference formulas.

    Only direct references are evaluated - ='Dig Sheet'!G8, =StakingReport!J2 -
    plus the phone lookup, which is the only other formula in the template. A
    formula that is anything else is left with whatever it already cached;
    Excel recalculates it on open regardless.
    """
    formulas = {}
    for sheet_name in CACHED_SHEETS:
        path = patcher._sheet_paths.get(sheet_name)
        if path is None:
            continue
        xml = patcher._text(path)
        formulas[sheet_name] = {
            ref: _unescape(text) for ref, text in _SIMPLE_FORMULA.findall(xml)
        }

    values: dict = {name: {} for name in formulas}
    resolving: set = set()

    def evaluate(sheet_name: str, ref: str):
        if ref in values.get(sheet_name, {}):
            return values[sheet_name][ref]
        formula = formulas.get(sheet_name, {}).get(ref)
        if formula is None:
            # A plain cell: whatever this app wrote there, else empty.
            if sheet_name == SHEET and ref in written:
                return written[ref]
            return None
        if (sheet_name, ref) in resolving:
            return None
        resolving.add((sheet_name, ref))
        try:
            result = _evaluate_formula(sheet_name, formula)
        finally:
            resolving.discard((sheet_name, ref))
        values[sheet_name][ref] = result
        return result

    def _evaluate_formula(sheet_name: str, formula: str):
        match = _XLOOKUP_PHONE.match(formula)
        if match:
            key = evaluate(sheet_name, match.group(1))
            return phones.get(_text(key), "")
        target = _DIG_SHEET_REF.match(formula)
        if target:
            return dig.cells.get(f"{target.group(1)}{target.group(2)}")
        target = _SHEET_REF.match(formula)
        if target:
            other = target.group(1)
            if other == DIG_SHEET:
                return dig.cells.get(f"{target.group(2)}{target.group(3)}")
            return evaluate(other, f"{target.group(2)}{target.group(3)}")
        raise _NotSimple()

    for sheet_name, group in formulas.items():
        for ref in list(group):
            try:
                evaluate(sheet_name, ref)
            except _NotSimple:
                continue
    return values


class _NotSimple(Exception):
    pass


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

_HIDDEN_HELPER_COLUMNS_FROM = 16  # column P onward is hidden on the Dig Sheet tab


# ---------------------------------------------------------------------------
# Keeping text inside its box
# ---------------------------------------------------------------------------
# A dig's directions, or an alignment sheet name, can be longer than the box the
# template gives it, and Excel simply cuts the text off at the edge. One-line
# cells are set to shrink to fit; the wrapped Directions / Staking Notes boxes
# get the largest font that fits, worked out here with Calibri's own metrics
# (Carlito, which is metric-compatible, when it is installed).

SHRINK_TO_FIT_CELLS = ("G4", "I3", "A6", "I6", "C11", "C19", "C20", "C21", "C22", "C23", "C24")
WRAPPED_BOXES = (CELL_DIRECTIONS, CELL_STAKING_NOTES)
_CALIBRI_FILES = (
    "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
    "/usr/share/fonts/truetype/crosextra/Carlito.ttf",
    "/usr/share/fonts/carlito/Carlito-Regular.ttf",
    "/usr/share/fonts/TTF/Carlito-Regular.ttf",
)
_CALIBRI = {"font": None, "tried": False}


def _text_width(text: str, size_pt: float) -> float:
    """Width in pixels (96 dpi) of ``text`` in Calibri at ``size_pt``."""
    pixels = size_pt * 96 / 72
    if not _CALIBRI["tried"]:
        _CALIBRI["tried"] = True
        for path in _CALIBRI_FILES:
            try:
                _CALIBRI["font"] = ImageFont.truetype(path, 100)
                break
            except OSError:
                continue
    font = _CALIBRI["font"]
    if font is not None:
        return font.getlength(text) * pixels / 100.0
    # No Carlito: Calibri averages about half an em a character; a little more
    # than that errs towards a smaller font rather than clipped text.
    return len(text) * pixels * 0.54


def wrapped_line_count(text: str, size_pt: float, width_px: float) -> int:
    lines = 0
    for paragraph in (text or "").split("\n"):
        current = ""
        count = 1
        for word in paragraph.split(" "):
            trial = f"{current} {word}" if current else word
            if _text_width(trial, size_pt) <= width_px:
                current = trial
                continue
            if current:
                count += 1
            # A single word wider than the box breaks across lines.
            overflow = _text_width(word, size_pt)
            while overflow > width_px:
                count += 1
                overflow -= width_px
            current = word if overflow <= width_px else ""
        lines += count
    return lines


def fit_font_size(text: str, width_px: float, height_pt: float,
                  base: float = 11.0, minimum: float = 6.0) -> float:
    """The largest font size, at most ``base``, at which wrapped text fits the box."""
    usable_width = max(width_px - 10, 20) / 1.04          # padding + a safety margin
    size = base
    while size > minimum:
        needed = wrapped_line_count(text, size, usable_width) * size * 1.37 + 4
        if needed <= height_pt:
            return size
        size -= 0.5
    return minimum


def fit_report_text(patcher: XlsmPatcher, sheet: str = SHEET) -> None:
    patcher.restyle(sheet, SHRINK_TO_FIT_CELLS, shrink=True)
    for ref in WRAPPED_BOXES:
        text = patcher.cell_text(sheet, ref)
        box = patcher.box_size(sheet, ref)
        if not text or box is None:
            continue
        size = fit_font_size(text, box[0], box[1])
        if size < 11.0:
            patcher.restyle(sheet, [ref], size=size)


def build_staking_report(
    template_bytes: bytes,
    dig: Dig,
    settings: ReportSettings,
    phones: Optional[dict] = None,
) -> tuple:
    """Return (.xlsm bytes, warnings) for a single dig."""
    patcher = XlsmPatcher(template_bytes)
    warnings = []

    names = patcher.sheet_names()
    for required in (SHEET, DIG_SHEET):
        if required not in names:
            raise ValueError(
                f"The template has no '{required}' sheet - found {', '.join(names)}."
            )

    # 1. The dig sheet, cell for cell, into the template's Dig Sheet tab.
    missing = patcher.set_values(DIG_SHEET, dig.cells, literal=True)
    visible = [ref for ref in missing if _column_index(_COLUMN_RE.match(ref).group(1))
               < _HIDDEN_HELPER_COLUMNS_FROM]
    if visible:
        warnings.append(
            "Could not place these dig sheet cells in the template: "
            + ", ".join(sorted(visible))
        )

    # 2. What a dig sheet cannot know.
    written = {}
    if settings.surveyor_name:
        # C11 only. The phone number in C12 is an XLOOKUP against the template's
        # own Contacts table keyed off this name - writing C12 would replace that
        # formula with a static value.
        written[CELL_SURVEYOR] = settings.surveyor_name
    written[CELL_DIRECTIONS] = dig.directions or ""
    if settings.staking_notes is not None:
        written[CELL_STAKING_NOTES] = settings.staking_notes

    # An empty Directions box is left exactly as the template has it.
    to_write = {ref: value for ref, value in written.items()
                if not (ref == CELL_DIRECTIONS and not value)}
    unplaced = patcher.set_values(SHEET, to_write)
    if unplaced:
        warnings.append(
            "Could not place these cells in the template: " + ", ".join(sorted(unplaced))
        )

    # 3. Formulas stay; the value each one shows is refreshed.
    for sheet_name, cached in computed_values(
        patcher, dig, written, phones or {}
    ).items():
        patcher.set_cached(sheet_name, cached)
    fit_report_text(patcher)
    patcher.force_full_recalc()

    if dig.aerial_image:
        slot = patcher.find_image_slot(SHEET)
        if slot is None:
            warnings.append(
                "No image placeholder found on the template's StakingReport sheet, "
                "so the aerial image was not inserted."
            )
        else:
            patcher.replace_image(slot, _encode_for_slot(dig.aerial_image, slot), "jpeg")

    return patcher.to_bytes(), warnings


def template_image_aspect(template_bytes: bytes) -> Optional[float]:
    """The aspect of the image slot in this template.

    Templates differ, so the aerial is rendered to the slot actually present
    rather than to a constant - otherwise the image is cropped to fit and the
    title card in its corner gets clipped.
    """
    try:
        slot = XlsmPatcher(template_bytes).find_image_slot(SHEET)
    except Exception:  # noqa: BLE001
        return None
    return slot.aspect if slot else None


def _encode_for_slot(image_bytes: bytes, slot) -> bytes:
    """Match the slot's aspect so Excel does not stretch the image."""
    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")

    target_aspect = slot.aspect
    width, height = image.size
    if abs((width / height) - target_aspect) > 0.005:
        if width / height > target_aspect:
            new_width = int(round(height * target_aspect))
            left = (width - new_width) // 2
            image = image.crop((left, 0, left + new_width, height))
        else:
            new_height = int(round(width / target_aspect))
            top = (height - new_height) // 2
            image = image.crop((0, top, width, top + new_height))

    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=90, subsampling=0, optimize=True)
    return buffer.getvalue()



# ==========================================================================
# REPORTS
# ==========================================================================
# The two extra deliverables Enterprise asks for with each dig:
#
#   * an Excavation Survey Report - an Adobe XFA form (.pdf). Its fields live in
#     an XML "datasets" packet inside the PDF, so filling it means replacing that
#     packet. It is replaced with an *incremental update* appended to the end of
#     the original file: the original bytes (and the signature that gives Adobe
#     Reader permission to fill and save the form) are left untouched, so the
#     dropdowns and the check box stay live and editable in Adobe.
#   * a Profile workbook (.xlsx) - the template's blank profile sheet with the
#     header block filled in; everything else on it is left as the template has it.
#
# Both read the same facts out of the dig sheet's title line, e.g.
#   Line ID 600-601-602 / Asmt ID 116,  East Leg Mainline, 8.625"  Kearney to Moberly

PIPELINE_CODE = "763"

_TITLE_RE = re.compile(
    r"Line\s*ID\s*(?P<lid>.+?)\s*/\s*Asmt\s*ID\s*(?P<aid>\d+)\s*,\s*(?P<rest>.+)$",
    re.I | re.S,
)


@dataclass
class AssetInfo:
    """What a dig sheet's title line says about the pipeline."""

    line_id: str = ""          # "600-601-602"
    assessment_id: str = ""    # "116"
    line_name: str = ""        # "East Leg Mainline"
    segment_name: str = ""     # '8.625" Kearney to Moberly'
    dig_number: str = ""       # "02A"

    @property
    def complete(self) -> bool:
        return all((self.line_id, self.assessment_id, self.line_name, self.segment_name))


def _squash(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def asset_info(dig: Dig) -> AssetInfo:
    """Parse the title line of the dig sheet (cell A5)."""
    info = AssetInfo(assessment_id=str(dig.package_id or ""), dig_number=dig.number or "")
    match = _TITLE_RE.search(_squash(dig.header))
    if not match:
        return info
    info.line_id = _squash(match.group("lid"))
    info.assessment_id = match.group("aid")
    # The line name is everything up to the next comma; the rest is the segment.
    name, _, segment = match.group("rest").partition(",")
    info.line_name = _squash(name)
    info.segment_name = _squash(segment)
    return info


def _asset_warnings(info: AssetInfo) -> list:
    missing = [label for label, value in (
        ("Line ID", info.line_id), ("assessment ID", info.assessment_id),
        ("line name", info.line_name), ("segment name", info.segment_name),
        ("dig number", info.dig_number)) if not value]
    if not missing:
        return []
    return [f"Could not read the {', '.join(missing)} from the dig sheet title, "
            "so those fields were left blank."]


# ---------------------------------------------------------------------------
# Excavation Survey Report (XFA PDF)
# ---------------------------------------------------------------------------

def _xfa_datasets(reader) -> tuple:
    """(object number, generation, current text) of the form's datasets packet."""
    form = reader.trailer["/Root"].get("/AcroForm")
    xfa = form.get("/XFA") if form is not None else None
    if not xfa:
        raise ValueError("This PDF has no XFA form data - is it the Excavation "
                         "Survey Report template?")
    for index in range(0, len(xfa) - 1, 2):
        if str(xfa[index]) == "datasets":
            reference = xfa.raw_get(index + 1) if hasattr(xfa, "raw_get") else xfa[index + 1]
            stream = reference.get_object()
            return reference.idnum, reference.generation, stream.get_data().decode("utf-8")
    raise ValueError("The PDF's form has no 'datasets' packet.")


def _stream_update_tail(original: bytes, reader, object_number: int,
                        generation: int, data: bytes) -> bytes:
    """The bytes to append to ``original`` to replace one stream object.

    The original bytes are not touched. A new copy of the stream object and a
    new cross-reference stream pointing at it are appended, which is how Adobe
    products save changes to a document that has usage rights (the signature
    covers the original bytes, and an appended update leaves that valid).
    """
    import struct

    trailer = reader.trailer
    size = int(trailer["/Size"])
    previous = int(re.search(rb"startxref\s+(\d+)\s*%%EOF\s*$", original[-200:]).group(1))
    root = trailer.raw_get("/Root") if hasattr(trailer, "raw_get") else trailer["/Root"]
    info = trailer.raw_get("/Info") if hasattr(trailer, "raw_get") else trailer.get("/Info")

    out = bytearray()
    if not original.endswith((b"\n", b"\r")):
        out += b"\r\n"

    stream_offset = len(original) + len(out)
    out += (f"{object_number} {generation} obj\n<</Length {len(data)}>>\nstream\n"
            ).encode("ascii") + data + b"\nendstream\nendobj\n"

    xref_number = size
    xref_offset = len(original) + len(out)
    rows = struct.pack(">BIH", 1, stream_offset, generation) + \
        struct.pack(">BIH", 1, xref_offset, 0)

    ids = trailer.get("/ID")
    id_entry = b""
    if ids:
        id_entry = b"/ID[" + b"".join(
            b"<" + bytes(item.original_bytes if hasattr(item, "original_bytes") else item).hex().encode() + b">"
            for item in ids) + b"]"
    dictionary = (
        f"<</Type/XRef/Size {size + 1}/W[1 4 2]/Index[{object_number} 1 {xref_number} 1]"
        f"/Root {root.idnum} {root.generation} R"
        + (f"/Info {info.idnum} {info.generation} R" if info is not None else "")
    ).encode("ascii") + id_entry + f"/Prev {previous}/Length {len(rows)}>>".encode("ascii")
    out += f"{xref_number} 0 obj\n".encode("ascii") + dictionary + b"\nstream\n" + rows \
        + b"\nendstream\nendobj\n"
    out += f"startxref\n{xref_offset}\n%%EOF\n".encode("ascii")
    return bytes(out)


def _fill_xfa_fields(datasets: str, values: dict) -> str:
    """Set the empty (or filled) elements named in ``values`` inside <xfa:data>.

    The packet also carries a data *description* section that repeats every
    element name, empty - only the data section may be touched.
    """
    cut = datasets.find("<dd:dataDescription")
    data, tail = (datasets, "") if cut < 0 else (datasets[:cut], datasets[cut:])
    for name, value in values.items():
        text = escape(str(value)) if value != "" else ""
        pattern = re.compile(
            r"<%(n)s\s*/>|<%(n)s\s*>[^<]*</%(n)s\s*>" % {"n": re.escape(name)}
        )
        replacement = f"<{name}>{text}</{name}>" if text else f"<{name}/>"
        data, count = pattern.subn(lambda _m: replacement, data, count=1)
        if count != 1:
            raise ValueError(f"The template has no '{name}' field to fill.")
    return data + tail


class AppendedPdf:
    """A filled PDF held as the template plus a few KB of appended update.

    The Excavation Survey Report template is nearly 4 MB and every dig gets its
    own copy, so keeping a hundred whole copies in memory is wasteful. The
    template is shared and only the tail differs.
    """

    def __init__(self, template: bytes, tail: bytes):
        self.template, self.tail = template, tail

    def __len__(self) -> int:
        return len(self.template) + len(self.tail)

    def to_bytes(self) -> bytes:
        return self.template + self.tail

    def write_to(self, handle) -> None:
        handle.write(self.template)
        handle.write(self.tail)


def excavation_field_values(dig: Dig) -> tuple:
    """({xfa element: value}, warnings) for one dig."""
    info = asset_info(dig)
    values = {
        "pipelineCode": PIPELINE_CODE,
        "pipelineName": info.line_name,
        "pipelineNumber": info.line_id,
        "assessmentsegmentAssessmentID": info.assessment_id,
        "assessmentsegmentName": info.segment_name,
        "excavationNumber": info.dig_number,
        # Each report is its own record. The template's ID is shared by every copy.
        "excavationSurveyReportID": str(uuid.uuid4()),
        # The template ships with another dig's remark in the comment box. Blank
        # it - a remark about the wrong weld in a signed report is worse than none.
        "excavationSurveyReportComment": "",
    }
    return values, _asset_warnings(info)


def build_excavation_report(template_bytes: bytes, dig: Dig, lazy: bool = False) -> tuple:
    """Return (PDF, warnings) - the template with the asset block filled in.

    The PDF is ``bytes``, or an ``AppendedPdf`` when ``lazy`` is true.

    The girth weld references, elevations, utilities and submit/validate boxes
    are left blank and editable; survey accuracy stays as the template has it.
    """
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(template_bytes))
    number, generation, original = _xfa_datasets(reader)

    values, warnings = excavation_field_values(dig)
    updated = _fill_xfa_fields(original, values).encode("utf-8")
    tail = _stream_update_tail(template_bytes, reader, number, generation, updated)
    return (AppendedPdf(template_bytes, tail) if lazy else template_bytes + tail), warnings


# ---------------------------------------------------------------------------
# Profile workbook
# ---------------------------------------------------------------------------

PROFILE_SHEET = "Blank"
PROFILE_SHEET_RENAMED = "Profile"


def build_profile_report(template_bytes: bytes, dig: Dig) -> tuple:
    """Return (.xlsx bytes, warnings): the template's blank sheet, header filled in."""
    info = asset_info(dig)
    warnings = _asset_warnings(info)

    patcher = XlsmPatcher(template_bytes)
    names = patcher.sheet_names()
    if PROFILE_SHEET not in names:
        raise ValueError(f"The profile template has no '{PROFILE_SHEET}' sheet - "
                         f"found {', '.join(names)}.")

    cells = {
        "E3": info.line_name,
        "E4": info.segment_name,
        "E5": (f"Line ID {info.line_id} /ASMT ID {info.assessment_id}"
               if info.line_id and info.assessment_id else None),
        "C9": info.dig_number or None,
        "U6": _number(dig.us_weld_id) if _number(dig.us_weld_id) is not None else (dig.us_weld_id or None),
        "U7": station_label(dig.us_weld_station_ft) if dig.us_weld_station_ft is not None else None,
        "U8": _number(dig.ds_weld_id) if _number(dig.ds_weld_id) is not None else (dig.ds_weld_id or None),
        "U9": station_label(dig.ds_weld_station_ft) if dig.ds_weld_station_ft is not None else None,
    }
    cells = {k: (int(v) if isinstance(v, float) and v.is_integer() else v) for k, v in cells.items()}
    missing = patcher.set_values(PROFILE_SHEET, cells, literal=True)
    if missing:
        warnings.append("Could not place these profile cells: " + ", ".join(sorted(missing)))

    patcher.keep_only_sheet(PROFILE_SHEET, rename_to=PROFILE_SHEET_RENAMED)
    patcher.force_full_recalc()
    return patcher.to_bytes(), warnings


# ==========================================================================
# ASBUILT
# ==========================================================================
# The As-Built drawing (.vsdx), one per dig.
#
# A Visio file is a zip of XML parts, so it is patched the way the xlsm is: the
# template's page1.xml is edited in place and nothing else is touched. What is
# filled, by shape ID in the Enterprise template:
#
#   264  title block (pipeline, section, LID/AID, dig #) - narrowed to sit between
#        the two logos, and the font shrunk per line so it never runs under one
#   193 / 222  D/S and U/S weld callouts (weld number, stationing)
#   294  the compass rose inside the Gibson box - turned so north is where it
#        really is when the pipe is drawn with flow running up the page
#   289  the lower-right job box (county ... alignment sheet; D.O.C, prepared by
#        and date are left for the person to fill in)
#
# The callouts and dimension shapes in the lower left are left exactly as the
# template has them.

VSDX_PAGE = "visio/pages/page1.xml"
TITLE_SHAPE, DS_WELD_SHAPE, US_WELD_SHAPE = "264", "193", "222"
COMPASS_SHAPE, JOB_BOX_SHAPE = "294", "289"
LOGO_SHAPES = ("224", "228")
TITLE_SIZES_PT = (24.0, 24.0, 18.0, 24.0)     # pipeline, section, LID/AID, dig #
TITLE_SIDE_GAP_IN = 0.1
TITLE_MARGIN_IN = 0.1
JOB_BOX_LINES = {
    "County": "county_state", "Legal": "legal_description", "Tract": "tract_number",
    "Aerial": "aerial_marker", "Milepost": "milepost", "Alignment": "alignment_sheet",
}
_CARLITO_BOLD = ["/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf",
                 "/usr/share/fonts/carlito/Carlito-Bold.ttf"]
_ARIAL_FILES = {False: ["/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"],
                True: ["/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
                       "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]}
_FONT_CACHE: dict = {}


def _measure(text: str, size_pt: float, files: list, key: str) -> float:
    """Width in inches of ``text`` at ``size_pt`` in the first font file that opens."""
    if key not in _FONT_CACHE:
        _FONT_CACHE[key] = None
        for path in files:
            try:
                _FONT_CACHE[key] = ImageFont.truetype(path, 200)
                break
            except OSError:
                continue
    font = _FONT_CACHE[key]
    if font is None:
        return len(text) * size_pt / 72.0 * 0.58      # a little wide, so it errs small
    return font.getlength(text) * size_pt / 200.0 / 72.0


def fit_size(text: str, base_pt: float, width_in: float, measure, minimum: float = 6.0) -> float:
    """The largest size up to ``base_pt`` at which ``text`` is no wider than ``width_in``."""
    full = measure(text, base_pt)
    if full <= width_in or full <= 0:
        return base_pt
    return max(minimum, round(base_pt * width_in / full * 2 - 0.5) / 2)


def _shape_span(xml: str, shape_id: str):
    match = re.search(rf"<Shape ID='{shape_id}'.*?</Shape>", xml, re.S)
    return (match.start(), match.end()) if match else None


def _cell_value(block: str, name: str) -> Optional[float]:
    match = re.search(rf"<Cell N='{name}' V='([^']*)'", block)
    try:
        return float(match.group(1)) if match else None
    except ValueError:
        return None


def _logo_gap(xml: str) -> Optional[tuple]:
    """(left edge, right edge) of the free space between the two header logos."""
    edges = []
    for shape_id in LOGO_SHAPES:
        span = _shape_span(xml, shape_id)
        if not span:
            return None
        block = xml[span[0]:span[1]]
        pin, width, loc = (_cell_value(block, n) for n in ("PinX", "Width", "LocPinX"))
        if None in (pin, width, loc):
            return None
        edges.append((pin - loc, pin - loc + width))
    edges.sort()
    return edges[0][1], edges[1][0]


def _split_title(info: "AssetInfo") -> tuple:
    name, segment = info.line_name, info.segment_name
    if not segment and ":" in name:
        name, _, segment = name.partition(":")
    return _squash(name), _squash(segment)


def _character_rows(sizes_pt: list, first_style: int = 21) -> str:
    rows = "".join(
        f"<Row IX='{i}'><Cell N='Style' V='{first_style}'/>"
        f"<Cell N='Size' V='{size / 72.0:.6f}' U='PT'/></Row>"
        for i, size in enumerate(sizes_pt))
    return f"<Section N='Character'>{rows}</Section>"


def _fill_title(xml: str, dig: "Dig", info: "AssetInfo") -> tuple:
    span = _shape_span(xml, TITLE_SHAPE)
    if not span:
        return xml, ["The as-built template has no title block (shape 264)."]
    block = xml[span[0]:span[1]]
    gap = _logo_gap(xml)
    left, right = gap if gap else (2.375, 6.323)
    width = max(right - left - 2 * TITLE_SIDE_GAP_IN, 1.0)
    centre = (left + right) / 2.0
    usable = width - 2 * TITLE_MARGIN_IN

    name, segment = _split_title(info)
    lid = f"LID: {info.line_id or '?'}  AID: {info.assessment_id or '?'}"
    number = f"Dig # {dig.number}" if dig.number else "Dig #"
    lines = [(name, TITLE_SIZES_PT[0]), (segment, TITLE_SIZES_PT[1]),
             (lid, TITLE_SIZES_PT[2]), (number, TITLE_SIZES_PT[3])]
    lines = [(t, b) for t, b in lines if t]

    sizes = [fit_size(t, b, usable, lambda x, y: _measure(x, y, _CARLITO_BOLD, "carlito-bold"))
             for t, b in lines]
    newline = "\r\n" if "\r\n" in xml else "\n"
    text = "".join(f"<cp IX='{i}'/>{escape(t)}{newline}" for i, (t, _) in enumerate(lines))
    cells = (f"<Cell N='PinX' V='{centre:.6f}'/>"
             + re.search(r"<Cell N='PinY' V='[^']*'/>", block).group(0)
             + f"<Cell N='Width' V='{width:.6f}'/><Cell N='LocPinX' V='{width / 2:.6f}'/>")
    new = (re.match(r"<Shape [^>]*>", block).group(0) + cells + _character_rows(sizes)
           + f"<Text><pp IX='0'/>{text}</Text></Shape>")
    return xml[:span[0]] + new + xml[span[1]:], []


def _fill_weld(xml: str, shape_id: str, label: str, number: str, station_ft) -> str:
    span = _shape_span(xml, shape_id)
    if not span:
        return xml
    block = xml[span[0]:span[1]]
    station = station_label(station_ft) if station_ft is not None else "XXX+XX"
    first = f"{label} Weld # {number or 'XXXX'}"
    newline = "\r\n" if "\r\n" in block else "\n"
    block = re.sub(r"(<Text>.*?>)[^<>]*?XXXX\r?\nXXX\+XX\r?\n(</Text>)",
                   lambda m: f"{m.group(1)}{escape(first)}{newline}{escape(station)}{newline}{m.group(2)}",
                   block, flags=re.S)
    return xml[:span[0]] + block + xml[span[1]:]


def _rotate_compass(xml: str, bearing: float) -> str:
    """Turn the rose so north lies where it does on the ground.

    The drawing runs flow up the page, so up is the flow bearing; north is then
    that far round the other way, and Visio angles count counter-clockwise.
    """
    match = re.search(rf"<Shape ID='{COMPASS_SHAPE}'[^>]*>", xml)
    if not match:
        return xml
    tag = match.group(0)
    angle = f"<Cell N='Angle' V='{math.radians(bearing):.10f}'/>"
    return xml[:match.end()] + angle + xml[match.end():]


def _master_text(zf: zipfile.ZipFile, master_id: str) -> Optional[str]:
    try:
        masters = zf.read("visio/masters/masters.xml").decode("utf-8")
        rels = zf.read("visio/masters/_rels/masters.xml.rels").decode("utf-8")
        block = re.search(rf"<Master ID='{master_id}'.*?</Master>", masters, re.S).group(0)
        rid = re.search(r"<Rel r:id='(\w+)'", block).group(1)
        target = re.search(rf'<Relationship [^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels)
        if not target:
            target = re.search(rf'<Relationship [^>]*Target="([^"]+)"[^>]*Id="{rid}"', rels)
        part = zf.read("visio/masters/" + target.group(1)).decode("utf-8")
        match = re.search(r"<Text>(.*?)</Text>", part, re.S)
        return match.group(1) if match else None
    except (KeyError, AttributeError):
        return None


def _job_box_values(dig: "Dig") -> dict:
    cells = dig.cells

    def pick(ref):
        value = cells.get(ref)
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        return "" if value is None else str(value).strip()

    milepost = _number(cells.get("C40"))
    return {
        "county_state": pick("C36"), "legal_description": pick("C37"),
        "tract_number": pick("C38"), "aerial_marker": pick("C39"),
        "milepost": f"{milepost:.2f}" if milepost is not None else "",
        "alignment_sheet": pick("C41"),
    }


def _fill_job_box(xml: str, zf: zipfile.ZipFile, dig: "Dig") -> tuple:
    span = _shape_span(xml, JOB_BOX_SHAPE)
    if not span:
        return xml, ["The as-built template has no job information box (shape 289)."]
    block = xml[span[0]:span[1]]
    master = re.search(r"Master='(\d+)'", block)
    template_text = _master_text(zf, master.group(1)) if master else None
    if not template_text:
        return xml, ["Could not read the job information box text from the as-built template."]

    values = _job_box_values(dig)
    box_length = 3.42 - 0.25
    sizes: dict = {}
    out_lines, warnings = [], []
    newline = "\r\n" if "\r\n" in template_text else "\n"
    for line in template_text.replace("\r\n", "\n").split("\n"):
        plain = re.sub(r"<[^>]*>", "", line)
        key = next((k for label, k in JOB_BOX_LINES.items() if plain.lstrip().startswith(label)), None)
        if key is not None:
            value = values.get(key, "")
            if value:
                line = re.sub(r"\?+(?:\.\?+)?", lambda m: escape(value), line, count=1)
                plain = re.sub(r"<[^>]*>", "", line)
            else:
                warnings.append(f"The dig sheet has no {key.replace('_', ' ')}; left as the template has it.")
            size = fit_size(plain, 10.0, box_length,
                            lambda x, y: _measure(x, y, _ARIAL_FILES[True], "arial-bold") * 0.96)
            if size < 10.0:
                for cp in re.findall(r"<cp IX='(\d+)'/>", line):
                    sizes[int(cp)] = size
        out_lines.append(line)
    rows = "".join(f"<Row IX='{i}'><Cell N='Size' V='{s / 72.0:.6f}' U='PT'/></Row>"
                   for i, s in sorted(sizes.items()))
    section = f"<Section N='Character'>{rows}</Section>" if rows else ""
    block = block.replace("</Shape>", f"{section}<Text>{newline.join(out_lines)}</Text></Shape>")
    return xml[:span[0]] + block + xml[span[1]:], warnings


def build_as_built(template_bytes: bytes, dig: "Dig") -> tuple:
    """Return (.vsdx bytes, warnings): the template with this dig's details filled in."""
    info = asset_info(dig)
    warnings = []
    if not info.line_id or not info.assessment_id:
        warnings.append("Could not read the Line ID / assessment ID from the dig sheet title.")
    source = zipfile.ZipFile(io.BytesIO(template_bytes))
    if VSDX_PAGE not in source.namelist():
        raise ValueError("The as-built template is not a Visio drawing (.vsdx).")
    xml = source.read(VSDX_PAGE).decode("utf-8")

    xml, problems = _fill_title(xml, dig, info)
    warnings.extend(problems)
    if not dig.us_weld_id or not dig.ds_weld_id:
        warnings.append("U/S or D/S weld number not found; that callout was left as the template has it.")
    xml = _fill_weld(xml, US_WELD_SHAPE, "U/S", dig.us_weld_id, dig.us_weld_station_ft)
    xml = _fill_weld(xml, DS_WELD_SHAPE, "D/S", dig.ds_weld_id, dig.ds_weld_station_ft)

    bearing = dig.flow_bearing
    if bearing is None:
        warnings.append("No coordinates to work out the flow direction; the compass was not turned.")
    else:
        xml = _rotate_compass(xml, bearing)
    xml, problems = _fill_job_box(xml, source, dig)
    warnings.extend(problems)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = xml.encode("utf-8") if item.filename == VSDX_PAGE else source.read(item.filename)
            target.writestr(item, data)
    return buffer.getvalue(), warnings



# ==========================================================================
# GENERATION
# ==========================================================================
# The whole run, kept out of the UI so it can be exercised by the tests.

XLSM_MIME = "application/vnd.ms-excel.sheet.macroEnabled.12"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _file_part(text: str) -> str:
    """A dig number made safe to use inside a file name."""
    return re.sub(r'[\\/:*?"<>|]+', "-", str(text or "")).strip() or "Unknown"


def folder_for(dig: Dig) -> str:
    """Each assessment (AID) gets its own folder in the output."""
    return f"AID {_file_part(dig.package_id)}"


def staking_report_name(dig: Dig) -> str:
    return f"{folder_for(dig)}/Staking Reports/02-03 Survey Staking Report_Dig#{_file_part(dig.number)}.xlsm"


def excavation_report_name(dig: Dig) -> str:
    return (f"{folder_for(dig)}/Excavation Survey Reports/"
            f"ExcavationSurvey Report_Dig {_file_part(dig.number)}.pdf")


def profile_report_name(dig: Dig) -> str:
    return f"{folder_for(dig)}/Profile Reports/Profile Dig #{_file_part(dig.number)}.xlsx"


def as_built_name(dig: Dig) -> str:
    return f"{folder_for(dig)}/As-Builts/As-Built Dig #{_file_part(dig.number)}.vsdx"


def cheat_sheet_name(package_id: str) -> str:
    return f"AID {_file_part(package_id)}/AID {_file_part(package_id)} - Dig Stake Cheat Sheet.xlsx"


@dataclass
class RunOptions:
    make_aerial: bool = True
    basemap: str = "Esri World Imagery"
    span_feet: float = 2600.0
    make_directions: bool = False
    token: str = ""
    auto_notes: bool = True


def generate_files(
    digs: list,
    template_bytes: bytes,
    cheat_templates: dict,
    pipeline: Optional[PipelineData],
    settings: ReportSettings,
    options: RunOptions,
    phones: Optional[dict] = None,
    progress=None,
    excavation_template: Optional[bytes] = None,
    profile_template: Optional[bytes] = None,
    as_built_template: Optional[bytes] = None,
) -> tuple:
    """Return ({path: bytes}, [issues]) for the chosen digs.

    Paths are relative to the zip root: one folder per AID, with the staking
    reports and cheat sheet in it and the excavation and profile reports in
    subfolders. The last two are made only when their template was uploaded.
    """
    outputs, issues = {}, []
    slot_aspect = template_image_aspect(template_bytes)
    total = max(1, len(digs))

    def tick(fraction, text):
        if progress is not None:
            progress(min(max(fraction, 0.0), 1.0), text)

    for index, dig in enumerate(digs, start=1):
        tick((index - 0.8) / total, f"{dig.name} (AID {dig.package_id}) - aerial image")
        dig.aerial_image = None
        if options.make_aerial:
            try:
                dig.aerial_image = render_aerial(
                    dig, pipeline, basemap=options.basemap,
                    span_feet=float(options.span_feet), token=options.token,
                    slot_aspect=slot_aspect,
                )
                if dig.aerial_image is None and dig.latitude is not None:
                    issues.append(f"{dig.name} (AID {dig.package_id}): aerial image could not be rendered.")
            except Exception as error:  # noqa: BLE001 - surfaced to the user
                issues.append(f"{dig.name} (AID {dig.package_id}): aerial image failed - {error}")

        if options.make_directions and options.token and dig.latitude is not None:
            tick((index - 0.45) / total, f"{dig.name} (AID {dig.package_id}) - directions")
            try:
                dig.directions = generate_directions(
                    float(dig.latitude), float(dig.longitude), options.token
                ).paragraph
            except Exception as error:  # noqa: BLE001
                issues.append(f"{dig.name} (AID {dig.package_id}): directions failed - {error}")

        tick((index - 0.15) / total, f"{dig.name} (AID {dig.package_id}) - staking report")
        try:
            report, report_warnings = build_staking_report(
                template_bytes, dig, settings, phones
            )
            outputs[staking_report_name(dig)] = report
            issues.extend(f"{dig.name} (AID {dig.package_id}): {w}" for w in report_warnings)
        except Exception as error:  # noqa: BLE001
            issues.append(f"{dig.name} (AID {dig.package_id}): staking report failed - {error}")

        if excavation_template:
            tick((index - 0.1) / total, f"{dig.name} (AID {dig.package_id}) - excavation report")
            try:
                pdf, pdf_warnings = build_excavation_report(excavation_template, dig, lazy=True)
                outputs[excavation_report_name(dig)] = pdf
                issues.extend(f"{dig.name} (AID {dig.package_id}): {w}" for w in pdf_warnings)
            except Exception as error:  # noqa: BLE001
                issues.append(f"{dig.name} (AID {dig.package_id}): excavation report failed - {error}")

        if profile_template:
            tick((index - 0.05) / total, f"{dig.name} (AID {dig.package_id}) - profile")
            try:
                sheet, sheet_warnings = build_profile_report(profile_template, dig)
                outputs[profile_report_name(dig)] = sheet
                # The same title line feeds the excavation report, which already warned.
                if not excavation_template:
                    issues.extend(f"{dig.name} (AID {dig.package_id}): {w}" for w in sheet_warnings)
            except Exception as error:  # noqa: BLE001
                issues.append(f"{dig.name} (AID {dig.package_id}): profile failed - {error}")

        if as_built_template:
            tick((index - 0.02) / total, f"{dig.name} (AID {dig.package_id}) - as-built")
            try:
                drawing, drawing_warnings = build_as_built(as_built_template, dig)
                outputs[as_built_name(dig)] = drawing
                issues.extend(f"{dig.name} (AID {dig.package_id}): as-built - {w}"
                              for w in drawing_warnings)
            except Exception as error:  # noqa: BLE001
                issues.append(f"{dig.name} (AID {dig.package_id}): as-built failed - {error}")

    # One cheat sheet per package: the packages are separate assessments, so
    # their odometers are not on one scale.
    packages: dict = {}
    for dig in digs:
        packages.setdefault(dig.package_id, []).append(dig)
    for package_id, group in packages.items():
        template = pick_cheat_template(cheat_templates, group)
        if template is None:
            issues.append(f"AID {package_id}: no cheat sheet template uploaded.")
            continue
        try:
            outputs[cheat_sheet_name(package_id)] = build_cheat_sheet(
                template, group, auto_notes=options.auto_notes,
            )
        except Exception as error:  # noqa: BLE001
            issues.append(f"AID {package_id}: cheat sheet failed - {error}")

    tick(1.0, "Done")
    return outputs, issues


# ==========================================================================
# STREAMLIT UI
# ==========================================================================

def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="Enterprise Dig File Generator", page_icon="🛠️",
                       layout="wide")

    for key in ("digs", "pipeline", "outputs", "template_info", "read_notes"):
        st.session_state.setdefault(key, None)

    def mapbox_token() -> str:
        try:
            return st.secrets["mapbox"]["token"]
        except Exception:  # noqa: BLE001
            return ""

    token = mapbox_token()

    # ------------------------------------------------------------------
    # Uploads
    # ------------------------------------------------------------------
    st.title("Enterprise Dig File Generator")
    st.caption(
        "Fills a staking report, an excavation survey report and a profile for "
        "every dig sheet in your dig packages, plus a cheat sheet for each package, "
        "in one folder per AID. On the staking reports, latitude, longitude, "
        "elevation, EDOC, survey date and photos are field measurements and are left blank."
    )

    left, right = st.columns(2)
    with left:
        template_file = st.file_uploader(
            "Staking report template (.xlsm)", type=["xlsm"], key="template",
            help="The Enterprise Dig Stake Template That Autopopulates.",
        )
        cheat_files = st.file_uploader(
            "Cheat sheet template(s) (.xlsx)", type=["xlsx"],
            accept_multiple_files=True, key="cheat",
            help="Upload the ascending one, the descending one, or both. With "
                 "both, each package gets the one that matches its stationing.",
        )
    with right:
        package_files = st.file_uploader(
            "Dig packages (.xlsm) - one or several", type=["xlsm", "xlsx"],
            accept_multiple_files=True, key="packages",
        )
        kmz_file = st.file_uploader(
            "Pipeline KMZ (centerline for the aerial image)", type=["kmz", "kml"],
            key="kmz",
        )

    left2, right2 = st.columns(2)
    with left2:
        excavation_file = st.file_uploader(
            "Excavation Survey Report template (.pdf) - optional", type=["pdf"],
            key="excavation",
            help="Enterprise's fillable Excavation Survey Report. One is filled "
                 "for every dig and stays editable in Adobe.",
        )
    with right2:
        profile_file = st.file_uploader(
            "Profile template (.xlsx) - optional", type=["xlsx"], key="profile",
            help="Enterprise's dig profile workbook. One is filled for every dig.",
        )
    as_built_file = st.file_uploader(
        "As-Built template (.vsdx) - optional", type=["vsdx"], key="as_built",
        help="Enterprise's As-Built drawing. One is filled for every dig, in an As-Builts folder.",
    )
    if excavation_file is not None:
        try:
            from pypdf import PdfReader
            _xfa_datasets(PdfReader(io.BytesIO(excavation_file.getvalue())))
        except Exception as error:  # noqa: BLE001
            st.warning(f"Excavation Survey Report template: {error}")
    if profile_file is not None:
        try:
            names = XlsmPatcher(profile_file.getvalue()).sheet_names()
            if PROFILE_SHEET not in names:
                st.warning(f"Profile template: no '{PROFILE_SHEET}' sheet - found "
                           f"{', '.join(names)}.")
        except Exception as error:  # noqa: BLE001
            st.warning(f"Profile template: not a readable workbook ({error}).")

    template_info: Optional[TemplateInfo] = None
    if template_file is not None:
        template_info = read_template_info(template_file.getvalue())
        for problem in template_info.problems:
            st.warning(f"Staking report template: {problem}")

    # ------------------------------------------------------------------
    # Sidebar
    # ------------------------------------------------------------------
    st.sidebar.header("Report details")
    stakers = template_info.stakers if template_info else []
    if stakers:
        choice = st.sidebar.selectbox(
            "Surveyed by", ["(leave blank)"] + stakers, index=0,
            help="The phone number fills itself in from the template's Contacts table.",
        )
        surveyor_name = "" if choice == "(leave blank)" else choice
    else:
        surveyor_name = st.sidebar.text_input(
            "Surveyed by", value="",
            help="The phone number fills itself in from the template's Contacts table.",
        )

    st.sidebar.header("Aerial image")
    make_aerial = st.sidebar.checkbox("Generate aerial images", value=True)
    basemap = st.sidebar.selectbox("Imagery", basemap_options(bool(token)), index=0)
    span_feet = st.sidebar.slider("Image width across the ground (ft)", 800, 6000, 2600, 200)

    st.sidebar.header("Directions")
    make_directions = st.sidebar.checkbox(
        "Generate driving directions", value=bool(token), disabled=not token,
        help="Uses the same Mapbox logic as the Dig Site Directions Generator. "
             "Add a token to .streamlit/secrets.toml to enable.",
    )
    if not token:
        st.sidebar.caption("No Mapbox token found - directions will be left blank.")

    st.sidebar.header("Cheat sheet")
    auto_notes = st.sidebar.checkbox("Auto-note digs close to the previous one", value=True)

    ready = bool(template_file and cheat_files and package_files)
    if not ready:
        st.info(
            "Upload the staking report template, a cheat sheet template and at "
            "least one dig package to begin."
        )

    # ------------------------------------------------------------------
    # Read the uploads
    # ------------------------------------------------------------------
    if st.button("Read uploads", type="primary", disabled=not ready):
        digs, problems = [], []
        with st.status("Reading dig packages...", expanded=False):
            for upload in package_files:
                try:
                    found = parse_dig_package(upload.getvalue(), upload.name)
                except Exception as error:  # noqa: BLE001 - surfaced to the user
                    problems.append(f"{upload.name}: {error}")
                    continue
                if not found:
                    problems.append(f"{upload.name}: no 'Dig (NN)' sheets were found.")
                digs.extend(found)

        pipeline = None
        if kmz_file is not None:
            try:
                pipeline = parse_kmz(kmz_file.getvalue(), kmz_file.name)
                if pipeline.is_empty:
                    problems.append(f"{kmz_file.name}: no lines or placemarks found.")
                    pipeline = None
            except Exception as error:  # noqa: BLE001
                problems.append(f"{kmz_file.name}: {error}")
        reconcile_with_kmz(digs, pipeline)
        if make_aerial and (pipeline is None or not pipeline.lines):
            problems.append(
                "No pipeline centerline loaded, so the aerial images will show the "
                "dig and its references but no pipe."
            )

        duplicates = {}
        for dig in digs:
            duplicates.setdefault((dig.package_id, dig.name), []).append(dig.sheet_name)
        for (package_id, name), sheets in duplicates.items():
            if len(sheets) > 1:
                problems.append(
                    f"AID {package_id}: {name} appears on more than one sheet "
                    f"({', '.join(sheets)}). Their reports would overwrite each other."
                )

        st.session_state.digs = digs
        st.session_state.pipeline = pipeline
        st.session_state.outputs = None
        st.session_state.read_notes = problems

    for problem in st.session_state.read_notes or []:
        st.warning(problem)

    digs = st.session_state.digs or []
    if not digs:
        _downloads(st)
        return

    # ------------------------------------------------------------------
    # Review
    # ------------------------------------------------------------------
    packages = sorted({d.package_id for d in digs})
    st.subheader(f"{len(digs)} dig sheets found in {len(packages)} package"
                 f"{'s' if len(packages) != 1 else ''}")

    shown = digs
    st.caption(
        "Every dig below is generated. Everything comes straight from the dig "
        "package - fix a wrong value in the package, not here."
    )

    table_key = "review"
    edited = st.data_editor(
        [
            {
                "AID": dig.package_id,
                "Dig": dig.name,
                "ODO": dig.odometer,
                "Station": station_label(dig.station_ft),
                "U/S ref": dig.upstream_reference,
                "U/S ft": _round2(dig.us_ref_distance),
                "D/S ref": dig.downstream_reference,
                "D/S ft": _round2(dig.ds_ref_distance),
                "County": dig.county_state,
                "Tract": dig.tract_number,
                "Alignment sheet": dig.alignment_sheet,
                "HCA": dig.hca,
                "Notes": dig.notes,
            }
            for dig in shown
        ],
        width="stretch",
        num_rows="fixed",
        key=table_key,
        disabled=[c for c in ("AID", "Dig", "ODO", "Station", "U/S ref", "U/S ft",
                              "D/S ref", "D/S ft", "County", "Tract",
                              "Alignment sheet", "HCA")],
        column_config={
            "Notes": st.column_config.TextColumn(
                "Cheat sheet notes",
                help="Goes in the Notes column of the cheat sheet.",
            ),
        },
    )

    warnings = [(d.name, d.package_id, w) for d in shown for w in d.warnings]
    if warnings:
        with st.expander(f"{len(warnings)} thing(s) to check", expanded=len(warnings) <= 12):
            for name, package_id, warning in warnings:
                st.write(f"**{name}** (AID {package_id}) - {warning}")

    selected = list(digs)
    for row, dig in zip(edited, shown):
        dig.notes = row.get("Notes") or ""

    # ------------------------------------------------------------------
    # Generate
    # ------------------------------------------------------------------
    if len(selected) > 25 and (make_aerial or make_directions):
        st.caption(
            f"All {len(selected)} digs are generated. Aerial images and directions "
            "are fetched one dig at a time, so this takes a few minutes."
        )

    if st.button(f"Generate files for all {len(selected)} digs", type="primary"):
        cheat_templates = {}
        for upload in cheat_files:
            flavour = cheat_sheet_flavour(upload.getvalue())
            cheat_templates.setdefault(flavour, upload.getvalue())

        # The Staking Notes box is left exactly as the template has it.
        settings = ReportSettings(surveyor_name=surveyor_name)
        options = RunOptions(
            make_aerial=make_aerial, basemap=basemap, span_feet=float(span_feet),
            make_directions=make_directions, token=token,
            auto_notes=auto_notes,
        )

        bar = st.progress(0.0, text="Starting...")
        outputs, issues = generate_files(
            selected,
            template_file.getvalue(),
            cheat_templates,
            st.session_state.pipeline,
            settings,
            options,
            phones=template_info.phones if template_info else {},
            progress=lambda fraction, text: bar.progress(fraction, text=text),
            excavation_template=excavation_file.getvalue() if excavation_file else None,
            profile_template=profile_file.getvalue() if profile_file else None,
            as_built_template=as_built_file.getvalue() if as_built_file else None,
        )
        bar.empty()
        st.session_state.outputs = outputs
        st.session_state.zip_path = None
        st.session_state.generated = [d.key for d in selected]
        for issue in issues:
            st.warning(issue)

    _downloads(st)


def write_zip(files: dict, handle) -> None:
    """Write {path: bytes | AppendedPdf} into a zip, one entry at a time."""
    with zipfile.ZipFile(handle, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in sorted(files):
            data = files[name]
            if isinstance(data, AppendedPdf):
                # PDFs barely compress, and a hundred of them take minutes to deflate.
                entry = zipfile.ZipInfo(name, date_time=_dt.datetime.now().timetuple()[:6])
                entry.compress_type = zipfile.ZIP_STORED
                with archive.open(entry, "w", force_zip64=True) as target:
                    data.write_to(target)
            else:
                archive.writestr(name, data)


def _downloads(st) -> None:
    outputs = st.session_state.outputs or {}
    if not outputs:
        return
    st.subheader("Files")

    # Built once per run, on disk: with the excavation PDFs this is hundreds of MB.
    path = st.session_state.get("zip_path")
    if not path or not os.path.exists(path):
        handle = tempfile.NamedTemporaryFile(suffix=".zip", delete=False)
        with handle:
            write_zip(outputs, handle)
        path = st.session_state.zip_path = handle.name
    megabytes = os.path.getsize(path) / 1_000_000

    with open(path, "rb") as zipped:
        st.download_button(
            f"Download everything (.zip, {megabytes:,.0f} MB)",
            data=zipped,
            file_name=f"Enterprise Dig Files {date.today():%Y-%m-%d}.zip",
            mime="application/zip",
            type="primary",
        )
    st.caption("One folder per AID. The cheat sheet is in the folder; staking, "
               "excavation and profile reports are in their own subfolders.")

    with st.expander(f"{len(outputs)} files in the zip", expanded=False):
        for name in sorted(outputs):
            st.text(name)

    keys = set(st.session_state.get("generated") or [])
    previews = [d for d in (st.session_state.digs or [])
                if d.key in keys and d.aerial_image]
    if previews:
        with st.expander("Aerial images", expanded=False):
            for dig in previews:
                st.image(dig.aerial_image, caption=f"{dig.name} (AID {dig.package_id})",
                         width="stretch")


if __name__ == "__main__":
    main()
