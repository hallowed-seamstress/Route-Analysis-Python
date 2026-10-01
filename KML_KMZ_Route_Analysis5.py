#!/usr/bin/env python3
"""
KML / KMZ Route Analysis
========================

Standalone GUI conversion of the analysis logic in
WTPL_DesignationPoints_RouteSections_BETA4.xlsm.

What it does
------------
1. Lets the user select one KML or KMZ file.
2. Detects every LineString and Polygon placemark.
3. Preselects all lines as routes and suggests long/narrow polygons.
4. Lets the user decide which items are routes, rename them, group alternatives
   such as A1/A2/A3 under corridor A, and enter a different design span for
   each route before analysis.
5. Converts a selected corridor-like polygon to an estimated centerline by
   splitting its boundary at the two ends and averaging the opposing sides.
6. Creates an Excel workbook with a checkbox-driven route dashboard, live
   selected-route totals, a selected-options map, route segments, designation
   points, projected corridor-junction angles, and analysis notes.

Analysis rules carried over from the workbook
---------------------------------------------
- Segment tangent count = integer part of segment length / design span.
- A segment exactly equal to one design span counts as zero tangents.
- Angle splits are user-configurable before analysis (two or three cutoffs).
- Default internal deflection angle 0 to less than 2 degrees: Tangent.
- Internal deflection angle 2 to less than 15 degrees: Light Angle.
- Internal deflection angle 15 to less than 45 degrees: Medium Angle.
- Internal deflection angle 45 degrees and above: Heavy Angle.
- Only the outside endpoints of the first and last selected corridor groups
  are Substation A-Frame points. Other option endpoints are corridor
  connection points.

Requirement
-----------
Python 3.10+ and XlsxWriter 3.2 or newer:

    py -m pip install --upgrade xlsxwriter

The script itself uses only the Python standard library for KML/KMZ parsing,
geodesic calculations, polygon-centerline estimation, and the GUI.
"""

from __future__ import annotations

import math
import re
import sys
import traceback
import warnings
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Sequence
from xml.etree import ElementTree as ET
from xml.parsers import expat

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError as exc:  # pragma: no cover - depends on local Python install
    raise SystemExit(
        "Tkinter is required. Install the standard Python distribution from "
        "python.org, then run this script again."
    ) from exc


EARTH_RADIUS_M = 6_371_008.8
DEFAULT_SPAN_M = 250.0
POLYGON_SUGGESTED_RATIO = 4.0
EPSILON_M = 0.01

ROUTE_COLORS = (
    "C00000", "004CC7", "008A3B", "9A3FC4", "E67E00", "007F86",
    "C2185B", "5B6A00", "6B4F2A", "455A64", "7B1FA2", "2E7D32",
    "AD1457", "1565C0", "EF6C00", "00695C", "4527A0", "558B2F",
    "6D4C41", "283593", "00838F", "9E9D24", "F4511E", "3949AB",
)


@dataclass
class LocalProjection:
    lon0: float
    lat0: float

    @property
    def cos_lat0(self) -> float:
        return max(1e-12, math.cos(math.radians(self.lat0)))

    def forward(self, lon: float, lat: float) -> tuple[float, float]:
        x = math.radians(lon - self.lon0) * EARTH_RADIUS_M * self.cos_lat0
        y = math.radians(lat - self.lat0) * EARTH_RADIUS_M
        return x, y

    def inverse(self, x: float, y: float) -> tuple[float, float]:
        lon = self.lon0 + math.degrees(x / (EARTH_RADIUS_M * self.cos_lat0))
        lat = self.lat0 + math.degrees(y / EARTH_RADIUS_M)
        return lon, lat


@dataclass
class GeometryCandidate:
    candidate_id: str
    placemark_name: str
    display_name: str
    folder_path: str
    geometry_type: str
    coordinates: list[tuple[float, float]]
    source_file: str
    polygon_ratio: float | None = None
    suggested: bool = False
    estimated_length_m: float = 0.0


@dataclass(frozen=True)
class AngleSplits:
    cutoffs: tuple[float, ...] = (2.0, 15.0, 45.0)

    def __post_init__(self) -> None:
        if (len(self.cutoffs) not in (2, 3)
                or any(not math.isfinite(v) or not 0 < v < 180 for v in self.cutoffs)
                or any(a >= b for a, b in zip(self.cutoffs, self.cutoffs[1:]))):
            raise ValueError("Enter two or three increasing angle cutoffs between 0 and 180, e.g. 3, 30 or 2, 15, 45.")

    @classmethod
    def parse(cls, text: str) -> AngleSplits:
        try:
            return cls(tuple(float(v.strip()) for v in text.split(",")))
        except ValueError:
            raise ValueError("Enter two or three increasing angle cutoffs between 0 and 180, e.g. 3, 30 or 2, 15, 45.") from None

    def bands(self) -> list[tuple[str, int, float, float | None]]:
        names = [("Tangent", 1), ("Medium Angle", 3), ("Heavy Angle", 4)]
        if len(self.cutoffs) == 3:
            names.insert(1, ("Light Angle", 2))
        edges = (0.0,) + self.cutoffs
        return [(name, code, edges[i], edges[i + 1] if i + 1 < len(edges) else None)
                for i, (name, code) in enumerate(names)]

    def range_label(self, code: int) -> str:
        for _, band_code, lower, upper in self.bands():
            if code == band_code:
                return f"{lower:g}°+" if upper is None else f"{lower:g}–<{upper:g}°"
        return "Not used"

    def preview(self) -> str:
        return " | ".join(f"{name}: {self.range_label(code)}" for name, code, _, _ in self.bands())


@dataclass
class RouteRequest:
    candidate: GeometryCandidate
    route_name: str
    span_m: float
    corridor_group: str = ""
    angle_splits: AngleSplits = field(default_factory=AngleSplits)


@dataclass
class RouteResult:
    request: RouteRequest
    coordinates: list[tuple[float, float]]
    segment_rows: list[dict] = field(default_factory=list)
    designation_rows: list[dict] = field(default_factory=list)
    total_length_m: float = 0.0
    total_tangents: int = 0
    tangent_dp_count: int = 0
    light_count: int = 0
    medium_count: int = 0
    heavy_count: int = 0
    substation_count: int = 0
    connection_count: int = 0
    polygon_note: str = ""


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def infer_corridor_group(route_name: str) -> str:
    """Infer A from names such as A1, A2, Route A3, or Alternative A4."""
    text = route_name.strip()
    prefixed = re.match(
        r"^(?:alternative|alt|route|corridor)\s*[-_:]?\s*([A-Za-z]+)\s*[-_ ]?\d+(?:\D|$)",
        text,
        flags=re.IGNORECASE,
    )
    compact = re.match(r"^([A-Za-z]+)\s*[-_ ]?\d+(?:\D|$)", text)
    match = prefixed or compact
    if match:
        return match.group(1).upper()
    return text or "UNGROUPED"


def natural_sort_key(text: str) -> tuple[tuple[int, object], ...]:
    parts = re.split(r"(\d+)", text.casefold())
    return tuple((1, int(part)) if part.isdigit() else (0, part) for part in parts if part)


def direct_child_text(element: ET.Element, wanted_name: str) -> str:
    for child in element:
        if local_name(child.tag) == wanted_name:
            return (child.text or "").strip()
    return ""


def first_descendant(element: ET.Element, wanted_name: str) -> ET.Element | None:
    for child in element.iter():
        if local_name(child.tag) == wanted_name:
            return child
    return None


def parse_coordinate_text(text: str | None) -> list[tuple[float, float]]:
    points: list[tuple[float, float]] = []
    if not text:
        return points
    for token in re.split(r"\s+", text.strip()):
        parts = token.split(",")
        if len(parts) < 2:
            continue
        try:
            lon = float(parts[0])
            lat = float(parts[1])
        except ValueError:
            continue
        if math.isfinite(lon) and math.isfinite(lat):
            points.append((lon, lat))
    return points


def parse_gx_track(track: ET.Element) -> list[tuple[float, float]]:
    points: list[tuple[float, float]] = []
    for element in track.iter():
        if local_name(element.tag) != "coord" or not element.text:
            continue
        parts = element.text.split()
        if len(parts) < 2:
            continue
        try:
            lon, lat = float(parts[0]), float(parts[1])
        except ValueError:
            continue
        if math.isfinite(lon) and math.isfinite(lat):
            points.append((lon, lat))
    return points


def parse_kml_xml(data: bytes) -> ET.Element:
    """Parse normally, repairing only missing, recognizable namespace bindings.

    Inspect XML names with Expat (namespace expansion disabled), so descriptions,
    comments and CDATA cannot be mistaken for prefixed element/attribute names.
    Inject declarations in memory; never rewrite the source KML or KMZ.
    """
    try:
        return ET.fromstring(data)
    except ET.ParseError as original:
        if original.code != expat.errors.codes[expat.errors.XML_ERROR_UNBOUND_PREFIX]:
            raise

    known = {
        "kml": "http://www.opengis.net/kml/2.2",
        "gx": "http://www.google.com/kml/ext/2.2",
        "atom": "http://www.w3.org/2005/Atom",
        "xal": "urn:oasis:names:tc:ciq:xsdschema:xAL:2.0",
        "xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsd": "http://www.w3.org/2001/XMLSchema",
    }
    parser = expat.ParserCreate()
    scopes: list[set[str]] = [{"xml"}]
    missing: set[str] = set()
    root_name = ""
    root_index = 0
    declared_encoding = "utf-8"

    def xml_decl(_version, encoding, _standalone) -> None:
        nonlocal declared_encoding
        if encoding:
            declared_encoding = encoding

    def reject_doctype(*_args) -> None:
        raise ValueError("Namespace repair is unavailable for XML with a DOCTYPE. Re-export as standard KML.")

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal root_name, root_index
        if not root_name:
            root_name = name
            root_index = parser.CurrentByteIndex
            # An arbitrary prefix on the KML root is identifiable by that root.
            if name.endswith(":kml"):
                known.setdefault(name.split(":", 1)[0], known["kml"])
        scope = scopes[-1].copy()
        scope.update(key.split(":", 1)[1] for key in attrs if key.startswith("xmlns:"))
        scopes.append(scope)
        for qualified in (name, *attrs):
            if ":" in qualified and not qualified.startswith("xmlns:"):
                prefix = qualified.split(":", 1)[0]
                if prefix not in scope:
                    missing.add(prefix)

    parser.XmlDeclHandler = xml_decl
    parser.StartDoctypeDeclHandler = reject_doctype
    parser.StartElementHandler = start
    parser.EndElementHandler = lambda _name: scopes.pop()
    try:
        parser.Parse(data, True)
    except expat.ExpatError as exc:
        raise ValueError(f"Could not inspect the KML XML for missing namespaces: {exc}") from exc

    unknown = missing.difference(known)
    if unknown:
        raise ValueError(
            "KML uses undeclared XML namespace prefix(es): " + ", ".join(sorted(unknown))
            + ". Their meaning cannot be inferred safely. Re-export the KML/KMZ with namespace declarations."
        )
    # Preserve the source encoding and byte positions, including UTF-16 files.
    if data.startswith((b"\xff\xfe", b"<\x00")):
        encoding = "utf-16-le"
    elif data.startswith((b"\xfe\xff", b"\x00<")):
        encoding = "utf-16-be"
    else:
        encoding = declared_encoding
    declarations = "".join(f' xmlns:{prefix}="{known[prefix]}"' for prefix in sorted(missing))
    insert_at = root_index + len(("<" + root_name).encode(encoding))
    repaired = data[:insert_at] + declarations.encode(encoding) + data[insert_at:]
    root = ET.fromstring(repaired)
    warnings.warn("KML namespace declarations repaired in memory: " + ", ".join(sorted(missing)),
                  UserWarning, stacklevel=2)
    return root


def load_kml_root(path: Path) -> tuple[ET.Element, str]:
    if path.suffix.lower() == ".kmz":
        with zipfile.ZipFile(path, "r") as archive:
            kml_names = [n for n in archive.namelist() if n.lower().endswith(".kml")]
            if not kml_names:
                raise ValueError("The KMZ does not contain a KML document.")
            chosen = next((n for n in kml_names if n.lower().endswith("doc.kml")), kml_names[0])
            data = archive.read(chosen)
            source_label = f"{path.name} / {chosen}"
    else:
        data = path.read_bytes()
        source_label = path.name

    try:
        return parse_kml_xml(data), source_label
    except ET.ParseError as exc:
        raise ValueError(f"Could not parse the KML XML: {exc}") from exc


def geometry_length_m(points: Sequence[tuple[float, float]]) -> float:
    return sum(haversine_m(a, b) for a, b in zip(points, points[1:]))


def polygon_elongation(points: Sequence[tuple[float, float]]) -> float:
    ring = normalize_polygon_ring(points)
    if len(ring) < 3:
        return 1.0
    projection = projection_for(ring)
    xy = [projection.forward(*point) for point in ring]
    cx = sum(x for x, _ in xy) / len(xy)
    cy = sum(y for _, y in xy) / len(xy)
    xx = sum((x - cx) ** 2 for x, _ in xy) / len(xy)
    yy = sum((y - cy) ** 2 for _, y in xy) / len(xy)
    xy_cov = sum((x - cx) * (y - cy) for x, y in xy) / len(xy)
    trace = xx + yy
    disc = math.sqrt(max(0.0, (xx - yy) ** 2 + 4.0 * xy_cov ** 2))
    major = max(0.0, (trace + disc) / 2.0)
    minor = max(0.0, (trace - disc) / 2.0)
    if minor <= 1e-9:
        return 999.0
    return min(999.0, math.sqrt(major / minor))


def polygon_axis_length_m(points: Sequence[tuple[float, float]]) -> float:
    """Return a fast length estimate without building the polygon centerline."""
    ring = normalize_polygon_ring(points)
    if len(ring) < 3:
        return 0.0
    projection = projection_for(ring)
    xy = [projection.forward(*point) for point in ring]
    axis = principal_axis(xy)
    projected = [x * axis[0] + y * axis[1] for x, y in xy]
    return max(projected) - min(projected)


def extract_candidates(path: Path) -> list[GeometryCandidate]:
    root, source_label = load_kml_root(path)
    candidates: list[GeometryCandidate] = []
    counter = 0

    def add_candidate(
        placemark_name: str,
        folder_path: str,
        geometry_type: str,
        points: list[tuple[float, float]],
        geometry_number: int,
    ) -> None:
        nonlocal counter
        minimum = 3 if geometry_type == "Polygon" else 2
        if len(points) < minimum:
            return
        counter += 1
        suffix = f" [{geometry_type} {geometry_number}]" if geometry_number > 1 else ""
        base = placemark_name or f"Unnamed {geometry_type} {counter}"
        display = base + suffix
        ratio = polygon_elongation(points) if geometry_type == "Polygon" else None
        suggested = geometry_type == "LineString" or bool(ratio and ratio >= POLYGON_SUGGESTED_RATIO)
        if geometry_type == "Polygon":
            try:
                # Keep the file-selection step responsive. The more involved
                # centerline interpolation is performed only for polygons the
                # user actually selects for analysis.
                estimated = polygon_axis_length_m(points)
            except Exception:
                estimated = 0.0
        else:
            estimated = geometry_length_m(remove_consecutive_duplicates(points))
        candidates.append(
            GeometryCandidate(
                candidate_id=f"G{counter:04d}",
                placemark_name=placemark_name or base,
                display_name=display,
                folder_path=folder_path,
                geometry_type=geometry_type,
                coordinates=points,
                source_file=source_label,
                polygon_ratio=ratio,
                suggested=suggested,
                estimated_length_m=estimated,
            )
        )

    def process_placemark(placemark: ET.Element, folders: list[str]) -> None:
        name = direct_child_text(placemark, "name")
        folder_path = " / ".join(folders)
        geometry_number = 0
        for element in placemark.iter():
            kind = local_name(element.tag)
            if kind == "LineString":
                coordinates_element = first_descendant(element, "coordinates")
                points = parse_coordinate_text(coordinates_element.text if coordinates_element is not None else None)
                geometry_number += 1
                add_candidate(name, folder_path, "LineString", points, geometry_number)
            elif kind == "Polygon":
                outer = first_descendant(element, "outerBoundaryIs")
                coordinates_element = first_descendant(outer, "coordinates") if outer is not None else None
                points = parse_coordinate_text(coordinates_element.text if coordinates_element is not None else None)
                geometry_number += 1
                add_candidate(name, folder_path, "Polygon", points, geometry_number)
            elif kind == "Track" and element.tag.startswith("{"):
                points = parse_gx_track(element)
                geometry_number += 1
                add_candidate(name, folder_path, "LineString", points, geometry_number)

    def walk(element: ET.Element, folders: list[str]) -> None:
        kind = local_name(element.tag)
        if kind == "Placemark":
            process_placemark(element, folders)
            return
        next_folders = folders
        if kind in ("Document", "Folder"):
            folder_name = direct_child_text(element, "name")
            if folder_name:
                next_folders = folders + [folder_name]
        for child in element:
            if local_name(child.tag) in ("Document", "Folder", "Placemark"):
                walk(child, next_folders)

    walk(root, [])
    return candidates


def projection_for(points: Sequence[tuple[float, float]]) -> LocalProjection:
    if not points:
        raise ValueError("No coordinates are available for projection.")
    return LocalProjection(
        lon0=sum(p[0] for p in points) / len(points),
        lat0=sum(p[1] for p in points) / len(points),
    )


def haversine_m(a: tuple[float, float], b: tuple[float, float]) -> float:
    lon1, lat1 = map(math.radians, a)
    lon2, lat2 = map(math.radians, b)
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    h = math.sin(dlat / 2.0) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2.0) ** 2
    return 2.0 * EARTH_RADIUS_M * math.asin(min(1.0, math.sqrt(max(0.0, h))))


def initial_bearing_deg(a: tuple[float, float], b: tuple[float, float]) -> float:
    lon1, lat1 = map(math.radians, a)
    lon2, lat2 = map(math.radians, b)
    dlon = lon2 - lon1
    y = math.sin(dlon) * math.cos(lat2)
    x = math.cos(lat1) * math.sin(lat2) - math.sin(lat1) * math.cos(lat2) * math.cos(dlon)
    return (math.degrees(math.atan2(y, x)) + 360.0) % 360.0


def deflection_angle_deg(
    previous: tuple[float, float],
    current: tuple[float, float],
    following: tuple[float, float],
) -> float:
    incoming = initial_bearing_deg(previous, current)
    outgoing = initial_bearing_deg(current, following)
    return abs((outgoing - incoming + 180.0) % 360.0 - 180.0)


def remove_consecutive_duplicates(
    points: Sequence[tuple[float, float]],
    tolerance_m: float = EPSILON_M,
) -> list[tuple[float, float]]:
    cleaned: list[tuple[float, float]] = []
    for point in points:
        if not cleaned or haversine_m(cleaned[-1], point) > tolerance_m:
            cleaned.append(point)
    return cleaned


def normalize_polygon_ring(points: Sequence[tuple[float, float]]) -> list[tuple[float, float]]:
    ring = remove_consecutive_duplicates(points)
    if len(ring) >= 2 and haversine_m(ring[0], ring[-1]) <= EPSILON_M:
        ring.pop()
    return ring


def principal_axis(xy: Sequence[tuple[float, float]]) -> tuple[float, float]:
    cx = sum(x for x, _ in xy) / len(xy)
    cy = sum(y for _, y in xy) / len(xy)
    xx = sum((x - cx) ** 2 for x, _ in xy)
    yy = sum((y - cy) ** 2 for _, y in xy)
    xy_cov = sum((x - cx) * (y - cy) for x, y in xy)
    angle = 0.5 * math.atan2(2.0 * xy_cov, xx - yy)
    return math.cos(angle), math.sin(angle)


def edge_endcap_score(
    a: tuple[float, float],
    b: tuple[float, float],
    axis: tuple[float, float],
    min_projection: float,
    max_projection: float,
    find_minimum: bool,
) -> float:
    ux, uy = axis
    mx, my = (a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0
    projection = mx * ux + my * uy
    projection_range = max(1e-9, max_projection - min_projection)
    normalized = (projection - min_projection) / projection_range
    dx, dy = b[0] - a[0], b[1] - a[1]
    length = math.hypot(dx, dy)
    perpendicularity = abs((dx * uy - dy * ux) / length) if length > 1e-9 else 0.0
    extreme_distance = normalized if find_minimum else 1.0 - normalized
    return extreme_distance + 0.08 * (1.0 - perpendicularity)


def insert_endcap_midpoints(
    ring: Sequence[tuple[float, float]],
    min_edge: int,
    max_edge: int,
) -> tuple[list[tuple[float, float]], int, int]:
    augmented: list[tuple[float, float]] = []
    min_index = -1
    max_index = -1
    count = len(ring)
    for index, point in enumerate(ring):
        augmented.append(point)
        midpoint = (
            (point[0] + ring[(index + 1) % count][0]) / 2.0,
            (point[1] + ring[(index + 1) % count][1]) / 2.0,
        )
        if index == min_edge:
            augmented.append(midpoint)
            min_index = len(augmented) - 1
        if index == max_edge:
            augmented.append(midpoint)
            max_index = len(augmented) - 1
    if min_index < 0 or max_index < 0:
        raise ValueError("Could not locate both polygon ends.")
    return augmented, min_index, max_index


def cyclic_path(
    points: Sequence[tuple[float, float]],
    start: int,
    end: int,
    step: int,
) -> list[tuple[float, float]]:
    result = [points[start]]
    index = start
    guard = 0
    while index != end:
        index = (index + step) % len(points)
        result.append(points[index])
        guard += 1
        if guard > len(points):
            raise ValueError("Polygon boundary traversal did not terminate.")
    return result


def resample_polyline_xy(
    points: Sequence[tuple[float, float]],
    sample_count: int,
) -> list[tuple[float, float]]:
    if sample_count <= 1:
        return [points[0]]
    cumulative = [0.0]
    for a, b in zip(points, points[1:]):
        cumulative.append(cumulative[-1] + math.hypot(b[0] - a[0], b[1] - a[1]))
    total = cumulative[-1]
    if total <= 1e-9:
        return [points[0]] * sample_count
    result: list[tuple[float, float]] = []
    segment = 0
    for sample_index in range(sample_count):
        target = total * sample_index / (sample_count - 1)
        while segment + 1 < len(cumulative) - 1 and cumulative[segment + 1] < target:
            segment += 1
        a = points[segment]
        b = points[segment + 1]
        span = cumulative[segment + 1] - cumulative[segment]
        fraction = 0.0 if span <= 1e-12 else (target - cumulative[segment]) / span
        result.append((a[0] + fraction * (b[0] - a[0]), a[1] + fraction * (b[1] - a[1])))
    return result


def simplify_centerline_xy(
    points: Sequence[tuple[float, float]],
    angle_tolerance_deg: float = 0.15,
) -> list[tuple[float, float]]:
    if len(points) <= 2:
        return list(points)
    output = [points[0]]
    for index in range(1, len(points) - 1):
        a = output[-1]
        b = points[index]
        c = points[index + 1]
        v1 = (b[0] - a[0], b[1] - a[1])
        v2 = (c[0] - b[0], c[1] - b[1])
        len1, len2 = math.hypot(*v1), math.hypot(*v2)
        if len1 <= EPSILON_M:
            continue
        if len2 <= EPSILON_M:
            continue
        turn = abs(math.degrees(math.atan2(v1[0] * v2[1] - v1[1] * v2[0], v1[0] * v2[0] + v1[1] * v2[1])))
        if turn >= angle_tolerance_deg:
            output.append(b)
    output.append(points[-1])
    return output


def polygon_centerline(points: Sequence[tuple[float, float]]) -> list[tuple[float, float]]:
    ring_lonlat = normalize_polygon_ring(points)
    if len(ring_lonlat) < 3:
        raise ValueError("A polygon route needs at least three distinct boundary points.")
    projection = projection_for(ring_lonlat)
    ring = [projection.forward(*point) for point in ring_lonlat]
    axis = principal_axis(ring)
    projected = [x * axis[0] + y * axis[1] for x, y in ring]
    min_projection, max_projection = min(projected), max(projected)
    if max_projection - min_projection <= EPSILON_M:
        raise ValueError("The polygon is too small to estimate a centerline.")

    min_edge = min(
        range(len(ring)),
        key=lambda i: edge_endcap_score(
            ring[i], ring[(i + 1) % len(ring)], axis,
            min_projection, max_projection, True,
        ),
    )
    max_edge = min(
        range(len(ring)),
        key=lambda i: edge_endcap_score(
            ring[i], ring[(i + 1) % len(ring)], axis,
            min_projection, max_projection, False,
        ),
    )
    if min_edge == max_edge:
        raise ValueError("The polygon ends could not be distinguished.")

    augmented, start, end = insert_endcap_midpoints(ring, min_edge, max_edge)
    side_a = cyclic_path(augmented, start, end, +1)
    side_b = cyclic_path(augmented, start, end, -1)
    sample_count = min(1000, max(3, len(side_a), len(side_b)))
    sampled_a = resample_polyline_xy(side_a, sample_count)
    sampled_b = resample_polyline_xy(side_b, sample_count)
    center_xy = [
        ((a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0)
        for a, b in zip(sampled_a, sampled_b)
    ]
    center_xy = simplify_centerline_xy(center_xy)
    center_lonlat = [projection.inverse(x, y) for x, y in center_xy]
    center_lonlat = remove_consecutive_duplicates(center_lonlat)
    if len(center_lonlat) < 2:
        raise ValueError("The estimated polygon centerline has fewer than two points.")
    return center_lonlat


def safe_identifier(text: str, fallback: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", text.strip()).strip("_")
    return cleaned[:40] or fallback


def tangent_count(length_m: float, span_m: float) -> int:
    if length_m <= 0.0:
        return 0
    if math.isclose(length_m, span_m, rel_tol=0.0, abs_tol=1e-7):
        return 0
    return int(length_m / span_m)


def classify_angle(angle: float, splits: AngleSplits | None = None) -> tuple[str, int]:
    magnitude = abs(angle)
    for name, code, _, upper in (splits or AngleSplits()).bands():
        if upper is None or magnitude < upper:
            return name, code
    raise AssertionError("Angle bands must include an open-ended final band")


def analyze_route(request: RouteRequest, route_number: int) -> RouteResult:
    candidate = request.candidate
    if not request.corridor_group.strip():
        request.corridor_group = infer_corridor_group(request.route_name)
    if candidate.geometry_type == "Polygon":
        coordinates = polygon_centerline(candidate.coordinates)
        ratio = candidate.polygon_ratio or 0.0
        polygon_note = (
            f"Estimated centerline from polygon boundary; elongation ratio {ratio:.2f}."
            + (" Review visually because the polygon is not strongly elongated." if ratio < 2.0 else "")
        )
    else:
        coordinates = remove_consecutive_duplicates(candidate.coordinates)
        polygon_note = ""
    if len(coordinates) < 2:
        raise ValueError(f"Route '{request.route_name}' has fewer than two distinct points.")

    result = RouteResult(request=request, coordinates=coordinates, polygon_note=polygon_note)
    route_code = safe_identifier(request.route_name, f"ROUTE_{route_number:02d}")
    designations = [f"{route_code}_DP{i + 1:04d}" for i in range(len(coordinates))]

    for index, (lon, lat) in enumerate(coordinates):
        if index == 0 or index == len(coordinates) - 1:
            angle_value: float | str = "n/a"
            category = "Corridor Connection Point"
            code: int | str = ""
            bucket = "Corridor Connection Point"
        else:
            angle_value = deflection_angle_deg(coordinates[index - 1], coordinates[index], coordinates[index + 1])
            category, code = classify_angle(angle_value, request.angle_splits)
            bucket = category
            if code == 1:
                result.tangent_dp_count += 1
            elif code == 2:
                result.light_count += 1
            elif code == 3:
                result.medium_count += 1
            else:
                result.heavy_count += 1
        result.designation_rows.append(
            {
                "Designation": designations[index],
                "Route": request.route_name,
                "Corridor Group": request.corridor_group,
                "Deflection Angle (degrees)": angle_value,
                "Longitude": lon,
                "Latitude": lat,
                "DE Category": category,
                "DE Code": code,
                "Summary Bucket": bucket,
                "Source Item": candidate.display_name,
                "Source Folder": candidate.folder_path,
            }
        )

    for index, (a, b) in enumerate(zip(coordinates, coordinates[1:])):
        length_m = haversine_m(a, b)
        tangents = tangent_count(length_m, request.span_m)
        result.total_length_m += length_m
        result.total_tangents += tangents
        result.segment_rows.append(
            {
                "Route": request.route_name,
                "Corridor Group": request.corridor_group,
                "Route Segment": f"{designations[index]}-{designations[index + 1]}",
                "Route Section": request.route_name,
                "Length (m)": length_m,
                "Design Span (m)": request.span_m,
                "Tangents": tangents,
                "Start Longitude": a[0],
                "Start Latitude": a[1],
                "End Longitude": b[0],
                "End Latitude": b[1],
                "Source Geometry": candidate.geometry_type,
            }
        )
    return result


def closest_distance_to_any(
    point: tuple[float, float],
    comparison_points: Sequence[tuple[float, float]],
) -> float:
    if not comparison_points:
        return 0.0
    return min(haversine_m(point, other) for other in comparison_points)


def exterior_endpoint_index(
    result: RouteResult,
    neighboring_endpoints: Sequence[tuple[float, float]],
) -> int:
    """Choose the endpoint farthest from the neighboring corridor group."""
    start_distance = closest_distance_to_any(result.coordinates[0], neighboring_endpoints)
    end_distance = closest_distance_to_any(result.coordinates[-1], neighboring_endpoints)
    return 0 if start_distance >= end_distance else -1


def set_endpoint_role(result: RouteResult, index: int, role: str) -> None:
    row = result.designation_rows[0 if index == 0 else -1]
    row["DE Category"] = role
    row["DE Code"] = ""
    row["Summary Bucket"] = role


def assign_endpoint_roles(results: Sequence[RouteResult]) -> None:
    """
    Count substations only at the two outside ends of the selected corridor.

    For example, with groups A through F, only the outside end of each A
    option and the outside end of each F option is a Substation A-Frame.
    Endpoints between corridor groups are Corridor Connection Points.
    """
    grouped: dict[str, list[RouteResult]] = {}
    for result in results:
        group = result.request.corridor_group.strip() or infer_corridor_group(result.request.route_name)
        group = group.upper()
        result.request.corridor_group = group
        grouped.setdefault(group, []).append(result)

    ordered_groups = sorted(grouped, key=natural_sort_key)
    for result in results:
        result.substation_count = 0
        result.connection_count = 2
        set_endpoint_role(result, 0, "Corridor Connection Point")
        set_endpoint_role(result, -1, "Corridor Connection Point")

    if not ordered_groups:
        return

    if len(ordered_groups) == 1:
        for result in grouped[ordered_groups[0]]:
            set_endpoint_role(result, 0, "Substation A-Frame")
            set_endpoint_role(result, -1, "Substation A-Frame")
            result.substation_count = 2
            result.connection_count = 0
        return

    first_group = grouped[ordered_groups[0]]
    second_group = grouped[ordered_groups[1]]
    second_endpoints = [
        point
        for result in second_group
        for point in (result.coordinates[0], result.coordinates[-1])
    ]
    for result in first_group:
        outside = exterior_endpoint_index(result, second_endpoints)
        set_endpoint_role(result, outside, "Substation A-Frame")
        result.substation_count = 1
        result.connection_count = 1

    last_group = grouped[ordered_groups[-1]]
    previous_group = grouped[ordered_groups[-2]]
    previous_endpoints = [
        point
        for result in previous_group
        for point in (result.coordinates[0], result.coordinates[-1])
    ]
    for result in last_group:
        outside = exterior_endpoint_index(result, previous_endpoints)
        set_endpoint_role(result, outside, "Substation A-Frame")
        result.substation_count = 1
        result.connection_count = 1


def analyze_requests(requests: Sequence[RouteRequest]) -> list[RouteResult]:
    results = [analyze_route(request, index + 1) for index, request in enumerate(requests)]
    assign_endpoint_roles(results)
    return results


@dataclass(frozen=True)
class CorridorJunction:
    upstream: RouteResult
    downstream: RouteResult
    upstream_end: int
    downstream_end: int
    gap_m: float
    angle_deg: float
    category: str
    code: int


def project_corridor_junctions(results: Sequence[RouteResult]) -> list[CorridorJunction]:
    """Project one junction for each option pair in neighboring corridor groups.

    Select the nearest non-substation endpoint pair, independent of KML direction.
    Compare the bearing arriving at the first end with the bearing departing
    the second end. A gap is reported, not filled with invented route geometry.
    Excel subsequently activates only the uniquely selected pair at each boundary.
    """
    grouped: dict[str, list[RouteResult]] = {}
    for result in results:
        grouped.setdefault(result.request.corridor_group, []).append(result)
    groups = sorted(grouped, key=natural_sort_key)
    junctions = []
    for left_group, right_group in zip(groups, groups[1:]):
        for left in grouped[left_group]:
            for right in grouped[right_group]:
                pairs = [
                    (haversine_m(left.coordinates[i], right.coordinates[j]), i, j)
                    for i in (0, -1) for j in (0, -1)
                    if left.designation_rows[i]["DE Category"] == "Corridor Connection Point"
                    and right.designation_rows[j]["DE Category"] == "Corridor Connection Point"
                ]
                if not pairs:
                    continue
                gap, i, j = min(pairs)
                left_inside = left.coordinates[1 if i == 0 else -2]
                right_inside = right.coordinates[1 if j == 0 else -2]
                incoming = (initial_bearing_deg(left.coordinates[i], left_inside) + 180.0) % 360.0
                outgoing = initial_bearing_deg(right.coordinates[j], right_inside)
                angle = abs((outgoing - incoming + 180.0) % 360.0 - 180.0)
                category, code = classify_angle(angle, left.request.angle_splits)
                junctions.append(CorridorJunction(left, right, i, j, gap, angle, category, code))
    return junctions


def excel_safe_text(value: object) -> object:
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        return "'" + value
    return value


def write_analysis_workbook(
    results: Sequence[RouteResult],
    output_path: Path,
    input_path: Path,
) -> None:
    try:
        import xlsxwriter
    except ImportError as exc:
        raise RuntimeError(
            "The 'xlsxwriter' package is required to create the interactive Excel report.\n\n"
            "Install it once from Command Prompt with:\n"
            "py -m pip install xlsxwriter"
        ) from exc
    if not hasattr(xlsxwriter.worksheet.Worksheet, "insert_checkbox"):
        raise RuntimeError(
            "A newer version of 'xlsxwriter' is needed for Excel checkboxes.\n\n"
            "Update it from Command Prompt with:\n"
            "py -m pip install --upgrade xlsxwriter"
        )

    ordered_results = sorted(
        results,
        key=lambda result: (
            natural_sort_key(result.request.corridor_group),
            natural_sort_key(result.request.route_name),
        ),
    )
    if not ordered_results:
        raise ValueError("No analyzed routes are available for the Excel report.")

    splits = ordered_results[0].request.angle_splits
    if any(r.request.angle_splits != splits for r in ordered_results):
        raise ValueError("All routes in one workbook must use the same angle splits.")
    ranges = {code: splits.range_label(code) for code in range(1, 5)}

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(output_path))
    workbook.set_calc_mode("auto")
    workbook.set_properties({
        "title": "KML / KMZ Route Analysis",
        "subject": "Interactive selected-route analysis",
        "comments": f"Generated from {input_path.name}",
    })

    navy = "#1F4E78"
    medium_blue = "#5B9BD5"
    light_blue = "#D9EAF7"
    pale_blue = "#EAF3F8"
    pale_green = "#E2F0D9"
    white = "#FFFFFF"
    gray = "#666666"
    light_border = "#D9E1F2"

    title_format = workbook.add_format({
        "font_name": "Aptos Display", "font_size": 17, "bold": True,
        "font_color": navy, "valign": "vcenter",
    })
    subtitle_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "italic": True,
        "font_color": gray,
    })
    instruction_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "font_color": "#404040",
        "bg_color": pale_blue, "text_wrap": True, "valign": "vcenter",
    })
    header_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "bold": True,
        "font_color": white, "bg_color": navy, "align": "center",
        "valign": "vcenter", "text_wrap": True,
        "bottom": 1, "bottom_color": light_border,
    })
    body_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "valign": "vcenter",
        "bottom": 1, "bottom_color": light_border,
    })
    text_wrap_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "valign": "top",
        "text_wrap": True, "bottom": 1, "bottom_color": light_border,
    })
    integer_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "num_format": "#,##0",
        "align": "right", "valign": "vcenter",
        "bottom": 1, "bottom_color": light_border,
    })
    decimal_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "num_format": "#,##0.00",
        "align": "right", "valign": "vcenter",
        "bottom": 1, "bottom_color": light_border,
    })
    coordinate_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "num_format": "0.00000000",
        "align": "right", "valign": "vcenter",
        "bottom": 1, "bottom_color": light_border,
    })
    checkbox_format = workbook.add_format({"align": "center", "valign": "vcenter"})
    selected_row_format = workbook.add_format({"bg_color": pale_green})
    summary_header_format = workbook.add_format({
        "font_name": "Aptos Display", "font_size": 12, "bold": True,
        "font_color": white, "bg_color": navy, "align": "left",
        "valign": "vcenter",
    })
    summary_label_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "bold": True,
        "bg_color": light_blue, "bottom": 1, "bottom_color": white,
    })
    summary_value_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 11, "bold": True,
        "bg_color": white, "align": "right", "num_format": "#,##0",
        "bottom": 1, "bottom_color": light_border,
    })
    summary_decimal_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 11, "bold": True,
        "bg_color": white, "align": "right", "num_format": "#,##0.00",
        "bottom": 1, "bottom_color": light_border,
    })
    note_format = workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "italic": True,
        "font_color": gray, "text_wrap": True, "valign": "top",
    })

    def write_headers(worksheet, row: int, headers: Sequence[str]) -> None:
        for column, value in enumerate(headers):
            worksheet.write(row, column, value, header_format)
        worksheet.set_row(row, 32)

    def finish_table(worksheet, header_row: int, last_row: int, last_col: int) -> None:
        worksheet.freeze_panes(header_row + 1, 0)
        if last_row > header_row:
            worksheet.autofilter(header_row, 0, last_row, last_col)
        worksheet.hide_gridlines(2)

    route_colors = {
        result.request.route_name: "#" + ROUTE_COLORS[index % len(ROUTE_COLORS)]
        for index, result in enumerate(ordered_results)
    }

    # Default to one complete alignment: the first natural option in each corridor group.
    default_selected: set[str] = set()
    seen_groups: set[str] = set()
    for result in ordered_results:
        group = result.request.corridor_group
        if group not in seen_groups:
            seen_groups.add(group)
            default_selected.add(result.request.route_name)

    dashboard = workbook.add_worksheet("Route Dashboard")
    dashboard.hide_gridlines(2)
    dashboard.set_zoom(85)
    dashboard.set_tab_color(medium_blue)
    dashboard.merge_range("A1:M1", "Route Selection Dashboard", title_format)
    dashboard.set_row(0, 27)
    dashboard.merge_range(
        "A2:M2",
        "Check the route options to include. The map and structure totals update from the checked options only. "
        "For one continuous alignment, select one option from each corridor group.",
        instruction_format,
    )
    dashboard.set_row(1, 34)
    dashboard.write("A3", f"Source: {input_path.name}", subtitle_format)
    dashboard.write("D3", f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subtitle_format)

    dashboard_headers = [
        "Route Option", "Select", "Corridor", "Span (m)", "Distance (km)",
        "Span Tangents", f"Tangent DP\n{ranges[1]}", f"Light Angle\n{ranges[2]}",
        f"Medium Angle\n{ranges[3]}", f"Heavy Angle\n{ranges[4]}", "Terminal /\nDead-End",
        "Connection\nEnds", "DP Total",
    ]
    dashboard_header_row = 3
    dashboard_first_row = dashboard_header_row + 1
    dashboard_last_row = dashboard_first_row + len(ordered_results) - 1
    write_headers(dashboard, dashboard_header_row, dashboard_headers)
    dashboard.set_column("A:A", 18)
    dashboard.set_column("B:B", 9)
    dashboard.set_column("C:C", 11)
    dashboard.set_column("D:E", 13)
    dashboard.set_column("F:M", 14)
    dashboard.set_column("N:N", 2.5)
    dashboard.set_column("O:O", 29)
    dashboard.set_column("P:P", 17)

    dashboard_row_by_route: dict[str, int] = {}
    for index, result in enumerate(ordered_results):
        row = dashboard_first_row + index
        request = result.request
        selected = request.route_name in default_selected
        dashboard_row_by_route[request.route_name] = row
        option_format = workbook.add_format({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": route_colors[request.route_name], "valign": "vcenter",
            "bottom": 1, "bottom_color": light_border,
        })
        dashboard.write(row, 0, excel_safe_text(request.route_name), option_format)
        dashboard.insert_checkbox(row, 1, selected, checkbox_format)
        dashboard.write(row, 2, excel_safe_text(request.corridor_group), body_format)
        dashboard.write_number(row, 3, request.span_m, decimal_format)
        dashboard.write_number(row, 4, result.total_length_m / 1000.0, decimal_format)
        dashboard.write_number(row, 5, result.total_tangents, integer_format)
        dashboard.write_number(row, 6, result.tangent_dp_count, integer_format)
        dashboard.write_number(row, 7, result.light_count, integer_format)
        dashboard.write_number(row, 8, result.medium_count, integer_format)
        dashboard.write_number(row, 9, result.heavy_count, integer_format)
        dashboard.write_number(row, 10, result.substation_count, integer_format)
        dashboard.write_number(row, 11, result.connection_count, integer_format)
        dashboard.write_number(row, 12, len(result.designation_rows), integer_format)
        dashboard.set_row(row, 24)

    dashboard.conditional_format(
        dashboard_first_row, 0, dashboard_last_row, len(dashboard_headers) - 1,
        {
            "type": "formula",
            "criteria": f"=$B{dashboard_first_row + 1}=TRUE",
            "format": selected_row_format,
        },
    )
    dashboard.autofilter(
        dashboard_header_row, 0, dashboard_last_row, len(dashboard_headers) - 1
    )
    dashboard.freeze_panes(dashboard_first_row, 0)

    first_excel_row = dashboard_first_row + 1
    last_excel_row = dashboard_last_row + 1
    selected_range = f"$B${first_excel_row}:$B${last_excel_row}"

    def sum_formula(column_letter: str) -> str:
        return (
            f"=SUMPRODUCT(--{selected_range},"
            f"${column_letter}${first_excel_row}:${column_letter}${last_excel_row})"
        )

    # Pair-dependent junctions remain separate from the internal per-route counts.
    junctions = project_corridor_junctions(ordered_results)
    junction_ws = workbook.add_worksheet("Corridor Junctions")
    junction_ws.merge_range("A1:N1", "Projected Corridor Junction Angles", title_format)
    junction_ws.merge_range("A2:N2",
        "One projected structure per active pair. Select exactly one option in each adjoining group. "
        "Endpoint gaps are reported; no connecting spans are added. Filter Active = 1 to see the current alignment.",
        instruction_format)
    junction_ws.set_row(1, 34)
    junction_headers = ["From Route", "To Route", "From Group", "To Group",
                        "From Endpoint", "To Endpoint", "Endpoint Gap (m)",
                        "Projected Angle (degrees)", "Category", "Code", "Active",
                        "From Selected", "To Selected", "Angle Range"]
    write_headers(junction_ws, 3, junction_headers)
    group_rows: dict[str, list[int]] = {}
    for result in ordered_results:
        group_rows.setdefault(result.request.corridor_group, []).append(
            dashboard_row_by_route[result.request.route_name] + 1)

    def group_count(group: str) -> str:
        rows = group_rows[group]
        return f"COUNTIF('Route Dashboard'!$B${min(rows)}:$B${max(rows)},TRUE)"

    junction_cached = {code: 0 for code in range(1, 5)}
    for row, junction in enumerate(junctions, 4):
        left = junction.upstream.request
        right = junction.downstream.request
        left_selected = left.route_name in default_selected
        right_selected = right.route_name in default_selected
        active = int(left_selected and right_selected)
        junction_cached[junction.code] += active
        values = [left.route_name, right.route_name, left.corridor_group, right.corridor_group,
                  junction.upstream.designation_rows[junction.upstream_end]["Designation"],
                  junction.downstream.designation_rows[junction.downstream_end]["Designation"],
                  junction.gap_m, junction.angle_deg, junction.category, junction.code]
        for column, value in enumerate(values):
            junction_ws.write(row, column, excel_safe_text(value),
                              decimal_format if column in (6, 7) else body_format)
        erow = row + 1
        junction_ws.write_formula(row, 10,
            f"=--AND(L{erow},M{erow},{group_count(left.corridor_group)}=1,{group_count(right.corridor_group)}=1)",
            integer_format, active)
        for col, request, checked in [(11, left, left_selected), (12, right, right_selected)]:
            drow = dashboard_row_by_route[request.route_name] + 1
            junction_ws.write_formula(row, col, f"='Route Dashboard'!$B${drow}", body_format, checked)
        junction_ws.write(row, 13, ranges[junction.code], body_format)
    junction_ws.set_column("A:F", 22)
    junction_ws.set_column("G:H", 22)
    junction_ws.set_column("I:I", 18)
    junction_ws.set_column("J:M", 13)
    junction_ws.set_column("N:N", 18)
    finish_table(junction_ws, 3, 3 + len(junctions), 13)

    def junction_sum(code: int | None = None) -> str:
        if not junctions:
            return "0"
        end = 4 + len(junctions)
        active_range = f"'Corridor Junctions'!$K$5:$K${end}"
        if code is None:
            return f"SUM({active_range})"
        return f"SUMIF('Corridor Junctions'!$J$5:$J${end},{code},{active_range})"

    junction_total = sum(junction_cached.values())

    selected_results = [
        result for result in ordered_results
        if result.request.route_name in default_selected
    ]
    summary_items = [
        ("Selected Route Options", f'=SUMPRODUCT(--{selected_range})', len(selected_results), summary_value_format),
        ("Total Distance (km)", sum_formula("E"), sum(r.total_length_m for r in selected_results) / 1000.0, summary_decimal_format),
        ("Span Tangents", sum_formula("F"), sum(r.total_tangents for r in selected_results), summary_value_format),
        (f"Tangent DPs ({ranges[1]})", sum_formula("G") + "+" + junction_sum(1), sum(r.tangent_dp_count for r in selected_results) + junction_cached[1], summary_value_format),
        (f"Light Angle DPs ({ranges[2]})", sum_formula("H") + "+" + junction_sum(2), sum(r.light_count for r in selected_results) + junction_cached[2], summary_value_format),
        (f"Medium Angle DPs ({ranges[3]})", sum_formula("I") + "+" + junction_sum(3), sum(r.medium_count for r in selected_results) + junction_cached[3], summary_value_format),
        (f"Heavy Angle DPs ({ranges[4]})", sum_formula("J") + "+" + junction_sum(4), sum(r.heavy_count for r in selected_results) + junction_cached[4], summary_value_format),
        (
            "Total Angle DPs",
            f"={sum_formula('H')[1:]}+{sum_formula('I')[1:]}+{sum_formula('J')[1:]}+{junction_sum(2)}+{junction_sum(3)}+{junction_sum(4)}",
            sum(r.light_count + r.medium_count + r.heavy_count for r in selected_results) + sum(junction_cached[c] for c in (2, 3, 4)),
            summary_value_format,
        ),
        ("Terminal / Dead-End Structures", sum_formula("K"), sum(r.substation_count for r in selected_results), summary_value_format),
        ("Corridor Connection Ends", sum_formula("L"), sum(r.connection_count for r in selected_results), summary_value_format),
        ("Projected Junction Structures", "=" + junction_sum(), junction_total, summary_value_format),
        ("Designation Points", sum_formula("M"), sum(len(r.designation_rows) for r in selected_results), summary_value_format),
        (
            "Estimated Total Structures",
            f"={sum_formula('F')[1:]}+{sum_formula('G')[1:]}+{sum_formula('H')[1:]}+{sum_formula('I')[1:]}+{sum_formula('J')[1:]}+{sum_formula('K')[1:]}+{junction_sum()}",
            sum(
                r.total_tangents + r.tangent_dp_count + r.light_count + r.medium_count
                + r.heavy_count + r.substation_count
                for r in selected_results
            ) + junction_total,
            summary_value_format,
        ),
    ]
    summary_start_row = 3
    dashboard.merge_range(summary_start_row, 14, summary_start_row, 15, "Selected Route Summary", summary_header_format)
    dashboard.set_row(summary_start_row, 24)
    for offset, (label, formula, cached_value, value_format) in enumerate(summary_items, 1):
        row = summary_start_row + offset
        dashboard.write(row, 14, label, summary_label_format)
        dashboard.write_formula(row, 15, formula, value_format, cached_value)
        dashboard.set_row(row, 22)
    dashboard.merge_range(
        summary_start_row + len(summary_items) + 2, 14,
        summary_start_row + len(summary_items) + 4, 15,
        "Category totals include one projected structure per active junction; route rows show internal points only. "
        "Connection Ends and Designation Points retain raw endpoint counts. See Corridor Junctions for angles and gaps.",
        note_format,
    )

    selection_checks = [f"{group_count(group)}<>1" for group in group_rows]
    selection_warning = '"Select one option per group; missing/ambiguous boundaries are excluded from junction totals."'
    selection_ok = '"Junction projection active: one option selected per group."'
    dashboard.merge_range("O23:P25", "", instruction_format)
    dashboard.write_formula("O23", f"=IF(OR({','.join(selection_checks)}),{selection_warning},{selection_ok})",
                            instruction_format, "Junction projection active: one option selected per group.")

    chart_data = workbook.add_worksheet("_Chart Data")
    chart_data_headers = [
        "Route Option", "Corridor", "Sequence", "Longitude", "Latitude",
        "Selected", "Plot Longitude", "Plot Latitude", "Point Type",
    ]
    write_headers(chart_data, 0, chart_data_headers)
    chart_data_row = 1
    chart_blocks: list[tuple[RouteResult, int, int, str]] = []
    for result in ordered_results:
        start_row = chart_data_row
        dashboard_row = dashboard_row_by_route[result.request.route_name]
        dashboard_excel_row = dashboard_row + 1
        selected = result.request.route_name in default_selected
        for sequence, point in enumerate(result.coordinates, 1):
            if sequence == 1:
                point_type = str(result.designation_rows[0]["DE Category"])
            elif sequence == len(result.coordinates):
                point_type = str(result.designation_rows[-1]["DE Category"])
            else:
                point_type = "Internal Designation Point"
            chart_data.write(chart_data_row, 0, excel_safe_text(result.request.route_name), body_format)
            chart_data.write(chart_data_row, 1, excel_safe_text(result.request.corridor_group), body_format)
            chart_data.write_number(chart_data_row, 2, sequence, integer_format)
            chart_data.write_number(chart_data_row, 3, point[0], coordinate_format)
            chart_data.write_number(chart_data_row, 4, point[1], coordinate_format)
            chart_data.write_formula(
                chart_data_row, 5,
                f"='Route Dashboard'!$B${dashboard_excel_row}",
                body_format, selected,
            )
            excel_row = chart_data_row + 1
            cached_x: object = point[0] if selected else "#N/A"
            cached_y: object = point[1] if selected else "#N/A"
            chart_data.write_formula(
                chart_data_row, 6, f"=IF(F{excel_row},D{excel_row},NA())",
                coordinate_format, cached_x,
            )
            chart_data.write_formula(
                chart_data_row, 7, f"=IF(F{excel_row},E{excel_row},NA())",
                coordinate_format, cached_y,
            )
            chart_data.write(chart_data_row, 8, point_type, body_format)
            chart_data_row += 1
        chart_blocks.append((
            result,
            start_row,
            chart_data_row - 1,
            route_colors[result.request.route_name],
        ))

    chart = workbook.add_chart({"type": "scatter", "subtype": "straight"})
    for result, start_row, end_row, color in chart_blocks:
        chart.add_series({
            "name": result.request.route_name,
            "categories": ["_Chart Data", start_row, 6, end_row, 6],
            "values": ["_Chart Data", start_row, 7, end_row, 7],
            "line": {"color": color, "width": 1.75},
            "marker": {"type": "none"},
        })
    chart.set_title({"name": "Selected Route Options", "name_font": {"size": 12, "bold": True}})
    chart.set_legend({"none": True})
    chart.set_x_axis({
        "name": "", "label_position": "none", "major_gridlines": {"visible": False},
        "line": {"none": True}, "major_tick_mark": "none", "minor_tick_mark": "none",
    })
    chart.set_y_axis({
        "name": "", "label_position": "none", "major_gridlines": {"visible": False},
        "line": {"none": True}, "major_tick_mark": "none", "minor_tick_mark": "none",
    })
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": white}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": white}})
    chart.set_size({"width": 770, "height": 500})
    dashboard.insert_chart("O27", chart, {"x_offset": 4, "y_offset": 4})
    chart_data.hide()

    summary_ws = workbook.add_worksheet("Route Summary")
    summary_ws.hide_gridlines(2)
    summary_ws.merge_range("A1:Q1", "KML / KMZ Route Analysis", title_format)
    summary_ws.write("A2", f"Source: {input_path.name}", subtitle_format)
    summary_ws.write("A3", f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subtitle_format)
    summary_headers = [
        "Route Option", "Corridor Group", "Source Item", "Folder", "Geometry", "Design Span (m)",
        "Total Length (m)", "Total Distance (km)", "Span Tangents",
        "DP Total", f"Tangent DP ({ranges[1]})", f"Light Angle ({ranges[2]})",
        f"Medium Angle ({ranges[3]})", f"Heavy Angle ({ranges[4]})",
        "Substation A-Frame", "Corridor Connections",
        "Polygon Centerline Note",
    ]
    write_headers(summary_ws, 4, summary_headers)
    for row_index, result in enumerate(ordered_results, 5):
        request = result.request
        values = [
            request.route_name,
            request.corridor_group,
            request.candidate.display_name,
            request.candidate.folder_path,
            request.candidate.geometry_type,
            request.span_m,
            result.total_length_m,
            result.total_length_m / 1000.0,
            result.total_tangents,
            len(result.designation_rows),
            result.tangent_dp_count,
            result.light_count,
            result.medium_count,
            result.heavy_count,
            result.substation_count,
            result.connection_count,
            result.polygon_note,
        ]
        for column, value in enumerate(values):
            if column in (5, 6, 7):
                summary_ws.write_number(row_index, column, float(value), decimal_format)
            elif column in range(8, 16):
                summary_ws.write_number(row_index, column, int(value), integer_format)
            elif column == 16:
                summary_ws.write(row_index, column, excel_safe_text(value), text_wrap_format)
            else:
                summary_ws.write(row_index, column, excel_safe_text(value), body_format)
    summary_widths = [18, 14, 26, 25, 14, 14, 16, 17, 14, 11, 19, 20, 22, 19, 20, 20, 52]
    for column, width in enumerate(summary_widths):
        summary_ws.set_column(column, column, width)
    finish_table(summary_ws, 4, 4 + len(ordered_results), len(summary_headers) - 1)

    segment_ws = workbook.add_worksheet("Route Segments")
    segment_headers = [
        "Route", "Corridor Group", "Route Segment", "Route Section", "Length (m)", "Design Span (m)",
        "Tangents", "Start Longitude", "Start Latitude", "End Longitude", "End Latitude",
        "Source Geometry",
    ]
    write_headers(segment_ws, 0, segment_headers)
    segment_row = 1
    for result in ordered_results:
        for record in result.segment_rows:
            for column, key in enumerate(segment_headers):
                value = record[key]
                if column in (4, 5):
                    segment_ws.write_number(segment_row, column, float(value), decimal_format)
                elif column == 6:
                    segment_ws.write_number(segment_row, column, int(value), integer_format)
                elif column in (7, 8, 9, 10):
                    segment_ws.write_number(segment_row, column, float(value), coordinate_format)
                else:
                    segment_ws.write(segment_row, column, excel_safe_text(value), body_format)
            segment_row += 1
    segment_widths = [18, 14, 42, 18, 14, 14, 11, 16, 15, 16, 15, 17]
    for column, width in enumerate(segment_widths):
        segment_ws.set_column(column, column, width)
    finish_table(segment_ws, 0, segment_row - 1, len(segment_headers) - 1)

    dp_ws = workbook.add_worksheet("Designation Points")
    dp_headers = [
        "Designation", "Route", "Corridor Group", "Deflection Angle (degrees)", "Longitude", "Latitude",
        "DE Category", "DE Code", "Summary Bucket", "Source Item", "Source Folder",
    ]
    write_headers(dp_ws, 0, dp_headers)
    dp_row = 1
    for result in ordered_results:
        for record in result.designation_rows:
            for column, key in enumerate(dp_headers):
                value = record[key]
                if column == 3 and isinstance(value, (int, float)):
                    dp_ws.write_number(dp_row, column, float(value), decimal_format)
                elif column in (4, 5):
                    dp_ws.write_number(dp_row, column, float(value), coordinate_format)
                elif column == 7 and isinstance(value, (int, float)):
                    dp_ws.write_number(dp_row, column, int(value), integer_format)
                else:
                    dp_ws.write(dp_row, column, excel_safe_text(value), body_format)
            dp_row += 1
    dp_widths = [25, 18, 14, 23, 16, 15, 25, 10, 25, 26, 28]
    for column, width in enumerate(dp_widths):
        dp_ws.set_column(column, column, width)
    finish_table(dp_ws, 0, dp_row - 1, len(dp_headers) - 1)

    notes_ws = workbook.add_worksheet("Analysis Notes")
    notes_ws.hide_gridlines(2)
    notes_ws.merge_range("A1:B1", "Analysis Rules and Assumptions", title_format)
    notes = [
        ("Source file", input_path.name),
        ("Route detection", "Every KML LineString is offered and preselected. Every Polygon is also offered; polygons with elongation ratio >= 4.0 are suggested."),
        ("Interactive dashboard", "The Route Dashboard contains native Excel checkboxes. Only checked options contribute to the map and selected-route summary."),
        ("Default selection", "The first naturally sorted option in each corridor group is checked when the workbook is created. Change the checkboxes to analyze another alignment."),
        ("Corridor alternatives", "Names such as A1, A2, and A3 are automatically grouped as alternatives for corridor A. The inferred group can be edited before analysis. Every included option receives its own checkbox and chart series."),
        ("Corridor order", "Corridor groups are ordered naturally by their group labels, so A precedes B and E precedes F."),
        ("Polygon centerline", "The polygon boundary is split at its estimated two endcaps. The two opposing boundary paths are resampled and averaged. Review unusual, branching, wide, self-intersecting, or non-corridor polygons visually."),
        ("Distance", "Segment distances are great-circle distances using Earth mean radius 6,371,008.8 m."),
        ("Deflection", "Magnitude of the direction change between the incoming and outgoing geodesic bearings at each internal route vertex."),
        ("Designation points", "Every non-duplicate vertex in the selected LineString or estimated polygon centerline is analyzed as a designation point."),
        ("Tangent DP", ranges[1]),
        ("Light Angle", ranges[2]),
        ("Medium Angle", ranges[3]),
        ("Heavy Angle", ranges[4]),
        ("Endpoints", "Only the outside endpoint of each option in the first and last selected corridor groups is a Substation A-Frame. Other option endpoints are Corridor Connection Points. The outside end is inferred as the endpoint farthest from the adjacent corridor group."),
        ("Span tangents", "Structures inserted along a segment from the integer part of Segment Length / Design Span. This is separate from a Tangent DP, which is an analyzed route vertex within the user-selected tangent angle range. A segment exactly equal to one design span is assigned zero, matching the source VBA rule (e.g., 150->0, 250->0, 251->1, 500->2 for a 250 m span)."),
        ("Deduplication", "Consecutive duplicate geometry points within 0.01 m are removed before analysis."),
        ("Route sections", "Each selected KML/KMZ item is analyzed independently as one route section. Pair-dependent junction projections are added only in the dashboard summary."),
        ("Projected junction angles", "For each pair of adjacent corridor groups in natural order, match the closest non-substation endpoints. Compare the local arriving bearing of the first route with the departing bearing of the next route, accounting for reversed KML geometry. Apply the chosen angle splits; count one projected structure per active pair, including a tangent junction."),
        ("Junction selection", "A pair is active only when both routes are checked and exactly one option is checked in each of their groups. Missing or multiple selections suppress junctions at the affected boundaries; no connection is inferred across an unchecked group. Internal route totals still follow all checked options."),
        ("Projection limits", "Endpoint gaps are shown on Corridor Junctions. The angle compares endpoint directions without adding gap length, connecting spans, or two separate tie-in structures. Review gaps and unusual corridor geometry before treating the estimate as a design count."),
        ("Junction counting", "Selected category totals and Estimated Total Structures include active projected junctions once. Total Angle DPs includes only non-tangent junctions. Raw Corridor Connection Ends, Designation Points, per-route counts, and source endpoint angle fields remain unchanged; pair-dependent angles are on Corridor Junctions."),
    ]
    write_headers(notes_ws, 2, ["Item", "Definition"])
    for row_index, (label, definition) in enumerate(notes, 3):
        notes_ws.write(row_index, 0, label, body_format)
        notes_ws.write(row_index, 1, definition, text_wrap_format)
        notes_ws.set_row(row_index, 38)
    notes_ws.set_column("A:A", 25)
    notes_ws.set_column("B:B", 110)
    finish_table(notes_ws, 2, 2 + len(notes), 1)

    try:
        workbook.close()
    except Exception:
        # Ensure a partially written workbook is not silently reported as complete.
        raise


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.inner.bind("<Configure>", self._configure_scroll_region)
        self.canvas.bind("<Configure>", self._configure_width)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

    def _configure_scroll_region(self, _event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _configure_width(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self.window, width=event.width)


class RouteSelectionDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, candidates: Sequence[GeometryCandidate]) -> None:
        super().__init__(parent)
        self.title("Select Routes and Enter Design Spans")
        self.geometry("1280x720")
        self.minsize(1050, 560)
        self.transient(parent)
        self.result: list[RouteRequest] | None = None
        self.candidates = list(candidates)
        self.rows: list[dict[str, object]] = []
        self.default_span = tk.StringVar(value=f"{DEFAULT_SPAN_M:g}")
        self.angle_split_text = tk.StringVar(value="2, 15, 45")
        self.angle_preview = tk.StringVar(value="")
        self.status = tk.StringVar(value="")
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.update_idletasks()
        self._center_over_parent(parent)
        self.lift()
        self.focus_force()
        self.grab_set()
        # A short topmost pulse prevents the modal window from opening behind
        # IDLE, VS Code, Spyder, or the file picker on Windows.
        try:
            self.attributes("-topmost", True)
            self.after(700, lambda: self.attributes("-topmost", False))
        except tk.TclError:
            pass

    def _center_over_parent(self, parent: tk.Misc) -> None:
        width = max(self.winfo_reqwidth(), 1050)
        height = max(self.winfo_reqheight(), 560)
        parent.update_idletasks()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - height) // 2)
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = min(max(0, x), max(0, screen_width - width))
        y = min(max(0, y), max(0, screen_height - height))
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="Choose what counts as a route, then enter the design span for each selected route.",
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            outer,
            text=(
                "All LineStrings are preselected. Long/narrow polygons are suggested and will be converted "
                "to an estimated centerline. Names such as A1, A2, and A3 are automatically grouped under "
                "corridor A. Uncheck anything that is not a route and correct the group if needed."
            ),
            wraplength=1100,
        ).pack(anchor="w", pady=(3, 10))

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(0, 8))
        ttk.Label(controls, text="Default span (m):").pack(side="left")
        ttk.Entry(controls, textvariable=self.default_span, width=12).pack(side="left", padx=(6, 4))
        ttk.Button(controls, text="Apply to Selected", command=self.apply_default_span).pack(side="left", padx=(0, 14))
        ttk.Button(controls, text="Select All Lines", command=self.select_lines).pack(side="left", padx=3)
        ttk.Button(controls, text="Select Suggested Polygons", command=self.select_suggested_polygons).pack(side="left", padx=3)
        ttk.Button(controls, text="Clear Selection", command=self.clear_selection).pack(side="left", padx=3)

        angle_controls = ttk.LabelFrame(outer, text="Angle splits (degrees) — applies to all routes", padding=8)
        angle_controls.pack(fill="x", pady=(0, 8))
        ttk.Label(angle_controls, text="Cutoffs:").pack(side="left")
        ttk.Entry(angle_controls, textvariable=self.angle_split_text, width=20).pack(side="left", padx=6)
        for preset in ("3, 30", "15, 45", "2, 15, 45"):
            ttk.Button(angle_controls, text=preset,
                       command=lambda value=preset: self.angle_split_text.set(value)).pack(side="left", padx=3)
        ttk.Label(outer, text="Enter 2 cutoffs for Tangent / Medium / Heavy, or 3 to include Light. Exact cutoffs enter the higher band.").pack(anchor="w")
        ttk.Label(outer, textvariable=self.angle_preview, wraplength=1000).pack(anchor="w", pady=(2, 6))
        self.angle_split_text.trace_add("write", self.update_angle_preview)
        self.update_angle_preview()

        header = ttk.Frame(outer, padding=(4, 5))
        header.pack(fill="x")
        header.columnconfigure(2, weight=1)
        header.columnconfigure(3, weight=1)
        widths = (7, 14, 25, 30, 13, 10, 15, 14)
        labels = (
            "Route?", "Type", "KML item / folder", "Output route name",
            "Corridor group", "Points", "Est. length", "Span (m)",
        )
        for column, (label, width) in enumerate(zip(labels, widths)):
            ttk.Label(header, text=label, width=width, font=("Segoe UI", 9, "bold")).grid(row=0, column=column, sticky="w", padx=3)

        scroll = ScrollableFrame(outer)
        scroll.pack(fill="both", expand=True)
        scroll.inner.columnconfigure(2, weight=1)
        scroll.inner.columnconfigure(3, weight=1)

        used_names: dict[str, int] = {}
        for row_index, candidate in enumerate(self.candidates):
            base_name = candidate.display_name.strip() or candidate.candidate_id
            used_names[base_name] = used_names.get(base_name, 0) + 1
            unique_name = base_name if used_names[base_name] == 1 else f"{base_name} {used_names[base_name]}"
            include_var = tk.BooleanVar(value=candidate.suggested)
            name_var = tk.StringVar(value=unique_name)
            group_var = tk.StringVar(value=infer_corridor_group(unique_name))
            span_var = tk.StringVar(value=f"{DEFAULT_SPAN_M:g}")
            row_frame = ttk.Frame(scroll.inner, padding=(4, 3))
            row_frame.grid(row=row_index, column=0, columnspan=8, sticky="ew")
            row_frame.columnconfigure(2, weight=1)
            row_frame.columnconfigure(3, weight=1)
            ttk.Checkbutton(row_frame, variable=include_var).grid(row=0, column=0, sticky="w", padx=3)
            kind_text = "Polygon -> centerline" if candidate.geometry_type == "Polygon" else "LineString"
            if candidate.geometry_type == "Polygon" and candidate.suggested:
                kind_text += " *"
            ttk.Label(row_frame, text=kind_text, width=22).grid(row=0, column=1, sticky="w", padx=3)
            item_text = candidate.display_name
            if candidate.folder_path:
                item_text += f"\n{candidate.folder_path}"
            ttk.Label(row_frame, text=item_text, wraplength=300).grid(row=0, column=2, sticky="ew", padx=3)
            ttk.Entry(row_frame, textvariable=name_var, width=30).grid(row=0, column=3, sticky="ew", padx=3)
            ttk.Entry(row_frame, textvariable=group_var, width=13).grid(row=0, column=4, sticky="w", padx=3)
            ttk.Label(row_frame, text=str(len(candidate.coordinates)), width=10).grid(row=0, column=5, sticky="w", padx=3)
            estimate = f"{candidate.estimated_length_m / 1000.0:.2f} km"
            if candidate.geometry_type == "Polygon" and candidate.polygon_ratio is not None:
                estimate += f" (R {candidate.polygon_ratio:.1f})"
            ttk.Label(row_frame, text=estimate, width=18).grid(row=0, column=6, sticky="w", padx=3)
            ttk.Entry(row_frame, textvariable=span_var, width=12).grid(row=0, column=7, sticky="w", padx=3)
            if row_index % 2:
                try:
                    row_frame.configure(style="Alternate.TFrame")
                except tk.TclError:
                    pass
            self.rows.append({
                "candidate": candidate,
                "include": include_var,
                "name": name_var,
                "group": group_var,
                "span": span_var,
            })

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(10, 0))
        ttk.Label(bottom, textvariable=self.status, foreground="#A00000").pack(side="left")
        ttk.Button(bottom, text="Cancel", command=self.cancel).pack(side="right", padx=(6, 0))
        ttk.Button(bottom, text="Create Analysis", command=self.accept).pack(side="right")

    def update_angle_preview(self, *_args: object) -> None:
        try:
            self.angle_preview.set(AngleSplits.parse(self.angle_split_text.get()).preview())
        except ValueError as exc:
            self.angle_preview.set(str(exc))

    def apply_default_span(self) -> None:
        try:
            value = float(self.default_span.get())
            if not math.isfinite(value) or value <= 0:
                raise ValueError
        except ValueError:
            self.status.set("Enter a valid positive default span.")
            return
        for row in self.rows:
            if bool(row["include"].get()):  # type: ignore[union-attr]
                row["span"].set(f"{value:g}")  # type: ignore[union-attr]
        self.status.set("")

    def select_lines(self) -> None:
        for row in self.rows:
            candidate: GeometryCandidate = row["candidate"]  # type: ignore[assignment]
            if candidate.geometry_type == "LineString":
                row["include"].set(True)  # type: ignore[union-attr]

    def select_suggested_polygons(self) -> None:
        for row in self.rows:
            candidate: GeometryCandidate = row["candidate"]  # type: ignore[assignment]
            if candidate.geometry_type == "Polygon" and candidate.suggested:
                row["include"].set(True)  # type: ignore[union-attr]

    def clear_selection(self) -> None:
        for row in self.rows:
            row["include"].set(False)  # type: ignore[union-attr]

    def accept(self) -> None:
        try:
            splits = AngleSplits.parse(self.angle_split_text.get())
        except ValueError as exc:
            self.status.set(str(exc))
            return
        requests: list[RouteRequest] = []
        names_seen: set[str] = set()
        for row in self.rows:
            if not bool(row["include"].get()):  # type: ignore[union-attr]
                continue
            candidate: GeometryCandidate = row["candidate"]  # type: ignore[assignment]
            route_name = str(row["name"].get()).strip()  # type: ignore[union-attr]
            if not route_name:
                self.status.set(f"Enter an output route name for '{candidate.display_name}'.")
                return
            normalized = route_name.casefold()
            if normalized in names_seen:
                self.status.set(f"Route names must be unique. '{route_name}' appears more than once.")
                return
            names_seen.add(normalized)
            corridor_group = str(row["group"].get()).strip()  # type: ignore[union-attr]
            if not corridor_group:
                self.status.set(f"Enter a corridor group for '{route_name}'.")
                return
            corridor_group = corridor_group.upper()
            try:
                span_m = float(str(row["span"].get()).replace(",", ""))  # type: ignore[union-attr]
                if not math.isfinite(span_m) or span_m <= 0:
                    raise ValueError
            except ValueError:
                self.status.set(f"Enter a valid positive span for '{route_name}'.")
                return
            requests.append(
                RouteRequest(
                    candidate=candidate,
                    route_name=route_name,
                    span_m=span_m,
                    corridor_group=corridor_group,
                    angle_splits=splits,
                )
            )
        if not requests:
            self.status.set("Select at least one item to analyze as a route.")
            return
        self.result = requests
        self.destroy()

    def cancel(self) -> None:
        self.result = None
        self.destroy()


def choose_input_file(root: tk.Tk) -> Path | None:
    selected = filedialog.askopenfilename(
        parent=root,
        title="Select KML or KMZ",
        filetypes=[("KML / KMZ", "*.kml *.kmz"), ("KML", "*.kml"), ("KMZ", "*.kmz"), ("All files", "*.*")],
    )
    return Path(selected) if selected else None


def choose_output_file(root: tk.Tk, input_path: Path) -> Path | None:
    default_name = f"{input_path.stem}_Route_Analysis.xlsx"
    selected = filedialog.asksaveasfilename(
        parent=root,
        title="Save Route Analysis Workbook",
        initialdir=str(input_path.parent),
        initialfile=default_name,
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
    )
    return Path(selected) if selected else None


def build_status_window(root: tk.Tk) -> tk.StringVar:
    root.title("KML / KMZ Route Analysis")
    root.geometry("520x170")
    root.resizable(False, False)
    outer = ttk.Frame(root, padding=22)
    outer.pack(fill="both", expand=True)
    ttk.Label(
        outer,
        text="KML / KMZ Route Analysis",
        font=("Segoe UI", 14, "bold"),
    ).pack(anchor="w")
    status = tk.StringVar(value="Select a KML or KMZ file to begin.")
    ttk.Label(outer, textvariable=status, wraplength=470).pack(anchor="w", pady=(12, 8))
    progress = ttk.Progressbar(outer, mode="indeterminate", length=465)
    progress.pack(fill="x")
    progress.start(12)
    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = max(0, (screen_width - 520) // 2)
    y = max(0, (screen_height - 170) // 2)
    root.geometry(f"520x170+{x}+{y}")
    return status


def show_root(root: tk.Tk) -> None:
    root.deiconify()
    root.lift()
    try:
        root.attributes("-topmost", True)
        root.after(500, lambda: root.attributes("-topmost", False))
    except tk.TclError:
        pass
    root.update()


def run_gui() -> None:
    root = tk.Tk()
    status = build_status_window(root)
    root.withdraw()
    try:
        input_path = choose_input_file(root)
        if input_path is None:
            return

        status.set(f"Reading {input_path.name} and finding route candidates...")
        show_root(root)
        candidates = extract_candidates(input_path)
        if not candidates:
            messagebox.showwarning(
                "No Route Candidates",
                "No LineString, gx:Track, or Polygon geometries were found in the selected file.",
                parent=root,
            )
            return

        line_count = sum(candidate.geometry_type == "LineString" for candidate in candidates)
        polygon_count = sum(candidate.geometry_type == "Polygon" for candidate in candidates)
        status.set(
            f"Found {line_count} line(s) and {polygon_count} polygon(s). "
            "Opening the route-selection window..."
        )
        root.update()
        dialog = RouteSelectionDialog(root, candidates)
        root.wait_window(dialog)
        requests = dialog.result
        if not requests:
            return

        status.set(f"{len(requests)} route(s) selected. Choose where to save the Excel analysis.")
        show_root(root)
        output_path = choose_output_file(root, input_path)
        if output_path is None:
            return

        status.set(f"Analyzing {len(requests)} route(s) and creating the Excel workbook...")
        show_root(root)
        results = analyze_requests(requests)
        write_analysis_workbook(results, output_path, input_path)
        messagebox.showinfo(
            "Analysis Complete",
            f"Analyzed {len(results)} route(s).\n\nSaved:\n{output_path}",
            parent=root,
        )
    except Exception as exc:
        details = traceback.format_exc()
        print(details, file=sys.stderr)
        show_root(root)
        messagebox.showerror(
            "Route Analysis Error",
            f"{exc}\n\nThe script stopped safely before creating the report.",
            parent=root,
        )
    finally:
        root.destroy()


if __name__ == "__main__":
    run_gui()
