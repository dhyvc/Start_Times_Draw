"""IOF XML v2/v3 course reader.

Parses IOF XML course data into a raw dict structure consumable by the mappers.
Handles both IOF XML 2.0 and 3.0 course elements.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Any


def _local(tag: str) -> str:
    """Return the tag name without its {namespace} prefix."""
    return tag.rpartition("}")[2]


def _children(elem: ET.Element, name: str) -> list[ET.Element]:
    """Direct children whose local name matches ``name``."""
    return [c for c in list(elem) if _local(c.tag) == name]


def _child(elem: ET.Element | None, name: str) -> ET.Element | None:
    """First direct child whose local name matches ``name`` (or None)."""
    if elem is None:
        return None
    for c in list(elem):
        if _local(c.tag) == name:
            return c
    return None


def _text(elem: ET.Element | None, name: str, default: str = "") -> str:
    """Stripped text of a direct child, or ``default`` when missing/empty."""
    node = _child(elem, name)
    if node is None or node.text is None:
        return default
    return node.text.strip()


def _parse_variation(var: ET.Element) -> dict[str, Any]:
    """Parse a single <CourseVariation> (v2) or <Course> treated as one (v3).

    Detects the style by inspecting <CourseControl> elements:
    * v2: children <Sequence> and <ControlCode>, plus separate
      <StartPointCode>/<FinishPointCode> siblings.
    * v3: <CourseControl type="Start|Control|Finish"> with a <Control> child;
      <Length>/<Climb> are direct children.
    """
    length_str = _text(var, "CourseLength") or _text(var, "Length")
    climb_str = _text(var, "CourseClimb") or _text(var, "Climb") or "0"

    start_code = _text(var, "StartPointCode")
    finish_code = _text(var, "FinishPointCode")

    course_controls = _children(var, "CourseControl")
    v3_style = any(cc.get("type") for cc in course_controls)

    control_codes: list[str] = []

    if v3_style:
        # Document order is the sequence; type attribute distinguishes roles.
        for cc in course_controls:
            ctype = (cc.get("type") or "").strip()
            ctrl_child = _child(cc, "Control")
            code = ctrl_child.text.strip() if ctrl_child is not None and ctrl_child.text else ""
            if not code:
                continue
            if ctype == "Start":
                start_code = code
            elif ctype == "Finish":
                finish_code = code
            else:
                # "Control" or any other non-start/finish type
                control_codes.append(code)
    else:
        # v2 style
        start_code = start_code or "0"
        finish_code = finish_code or "F"
        control_refs: list[tuple[str, str]] = []
        for cc in course_controls:
            seq = _text(cc, "Sequence")
            code = _text(cc, "ControlCode")
            if code:
                control_refs.append((seq, code))
        sorted_controls = sorted(
            control_refs, key=lambda x: int(x[0]) if x[0].isdigit() else 0
        )
        control_codes = [code for _, code in sorted_controls]

    return {
        "length": length_str,
        "climb": climb_str,
        "start_control": start_code,
        "finish_control": finish_code,
        "controls": control_codes,
    }


def _parse_course(elem: ET.Element) -> dict[str, Any]:
    """Parse a <Course> element into a raw dict."""
    name = _text(elem, "CourseName") or _text(elem, "Name")
    short_name = _text(elem, "ClassShortName") or name

    variations = _children(elem, "CourseVariation")
    if not variations:
        variations = [elem]

    parsed_variations = [_parse_variation(var) for var in variations]

    return {
        "name": name,
        "short_name": short_name,
        "variations": parsed_variations,
    }


def read(path: str) -> list[dict[str, Any]]:
    """Read an IOF XML v2/v3 file and return a list of raw course dicts.

    Each dict has ``name``, ``short_name`` and ``variations`` (a list of
    per-variation dicts with length, climb, controls, start/finish codes).
    """
    tree = ET.parse(path)
    root = tree.getroot()

    courses: list[dict[str, Any]] = []
    for course_elem in root.iter():
        if _local(course_elem.tag) == "Course":
            courses.append(_parse_course(course_elem))

    return courses
