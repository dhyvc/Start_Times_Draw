"""courses.dat file writer for Mulka2.

Produces the Mulka2 courses.dat CSV format:

    Course,Length,Climb,Controls,S,1,2,...,N,F
    <name>,<length>,<climb>,<N>,<start>,<ctrl1>,...,<ctrlN>,<finish>

Optionally includes a ``Short Name`` column:

    Course,Short Name,Length,Climb,Controls,S,1,2,...,N,F
    <name>,<short_name>,<length>,<climb>,<N>,<start>,<ctrl1>,...,<ctrlN>,<finish>

The ``S`` column reflects the global ``punch_at_start`` setting:

* ``punch_at_start=True`` (default, checkbox checked): start column is ``0``
* ``punch_at_start=False`` (checkbox cleared): start column is empty

The ``F`` column is always literal ``F``.
"""

from __future__ import annotations

import csv
from io import StringIO

from ...models import NormalizedCourse

FINISH_COL = "F"


def write(
    courses: list[NormalizedCourse],
    include_short_name: bool = True,
    punch_at_start: bool = True,
) -> str:
    """Serialize courses to courses.dat CSV text.

    Parameters
    ----------
    courses:
        NormalizedCourse rows to serialize.
    include_short_name:
        When True, emit a ``Short Name`` column after ``Course`` (matches the
        reference script and the Downloads sample). When False, emit the compact
        Mulka2 sample format without that column.
    punch_at_start:
        Global checkbox (checked by default). When True the start column is
        ``0``; when False it is empty for every course.
    """
    max_controls = max((len(c.controls) for c in courses), default=0)
    start_col = "0" if punch_at_start else ""

    header: list[str] = ["Course"]
    if include_short_name:
        header.append("Short Name")
    header += ["Length", "Climb", "Controls", "S"]
    header += [str(i) for i in range(1, max_controls + 1)]
    header.append("F")

    buffer = StringIO()
    writer = csv.writer(buffer, lineterminator="\r\n")
    writer.writerow(header)

    for course in courses:
        row: list[str] = [course.name]
        if include_short_name:
            row.append(course.short_name)
        row += [
            course.length,
            course.climb,
            str(len(course.controls)),
            start_col,
        ]
        row += course.controls
        row += [""] * (max_controls - len(course.controls))
        row.append(FINISH_COL)
        writer.writerow(row)

    return buffer.getvalue()


def write_file(
    courses: list[NormalizedCourse],
    path: str,
    include_short_name: bool = True,
    punch_at_start: bool = True,
) -> None:
    """Write courses to a courses.dat file at ``path``."""
    with open(path, "w", encoding="utf-8", newline="") as fh:
        fh.write(write(courses, include_short_name=include_short_name, punch_at_start=punch_at_start))
