import csv
import io

from competition.courses.models import NormalizedCourse
from competition.courses.adapters.iof_xml.reader import read as read_iof_xml
from competition.courses.adapters.iof_xml.mappers import map_course
from competition.courses.adapters.iof_xml.writer import write


# ---------------- courses.dat writer ----------------

def _courses():
    return [
        NormalizedCourse(
            name="Men", short_name="Men", length="4960", climb="310",
            controls=["54", "55", "56", "74", "75", "77", "71", "85", "87",
                      "88", "96", "47", "57"],
        ),
        NormalizedCourse(
            name="Women", short_name="Women", length="4030", climb="275",
            controls=["51", "52", "53", "74", "75", "76", "78", "58", "61",
                      "90", "97", "43", "57"],
        ),
    ]


def test_writer_matches_mulka2_sample():
    """Compact format without Short Name column matches Mulka2 sample."""
    text = write(_courses(), include_short_name=False, punch_at_start=True)
    expected = (
        "Course,Length,Climb,Controls,S,1,2,3,4,5,6,7,8,9,10,11,12,13,F\r\n"
        "Men,4960,310,13,0,54,55,56,74,75,77,71,85,87,88,96,47,57,F\r\n"
        "Women,4030,275,13,0,51,52,53,74,75,76,78,58,61,90,97,43,57,F\r\n"
    )
    assert text == expected


def test_writer_matches_downloads_sample():
    """Format with Short Name column matches the Downloads sample."""
    courses = [
        NormalizedCourse(
            name="Long", short_name="Long", length="8010", climb="0",
            controls=["31", "32", "33", "34", "35", "36", "37", "38", "78",
                      "40", "41", "42", "43", "44", "45", "46", "47", "48",
                      "49", "50", "51", "52", "53", "54", "55", "100"],
        ),
        NormalizedCourse(
            name="Med Plus", short_name="Medium Plus", length="7150", climb="0",
            controls=["57", "31", "34", "35", "45", "44", "43", "42", "41",
                      "40", "78", "48", "47", "59", "53", "52", "51", "50",
                      "69", "67", "55", "32", "56", "100"],
        ),
    ]
    text = write(courses, include_short_name=True, punch_at_start=True)
    rows = list(csv.reader(io.StringIO(text)))
    assert rows[0] == ["Course", "Short Name", "Length", "Climb", "Controls", "S"] + \
        [str(i) for i in range(1, 27)] + ["F"]
    assert rows[1] == ["Long", "Long", "8010", "0", "26", "0"] + \
        ["31", "32", "33", "34", "35", "36", "37", "38", "78", "40", "41",
         "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52",
         "53", "54", "55", "100", "F"]
    assert rows[2] == ["Med Plus", "Medium Plus", "7150", "0", "24", "0"] + \
        ["57", "31", "34", "35", "45", "44", "43", "42", "41", "40", "78",
         "48", "47", "59", "53", "52", "51", "50", "69", "67", "55", "32",
         "56", "100", "", "", "F"]


def test_punch_at_start_default_true():
    """Global checkbox is checked by default: start column is '0'."""
    text = write(_courses(), include_short_name=False)
    rows = list(csv.reader(io.StringIO(text)))
    # Row layout: Course,Length,Climb,Controls,S,1,...
    assert rows[1][4] == "0"
    assert rows[2][4] == "0"


def test_punch_at_start_false():
    """When checkbox is cleared, start column is empty for every course."""
    text = write(_courses(), include_short_name=False, punch_at_start=False)
    rows = list(csv.reader(io.StringIO(text)))
    assert rows[1][4] == ""
    assert rows[2][4] == ""


def test_padding_shorter_courses():
    """Courses with fewer controls are padded with empty columns."""
    courses = [
        NormalizedCourse(name="Short", short_name="Short", length="1940", climb="0",
                         controls=["70", "61", "71", "72", "73", "74", "75", "56", "100"]),
        NormalizedCourse(name="Long", short_name="Long", length="8010", climb="0",
                         controls=["31", "32", "33"]),
    ]
    text = write(courses, include_short_name=False, punch_at_start=True)
    rows = list(csv.reader(io.StringIO(text)))
    header = rows[0]
    max_controls = len(header) - 6  # Course,Length,Climb,Controls,S,F
    assert max_controls == 9
    # Long row: controls start at index 5
    long_row = rows[2]
    assert long_row[5:8] == ["31", "32", "33"]
    assert long_row[8:-1] == [""] * (max_controls - 3)
    assert long_row[-1] == "F"


def test_plus_name_normalization():
    """'+' in names is normalized to ' Plus'."""
    courses = [
        NormalizedCourse(name="Short+ H", short_name="Short+ M", length="4250",
                         climb="0", controls=["65", "55", "100"]),
    ]
    # Normalization happens in the mapper; writer preserves what it's given.
    text = write(courses, include_short_name=True, punch_at_start=True)
    rows = list(csv.reader(io.StringIO(text)))
    assert rows[1][0] == "Short+ H"
    assert rows[1][1] == "Short+ M"


def test_empty_controls():
    """A course with no controls still renders a valid row."""
    courses = [
        NormalizedCourse(name="Sprint", short_name="Sprint", length="1000",
                         climb="50", controls=[]),
    ]
    text = write(courses, include_short_name=False, punch_at_start=True)
    rows = list(csv.reader(io.StringIO(text)))
    assert rows[0] == ["Course", "Length", "Climb", "Controls", "S", "F"]
    assert rows[1] == ["Sprint", "1000", "50", "0", "0", "F"]


# ---------------- IOF XML reader + mapper ----------------

_V2_XML = b"""<?xml version="1.0" encoding="UTF-8" ?>
<!DOCTYPE CourseData SYSTEM "IOFdata_OCAD.dtd">
<CourseData>
  <IOFVersion version="2.0.3" />
  <Control>
    <ControlCode> 31 </ControlCode>
  </Control>
  <Control>
    <ControlCode> 32 </ControlCode>
  </Control>
  <Control>
    <ControlCode> 100 </ControlCode>
  </Control>
  <Course>
    <CourseName> Long </CourseName>
    <ClassShortName> Long </ClassShortName>
    <CourseVariation>
      <CourseLength> 8010 </CourseLength>
      <CourseClimb>  </CourseClimb>
      <StartPointCode> S1 </StartPointCode>
      <CourseControl>
        <Sequence> 1 </Sequence>
        <ControlCode> 31 </ControlCode>
      </CourseControl>
      <CourseControl>
        <Sequence> 2 </Sequence>
        <ControlCode> 32 </ControlCode>
      </CourseControl>
      <CourseControl>
        <Sequence> 3 </Sequence>
        <ControlCode> 100 </ControlCode>
      </CourseControl>
      <FinishPointCode> F1 </FinishPointCode>
    </CourseVariation>
  </Course>
  <Course>
    <CourseName> Short+ </CourseName>
    <ClassShortName> Short M </ClassShortName>
    <CourseVariation>
      <CourseLength> 4250 </CourseLength>
      <CourseClimb>  </CourseClimb>
      <CourseControl>
        <Sequence> 1 </Sequence>
        <ControlCode> 65 </ControlCode>
      </CourseControl>
      <CourseControl>
        <Sequence> 2 </Sequence>
        <ControlCode> 100 </ControlCode>
      </CourseControl>
      <FinishPointCode> F1 </FinishPointCode>
    </CourseVariation>
  </Course>
</CourseData>"""


def test_iof_reader_parses_v2_courses():
    raw_courses = read_iof_xml(io.BytesIO(_V2_XML))
    assert len(raw_courses) == 2
    assert raw_courses[0]["name"] == "Long"
    assert raw_courses[0]["short_name"] == "Long"
    assert len(raw_courses[0]["variations"]) == 1
    var = raw_courses[0]["variations"][0]
    assert var["length"] == "8010"
    assert var["controls"] == ["31", "32", "100"]

    assert raw_courses[1]["name"] == "Short+"
    assert raw_courses[1]["short_name"] == "Short M"
    assert raw_courses[1]["variations"][0]["controls"] == ["65", "100"]


def test_iof_mapper_normalizes_plus_and_controls():
    raw_courses = read_iof_xml(io.BytesIO(_V2_XML))
    rows = []
    for raw in raw_courses:
        rows.extend(map_course(raw))
    assert len(rows) == 2

    long_row = rows[0]
    assert long_row.name == "Long"
    assert long_row.short_name == "Long"
    assert long_row.length == "8010"
    assert long_row.controls == ["31", "32", "100"]

    short_row = rows[1]
    assert short_row.name == "Short Plus"
    assert short_row.short_name == "Short M"
    assert short_row.controls == ["65", "100"]


def test_iof_end_to_end_writer_matches_expected():
    raw_courses = read_iof_xml(io.BytesIO(_V2_XML))
    rows = []
    for raw in raw_courses:
        rows.extend(map_course(raw))

    text = write(rows, include_short_name=True, punch_at_start=True)
    lines = text.split("\r\n")
    assert lines[0] == ("Course,Short Name,Length,Climb,Controls,S,1,2,3,F")
    assert lines[1] == ("Long,Long,8010,0,3,0,31,32,100,F")
    assert lines[2] == ("Short Plus,Short M,4250,0,2,0,65,100,,F")


# ---------------- IOF XML v3 ----------------

_V3_XML = b"""<?xml version="1.0" encoding="UTF-8" ?>
<CourseData xmlns="http://www.orienteering.org/datastandard/3.0" iofVersion="3.0">
  <Event><Name>Valley of the Cross 2026</Name></Event>
  <RaceCourseData>
    <Control type="Start">
      <Id>S1</Id>
    </Control>
    <Control type="Start">
      <Id>S2</Id>
    </Control>
    <Control type="Finish">
      <Id>F1</Id>
    </Control>
    <Course>
      <Name>Long</Name>
      <Length>2810</Length>
      <Climb>0</Climb>
      <CourseControl type="Start">
        <Control>S1</Control>
      </CourseControl>
      <CourseControl type="Control">
        <Control>31</Control>
      </CourseControl>
      <CourseControl type="Control">
        <Control>32</Control>
      </CourseControl>
      <CourseControl type="Control">
        <Control>100</Control>
      </CourseControl>
      <CourseControl type="Finish">
        <Control>F1</Control>
      </CourseControl>
    </Course>
    <Course>
      <Name>Short</Name>
      <Length>2050</Length>
      <Climb>0</Climb>
      <CourseControl type="Start">
        <Control>S1</Control>
      </CourseControl>
      <CourseControl type="Control">
        <Control>51</Control>
      </CourseControl>
      <CourseControl type="Control">
        <Control>32</Control>
      </CourseControl>
      <CourseControl type="Finish">
        <Control>F1</Control>
      </CourseControl>
    </Course>
    <Course>
      <Name>With Fun</Name>
      <Length>2290</Length>
      <Climb>0</Climb>
      <CourseControl type="Start">
        <Control>S2</Control>
      </CourseControl>
      <CourseControl type="Control">
        <Control>1</Control>
      </CourseControl>
      <CourseControl type="Control">
        <Control>2</Control>
      </CourseControl>
      <CourseControl type="Start">
        <Control>S2</Control>
      </CourseControl>
    </Course>
  </RaceCourseData>
</CourseData>"""


def test_iof_reader_parses_v3_courses():
    raw_courses = read_iof_xml(io.BytesIO(_V3_XML))
    assert len(raw_courses) == 3

    # v3 uses <Name> not <CourseName>, short_name falls back to name
    assert raw_courses[0]["name"] == "Long"
    assert raw_courses[0]["short_name"] == "Long"
    var = raw_courses[0]["variations"][0]
    assert var["length"] == "2810"
    assert var["start_control"] == "S1"
    assert var["finish_control"] == "F1"
    assert var["controls"] == ["31", "32", "100"]

    assert raw_courses[1]["name"] == "Short"
    assert raw_courses[1]["variations"][0]["controls"] == ["51", "32"]
    assert raw_courses[1]["variations"][0]["start_control"] == "S1"
    assert raw_courses[1]["variations"][0]["finish_control"] == "F1"

    # Multiple Start elements: last one wins; extra Start ignored as control
    assert raw_courses[2]["name"] == "With Fun"
    assert raw_courses[2]["variations"][0]["start_control"] == "S2"
    assert raw_courses[2]["variations"][0]["controls"] == ["1", "2"]


def test_iof_v3_end_to_end_writer():
    raw_courses = read_iof_xml(io.BytesIO(_V3_XML))
    rows = []
    for raw in raw_courses:
        rows.extend(map_course(raw))

    text = write(rows, include_short_name=True, punch_at_start=True)
    lines = text.split("\r\n")
    # Long has 3 controls -> max 3 control columns; shorter courses are padded
    assert lines[0] == ("Course,Short Name,Length,Climb,Controls,S,1,2,3,F")
    assert lines[1] == ("Long,Long,2810,0,3,0,31,32,100,F")
    assert lines[2] == ("Short,Short,2050,0,2,0,51,32,,F")
    assert lines[3] == ("With Fun,With Fun,2290,0,2,0,1,2,,F")
