from dataclasses import dataclass, field


@dataclass
class NormalizedCourse:
    """Shared course contract used across input adapters and the courses.dat writer."""

    name: str
    short_name: str
    length: str
    climb: str
    controls: list[str] = field(default_factory=list)
