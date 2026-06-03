from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Union


@dataclass
class GlazingItem:
    level: str | None
    tag: str
    description: str
    count: int
    width_raw: str
    height_raw: str
    width_inches: float | None
    height_inches: float | None
    area_sf: float | None
    remarks: str | None
    noa: str | None
    brand_product: str | None
    confidence: float
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class DoorItem:
    level: str | None
    door_number: str
    location: str
    quantity: int
    width_raw: str
    height_raw: str
    width_inches: float | None
    height_inches: float | None
    panels: int | None
    fixed_panels: int | None
    jamb: str | None
    type: str | None
    material: str | None
    hardware: str | None
    remarks: str | None
    confidence: float
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class GlazingSchedule:
    project_address: str | None
    source_page: int | None
    items: list[GlazingItem]
    schedule_type: str = "glazing"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class DoorSchedule:
    project_address: str | None
    source_page: int | None
    items: list[DoorItem]
    schedule_type: str = "door"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


Schedule = Union[GlazingSchedule, DoorSchedule]


@dataclass
class ExtractionResult:
    schedules: list[Schedule] = field(default_factory=list)
    audit: list[dict[str, Any]] = field(default_factory=list)

    def extend(self, other: "ExtractionResult") -> None:
        self.schedules.extend(other.schedules)
        self.audit.extend(other.audit)

    def to_dict(self) -> list[dict[str, Any]]:
        return [schedule.to_dict() for schedule in self.schedules]
