from dataclasses import dataclass
from enum import Enum


class Action(str, Enum):
    INCREASE = "increase"
    DECREASE = "decrease"


class RoundingMode(str, Enum):
    NONE = "none"
    UP = "up"
    DOWN = "down"


@dataclass(frozen=True)
class ProcessingOptions:
    input_path: str
    output_path: str
    column: str
    percentage: float
    action: Action
    rounding: RoundingMode = RoundingMode.NONE


@dataclass(frozen=True)
class ProcessingResult:
    sheet_name: str
    processed_cells: int
    skipped_empty: int
    skipped_formula: int
    skipped_non_numeric: int
