# Copyright (C) 2023 Shashi Jeevan Midde Peddanna
#
# This file is released under the "GNU General Public License v3.0". 
# Please see the LICENSE file that should have been included as 
# part of this package.
from dataclasses import dataclass
from dataclasses_json import dataclass_json
from typing import List

@dataclass_json
@dataclass
class Range:
    sheet_id: int
    start_row_index: int
    end_row_index: int
    start_column_index: int
    end_column_index: int

@dataclass_json
@dataclass
class RGBColor:
    red: int
    green: int
    blue: int

@dataclass_json
@dataclass
class BackgroundColorStyle:
    rgb_color: RGBColor

@dataclass_json
@dataclass
class UserEnteredFormat:
    background_color_style: BackgroundColorStyle

@dataclass_json
@dataclass
class Value:
    user_entered_format: UserEnteredFormat

@dataclass_json
@dataclass
class Row:
    values: List[Value]

@dataclass_json
@dataclass
class UpdateCellsClass:
    range: Range
    rows: List[Row]
    fields: str

@dataclass_json
@dataclass
class Request:
    update_cells: UpdateCellsClass

@dataclass_json
@dataclass
class UpdateCells:
    requests: List[Request]

