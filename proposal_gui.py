"""GUI workflow that builds cost proposals from Excel data and a Word template."""

# pylint: disable=too-many-lines

from __future__ import annotations

import argparse
import datetime
import json
import math
import os
import re
import shutil
import sys
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Generator, Mapping, Sequence, TypedDict, cast
from zipfile import ZipFile
from html import unescape as html_unescape
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet


_RESOURCE_ROOTS = [
    Path(r"C:\Proposal Creator"),
    Path(r"C:\Proposal Writer"),
]


def _default_resource_path(filename: str) -> Path:
    """Return the preferred resource path, prioritizing the canonical Proposal Writer folder."""

    for root in _RESOURCE_ROOTS:
        preferred = root / filename
        if preferred.exists():
            return preferred
    return Path(__file__).resolve().parent / filename


DEFAULT_WORKBOOK = _default_resource_path("Proposal Creator Key.xlsx")
DEFAULT_SCOPE_WORKBOOK = _default_resource_path("Civil Engineering Services.xlsx")
DEFAULT_CLIENT_WORKBOOK = _default_resource_path("Monday Export.xlsx")
CONFIG_PATH = _RESOURCE_ROOTS[1] / "proposal_config.json"


def _default_template_path() -> Path:
    preferred = _default_resource_path("PROPOSAL FOR PROFESSIONAL SERVICES.dotx")
    if preferred.exists():
        return preferred
    return _default_resource_path("Civil Proposal Template.dotx")


DEFAULT_TEMPLATE = _default_template_path()


def _load_config() -> dict[str, str]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _save_config(data: dict[str, str]) -> None:
    CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")


@dataclass(frozen=True)
class ProjectType:
    """Metadata describing an available project type."""

    id: int
    name: str
    description: str | None = None


@dataclass(frozen=True)
class ServiceType:
    """Represents a service category in the proposal key."""

    id: int
    name: str
    description: str | None = None


@dataclass(frozen=True)
class ProjectManager:
    """Stores the primary contact assigned to a proposal."""

    name: str
    title: str | None = None


@dataclass(frozen=True)
class MiscOption:
    """Scale + acres per sheet pairing from the misc worksheet."""

    scale: str
    acres_per_sheet: str


@dataclass(frozen=True)
class ScopeItem:
    """A selectable scope entry sourced from the Civil Engineering workbook."""

    id: int
    name: str
    description: str
    project_type_id: int | None
    service_type_id: int | None
    units: str | None
    fee: float | None


SCOPE_STATUS_REQUIRED = "Required"
SCOPE_STATUS_REQUESTED = "Requested By Client"
SCOPE_STATUS_OPTIONS = [SCOPE_STATUS_REQUIRED, SCOPE_STATUS_REQUESTED]


@dataclass(frozen=True)
class ScopeSelection:
    """User-selected scope with status."""

    scope: ScopeItem
    status: str = SCOPE_STATUS_REQUIRED
    show_fee_breakout: bool = False


@dataclass(frozen=True)
class Assumption:
    """Proposal assumption or exclusion."""

    text: str
    is_additional: bool = False


@dataclass(frozen=True)
class ClientInfo:
    """Client or opportunity details sourced from the Monday export."""

    name: str
    requested_on: str | None = None
    project_name: str | None = None
    project_number: str | None = None
    pm: str | None = None
    contact_name: str | None = None
    company: str | None = None
    contact_email: str | None = None
    contact_phone: str | None = None
    address: str | None = None
    description: str | None = None
    communication_method: str | None = None


@dataclass(frozen=True)
class ProposalData:
    """Container for every lookup dataset needed by the wizard."""

    project_types: list[ProjectType]
    services: list[ServiceType]
    project_managers: list[ProjectManager]
    misc_options: list[MiscOption]
    scope_items: list[ScopeItem]
    assumptions: list[Assumption]
    clients: list[ClientInfo]
    fee_ref_blocks: dict[str, FeeRefBlock]
    average_multiplier: float


@dataclass(frozen=True)
class ProjectMetrics:
    """Numeric inputs captured in the first step of the wizard."""

    developed_area: float
    roadway_length: float
    scale: MiscOption
    average_multiplier: float


@dataclass(frozen=True)
class FeeRow:
    """A calculated fee line for the proposal."""

    scope: ScopeItem
    unit: str
    units: float
    unit_cost: float
    total: float
    status: str = SCOPE_STATUS_REQUIRED


@dataclass(frozen=True)
class FeeRefPosition:
    """Editable position row from the Fee_Ref sheet."""

    title: str
    rate: float
    hours: float


@dataclass(frozen=True)
class FeeRefBlock:
    """Fee_Ref block for a specific project type."""

    project_type: str
    positions: list[FeeRefPosition]


def _fee_rows_default() -> list["FeeRow"]:
    return []


def _fee_totals_default() -> dict[int, float]:
    return {}


def _fee_ref_totals_default() -> dict[str, float]:
    return {}


def _all_services_default() -> list[ServiceType]:
    return []


@dataclass(frozen=True)
class ProposalContext:
    """Final user selections that feed the Word template."""

    client: ClientInfo | None
    project_types: list[ProjectType]
    services: list[ServiceType]
    project_managers: list[ProjectManager]
    scale: MiscOption
    scopes: list[ScopeSelection]
    assumptions: list[Assumption]
    metrics: ProjectMetrics | None = None
    fee_rows: list["FeeRow"] = field(default_factory=_fee_rows_default)
    fee_totals: dict[int, float] = field(default_factory=_fee_totals_default)
    all_services: list[ServiceType] = field(default_factory=_all_services_default)
    fee_ref_totals: dict[str, float] = field(default_factory=_fee_ref_totals_default)
    pm_name: str | None = None
    pm_title: str | None = None


@dataclass(frozen=True)
class OptionItem:
    """Listbox entry in the wizard."""

    label: str
    value: object


def _client_label(client: ClientInfo) -> str:
    """Produce a concise label for client list selections."""

    base = client.project_name or client.name
    if client.contact_name:
        return f"{base} - {client.contact_name}"
    return base


@dataclass
class WizardStep:
    """Declarative definition of each wizard step."""

    key: str
    title: str
    description: str
    options: list[OptionItem]
    allow_multiple: bool = True
DEMO_PROJECT_TYPES = [
    ProjectType(1, "Residential Subdivision"),
    ProjectType(2, "Commercial Small (< 3 ac)"),
]
DEMO_SERVICES = [
    ServiceType(1, "Planning Services"),
    ServiceType(3, "Civil Engineering Services"),
]
DEMO_PROJECT_MANAGERS = [
    ProjectManager("Rob Foster", "Project Manager"),
    ProjectManager("Skylar Reese", "Associate PM"),
]
DEMO_MISC_OPTIONS = [
    MiscOption('1"=10\'', "1"),
    MiscOption('1"=20\'', "4.2"),
]
DEMO_SCOPE_ITEMS = [
    ScopeItem(
        id=100,
        name="Site Plan",
        description="Prepare a detailed site plan suitable for AHJ review.",
        project_type_id=None,
        service_type_id=3,
        units="EA",
        fee=4500.0,
    ),
    ScopeItem(
        id=101,
        name="Grading Plan",
        description="Develop grading plan and submit to AHJ.",
        project_type_id=None,
        service_type_id=3,
        units="EA",
        fee=3200.0,
    ),
    ScopeItem(
        id=102,
        name="Drainage Report",
        description="Prepare narrative drainage report for the proposed improvements.",
        project_type_id=None,
        service_type_id=4,
        units="EA",
        fee=None,
    ),
]
DEMO_ASSUMPTIONS = [
    Assumption("Client to provide survey control data."),
    Assumption("City review fees are excluded and will be paid directly by the client."),
    Assumption("Additional coordination meetings billed hourly upon request.", is_additional=True),
]
DEMO_FEE_REF_BLOCKS = {
    "Commercial Small (< 3 ac)": FeeRefBlock(
        project_type="Commercial Small (< 3 ac)",
        positions=[
            FeeRefPosition("Principal", 225.0, 0.5),
            FeeRefPosition("PM", 175.0, 4.0),
            FeeRefPosition("Proj. Engr.", 150.0, 3.0),
        ],
    )
}
DEMO_CLIENTS = [
    ClientInfo(
        name="JAW HEADQUARTERS",
        project_name="JAW HQ Expansion",
        contact_name="Pat Stone",
        company="JAW Corp",
        contact_email="pat.stone@example.com",
        contact_phone="555-0100",
        address="123 Main St, Austin, TX",
        description="HQ expansion, parking improvements, and streetscape upgrades.",
        pm="Rob Foster",
    )
]


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Build the CLI argument parser and return parsed arguments."""

    parser = argparse.ArgumentParser(
        description="GUI assistant that builds cost proposals from Excel-based scope libraries."
    )
    parser.add_argument(
        "--data",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help="Path to 'Proposal Creator Key.xlsx'.",
    )
    parser.add_argument(
        "--scope-data",
        type=Path,
        default=DEFAULT_SCOPE_WORKBOOK,
        help="Path to 'Civil Engineering Services.xlsx'.",
    )
    parser.add_argument(
        "--client-data",
        type=Path,
        default=DEFAULT_CLIENT_WORKBOOK,
        help="Path to 'Monday Export.xlsx' containing client requests.",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help="Path to 'Civil Proposal Template.dotx'.",
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help="Launch the GUI with built-in sample datasets (skips Excel reading).",
    )
    return parser.parse_args(argv)
def load_proposal_data(key_workbook: Path, scope_workbook: Path, client_workbook: Path) -> ProposalData:
    """Read the key + scope workbooks and convert them into typed lists."""

    if not key_workbook.exists():
        raise FileNotFoundError(f"Could not find '{key_workbook}'.")
    if not scope_workbook.exists():
        raise FileNotFoundError(f"Could not find '{scope_workbook}'.")

    key_book = load_workbook(key_workbook, data_only=True, read_only=True)
    try:
        try:
            project_types = _read_project_types(key_book["Project_Types"])
            services = _read_services(key_book["Services"])
            managers = _read_project_managers(key_book["PMs"])
            misc_sheet = (
                key_book["SheetsbyAcre"]
                if "SheetsbyAcre" in key_book.sheetnames
                else key_book["Misc"]
            )
            misc_options = _read_misc_options(misc_sheet)
        except KeyError as exc:
            raise ValueError(
                "The proposal key workbook must contain "
                "Project_Types, Services, PMs, and a SheetsbyAcre or Misc sheet."
            ) from exc
    finally:
        key_book.close()

    scope_book = load_workbook(scope_workbook, data_only=True, read_only=True)
    try:
        try:
            scope_items = _read_scope_items(scope_book["Scopes"])
            assumptions = _read_assumptions(scope_book["Assumptions"])
            fee_ref_blocks = _read_fee_ref_blocks(scope_workbook)
            average_multiplier = _read_average_multiplier(scope_workbook)
        except KeyError as exc:
            raise ValueError(
                "The Civil Engineering Services workbook must contain "
                "Scopes and Assumptions sheets."
            ) from exc
    finally:
        scope_book.close()

    clients = _read_clients(client_workbook)

    return ProposalData(
        project_types=project_types,
        services=services,
        project_managers=managers,
        misc_options=misc_options,
        scope_items=scope_items,
        assumptions=assumptions,
        clients=clients,
        fee_ref_blocks=fee_ref_blocks,
        average_multiplier=average_multiplier,
    )


def _iter_rows(sheet: Worksheet, start_row: int, max_empty: int = 50) -> Generator[tuple[object, ...], None, None]:
    """Yield worksheet rows, stopping after too many blank entries."""

    empty = 0
    for row in sheet.iter_rows(min_row=start_row, values_only=True):
        if all(cell in (None, "") for cell in row):
            empty += 1
            if empty >= max_empty:
                break
            continue
        empty = 0
        yield row


def _clean_cell(value: object | None) -> str:
    """Return a trimmed string value for worksheet cells."""

    if value in (None, ""):
        return ""
    return str(value).strip()


def _safe_int(value: object | None) -> int | None:
    """Convert worksheet values to integers when possible."""

    if value in (None, ""):
        return None
    try:
        return int(float(cast(str, value)))
    except (TypeError, ValueError):
        return None


def _as_float(value: object | None) -> float | None:
    """Convert worksheet values to floats when possible."""

    if value in (None, ""):
        return None
    try:
        return float(cast(str, value))
    except (TypeError, ValueError):
        return None


def _read_project_types(sheet: Worksheet) -> list[ProjectType]:
    items: list[ProjectType] = []
    for row in _iter_rows(sheet, start_row=2):
        project_id = _safe_int(row[0])
        name = _clean_cell(row[1])
        if project_id is None or not name:
            continue
        items.append(ProjectType(project_id, name, _clean_cell(row[2]) or None))
    if not items:
        raise ValueError("No project types were found in the workbook.")
    return items


def _read_services(sheet: Worksheet) -> list[ServiceType]:
    items: list[ServiceType] = []
    for row in _iter_rows(sheet, start_row=2):
        service_id = _safe_int(row[0])
        name = _clean_cell(row[1])
        if service_id is None or not name:
            continue
        items.append(ServiceType(service_id, name, _clean_cell(row[2]) or None))
    if not items:
        raise ValueError("No service types were found in the workbook.")
    return items


def _read_project_managers(sheet: Worksheet) -> list[ProjectManager]:
    items: list[ProjectManager] = []
    for row in _iter_rows(sheet, start_row=2):
        name = _clean_cell(row[0])
        if not name:
            continue
        items.append(ProjectManager(name=name, title=_clean_cell(row[1]) or None))
    if not items:
        raise ValueError("No project managers were found in the workbook.")
    return items


def _read_misc_options(sheet: Worksheet) -> list[MiscOption]:
    options: list[MiscOption] = []
    for row in _iter_rows(sheet, start_row=2):
        scale = _clean_cell(row[0])
        acres = _clean_cell(row[1])
        if scale and acres:
            options.append(MiscOption(scale=scale, acres_per_sheet=acres))
    if not options:
        raise ValueError("No entries were detected in the Misc sheet.")
    return options


def _read_scope_items(sheet: Worksheet) -> list[ScopeItem]:
    items: list[ScopeItem] = []
    for row in _iter_rows(sheet, start_row=2):
        scope_id = _safe_int(row[0])
        name = _clean_cell(row[1])
        description = _clean_cell(row[2])
        if scope_id is None or not name or not description:
            continue
        items.append(
            ScopeItem(
                id=scope_id,
                name=name,
                description=description,
                project_type_id=_safe_int(row[3]),
                service_type_id=_safe_int(row[4]),
                units=_clean_cell(row[5]) or None,
                fee=_as_float(row[6]),
            )
        )
    if not items:
        raise ValueError("No scope items were found in the Civil Engineering Services workbook.")
    return items


def _read_assumptions(sheet: Worksheet) -> list[Assumption]:
    results: list[Assumption] = []
    for row in _iter_rows(sheet, start_row=2):
        statement = _clean_cell(row[1])
        if not statement:
            continue
        results.append(Assumption(text=statement, is_additional=bool(_clean_cell(row[0]))))
    if not results:
        raise ValueError("No assumptions were defined in the Civil Engineering Services workbook.")
    return results


def _read_fee_ref_blocks(workbook_path: Path) -> dict[str, FeeRefBlock]:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Could not find '{workbook_path}'.")

    formula_book = load_workbook(workbook_path, data_only=False, read_only=True)
    value_book = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        formula_sheet = formula_book["Fee_Ref"]
        value_sheet = value_book["Fee_Ref"]
    except KeyError as exc:
        formula_book.close()
        value_book.close()
        raise ValueError("Fee_Ref sheet is missing from Civil Engineering Services.xlsx.") from exc

    def _cell_value(sheet: Worksheet, row: int, col: int) -> str:
        return _clean_cell(sheet.cell(row=row, column=col).value)

    def _float_value(sheet: Worksheet, row: int, col: int) -> float:
        return _as_float(sheet.cell(row=row, column=col).value) or 0.0

    blocks: dict[str, FeeRefBlock] = {}
    empty_rows = 0
    for row in range(4, 200):
        name = _cell_value(formula_sheet, row, 1)
        if not name:
            empty_rows += 1
            if empty_rows > 15:
                break
            continue
        empty_rows = 0
        formula = _clean_cell(formula_sheet.cell(row=row, column=3).value)
        match = re.search(r"=([A-Z]+)(\d+)", formula)
        if not match:
            continue
        ref_col_letters, ref_row_text = match.groups()
        ref_row = int(ref_row_text)
        ref_col = column_index_from_string(ref_col_letters)
        start_col = ref_col - 3
        if start_col < 1:
            continue
        header_row = None
        for search_row in range(ref_row, 1, -1):
            if _cell_value(formula_sheet, search_row, start_col).lower() == "position":
                header_row = search_row
                break
        if header_row is None:
            continue

        positions: list[FeeRefPosition] = []
        for data_row in range(header_row + 1, header_row + 60):
            title = _cell_value(value_sheet, data_row, start_col)
            if not title:
                continue
            lowered = title.lower()
            if "blended rate" in lowered or "typical cost per sheet" in lowered:
                break
            rate = _float_value(value_sheet, data_row, start_col + 1)
            hours = _float_value(value_sheet, data_row, start_col + 3)
            positions.append(FeeRefPosition(title=title, rate=rate, hours=hours))
        if positions:
            blocks[name] = FeeRefBlock(project_type=name, positions=positions)

    formula_book.close()
    value_book.close()
    return blocks


def _read_average_multiplier(workbook_path: Path) -> float:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Could not find '{workbook_path}'.")

    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        sheet = workbook["Fee_Ref"]
        value = sheet["N1"].value
    except KeyError:
        workbook.close()
        return 1.0
    workbook.close()
    try:
        return float(value)
    except (TypeError, ValueError):
        return 1.0


def _read_clients(workbook_path: Path) -> list[ClientInfo]:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Could not find '{workbook_path}'.")

    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        if sheet is None:
            raise ValueError("The Monday export workbook does not include an active worksheet.")
        rows = [[_clean_cell(cell) for cell in row] for row in sheet.iter_rows(values_only=True)]
        header: list[str] | None = None
        start_index = 0
        for idx, row in enumerate(rows):
            if not any(row):
                continue
            if "Name" in row and "Contact Email" in row:
                header = row
                start_index = idx + 1
                break
        if header is None:
            raise ValueError(
                "Unable to locate the header row in the Monday export. "
                "Ensure the sheet contains a row with 'Name' and 'Contact Email'."
            )
        name_index = header.index("Name")
        clients: list[ClientInfo] = []
        empty_rows = 0
        for row in rows[start_index:]:
            if not any(row):
                empty_rows += 1
                if empty_rows > 20:
                    break
                continue
            empty_rows = 0
            if len(row) <= name_index:
                continue
            name = row[name_index]
            if not name:
                continue
            data = {header[i]: row[i] for i in range(min(len(header), len(row))) if header[i]}
            clients.append(
                ClientInfo(
                    name=data.get("Name", name),
                    requested_on=data.get("Requested on") or None,
                    project_name=data.get("Project Name") or None,
                    project_number=data.get("Project Number") or None,
                    pm=data.get("PM") or None,
                    contact_name=data.get("Contact Name") or None,
                    company=data.get("Company/Owner") or None,
                    contact_email=data.get("Contact Email") or data.get("Email") or None,
                    contact_phone=data.get("Contact Phone Number") or None,
                    address=data.get("Contact Address") or data.get("Location") or None,
                    description=data.get("Description of project") or None,
                    communication_method=data.get("Preferred Method of Communication:") or None,
                )
            )
        if not clients:
            raise ValueError("No client records were found in the Monday export.")
        return clients
    finally:
        workbook.close()


def format_currency(amount: float | None) -> str:
    """Render a currency string (or TBD when the amount is missing)."""

    if amount is None:
        return "TBD"
    return f"${amount:,.2f}"


def _last_name(full_name: str | None) -> str | None:
    """Return the last name from a contact name."""

    if not full_name:
        return None
    parts = full_name.strip().split()
    if not parts:
        return None
    return parts[-1]


def _safe_float(value: str) -> float | None:
    try:
        return float(value.strip())
    except (AttributeError, ValueError):
        return None


def _safe_int_ceil(value: float | None) -> int | None:
    if value is None:
        return None
    return int(math.ceil(value))


def _calculate_units(
    unit: str,
    developed_area: float,
    acres_per_sheet: float,
    roadway_length: float,
) -> int:
    unit_upper = unit.upper()
    unit_key = re.sub(r"[^A-Z]", "", unit_upper)
    if "SHEET" in unit_upper:
        return _safe_int_ceil(developed_area / acres_per_sheet) or 1
    if unit_key == "PP":
        return _safe_int_ceil(roadway_length / 800) or 1
    if unit_key == "EA":
        return 1
    return 1


def _sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value.strip())
    return cleaned or "Proposal"


def _cached_source_for_key(key: str, cache_dir: Path) -> Path | None:
    if key == "template":
        candidates = [
            cache_dir / "template.dotx",
            cache_dir / "template.docx",
        ]
        existing = [path for path in candidates if path.exists()]
        if not existing:
            return None
        return max(existing, key=lambda path: path.stat().st_mtime)
    candidate = cache_dir / f"{key}.xlsx"
    if candidate.exists():
        return candidate
    return None


def _export_excel_iteration(
    sources: dict[str, Path],
    output_dir: Path,
    cache_dir: Path,
) -> tuple[list[Path], list[str]]:
    if not output_dir.exists():
        raise FileNotFoundError(f"Export folder '{output_dir}' was not found.")
    if not output_dir.is_dir():
        raise ValueError(f"Export path '{output_dir}' is not a folder.")
    cache_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    created: list[Path] = []
    used_cache: list[str] = []
    for key in ("data", "scope_data", "client_data", "template"):
        source = sources.get(key)
        if source is None:
            continue
        display_stem = source.stem
        if source.exists():
            suffix = source.suffix.lower()
            copy_source = source
        else:
            cached = _cached_source_for_key(key, cache_dir)
            if cached is None:
                raise FileNotFoundError(f"Could not find '{source}'.")
            copy_source = cached
            suffix = cached.suffix.lower()
            used_cache.append(key)
        if key == "template":
            if suffix not in (".dotx", ".docx"):
                raise ValueError(f"'{source.name}' is not a Word template file.")
        else:
            if suffix != ".xlsx":
                raise ValueError(f"'{source.name}' is not an Excel .xlsx file.")

        base_name = _sanitize_filename(f"{display_stem} - {timestamp}")
        target = output_dir / f"{base_name}{suffix}"
        if target.exists():
            for idx in range(1, 1000):
                candidate = output_dir / f"{base_name} ({idx}){suffix}"
                if not candidate.exists():
                    target = candidate
                    break
        shutil.copyfile(copy_source, target)
        created.append(target)

        cache_target = cache_dir / f"{key}{suffix}"
        shutil.copyfile(copy_source, cache_target)

    if not created:
        raise ValueError("No source files were selected to export.")
    return created, used_cache


def _lookup_pm_title(pm_name: str | None, managers: list[ProjectManager]) -> str | None:
    if not pm_name:
        return None
    for manager in managers:
        if manager.name.lower() == pm_name.lower():
            return manager.title
    return None


def _build_placeholder_replacements(context: ProposalContext) -> dict[str, str | None]:
    client = context.client
    contact_name = client.contact_name if client else None
    pm_name = context.pm_name
    pm_title = context.pm_title
    average_multiplier = ""
    if context.metrics is not None:
        average_multiplier = f"{context.metrics.average_multiplier:.2f}"
    if not pm_name and context.project_managers:
        pm_name = context.project_managers[0].name
    if not pm_title and context.project_managers:
        pm_title = context.project_managers[0].title
    replacements: dict[str, str | None] = {
        "Contact Name": contact_name or "",
        "Company/Owner": (client.company if client else "") or "",
        "Contact Address": (client.address if client else "") or "",
        "Last Name": _last_name(contact_name) or "",
        "Description of project": (client.description if client else "") or "",
        "PM": pm_name or "",
        "PM Title": pm_title or "",
        "Project Name": (client.project_name if client else "") or (client.name if client else "") or "",
        "Average Multiplier": average_multiplier,
    }
    return replacements


def generate_proposal_document(
    template_path: Path,
    output_path: Path,
    context: ProposalContext,
    replacements: Mapping[str, str | None],
) -> Path:
    """Produce a Word document that contains the selected scope details."""

    if not template_path.exists():
        raise FileNotFoundError(f"Template '{template_path}' was not found.")

    output_path = output_path.with_suffix(".docx")
    shutil.copyfile(template_path, output_path)

    with ZipFile(output_path, "r") as doc:
        contents = {name: doc.read(name) for name in doc.namelist()}

    for name, raw in contents.items():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        xml = raw.decode("utf-8")
        xml = _replace_placeholder_block(
            xml,
            "<Assumptions>",
            _build_assumption_lines(context),
        )
        xml = _replace_placeholder_block(
            xml,
            "<Scope>",
            _build_scope_lines(context),
        )
        xml = _replace_placeholder_block(
            xml,
            "<Add_Scope>",
            _build_additional_services_lines(context),
        )
        fee_table_xml = _build_fee_table_xml(context)
        if fee_table_xml:
            xml = _replace_placeholder_with_xml(xml, "<Fee Table>", fee_table_xml)
        xml = _replace_inline_placeholders(xml, replacements)
        contents[name] = xml.encode("utf-8")

    content_type_key = "[Content_Types].xml"
    if content_type_key in contents:
        type_xml = contents[content_type_key].decode("utf-8")
        contents[content_type_key] = _ensure_doc_content_types(type_xml).encode("utf-8")

    with ZipFile(output_path, "w") as doc:
        for name, data in contents.items():
            doc.writestr(name, data)

    return output_path


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> Path:
    """Convert a Word document to PDF using the installed Word application."""

    try:
        import win32com.client  # type: ignore[import]
    except ImportError as exc:
        raise RuntimeError(
            "win32com is required to export proposals as PDF files. "
            "Please install pywin32 and try again."
        ) from exc

    docx_path = docx_path.resolve()
    pdf_path = pdf_path.with_suffix(".pdf").resolve()
    word: Any = win32com.client.DispatchEx("Word.Application")  # type: ignore[attr-defined]
    word.Visible = False
    doc: Any = word.Documents.Open(  # type: ignore[attr-defined]
        str(docx_path),
        False,
        True,  # ReadOnly
    )
    try:
        doc.ExportAsFixedFormat(  # type: ignore[attr-defined]
            OutputFileName=str(pdf_path),
            ExportFormat=17,  # wdExportFormatPDF
        )
    finally:
        close_method = cast(Callable[[bool], None], getattr(doc, "Close"))  # type: ignore[arg-type]
        close_method(False)
        quit_method = cast(Callable[[], None], getattr(word, "Quit"))  # type: ignore[arg-type]
        quit_method()

    return pdf_path


def _replace_placeholder_block(xml: str, placeholder: str, lines: list[str]) -> str:
    """Replace the placeholder paragraph with new XML content."""

    if not lines or all(line == "" for line in lines):
        return xml
    for match in re.finditer(r"<w:p[^>]*>.*?</w:p>", xml, re.DOTALL):
        paragraph = match.group(0)
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", paragraph))
        text = html_unescape(text)
        if placeholder not in text:
            continue
        ppr_match = re.search(r"<w:pPr>.*?</w:pPr>", paragraph, re.DOTALL)
        paragraph_props = ppr_match.group(0) if ppr_match else ""
        replacement_xml = _paragraphs_xml(lines, paragraph_props)
        return xml[: match.start()] + replacement_xml + xml[match.end() :]
    return xml


def _replace_placeholder_with_xml(xml: str, placeholder: str, replacement_xml: str) -> str:
    """Replace the placeholder paragraph with raw XML content."""

    for match in re.finditer(r"<w:p[^>]*>.*?</w:p>", xml, re.DOTALL):
        paragraph = match.group(0)
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", paragraph))
        text = html_unescape(text)
        if placeholder not in text:
            continue
        return xml[: match.start()] + replacement_xml + xml[match.end() :]
    return xml


def _paragraphs_xml(lines: list[str], paragraph_props: str = "") -> str:
    """Build very simple paragraphs for insertion into the Word template."""

    if not lines:
        lines = [""]
    paragraphs: list[str] = []
    for line in lines:
        if line:
            run_pr = ""
            if line == "Requested by Client Items":
                run_pr = "<w:rPr><w:b/><w:sz w:val=\"26\"/></w:rPr>"
            paragraphs.append(
                "<w:p>"
                f"{paragraph_props}"
                f"<w:r>{run_pr}<w:t xml:space=\"preserve\">"
                f"{escape(line)}"
                "</w:t></w:r></w:p>"
            )
        else:
            paragraphs.append(f"<w:p>{paragraph_props}</w:p>")
    return "".join(paragraphs)


def _replace_inline_placeholders(xml: str, replacements: Mapping[str, str | None]) -> str:
    """Replace inline placeholders when replacement text is available."""

    for placeholder, value in replacements.items():
        if value is None or value == "":
            continue
        safe_value: str = value
        placeholder_pattern = "(?:<[^>]+>)*" + "".join(
            rf"{re.escape(char)}(?:<[^>]+>)*" for char in placeholder
        )
        pattern = re.compile(rf"&lt;{placeholder_pattern}&gt;", re.DOTALL)

        def _replacement(match: re.Match[str]) -> str:
            fragment = match.group(0)
            rpr_match = re.search(r"<w:rPr>.*?</w:rPr>", fragment, re.DOTALL)
            rpr = rpr_match.group(0) if rpr_match else ""
            return (
                f"<w:r>{rpr}<w:t xml:space=\"preserve\">"
                f"{escape(safe_value)}</w:t></w:r>"
            )

        xml = pattern.sub(_replacement, xml)
    return xml


def _scope_included_in_base(status: str) -> bool:
    return status == SCOPE_STATUS_REQUIRED


def _scope_is_requested(status: str) -> bool:
    return status == SCOPE_STATUS_REQUESTED


def _build_scope_lines(context: ProposalContext) -> list[str]:
    """Return the textual scope summary fed into the template."""

    if not context.scopes:
        return []
    lines: list[str] = []
    required_scopes = [selection for selection in context.scopes if not _scope_is_requested(selection.status)]
    requested_scopes = [selection for selection in context.scopes if _scope_is_requested(selection.status)]

    scope_index = 1
    for selection in required_scopes:
        scope = selection.scope
        details: list[str] = []
        if scope.units:
            details.append(scope.units)
        header = f"{scope_index}. {scope.name}"
        if details:
            header = f"{header} ({' | '.join(details)})"
        lines.append(header)
        if scope.description:
            lines.extend(scope.description.strip().splitlines())
        lines.append("")
        scope_index += 1

    if requested_scopes:
        lines.append("Requested by Client Items")
        lines.append("")
        for selection in requested_scopes:
            scope = selection.scope
            details = []
            if scope.units:
                details.append(scope.units)
            header = f"{scope_index}. {scope.name} (Requested by Client)"
            if details:
                header = f"{header} ({' | '.join(details)})"
            lines.append(header)
            if scope.description:
                lines.extend(scope.description.strip().splitlines())
            lines.append("")
            scope_index += 1
    return lines


def _build_assumption_lines(context: ProposalContext) -> list[str]:
    """Return assumption bullet strings."""

    if not context.assumptions:
        return []
    return [f"- {assumption.text}" for assumption in context.assumptions]


def _build_additional_services_lines(context: ProposalContext) -> list[str]:
    """Return explanatory text for optional services."""

    return []


def _table_xml(rows: list[list[str]], header_rows: int = 1, merge_header: bool = False) -> str:
    """Build a simple Word table from rows of text."""

    if not rows:
        return ""
    col_count = len(rows[0])
    table_width = 10800
    if col_count == 3:
        col_widths = [7200, 1800, 1800]
    else:
        col_widths = [int(table_width / col_count)] * col_count
    table_parts = [
        '<w:tbl>',
        "<w:tblPr>"
        "<w:tblStyle w:val=\"TableGrid\"/>"
        "<w:jc w:val=\"center\"/>"
        f"<w:tblW w:w=\"{table_width}\" w:type=\"dxa\"/>"
        "<w:tblLayout w:type=\"fixed\"/>"
        "<w:tblBorders>"
        "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>"
        "</w:tblBorders>"
        "</w:tblPr>",
        "<w:tblGrid>"
        + "".join(f'<w:gridCol w:w="{width}"/>' for width in col_widths)
        + "</w:tblGrid>",
    ]
    for row_index, row in enumerate(rows):
        table_parts.append("<w:tr>")
        if merge_header and row_index < header_rows:
            title = row[0] if row else ""
            run_pr = "<w:rPr><w:b/></w:rPr>"
            table_parts.append(
                "<w:tc>"
                f"<w:tcPr><w:tcW w:w=\"{sum(col_widths)}\" w:type=\"dxa\"/>"
                "<w:vAlign w:val=\"center\"/>"
                f"<w:gridSpan w:val=\"{col_count}\"/></w:tcPr>"
                "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>"
                f"<w:r>{run_pr}<w:t xml:space=\"preserve\">{escape(title)}</w:t></w:r>"
                "</w:p>"
                "</w:tc>"
            )
        else:
            for col_idx, cell in enumerate(row):
                cell_text = cell
                extra_bold = False
                if cell_text.startswith("%%B%%"):
                    cell_text = cell_text[len("%%B%%") :]
                    extra_bold = True
                bold = row_index < header_rows
                if bold or extra_bold:
                    run_pr = "<w:rPr><w:b/></w:rPr>"
                else:
                    run_pr = ""
                align = "center" if col_idx > 0 else "left"
                table_parts.append(
                    "<w:tc>"
                    f"<w:tcPr><w:tcW w:w=\"{col_widths[col_idx]}\" w:type=\"dxa\"/>"
                    "<w:vAlign w:val=\"center\"/></w:tcPr>"
                    f"<w:p><w:pPr><w:jc w:val=\"{align}\"/></w:pPr>"
                    f"<w:r>{run_pr}<w:t xml:space=\"preserve\">{escape(cell_text)}</w:t></w:r>"
                    "</w:p>"
                    "</w:tc>"
                )
        table_parts.append("</w:tr>")
    table_parts.append("</w:tbl>")
    return "".join(table_parts)


def _build_fee_table_xml(context: ProposalContext) -> str:
    """Build the detailed cost breakdown table as Word XML."""

    if not context.fee_rows:
        return ""
    fee_by_scope_id = {row.scope.id: row.total for row in context.fee_rows}
    requested_scopes = [selection for selection in context.scopes if _scope_is_requested(selection.status)]
    fee_breakout = [selection for selection in context.scopes if selection.show_fee_breakout]

    breakdown_rows: list[list[str]] = [["Detailed Cost Breakdown", "", ""]]
    breakdown_total = 0.0
    fee_ref_total = sum(context.fee_ref_totals.values()) if context.fee_ref_totals else 0.0
    for service in context.services:
        total = context.fee_totals.get(service.id, 0.0)
        if fee_ref_total and "civil engineering" in service.name.lower():
            total += fee_ref_total
        if total <= 0:
            continue
        amount = format_currency(total)
        breakdown_rows.append([f"%%B%%{service.name}", amount, "No Tax"])
        breakdown_total += total
    breakdown_rows.append(["%%B%%Total", format_currency(breakdown_total), "No Tax"])

    requested_rows: list[list[str]] = []
    for selection in requested_scopes:
        amount = fee_by_scope_id.get(selection.scope.id, selection.scope.fee)
        if amount is None or amount <= 0:
            continue
        requested_rows.append([selection.scope.name, format_currency(amount), "No Tax"])
    if requested_rows:
        breakdown_rows.append(["%%B%%If Requested Services", "", ""])
        breakdown_rows.extend([[f"%%B%%     * {row[0]}", row[1], row[2]] for row in requested_rows])

    nte_rows: list[list[str]] = []
    for selection in fee_breakout:
        amount = fee_by_scope_id.get(selection.scope.id, selection.scope.fee)
        if amount is None or amount <= 0:
            continue
        nte_rows.append([selection.scope.name, format_currency(amount), "No Tax"])
    if nte_rows:
        breakdown_rows.append(["%%B%%Hourly Not-To-Exceed Services", "", ""])
        breakdown_rows.extend([[f"%%B%%     * {row[0]}", row[1], row[2]] for row in nte_rows])

    breakdown_table = _table_xml(breakdown_rows, merge_header=True)
    note = "* Items denoted with an asterisk are not included in the total cost."
    return breakdown_table + _paragraphs_xml(["", note])


def _client_summary_lines(client: ClientInfo) -> list[str]:
    """Return a human-readable summary for the client selection."""

    lines: list[str] = []
    primary = client.project_name or client.name
    lines.append(primary)
    if client.company and client.company != client.name:
        lines.append(f"Company: {client.company}")
    if client.contact_name:
        contact = f"Contact: {client.contact_name}"
        bits: list[str] = []
        if client.contact_phone:
            bits.append(client.contact_phone)
        if client.contact_email:
            bits.append(client.contact_email)
        if bits:
            contact = f"{contact} ({' | '.join(bits)})"
        lines.append(contact)
    if client.address:
        lines.append(f"Address: {client.address}")
    if client.requested_on:
        lines.append(f"Requested: {client.requested_on}")
    if client.communication_method:
        lines.append(f"Preferred communication: {client.communication_method}")
    if client.description:
        lines.append("")
        lines.append(client.description)
    return lines


def _ensure_doc_content_types(xml: str) -> str:
    """Ensure the package advertises a .docx main content type."""

    template_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
    )
    document_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    )
    if template_type in xml:
        xml = xml.replace(template_type, document_type)
    return xml
class ProposalApp(tk.Tk):  # pylint: disable=too-many-instance-attributes
    """Tkinter wizard that guides the user through every proposal choice."""

    def __init__(
        self,
        data: ProposalData,
        key_path: Path | None,
        scope_path: Path | None,
        client_path: Path | None,
        template_path: Path,
    ) -> None:
        super().__init__()
        self.data = data
        self.key_path = key_path
        self.scope_path = scope_path
        self.client_path = client_path
        self.template_path = template_path

        self.title("Proposal Writer")
        self.minsize(width=620, height=560)
        self.geometry("900x700")
        self.configure(padx=16, pady=16)

        scope_step = WizardStep(
            key="scopes",
            title="Scope Items",
            description="Select the scope items that will be included in the base services.",
            options=[],
        )
        assumption_options = [
            OptionItem(
                label=f"{'Optional - ' if assumption.is_additional else ''}{assumption.text}",
                value=assumption,
            )
            for assumption in data.assumptions
        ]
        assumption_step = WizardStep(
            key="assumptions",
            title="Assumptions",
            description="Choose the assumptions and exclusions that apply to this proposal.",
            options=assumption_options,
        )

        initial_steps: list[WizardStep] = []
        if data.clients:
            initial_steps.append(
                WizardStep(
                    key="client",
                    title="Client Request",
                    description="Select the client information pulled from Monday.com.",
                    options=[OptionItem(label=_client_label(client), value=client) for client in data.clients],
                    allow_multiple=False,
                )
            )

        initial_steps.extend(
            [
                WizardStep(
                    key="project_types",
                    title="Project Types",
                    description="Select every project type that applies to this opportunity.",
                    options=[OptionItem(label=pt.name, value=pt) for pt in data.project_types],
                ),
                WizardStep(
                    key="services",
                    title="Services",
                    description="Choose the services included in the scope.",
                    options=[OptionItem(label=svc.name, value=svc) for svc in data.services],
                ),
                WizardStep(
                    key="project_managers",
                    title="Project Managers",
                    description="Assign the project managers responsible for delivery.",
                    options=[
                        OptionItem(
                            label=f"{pm.name} ({pm.title})" if pm.title else pm.name,
                            value=pm,
                        )
                        for pm in data.project_managers
                    ],
                ),
                WizardStep(
                    key="scale",
                    title="Scale & Acres per Sheet",
                    description=(
                        "Pick the drawing scale and acres-per-sheet applicable to this proposal."
                    ),
                    options=[
                        OptionItem(
                            label=f"{misc.scale} - {misc.acres_per_sheet} acres/sheet", value=misc
                        )
                        for misc in data.misc_options
                    ],
                    allow_multiple=False,
                ),
                scope_step,
                assumption_step,
            ]
        )

        self.steps = initial_steps

        self.current_step_index = 0
        self.selections: dict[str, list[OptionItem]] = {step.key: [] for step in self.steps}
        self.active_options: list[OptionItem] = []
        self._scope_notice_shown = False

        default_assumptions = [
            option
            for option in assumption_options
            if not cast(Assumption, option.value).is_additional
        ]
        if default_assumptions:
            self.selections["assumptions"] = default_assumptions.copy()

        self._build_widgets()
        self._load_step(0)

    def _build_widgets(self) -> None:
        self.heading_label = ttk.Label(self, font=("Segoe UI", 14, "bold"))
        self.heading_label.pack(fill="x")

        self.description_label = ttk.Label(self, wraplength=580, justify="left")
        self.description_label.pack(fill="x", pady=(4, 12))

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(
            list_frame,
            selectmode=tk.MULTIPLE,
            height=12,
            exportselection=False,
        )
        self.listbox.pack(side=tk.LEFT, fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._update_detail_panel)
        self.listbox.bind("<Button-1>", self._toggle_listbox_selection, add="+")

        listbox_yview: Callable[..., None] = cast(
            Callable[..., None], getattr(self.listbox, "yview")
        )
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox_yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        detail_frame = ttk.LabelFrame(self, text="Details")
        detail_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.detail_text = tk.Text(detail_frame, height=8, wrap="word")
        self.detail_text.pack(fill="both", expand=True)
        self.detail_text.configure(state="disabled")

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", pady=(12, 0))

        self.back_button = ttk.Button(button_frame, text="Back", command=self._handle_back)
        self.back_button.pack(side=tk.LEFT)

        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(
            side=tk.RIGHT, padx=(8, 0)
        )

        self.next_button = ttk.Button(button_frame, text="Next", command=self._handle_next)
        self.next_button.pack(side=tk.RIGHT)

        self.status_label = ttk.Label(self, foreground="#555555")
        self.status_label.pack(fill="x", pady=(12, 0))
    def _load_step(self, index: int) -> None:  # pylint: disable=too-many-locals
        if index < 0 or index >= len(self.steps):
            return
        self.current_step_index = index
        step = self.steps[index]

        if step.key == "scopes":
            scope_options = self._build_scope_options()
            if not scope_options:
                self.selections["scopes"] = []
                if not self._scope_notice_shown:
                    messagebox.showinfo(
                        "Scope Library",
                        "No scope items match the selected services/project types. "
                        "Update your selections to include scope-driven services.",
                        parent=self,
                    )
                    self._scope_notice_shown = True
                self._load_step(index + 1)
                return
            step.options = scope_options

        self.heading_label.config(text=step.title)
        self.description_label.config(text=step.description)
        selectmode = tk.MULTIPLE if step.allow_multiple else tk.SINGLE
        if str(self.listbox["selectmode"]) != selectmode:
            self.listbox.configure(selectmode=selectmode)

        self.listbox.delete(0, tk.END)
        self.active_options = step.options
        for option in step.options:
            self.listbox.insert(tk.END, option.label)

        height = max(6, min(18, len(step.options)))
        self.listbox.configure(height=height)

        previous = self.selections.get(step.key, [])
        self.listbox.selection_clear(0, tk.END)
        if previous:
            labels = {opt.label for opt in previous}
            for idx, option in enumerate(step.options):
                if option.label in labels:
                    self.listbox.selection_set(idx)

        is_first = index == 0
        is_last = index == len(self.steps) - 1
        self.back_button.configure(state=tk.DISABLED if is_first else tk.NORMAL)
        self.next_button.config(text="Finish" if is_last else "Next")

        source_parts: list[str] = []
        if self.key_path:
            source_parts.append(self.key_path.name)
        if self.scope_path:
            source_parts.append(self.scope_path.name)
        source_label = " - ".join(source_parts) if source_parts else "demo data"
        status = (
            f"Step {index + 1} of {len(self.steps)} - {len(step.options)} option(s) - "
            f"Source: {source_label}"
        )
        self.status_label.config(text=status)
        self._set_detail_text("")

    def _handle_next(self) -> None:
        if not self._capture_selection():
            return
        if self.current_step_index == len(self.steps) - 1:
            context = self._build_context()
            SelectionSummary(self, context, self.template_path)
            return
        self._load_step(self.current_step_index + 1)

    def _handle_back(self) -> None:
        if self.current_step_index == 0:
            return
        self._load_step(self.current_step_index - 1)

    def _capture_selection(self) -> bool:
        step = self.steps[self.current_step_index]
        listbox_selection_attr = getattr(self.listbox, "curselection")
        listbox_selection = cast(Callable[[], tuple[int, ...]], listbox_selection_attr)
        indices = listbox_selection()
        if not indices:
            message = (
                "Please select at least one option to continue."
                if step.allow_multiple
                else "Please select an option to continue."
            )
            messagebox.showinfo(step.title, message, parent=self)
            return False
        self.selections[step.key] = [step.options[i] for i in indices]
        return True

    def _build_scope_options(self) -> list[OptionItem]:
        selected_services = {
            cast(ServiceType, option.value).id for option in self.selections.get("services", [])
        }
        selected_projects = {
            cast(ProjectType, option.value).id
            for option in self.selections.get("project_types", [])
        }
        options: list[OptionItem] = []
        for scope in self.data.scope_items:
            service_match = (
                not selected_services
                or scope.service_type_id in selected_services
                or scope.service_type_id in (None, 0)
            )
            project_match = (
                not selected_projects
                or scope.project_type_id in selected_projects
                or scope.project_type_id in (None, 0)
            )
            if service_match and project_match:
                label = scope.name
                if scope.fee is not None:
                    label = f"{label} ({format_currency(scope.fee)})"
                options.append(OptionItem(label=label, value=scope))
        return options

    def _update_detail_panel(self, *_event: tk.Event) -> None:  # pylint: disable=too-many-locals
        step = self.steps[self.current_step_index]
        curselection_attr = getattr(self.listbox, "curselection")
        indices = cast(tuple[int, ...], curselection_attr())
        if not indices:
            self._set_detail_text("")
            return
        option = step.options[indices[-1]]
        text = option.label

        if step.key == "project_types":
            project = cast(ProjectType, option.value)
            description = project.description or "No description available."
            text = f"{project.name}\n\n{description}"
        elif step.key == "services":
            service = cast(ServiceType, option.value)
            description = service.description or "No description available."
            text = f"{service.name}\n\n{description}"
        elif step.key == "project_managers":
            manager = cast(ProjectManager, option.value)
            title = f" ({manager.title})" if manager.title else ""
            text = f"{manager.name}{title}"
        elif step.key == "scopes":
            scope = cast(ScopeItem, option.value)
            details: list[str] = []
            if scope.units:
                details.append(scope.units)
            if scope.fee is not None:
                details.append(format_currency(scope.fee))
            header = scope.name
            if details:
                header = f"{header} ({' | '.join(details)})"
            text = f"{header}\n\n{scope.description}"
        elif step.key == "assumptions":
            assumption = cast(Assumption, option.value)
            prefix = "Optional\n\n" if assumption.is_additional else ""
            text = f"{prefix}{assumption.text}"
        elif step.key == "client":
            client = cast(ClientInfo, option.value)
            text = "\n".join(_client_summary_lines(client))

        self._set_detail_text(text)

    def _set_detail_text(self, value: str) -> None:
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", tk.END)
        if value:
            self.detail_text.insert("1.0", value)
        self.detail_text.configure(state="disabled")

    def _build_context(self) -> ProposalContext:
        def selected_values(key: str) -> list[object]:
            return [option.value for option in self.selections.get(key, [])]

        client_values = selected_values("client")
        client = cast(ClientInfo, client_values[0]) if client_values else None
        project_types = [cast(ProjectType, value) for value in selected_values("project_types")]
        services = [cast(ServiceType, value) for value in selected_values("services")]
        managers = [cast(ProjectManager, value) for value in selected_values("project_managers")]
        scopes = [
            ScopeSelection(scope=cast(ScopeItem, value))
            for value in selected_values("scopes")
        ]
        assumptions = [cast(Assumption, value) for value in selected_values("assumptions")]
        scale_values = selected_values("scale")
        scale = cast(MiscOption, scale_values[0]) if scale_values else self.data.misc_options[0]
        return ProposalContext(
            client=client,
            project_types=project_types,
            services=services,
            project_managers=managers,
            scale=scale,
            scopes=scopes,
            assumptions=assumptions,
        )

    def _toggle_listbox_selection(self, event: tk.Event) -> str | None:
        """Allow single clicks to toggle selection without modifier keys."""

        widget = cast(tk.Listbox, event.widget)
        nearest_fn = cast(Callable[[int], int], widget.nearest)  # type: ignore[attr-defined]
        includes_fn = cast(Callable[[int], bool], widget.selection_includes)  # type: ignore[attr-defined]
        index = nearest_fn(int(event.y))
        if index < 0:
            return "break"
        if includes_fn(index):
            widget.selection_clear(index)
        else:
            widget.selection_set(index)
        self._update_detail_panel()
        return "break"
class SelectionSummary(tk.Toplevel):
    """Final summary dialog that allows users to export the document."""

    def __init__(self, parent: tk.Misc, context: ProposalContext, template_path: Path) -> None:
        super().__init__(parent)
        self.context = context
        self.template_path = template_path
        self.title("Proposal Summary")
        self.configure(padx=16, pady=16)
        self.geometry("900x700")
        self.resizable(True, True)

        ttk.Label(
            self,
            text="Review the selections below, then generate the draft proposal from the template.",
            wraplength=560,
            justify="left",
        ).pack(fill="x")

        summary_container = ttk.Frame(self)
        summary_container.pack(fill="both", expand=True, pady=(8, 12))

        canvas = tk.Canvas(summary_container, borderwidth=0, highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill="both", expand=True)
        canvas_yview: Callable[..., None] = cast(Callable[..., None], getattr(canvas, "yview"))
        scrollbar = ttk.Scrollbar(summary_container, orient="vertical", command=canvas_yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        self._summary_frame = ttk.Frame(canvas)
        frame_window = canvas.create_window((0, 0), window=self._summary_frame, anchor="nw")

        def _sync_scroll_region(_event: tk.Event) -> None:  # pragma: no cover - GUI callback
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_frame_width(event: tk.Event) -> None:  # pragma: no cover - GUI callback
            canvas.itemconfigure(frame_window, width=event.width)

        self._summary_frame.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _sync_frame_width)

        self._add_section(
            self._summary_frame,
            "Project Types",
            [pt.name for pt in context.project_types] or ["(None)"],
        )
        self._add_section(
            self._summary_frame,
            "Services",
            [service.name for service in context.services] or ["(None)"],
        )
        self._add_section(
            self._summary_frame,
            "Project Managers",
            [
                pm.name if not pm.title else f"{pm.name} ({pm.title})"
                for pm in context.project_managers
            ]
            or ["(None)"],
        )
        self._add_section(
            self._summary_frame,
            "Scale",
            [f"{context.scale.scale} ({context.scale.acres_per_sheet} acres/sheet)"],
        )

        fee_by_scope_id = {row.scope.id: row.total for row in context.fee_rows}
        scope_lines: list[str] = []
        for selection in context.scopes:
            amount = fee_by_scope_id.get(selection.scope.id, selection.scope.fee)
            amount_label = format_currency(amount) if amount is not None else "TBD"
            breakout_label = " + Fee Breakdown" if selection.show_fee_breakout else ""
            scope_lines.append(
                f"{selection.scope.name} ({selection.status}{breakout_label}) - {amount_label}"
            )
        if not scope_lines:
            scope_lines = ["(No scopes selected)"]
        self._add_section(self._summary_frame, "Scopes", scope_lines)

        assumption_lines = [assumption.text for assumption in context.assumptions] or ["(None)"]
        self._add_section(self._summary_frame, "Assumptions", assumption_lines)

        total_fee = 0.0
        for selection in context.scopes:
            if not _scope_included_in_base(selection.status):
                continue
            amount = fee_by_scope_id.get(selection.scope.id, selection.scope.fee)
            if amount is not None:
                total_fee += amount
        ttk.Label(
            self,
            text=f"Estimated Total (excluding TBD items): {format_currency(total_fee)}",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            self,
            text=f"Template: {template_path}",
            foreground="#555555",
        ).pack(anchor="w")

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", pady=(12, 0))

        ttk.Button(button_frame, text="Close", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Generate Proposal", command=self._generate_document).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

    def _add_section(self, parent: tk.Misc, title: str, lines: list[str]) -> None:
        group = ttk.LabelFrame(parent, text=title)
        group.pack(fill="x", pady=4)
        for line in lines:
            ttk.Label(group, text=f"- {line}", wraplength=540, justify="left").pack(
                anchor="w", padx=8, pady=2
            )

    def _generate_document(self) -> None:
        target = filedialog.asksaveasfilename(
            title="Save Proposal Letter As",
            defaultextension=".docx",
            initialfile=f"Proposal_{datetime.date.today().isoformat()}",
            filetypes=[
                ("Word Document", "*.docx"),
                ("All Files", "*.*"),
            ],
        )
        if not target:
            return
        output_path = Path(target)
        try:
            replacements = _build_placeholder_replacements(self.context)
            saved_path = generate_proposal_document(
                self.template_path,
                output_path,
                self.context,
                replacements,
            )
        except (OSError, ValueError) as exc:
            messagebox.showerror(
                "Proposal Writer",
                f"Failed to generate the proposal: {exc}",
                parent=self,
            )
            return
        pdf_note = ""
        try:
            pdf_path = convert_docx_to_pdf(saved_path, saved_path)
            pdf_note = f"\nPDF generated at:\n{pdf_path}"
        except Exception as exc:  # noqa: BLE001 - surface automation issues to the user
            pdf_note = f"\nPDF export skipped: {exc}"
        messagebox.showinfo(
            "Proposal Writer",
            f"Proposal generated successfully:\n{saved_path}{pdf_note}",
            parent=self,
        )


class ProjectInfoWindow(tk.Toplevel):
    """Collect project metrics, project type, services, and client."""

    def __init__(
        self,
        parent: tk.Misc,
        data: ProposalData,
        on_next: Callable[
            [
                ClientInfo | None,
                list[ProjectType],
                list[ServiceType],
                ProjectMetrics,
                list[ProjectManager],
                list[MiscOption],
                str,
                str,
            ],
            None,
        ],
        on_manage_files: Callable[[], bool] | None = None,
        selected_client: ClientInfo | None = None,
        selected_project_names: list[str] | None = None,
        selected_scale_labels: list[str] | None = None,
        selected_service_indices: list[int] | None = None,
        selected_pm_names: list[str] | None = None,
    ) -> None:
        super().__init__(parent)
        self.data = data
        self.on_next = on_next
        self.on_manage_files = on_manage_files
        self.title("Proposal Inputs")
        self.configure(padx=16, pady=16)
        self.resizable(True, True)
        self.geometry("800x650")

        if self.on_manage_files:
            menu = tk.Menu(self)
            file_menu = tk.Menu(menu, tearoff=0)
            file_menu.add_command(label="Manage Files", command=self._handle_manage_files)
            menu.add_cascade(label="File", menu=file_menu)
            self.config(menu=menu)

        ttk.Label(self, text="Enter the project details below.").pack(anchor="w")

        form = ttk.Frame(self)
        form.pack(fill="x", pady=(8, 12))

        if self.on_manage_files:
            manage_frame = ttk.Frame(self)
            manage_frame.pack(fill="x", pady=(0, 8))
            ttk.Button(
                manage_frame,
                text="Update Source Files",
                command=self._handle_manage_files,
            ).pack(side=tk.RIGHT)

        self._client_var = tk.StringVar()
        self._project_type_var = tk.StringVar()
        self._scale_var = tk.StringVar()
        self._developed_area_var = tk.StringVar()
        self._roadway_length_var = tk.StringVar()
        self._pm_var = tk.StringVar()
        self._pm_title_var = tk.StringVar()

        self._client_map = {
            _client_label(client): client for client in data.clients
        }
        self._project_type_map = {pt.name: pt for pt in data.project_types}
        self._scale_map = {
            f"{misc.scale} ({misc.acres_per_sheet} acres/sheet)": misc
            for misc in data.misc_options
        }
        self._pm_map = {pm.name: pm for pm in data.project_managers}

        ttk.Label(form, text="Client Request").grid(row=0, column=0, sticky="w", pady=4)
        client_combo = ttk.Combobox(
            form,
            textvariable=self._client_var,
            values=list(self._client_map.keys()),
            state="readonly",
            width=40,
        )
        client_combo.grid(row=0, column=1, sticky="ew", pady=4)
        if self._client_map:
            client_combo.current(0)
        if selected_client:
            label = _client_label(selected_client)
            if label in self._client_map:
                client_combo.set(label)

        ttk.Label(form, text="Project Types").grid(row=1, column=0, sticky="w", pady=4)
        project_frame = ttk.Frame(form)
        project_frame.grid(row=1, column=1, sticky="ew", pady=4)
        self._project_listbox = tk.Listbox(
            project_frame,
            selectmode=tk.MULTIPLE,
            height=4,
            exportselection=False,
        )
        self._project_listbox.pack(side=tk.LEFT, fill="both", expand=True)
        for name in self._project_type_map.keys():
            self._project_listbox.insert(tk.END, name)
        if selected_project_names:
            name_set = set(selected_project_names)
            for idx, name in enumerate(self._project_type_map.keys()):
                if name in name_set:
                    self._project_listbox.selection_set(idx)
        project_yview: Callable[..., None] = cast(
            Callable[..., None], getattr(self._project_listbox, "yview")
        )
        project_scroll = ttk.Scrollbar(project_frame, orient="vertical", command=project_yview)
        project_scroll.pack(side=tk.RIGHT, fill="y")
        self._project_listbox.configure(yscrollcommand=project_scroll.set)

        ttk.Label(form, text="Plan Sheet Scales").grid(row=2, column=0, sticky="w", pady=4)
        scale_frame = ttk.Frame(form)
        scale_frame.grid(row=2, column=1, sticky="ew", pady=4)
        self._scale_listbox = tk.Listbox(
            scale_frame,
            selectmode=tk.MULTIPLE,
            height=4,
            exportselection=False,
        )
        self._scale_listbox.pack(side=tk.LEFT, fill="both", expand=True)
        for name in self._scale_map.keys():
            self._scale_listbox.insert(tk.END, name)
        if selected_scale_labels:
            label_set = set(selected_scale_labels)
            for idx, label in enumerate(self._scale_map.keys()):
                if label in label_set:
                    self._scale_listbox.selection_set(idx)
        scale_yview: Callable[..., None] = cast(
            Callable[..., None], getattr(self._scale_listbox, "yview")
        )
        scale_scroll = ttk.Scrollbar(scale_frame, orient="vertical", command=scale_yview)
        scale_scroll.pack(side=tk.RIGHT, fill="y")
        self._scale_listbox.configure(yscrollcommand=scale_scroll.set)

        ttk.Label(form, text="Project Manager (PM)").grid(
            row=3, column=0, sticky="w", pady=4
        )
        pm_frame = ttk.Frame(form)
        pm_frame.grid(row=3, column=1, sticky="ew", pady=4)
        self._pm_listbox = tk.Listbox(
            pm_frame,
            selectmode=tk.MULTIPLE,
            height=4,
            exportselection=False,
        )
        self._pm_listbox.pack(side=tk.LEFT, fill="both", expand=True)
        for name in self._pm_map.keys():
            self._pm_listbox.insert(tk.END, name)
        if selected_pm_names:
            name_set = set(selected_pm_names)
            for idx, name in enumerate(self._pm_map.keys()):
                if name in name_set:
                    self._pm_listbox.selection_set(idx)
        pm_yview: Callable[..., None] = cast(
            Callable[..., None], getattr(self._pm_listbox, "yview")
        )
        pm_scroll = ttk.Scrollbar(pm_frame, orient="vertical", command=pm_yview)
        pm_scroll.pack(side=tk.RIGHT, fill="y")
        self._pm_listbox.configure(yscrollcommand=pm_scroll.set)

        ttk.Label(form, text="Project Manager Title").grid(
            row=4, column=0, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self._pm_title_var, width=30, state="readonly").grid(
            row=4, column=1, sticky="w", pady=4
        )

        ttk.Label(form, text="Site's Developed Area (acres)").grid(
            row=5, column=0, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self._developed_area_var, width=20).grid(
            row=5, column=1, sticky="w", pady=4
        )

        ttk.Label(form, text="Length of Roadway (ft)").grid(
            row=6, column=0, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self._roadway_length_var, width=20).grid(
            row=6, column=1, sticky="w", pady=4
        )

        ttk.Label(self, text="Select Services").pack(anchor="w")
        service_frame = ttk.Frame(self)
        service_frame.pack(fill="both", pady=(4, 12))

        self._service_listbox = tk.Listbox(
            service_frame,
            selectmode=tk.MULTIPLE,
            height=8,
            exportselection=False,
        )
        self._service_listbox.pack(side=tk.LEFT, fill="both", expand=True)
        for service in data.services:
            self._service_listbox.insert(tk.END, service.name)
        if selected_service_indices:
            for idx in selected_service_indices:
                self._service_listbox.selection_set(idx)

        listbox_yview: Callable[..., None] = cast(
            Callable[..., None], getattr(self._service_listbox, "yview")
        )
        scrollbar = ttk.Scrollbar(service_frame, orient="vertical", command=listbox_yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self._service_listbox.configure(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Next", command=self._handle_next).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

        def _selected_pm_titles() -> str:
            curselection_fn = cast(
                Callable[[], tuple[int, ...]], getattr(self._pm_listbox, "curselection")
            )
            indices = curselection_fn()
            titles: list[str] = []
            for idx in indices:
                get_fn = cast(Callable[[int], str], getattr(self._pm_listbox, "get"))
                name = get_fn(idx)
                pm = self._pm_map.get(name)
                if pm and pm.title:
                    titles.append(pm.title)
            return ", ".join(titles)

        def _sync_pm_title(*_args: object) -> None:
            self._pm_title_var.set(_selected_pm_titles())

        self._pm_listbox.bind("<<ListboxSelect>>", _sync_pm_title)

        def _sync_pm_from_client(*_args: object) -> None:
            client = self._client_map.get(self._client_var.get())
            if not client or not client.pm:
                return
            if client.pm in self._pm_map:
                try:
                    idx = list(self._pm_map.keys()).index(client.pm)
                except ValueError:
                    return
                self._pm_listbox.selection_clear(0, tk.END)
                self._pm_listbox.selection_set(idx)
                _sync_pm_title()

        self._client_var.trace_add("write", _sync_pm_from_client)
        if self._client_map:
            _sync_pm_from_client()

    def _handle_next(self) -> None:
        project_indices = cast(
            Callable[[], tuple[int, ...]], getattr(self._project_listbox, "curselection")
        )()
        if not project_indices:
            messagebox.showinfo("Project Types", "Select at least one project type.", parent=self)
            return

        scale_indices = cast(
            Callable[[], tuple[int, ...]], getattr(self._scale_listbox, "curselection")
        )()
        if not scale_indices:
            messagebox.showinfo(
                "Plan Sheet Scales", "Select at least one plan sheet scale.", parent=self
            )
            return

        developed_area = _safe_float(self._developed_area_var.get())
        if developed_area is None:
            messagebox.showinfo(
                "Developed Area", "Enter a numeric value for developed area.", parent=self
            )
            return

        roadway_length = _safe_float(self._roadway_length_var.get() or "0") or 0.0

        average_multiplier = self.data.average_multiplier

        selected_services: list[ServiceType] = []
        curselection_fn = cast(
            Callable[[], tuple[int, ...]], getattr(self._service_listbox, "curselection")
        )
        service_indices = curselection_fn()
        for index in service_indices:
            selected_services.append(self.data.services[index])
        if not selected_services:
            messagebox.showinfo("Services", "Select at least one service.", parent=self)
            return

        project_get = cast(Callable[[int], str], getattr(self._project_listbox, "get"))
        selected_projects = [
            self._project_type_map[project_get(i)] for i in project_indices
        ]
        scale_get = cast(Callable[[int], str], getattr(self._scale_listbox, "get"))
        selected_scales = [self._scale_map[scale_get(i)] for i in scale_indices]
        scale = selected_scales[0]
        client = self._client_map.get(self._client_var.get())
        metrics = ProjectMetrics(
            developed_area=developed_area,
            roadway_length=roadway_length,
            scale=scale,
            average_multiplier=average_multiplier,
        )
        pm_indices = cast(
            Callable[[], tuple[int, ...]], getattr(self._pm_listbox, "curselection")
        )()
        if not pm_indices:
            messagebox.showinfo("PM", "Please select a project manager.", parent=self)
            return
        pm_get = cast(Callable[[int], str], getattr(self._pm_listbox, "get"))
        selected_pms = [self._pm_map[pm_get(i)] for i in pm_indices]
        pm_name = ", ".join(pm.name for pm in selected_pms)
        pm_title = ", ".join(pm.title for pm in selected_pms if pm.title)
        self.on_next(
            client,
            selected_projects,
            selected_services,
            metrics,
            selected_pms,
            selected_scales,
            pm_name,
            pm_title,
        )
        self.destroy()

    def _handle_manage_files(self) -> None:
        if not self.on_manage_files:
            return
        reloaded = self.on_manage_files()
        if reloaded:
            self.destroy()


class AssumptionsWindow(tk.Toplevel):
    """Select assumptions from the Civil Engineering Services workbook."""

    def __init__(
        self,
        parent: tk.Misc,
        assumptions: list[Assumption],
        on_next: Callable[[list[Assumption]], None],
        on_back: Callable[[], None],
        selected: list[Assumption] | None = None,
    ) -> None:
        super().__init__(parent)
        self.assumptions = assumptions
        self.on_next = on_next
        self.on_back = on_back
        self.title("Assumptions")
        self.configure(padx=16, pady=16)
        self.resizable(True, True)
        self.geometry("800x650")

        ttk.Label(self, text="Select the assumptions to include.").pack(anchor="w")

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, pady=(8, 12))

        self._listbox = tk.Listbox(
            list_frame,
            selectmode=tk.MULTIPLE,
            height=10,
            exportselection=False,
        )
        self._listbox.pack(side=tk.LEFT, fill="both", expand=True)
        for assumption in assumptions:
            label = f"{'Optional - ' if assumption.is_additional else ''}{assumption.text}"
            self._listbox.insert(tk.END, label)

        listbox_yview: Callable[..., None] = cast(
            Callable[..., None], getattr(self._listbox, "yview")
        )
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox_yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self._listbox.configure(yscrollcommand=scrollbar.set)

        if selected:
            selected_texts = {assumption.text for assumption in selected}
            for idx, assumption in enumerate(assumptions):
                if assumption.text in selected_texts:
                    self._listbox.selection_set(idx)
        else:
            for idx, assumption in enumerate(assumptions):
                if not assumption.is_additional:
                    self._listbox.selection_set(idx)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="Back", command=self._handle_back).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Next", command=self._handle_next).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

    def _handle_next(self) -> None:
        curselection_fn = cast(Callable[[], tuple[int, ...]], getattr(self._listbox, "curselection"))
        indices = curselection_fn()
        selected = [self.assumptions[i] for i in indices]
        if not selected:
            messagebox.showinfo(
                "Assumptions",
                "Select at least one assumption to continue.",
                parent=self,
            )
            return
        self.on_next(selected)
        self.destroy()

    def _handle_back(self) -> None:
        self.on_back()
        self.destroy()


class ScopeWindow(tk.Toplevel):
    """Select scopes filtered by project type + services."""

    def __init__(
        self,
        parent: tk.Misc,
        scopes: list[ScopeItem],
        project_types: list[ProjectType],
        services: list[ServiceType],
        on_next: Callable[[list[ScopeSelection]], None],
        on_back: Callable[[], None],
        selected_scope_states: dict[int, tuple[str, bool]] | None = None,
    ) -> None:
        super().__init__(parent)
        self.on_next = on_next
        self.on_back = on_back
        self.title("Scopes")
        self.configure(padx=16, pady=16)
        self.resizable(True, True)
        self.geometry("820x680")

        ttk.Label(
            self,
            text=(
                "Select the scopes to include, mark Required or Requested By Client, "
                "and optionally show them in the Fee Breakdown table."
            ),
            wraplength=720,
            justify="left",
        ).pack(anchor="w")

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, pady=(8, 12))

        service_ids = {service.id for service in services}
        project_ids = {project.id for project in project_types}
        filtered_scopes = [
            scope
            for scope in scopes
            if (scope.service_type_id in service_ids or scope.service_type_id in (None, 0))
            and (scope.project_type_id in project_ids or scope.project_type_id in (None, 0))
        ]
        self._scopes = filtered_scopes

        canvas = tk.Canvas(list_frame, borderwidth=0, highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill="both", expand=True)
        canvas_yview: Callable[..., None] = cast(Callable[..., None], getattr(canvas, "yview"))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas_yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        list_container = ttk.Frame(canvas)
        frame_window = canvas.create_window((0, 0), window=list_container, anchor="nw")

        def _sync_scroll_region(_event: tk.Event) -> None:  # pragma: no cover - GUI callback
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_frame_width(event: tk.Event) -> None:  # pragma: no cover - GUI callback
            canvas.itemconfigure(frame_window, width=event.width)

        list_container.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _sync_frame_width)
        canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))
        canvas.bind("<Button-4>", lambda _event: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda _event: canvas.yview_scroll(1, "units"))

        header = ttk.Frame(list_container)
        header.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 6))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Scope", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(header, text="Status", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=1, sticky="e", padx=(8, 0)
        )
        ttk.Label(header, text="Fee Breakdown", font=("Segoe UI", 10, "bold"), anchor="center").grid(
            row=0, column=2, sticky="ew", padx=(8, 0)
        )

        self._rows: list[dict[str, object]] = []
        selected_scope_states = selected_scope_states or {}
        for idx, scope in enumerate(filtered_scopes, start=1):
            row = ttk.Frame(list_container)
            row.grid(row=idx, column=0, sticky="ew", padx=2, pady=2)
            row.columnconfigure(0, weight=1)

            stored_status, stored_breakout = selected_scope_states.get(
                scope.id, (SCOPE_STATUS_REQUIRED, False)
            )
            status_default = stored_status
            selected_default = scope.id in selected_scope_states
            selected_var = tk.BooleanVar(value=selected_default)
            status_var = tk.StringVar(value=status_default)
            fee_breakout_var = tk.BooleanVar(value=stored_breakout)

            ttk.Checkbutton(row, text=scope.name, variable=selected_var).grid(
                row=0, column=0, sticky="w"
            )
            status_box = ttk.Combobox(
                row,
                textvariable=status_var,
                values=SCOPE_STATUS_OPTIONS,
                state="readonly",
                width=16,
            )
            status_box.grid(row=0, column=1, sticky="e", padx=(8, 0))
            ttk.Checkbutton(row, variable=fee_breakout_var).grid(
                row=0, column=2, sticky="e", padx=(8, 0)
            )

            self._rows.append(
                {
                    "scope": scope,
                    "selected_var": selected_var,
                    "status_var": status_var,
                    "fee_breakout_var": fee_breakout_var,
                }
            )

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="Select All", command=self._handle_select_all).pack(
            side=tk.LEFT
        )
        ttk.Button(
            button_frame, text="Select All Fee Breakdown", command=self._handle_select_all_fee_breakdown
        ).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(button_frame, text="Back", command=self._handle_back).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Next", command=self._handle_next).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

    def _handle_next(self) -> None:
        selected: list[ScopeSelection] = []
        for row in self._rows:
            selected_var = cast(tk.BooleanVar, row["selected_var"])
            if not selected_var.get():
                continue
            scope = cast(ScopeItem, row["scope"])
            status_var = cast(tk.StringVar, row["status_var"])
            fee_breakout_var = cast(tk.BooleanVar, row["fee_breakout_var"])
            status = status_var.get() or SCOPE_STATUS_REQUIRED
            selected.append(
                ScopeSelection(
                    scope=scope,
                    status=status,
                    show_fee_breakout=fee_breakout_var.get(),
                )
            )
        if not selected:
            messagebox.showinfo(
                "Scopes", "Select at least one scope to continue.", parent=self
            )
            return
        self.on_next(selected)
        self.destroy()

    def _handle_select_all(self) -> None:
        for row in self._rows:
            selected_var = cast(tk.BooleanVar, row["selected_var"])
            selected_var.set(True)

    def _handle_select_all_fee_breakdown(self) -> None:
        for row in self._rows:
            fee_breakout_var = cast(tk.BooleanVar, row["fee_breakout_var"])
            fee_breakout_var.set(True)

    def _handle_back(self) -> None:
        self.on_back()
        self.destroy()


class FeeRefWindow(tk.Toplevel):
    """Edit Fee_Ref rate/hour inputs per project type."""

    def __init__(
        self,
        parent: tk.Misc,
        fee_ref_blocks: dict[str, FeeRefBlock],
        project_types: list[ProjectType],
        average_multiplier: float,
        on_next: Callable[[dict[str, float], float], None],
        on_back: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.on_next = on_next
        self.on_back = on_back
        self.title("Fee_Ref Editor")
        self.configure(padx=16, pady=16)
        self.resizable(True, True)
        self.geometry("900x700")

        ttk.Label(
            self,
            text="Review and edit rates/hours for each project type.",
            wraplength=760,
            justify="left",
        ).pack(anchor="w")

        multiplier_frame = ttk.Frame(self)
        multiplier_frame.pack(anchor="w", pady=(8, 4))
        ttk.Label(multiplier_frame, text="Average Multiplier").pack(side=tk.LEFT)
        self._avg_multiplier_var = tk.StringVar(
            value=f"{average_multiplier:.2f}" if average_multiplier else "1.00"
        )
        ttk.Entry(
            multiplier_frame,
            textvariable=self._avg_multiplier_var,
            width=10,
        ).pack(side=tk.LEFT, padx=(8, 0))

        self._entries: dict[str, list[dict[str, object]]] = {}
        self._total_vars: dict[str, tk.StringVar] = {}

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, pady=(8, 12))

        selected_names = [pt.name for pt in project_types]
        for name in selected_names:
            block = fee_ref_blocks.get(name)
            if not block:
                continue
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=name)

            header = ttk.Frame(frame)
            header.pack(fill="x", pady=(4, 6))
            column_widths = [30, 12, 12, 10, 14]
            for idx, (title, width) in enumerate(
                zip(["Position", "Rate", "Hours", "%", "Cost"], column_widths)
            ):
                ttk.Label(header, text=title, width=width, anchor="w").grid(
                    row=0, column=idx, sticky="w"
                )

            rows_container = ttk.Frame(frame)
            rows_container.pack(fill="both", expand=True)

            row_entries: list[dict[str, object]] = []
            for row_idx, position in enumerate(block.positions):
                ttk.Label(
                    rows_container,
                    text=position.title,
                    width=column_widths[0],
                    anchor="w",
                ).grid(
                    row=row_idx, column=0, sticky="w", pady=2
                )
                rate_var = tk.StringVar(value=str(position.rate))
                hours_var = tk.StringVar(value=str(position.hours))
                multiplier_value = average_multiplier or 1.0
                cost_rate = position.rate / multiplier_value
                percent_var = tk.StringVar(value="")
                cost_var = tk.StringVar(value=format_currency(cost_rate))
                ttk.Entry(rows_container, textvariable=rate_var, width=column_widths[1]).grid(
                    row=row_idx, column=1, sticky="w", pady=2
                )
                ttk.Entry(rows_container, textvariable=hours_var, width=column_widths[2]).grid(
                    row=row_idx, column=2, sticky="w", pady=2
                )
                ttk.Label(
                    rows_container,
                    textvariable=percent_var,
                    width=column_widths[3],
                    anchor="w",
                ).grid(
                    row=row_idx, column=3, sticky="w", pady=2
                )
                ttk.Label(
                    rows_container,
                    textvariable=cost_var,
                    width=column_widths[4],
                    anchor="w",
                ).grid(
                    row=row_idx, column=4, sticky="w", pady=2
                )
                row_entries.append(
                    {
                        "rate_var": rate_var,
                        "hours_var": hours_var,
                        "percent_var": percent_var,
                        "cost_var": cost_var,
                    }
                )

            total_var = tk.StringVar(value=format_currency(0.0))
            self._total_vars[name] = total_var
            ttk.Label(frame, text="Typical Cost Per Sheet", font=("Segoe UI", 10, "bold")).pack(
                anchor="w", pady=(8, 0)
            )
            ttk.Label(frame, textvariable=total_var, font=("Segoe UI", 10, "bold")).pack(
                anchor="w"
            )

            self._entries[name] = row_entries
            self._recalculate_totals(name)

            def _make_recalc(key: str) -> Callable[..., None]:
                def _recalc(*_args: object) -> None:
                    self._recalculate_totals(key)

                return _recalc

            recalc_callback = _make_recalc(name)
            for row in row_entries:
                cast(tk.StringVar, row["rate_var"]).trace_add("write", recalc_callback)
                cast(tk.StringVar, row["hours_var"]).trace_add("write", recalc_callback)

        self._avg_multiplier_var.trace_add("write", self._handle_multiplier_change)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="Back", command=self._handle_back).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Next", command=self._handle_next).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

    def _recalculate_totals(self, key: str) -> None:
        multiplier = _safe_float(self._avg_multiplier_var.get())
        if multiplier is None or multiplier == 0:
            multiplier = 1.0
        rows = self._entries.get(key, [])
        total_hours = sum(
            _safe_float(cast(tk.StringVar, row["hours_var"]).get()) or 0.0 for row in rows
        )
        if total_hours == 0:
            total_hours = 1.0
        total = 0.0
        for row in rows:
            rate = _safe_float(cast(tk.StringVar, row["rate_var"]).get()) or 0.0
            hours = _safe_float(cast(tk.StringVar, row["hours_var"]).get()) or 0.0
            cost_rate = rate / multiplier
            row_total = rate * hours
            cast(tk.StringVar, row["cost_var"]).set(format_currency(cost_rate))
            percent = hours / total_hours
            cast(tk.StringVar, row["percent_var"]).set(f"{percent:.1%}")
            total += row_total
        total_var = self._total_vars.get(key)
        if total_var:
            total_var.set(format_currency(total))

    def _handle_multiplier_change(self, *_args: object) -> None:
        for key in self._entries:
            self._recalculate_totals(key)

    def _handle_next(self) -> None:
        multiplier = _safe_float(self._avg_multiplier_var.get())
        if multiplier is None or multiplier == 0:
            messagebox.showinfo(
                "Average Multiplier",
                "Enter a numeric value for the average multiplier.",
                parent=self,
            )
            return
        totals: dict[str, float] = {}
        for key, rows in self._entries.items():
            total = 0.0
            for row in rows:
                rate = _safe_float(cast(tk.StringVar, row["rate_var"]).get())
                hours = _safe_float(cast(tk.StringVar, row["hours_var"]).get())
                if rate is None or hours is None:
                    messagebox.showinfo(
                        "Fee_Ref",
                        "All rate and hour values must be numeric.",
                        parent=self,
                    )
                    return
                total += rate * hours
            totals[key] = total
        self.on_next(totals, multiplier)
        self.destroy()

    def _handle_back(self) -> None:
        self.on_back()
        self.destroy()


class FileManagerWindow(tk.Toplevel):
    """Manage source file locations for the proposal builder."""

    def __init__(
        self,
        parent: tk.Misc,
        initial_paths: dict[str, Path | None],
        on_save: Callable[[dict[str, Path]], bool],
    ) -> None:
        super().__init__(parent)
        self.on_save = on_save
        self.updated = False
        self.title("File Manager")
        self.configure(padx=16, pady=16)
        self.resizable(False, False)

        ttk.Label(
            self,
            text="Select the source files to use. These will stay active until updated again.",
            wraplength=520,
            justify="left",
        ).pack(anchor="w")

        self._vars: dict[str, tk.StringVar] = {
            "data": tk.StringVar(value=str(initial_paths.get("data") or "")),
            "scope_data": tk.StringVar(value=str(initial_paths.get("scope_data") or "")),
            "client_data": tk.StringVar(value=str(initial_paths.get("client_data") or "")),
            "template": tk.StringVar(value=str(initial_paths.get("template") or "")),
        }

        form = ttk.Frame(self)
        form.pack(fill="x", pady=(12, 12))

        self._add_picker(
            form,
            row=0,
            label="Proposal Creator Key.xlsx",
            var=self._vars["data"],
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        self._add_picker(
            form,
            row=1,
            label="Civil Engineering Services.xlsx",
            var=self._vars["scope_data"],
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        self._add_picker(
            form,
            row=2,
            label="Monday Export.xlsx",
            var=self._vars["client_data"],
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        self._add_picker(
            form,
            row=3,
            label="Proposal Template (.dotx/.docx)",
            var=self._vars["template"],
            filetypes=[("Word Files", "*.dotx;*.docx"), ("All Files", "*.*")],
        )

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x")
        ttk.Button(
            button_frame,
            text="Export Source Copies",
            command=self._handle_export,
        ).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Save", command=self._handle_save).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

    def _add_picker(
        self,
        parent: tk.Misc,
        row: int,
        label: str,
        var: tk.StringVar,
        filetypes: list[tuple[str, str]],
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var, width=60).grid(
            row=row, column=1, sticky="ew", pady=4
        )
        ttk.Button(
            parent,
            text="Browse",
            command=lambda: self._browse(var, filetypes),
        ).grid(row=row, column=2, sticky="w", padx=(8, 0), pady=4)
        parent.grid_columnconfigure(1, weight=1)

    def _browse(self, var: tk.StringVar, filetypes: list[tuple[str, str]]) -> None:
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            var.set(filename)

    def _handle_save(self) -> None:
        paths: dict[str, Path] = {}
        for key, var in self._vars.items():
            value = var.get().strip()
            if not value:
                messagebox.showinfo("File Manager", "All file paths are required.", parent=self)
                return
            paths[key] = Path(value)

        if self.on_save(paths):
            self.updated = True
            self.destroy()

    def _handle_export(self) -> None:
        sources: dict[str, Path] = {}
        for key in ("data", "scope_data", "client_data", "template"):
            value = self._vars[key].get().strip()
            if not value:
                messagebox.showinfo(
                    "File Manager",
                    "Select the three Excel files and template before exporting.",
                    parent=self,
                )
                return
            sources[key] = Path(value)

        destination = filedialog.askdirectory(title="Choose Export Folder")
        if not destination:
            return

        try:
            cache_dir = _RESOURCE_ROOTS[1] / "export_cache"
            exported, used_cache = _export_excel_iteration(
                sources,
                Path(destination),
                cache_dir,
            )
        except (OSError, ValueError) as exc:
            messagebox.showerror("File Manager", f"Export failed: {exc}", parent=self)
            return

        cache_note = ""
        if used_cache:
            labels = ", ".join(used_cache)
            cache_note = f"\nUsed cached copies for: {labels}"
        messagebox.showinfo(
            "File Manager",
            f"Exported {len(exported)} source copies to:\n{destination}{cache_note}",
            parent=self,
        )

class FeeTableWindow(tk.Toplevel):
    """Editable fee table for scope items."""

    def __init__(
        self,
        parent: tk.Misc,
        scopes: list[ScopeSelection],
        services: list[ServiceType],
        metrics: ProjectMetrics,
        on_next: Callable[[list[FeeRow], dict[int, float]], None],
        on_back: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.scopes = scopes
        self.services = services
        self.metrics = metrics
        self.on_next = on_next
        self.on_back = on_back
        self.title("Fee Table")
        self.configure(padx=16, pady=16)
        self.resizable(True, True)
        self.geometry("900x700")

        ttk.Label(
            self,
            text="Review and adjust the fee table before generating the proposal.",
            wraplength=560,
            justify="left",
        ).pack(anchor="w")

        header = ttk.Frame(self)
        header.pack(fill="x", pady=(8, 4))
        column_widths = [34, 8, 10, 12, 12]
        for idx, (title, width) in enumerate(
            zip(["Scope", "Unit", "Units", "Unit Cost", "Total"], column_widths)
        ):
            ttk.Label(header, text=title, width=width, anchor="w").grid(row=0, column=idx)

        class _FeeRowEntry(TypedDict):
            scope: ScopeItem
            status: str
            unit: str
            units_var: tk.StringVar
            unit_cost_var: tk.StringVar
            total_var: tk.StringVar

        self._row_entries: list[_FeeRowEntry] = []

        table_container = ttk.Frame(self)
        table_container.pack(fill="both", expand=True)

        canvas = tk.Canvas(table_container, borderwidth=0, highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill="both", expand=True)
        canvas_yview: Callable[..., None] = cast(Callable[..., None], getattr(canvas, "yview"))
        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=canvas_yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        table_frame = ttk.Frame(canvas)
        frame_window = canvas.create_window((0, 0), window=table_frame, anchor="nw")

        def _sync_scroll_region(_event: tk.Event) -> None:  # pragma: no cover - GUI callback
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_frame_width(event: tk.Event) -> None:  # pragma: no cover - GUI callback
            canvas.itemconfigure(frame_window, width=event.width)

        table_frame.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _sync_frame_width)

        acres_per_sheet = _safe_float(self.metrics.scale.acres_per_sheet)
        acres_per_sheet = acres_per_sheet or 1.0

        for row_idx, selection in enumerate(scopes):
            scope = selection.scope
            unit = (scope.units or "EA").strip()
            units_value = float(
                _calculate_units(
                    unit,
                    self.metrics.developed_area,
                    acres_per_sheet,
                    self.metrics.roadway_length,
                )
            )

            unit_cost = scope.fee or 0.0
            total_value = units_value * unit_cost

            label = scope.name if selection.status == SCOPE_STATUS_REQUIRED else f"{scope.name} ({selection.status})"
            if selection.show_fee_breakout:
                label = f"{label} [Fee Breakdown]"
            ttk.Label(
                table_frame,
                text=label,
                wraplength=240,
                width=column_widths[0],
                anchor="w",
            ).grid(
                row=row_idx, column=0, sticky="w", pady=2
            )
            ttk.Label(table_frame, text=unit, width=column_widths[1], anchor="w").grid(
                row=row_idx, column=1, sticky="w", pady=2
            )

            units_var = tk.StringVar(value=str(units_value))
            unit_cost_var = tk.StringVar(value=str(unit_cost))
            total_var = tk.StringVar(value=format_currency(total_value))

            ttk.Entry(table_frame, textvariable=units_var, width=column_widths[2]).grid(
                row=row_idx, column=2, sticky="w", pady=2
            )
            ttk.Entry(
                table_frame,
                textvariable=unit_cost_var,
                width=column_widths[3],
            ).grid(
                row=row_idx, column=3, sticky="w", pady=2
            )
            ttk.Label(
                table_frame,
                textvariable=total_var,
                width=column_widths[4],
                anchor="w",
            ).grid(
                row=row_idx, column=4, sticky="w", pady=2
            )

            row_data: _FeeRowEntry = {
                "scope": scope,
                "status": selection.status,
                "unit": unit,
                "units_var": units_var,
                "unit_cost_var": unit_cost_var,
                "total_var": total_var,
            }
            self._row_entries.append(row_data)

            units_var.trace_add("write", self._recalculate_totals)
            unit_cost_var.trace_add("write", self._recalculate_totals)

        self._totals_frame = ttk.Frame(self)
        self._totals_frame.pack(fill="x", pady=(12, 0))
        self._total_labels: dict[int, tk.StringVar] = {}
        ttk.Label(self._totals_frame, text="Totals by Service Type").pack(anchor="w")
        for service in services:
            total_var = tk.StringVar(value=format_currency(0.0))
            self._total_labels[service.id] = total_var
            ttk.Label(
                self._totals_frame,
                text=f"{service.name}:",
                width=35,
                anchor="w",
            ).pack(anchor="w")
            ttk.Label(self._totals_frame, textvariable=total_var).pack(anchor="w")

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", pady=(12, 0))
        ttk.Button(button_frame, text="Back", command=self._handle_back).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Next", command=self._handle_next).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

        self._recalculate_totals()

    def _recalculate_totals(self, *_args: object) -> None:
        service_totals = {service.id: 0.0 for service in self.services}
        for row in self._row_entries:
            units = _safe_float(row["units_var"].get()) or 0.0
            unit_cost = _safe_float(row["unit_cost_var"].get()) or 0.0
            total = units * unit_cost
            row["total_var"].set(format_currency(total))
            scope = row["scope"]
            status = row["status"]
            if scope.service_type_id is not None and _scope_included_in_base(status):
                service_totals[scope.service_type_id] = (
                    service_totals.get(scope.service_type_id, 0.0) + total
                )
        for service_id, total in service_totals.items():
            total_var = self._total_labels.get(service_id)
            if total_var:
                total_var.set(format_currency(total))

    def _handle_next(self) -> None:
        fee_rows: list[FeeRow] = []
        fee_totals: dict[int, float] = {service.id: 0.0 for service in self.services}
        for row in self._row_entries:
            scope = row["scope"]
            status = row["status"]
            units = _safe_float(row["units_var"].get())
            unit_cost = _safe_float(row["unit_cost_var"].get())
            if units is None or unit_cost is None:
                messagebox.showinfo(
                    "Fee Table",
                    "All rows must have numeric units and unit costs.",
                    parent=self,
                )
                return
            total = units * unit_cost
            fee_rows.append(
                FeeRow(
                    scope=scope,
                    unit=row["unit"],
                    units=units,
                    unit_cost=unit_cost,
                    total=total,
                    status=status,
                )
            )
            if scope.service_type_id is not None and _scope_included_in_base(status):
                fee_totals[scope.service_type_id] = fee_totals.get(
                    scope.service_type_id, 0.0
                ) + total
        self.on_next(fee_rows, fee_totals)
        try:
            self.destroy()
        except tk.TclError:
            pass

    def _handle_back(self) -> None:
        self.on_back()
        self.destroy()


class ProposalWizard(tk.Tk):
    """Drive the multi-step proposal wizard workflow."""

    def __init__(
        self,
        data: ProposalData,
        key_path: Path | None,
        scope_path: Path | None,
        client_path: Path | None,
        template_path: Path,
    ) -> None:
        super().__init__()
        self.data = data
        self.key_path = key_path
        self.scope_path = scope_path
        self.client_path = client_path
        self.template_path = template_path
        self.withdraw()

        creator_root = _RESOURCE_ROOTS[0]
        if creator_root.exists():
            try:
                os.startfile(creator_root)  # type: ignore[attr-defined]
            except OSError:
                pass

        self._client: ClientInfo | None = None
        self._project_types: list[ProjectType] = []
        self._services: list[ServiceType] = []
        self._metrics: ProjectMetrics | None = None
        self._project_managers: list[ProjectManager] = []
        self._selected_scales: list[MiscOption] = []
        self._pm_name: str | None = None
        self._pm_title: str | None = None
        self._fee_ref_totals: dict[str, float] = {}
        self._fee_ref_blocks: dict[str, FeeRefBlock] = {}
        self._assumptions: list[Assumption] = []
        self._scopes: list[ScopeSelection] = []
        self._fee_rows: list[FeeRow] = []
        self._fee_totals: dict[int, float] = {}
        self._selected_project_names: list[str] = []
        self._selected_scale_labels: list[str] = []
        self._selected_service_indices: list[int] = []
        self._selected_pm_names: list[str] = []
        self._selected_scope_states: dict[int, tuple[str, bool]] = {}

        ProjectInfoWindow(self, data, self._handle_project_info, self._open_file_manager)

    def _handle_project_info(
        self,
        client: ClientInfo | None,
        project_types: list[ProjectType],
        services: list[ServiceType],
        metrics: ProjectMetrics,
        project_managers: list[ProjectManager],
        scales: list[MiscOption],
        pm_name: str,
        pm_title: str,
    ) -> None:
        self._client = client
        self._project_types = project_types
        self._services = services
        self._metrics = metrics
        self._project_managers = project_managers
        self._selected_scales = scales
        self._pm_name = pm_name
        self._pm_title = pm_title
        self._selected_project_names = [pt.name for pt in project_types]
        self._selected_scale_labels = [
            f"{scale.scale} ({scale.acres_per_sheet} acres/sheet)" for scale in scales
        ]
        self._selected_service_indices = [
            idx
            for idx, svc in enumerate(self.data.services)
            if svc in services
        ]
        self._selected_pm_names = [pm.name for pm in project_managers]
        AssumptionsWindow(
            self,
            self.data.assumptions,
            self._handle_assumptions,
            self._show_project_info,
            self._assumptions,
        )

    def _handle_assumptions(self, assumptions: list[Assumption]) -> None:
        self._assumptions = assumptions
        if not self._project_types:
            return
        ScopeWindow(
            self,
            self.data.scope_items,
            self._project_types,
            self._services,
            self._handle_scopes,
            self._show_assumptions,
            self._selected_scope_states,
        )

    def _handle_scopes(self, scopes: list[ScopeSelection]) -> None:
        self._scopes = scopes
        self._selected_scope_states = {
            selection.scope.id: (selection.status, selection.show_fee_breakout)
            for selection in scopes
        }
        if not self._metrics:
            return
        self._fee_ref_blocks = {
            name: block
            for name, block in self.data.fee_ref_blocks.items()
            if name in {pt.name for pt in self._project_types}
        }
        if self._fee_ref_blocks:
            FeeRefWindow(
                self,
                self._fee_ref_blocks,
                self._project_types,
                self._metrics.average_multiplier,
                self._handle_fee_ref,
                self._show_scopes,
            )
            return
        FeeTableWindow(
            self,
            scopes,
            self._services,
            self._metrics,
            self._handle_fee_table,
            self._show_scopes,
        )

    def _handle_fee_ref(self, totals: dict[str, float], average_multiplier: float) -> None:
        self._fee_ref_totals = totals
        if not self._metrics:
            return
        self._metrics = ProjectMetrics(
            developed_area=self._metrics.developed_area,
            roadway_length=self._metrics.roadway_length,
            scale=self._metrics.scale,
            average_multiplier=average_multiplier,
        )
        FeeTableWindow(
            self,
            self._scopes,
            self._services,
            self._metrics,
            self._handle_fee_table,
            self._show_fee_ref_or_scopes,
        )

    def _handle_fee_table(self, fee_rows: list[FeeRow], fee_totals: dict[int, float]) -> None:
        self._fee_rows = fee_rows
        self._fee_totals = fee_totals
        self._finalize_proposal()

    def _finalize_proposal(self) -> None:
        if not self._metrics or not self._project_types:
            return

        pm_name = self._pm_name or (self._client.pm if self._client else None)
        pm_title = self._pm_title or _lookup_pm_title(pm_name, self.data.project_managers)

        context = ProposalContext(
            client=self._client,
            project_types=self._project_types,
            services=self._services,
            all_services=self.data.services,
            project_managers=self._project_managers,
            scale=self._metrics.scale,
            scopes=self._scopes,
            assumptions=self._assumptions,
            metrics=self._metrics,
            fee_rows=self._fee_rows,
            fee_totals=self._fee_totals,
            fee_ref_totals=self._fee_ref_totals,
            pm_name=pm_name,
            pm_title=pm_title,
        )

        replacements = _build_placeholder_replacements(context)
        project_name = replacements.get("Project Name") or "Proposal"
        filename = _sanitize_filename(f"{project_name} Proposal")
        target = filedialog.asksaveasfilename(
            title="Save Proposal Letter As",
            defaultextension=".docx",
            initialfile=f"{filename}.docx",
            filetypes=[
                ("Word Document", "*.docx"),
                ("All Files", "*.*"),
            ],
        )
        if not target:
            return
        output_path = Path(target)
        try:
            generate_proposal_document(
                self.template_path,
                output_path,
                context,
                replacements,
            )
        except (OSError, ValueError) as exc:
            messagebox.showerror("Proposal Writer", f"Failed to save proposal: {exc}")
            return

        messagebox.showinfo(
            "Proposal Writer",
            f"Completed!\nProposal saved to:\n{output_path}",
        )
        self.destroy()

    def _open_file_manager(self) -> bool:
        initial: dict[str, Path | None] = {
            "data": self.key_path,
            "scope_data": self.scope_path,
            "client_data": self.client_path,
            "template": self.template_path,
        }
        dialog = FileManagerWindow(self, initial, self._apply_file_paths)
        dialog.grab_set()
        self.wait_window(dialog)
        return dialog.updated

    def _apply_file_paths(self, paths: dict[str, Path]) -> bool:
        updates = {key: str(path) for key, path in paths.items()}
        try:
            data = load_proposal_data(
                Path(updates["data"]),
                Path(updates["scope_data"]),
                Path(updates["client_data"]),
            )
        except (OSError, ValueError) as exc:
            messagebox.showerror("Proposal Writer", f"Failed to load files: {exc}")
            return False

        _save_config(updates)
        self.data = data
        self.key_path = Path(updates["data"])
        self.scope_path = Path(updates["scope_data"])
        self.client_path = Path(updates["client_data"])
        self.template_path = Path(updates["template"])

        self._client = None
        self._project_types = []
        self._services = []
        self._metrics = None
        self._project_managers = []
        self._selected_scales = []
        self._pm_name = None
        self._pm_title = None
        self._fee_ref_totals = {}
        self._fee_ref_blocks = {}
        self._assumptions = []
        self._scopes = []
        self._fee_rows = []
        self._fee_totals = {}
        self._selected_project_names = []
        self._selected_scale_labels = []
        self._selected_service_indices = []
        self._selected_pm_names = []
        self._selected_scope_states = {}

        return True

    def _show_project_info(self) -> None:
        ProjectInfoWindow(
            self,
            self.data,
            self._handle_project_info,
            self._open_file_manager,
            selected_client=self._client,
            selected_project_names=self._selected_project_names,
            selected_scale_labels=self._selected_scale_labels,
            selected_service_indices=self._selected_service_indices,
            selected_pm_names=self._selected_pm_names,
        )

    def _show_assumptions(self) -> None:
        AssumptionsWindow(
            self,
            self.data.assumptions,
            self._handle_assumptions,
            self._show_project_info,
            self._assumptions,
        )

    def _show_scopes(self) -> None:
        if not self._project_types:
            return
        ScopeWindow(
            self,
            self.data.scope_items,
            self._project_types,
            self._services,
            self._handle_scopes,
            self._show_assumptions,
            self._selected_scope_states,
        )

    def _show_fee_ref_or_scopes(self) -> None:
        if self._fee_ref_blocks:
            FeeRefWindow(
                self,
                self._fee_ref_blocks,
                self._project_types,
                self._metrics.average_multiplier if self._metrics else self.data.average_multiplier,
                self._handle_fee_ref,
                self._show_scopes,
            )
            return
        self._show_scopes()
def _demo_data() -> ProposalData:
    """Provide deterministic demo datasets for local testing."""

    return ProposalData(
        project_types=DEMO_PROJECT_TYPES.copy(),
        services=DEMO_SERVICES.copy(),
        project_managers=DEMO_PROJECT_MANAGERS.copy(),
        misc_options=DEMO_MISC_OPTIONS.copy(),
        scope_items=DEMO_SCOPE_ITEMS.copy(),
        assumptions=DEMO_ASSUMPTIONS.copy(),
        clients=DEMO_CLIENTS.copy(),
        fee_ref_blocks=DEMO_FEE_REF_BLOCKS.copy(),
        average_multiplier=3.18,
    )


def main(argv: Sequence[str] | None = None) -> int:
    """CLI entry point."""

    args = parse_args(argv)
    config = _load_config()
    if config.get("data"):
        args.data = Path(config["data"])
    if config.get("scope_data"):
        args.scope_data = Path(config["scope_data"])
    if config.get("client_data"):
        args.client_data = Path(config["client_data"])
    if config.get("template"):
        args.template = Path(config["template"])

    if args.demo:
        data = _demo_data()
        key_path: Path | None = None
        scope_path: Path | None = None
    else:
        try:
            data = load_proposal_data(args.data, args.scope_data, args.client_data)
            key_path = args.data
            scope_path = args.scope_data
        except (OSError, ValueError) as exc:
            messagebox.showerror("Proposal Writer", str(exc))
            return 1

    app = ProposalWizard(data, key_path, scope_path, args.client_data, args.template)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
