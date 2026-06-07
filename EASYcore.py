from __future__ import annotations

import json
import logging
import math
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import matplotlib
import numpy as np
import pandas as pd
import skrf as rf


matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import EngFormatter
from openpyxl.drawing.image import Image as ExcelImage


FREQUENCY_UNITS = {
    "hz": 1.0,
    "khz": 1e3,
    "mhz": 1e6,
    "ghz": 1e9,
}
PARAM_INDEX = {
    "Sdd11": (0, 0),
    "Sdd22": (1, 1),
    "Sdd21": (1, 0),
    "Scd22": (3, 1),
    "Sdc22": (1, 3),
    "Scc21": (3, 2),
}
PLOT_SPECS = [
    ("Sdd11_Sdd22.png", "Sdd11 and Sdd22", ["Sdd11", "Sdd22"]),
    ("Sdd21.png", "Sdd21", ["Sdd21"]),
    ("Scd22.png", "Scd22", ["Scd22"]),
    ("Sdc22.png", "Sdc22", ["Sdc22"]),
    ("Scc21.png", "Scc21", ["Scc21"]),
]


@dataclass(frozen=True)
class PlotLimits:
    x_min_hz: float | None = None
    x_max_hz: float | None = None
    y_min_db: float | None = None
    y_max_db: float | None = None


@dataclass(frozen=True)
class PlotLimitSettings:
    global_limits: PlotLimits
    per_plot: dict[str, PlotLimits]


@dataclass(frozen=True)
class TitleSettings:
    enabled: bool = False
    format: str = "{plot_title}"


@dataclass(frozen=True)
class LegendSettings:
    mode: str = "off"


@dataclass(frozen=True)
class AverageLineSettings:
    enabled: bool
    color: str
    linewidth: float
    label: str


@dataclass(frozen=True)
class PlotAverageSettings:
    direct_db: AverageLineSettings
    linear_magnitude_db: AverageLineSettings


@dataclass(frozen=True)
class XtkPlotSettings:
    enabled: bool = True
    folder_name: str = "XTK"
    parameter: str = "Sdd21"
    plot_types: list[str] | None = None


@dataclass(frozen=True)
class FrequencyTrimSettings:
    enabled: bool = True
    margin_ratio: float = 1.05
    max_hz: float | None = None


@dataclass(frozen=True)
class AdditionalExcelTarget:
    groups: list[str]
    parameter: str
    frequencies_hz: list[float]
    plot_limits: PlotLimits | None = None


@dataclass(frozen=True)
class ProcessedNetwork:
    file_path: Path
    label: str
    frequency_hz: np.ndarray
    single_ended_s: np.ndarray
    mixed_mode_s: np.ndarray


@dataclass(frozen=True)
class FrequencyEstimate:
    value: complex
    method: str
    lower_frequency_hz: float
    upper_frequency_hz: float


@dataclass(frozen=True)
class LimitOverlaySettings:
    enabled: bool = False


def load_config(config_path: Path | None) -> tuple[dict[str, Any], Path | None]:
    if config_path is None:
        return {}, None
    if not config_path.exists():
        raise FileNotFoundError(f"Config file does not exist: {config_path}")
    with config_path.open("r", encoding="utf-8") as file:
        config = json.load(file)
    if not isinstance(config, dict):
        raise ValueError("Config file root must be a JSON object.")
    return config, config_path


def parse_frequency(value: Any) -> float:
    if isinstance(value, (int, float)):
        frequency_hz = float(value)
    elif isinstance(value, str):
        text = value.strip()
        match = re.fullmatch(
            r"([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)\s*([a-zA-Z]*)",
            text,
        )
        if not match:
            raise ValueError(f"Invalid frequency value: {value!r}")
        number = float(match.group(1))
        unit = match.group(2).lower() or "hz"
        if unit not in FREQUENCY_UNITS:
            raise ValueError("Unsupported frequency unit. Use Hz, kHz, MHz, or GHz.")
        frequency_hz = number * FREQUENCY_UNITS[unit]
    else:
        raise TypeError(f"Frequency must be a number or string, got {type(value).__name__}.")
    if frequency_hz <= 0 or not math.isfinite(frequency_hz):
        raise ValueError(f"Frequency must be a positive finite value, got {value!r}.")
    return frequency_hz


def optional_float(value: Any, field_name: str) -> float | None:
    if value is None:
        return None
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"{field_name} must be finite.")
    return number


def parse_limit_pair(raw_limits: dict[str, Any], field_name: str, parser) -> tuple[float | None, float | None]:
    if not isinstance(raw_limits, dict):
        raise ValueError(f"{field_name} must be a JSON object.")
    min_value = raw_limits.get("min")
    max_value = raw_limits.get("max")
    parsed_min = parser(min_value) if min_value is not None else None
    parsed_max = parser(max_value) if max_value is not None else None
    if parsed_min is not None and parsed_max is not None and parsed_min >= parsed_max:
        raise ValueError(f"{field_name}.min must be smaller than {field_name}.max.")
    return parsed_min, parsed_max


def parse_plot_limit_block(raw_limits: dict[str, Any], field_name: str) -> PlotLimits:
    if not isinstance(raw_limits, dict):
        raise ValueError(f"{field_name} must be a JSON object.")
    x_min_hz = x_max_hz = y_min_db = y_max_db = None
    if "x_frequency" in raw_limits:
        x_min_hz, x_max_hz = parse_limit_pair(
            raw_limits["x_frequency"], f"{field_name}.x_frequency", parse_frequency
        )
    if "y_db" in raw_limits:
        y_min_db, y_max_db = parse_limit_pair(
            raw_limits["y_db"],
            f"{field_name}.y_db",
            lambda raw: optional_float(raw, f"{field_name}.y_db"),
        )
    return PlotLimits(x_min_hz, x_max_hz, y_min_db, y_max_db)


def resolve_plot_limits(config: dict[str, Any]) -> PlotLimitSettings:
    raw = config.get("plot_limits", {})
    if not isinstance(raw, dict):
        raise ValueError("plot_limits must be a JSON object.")
    global_limits = (
        PlotLimits()
        if not isinstance(raw.get("global"), dict)
        else parse_plot_limit_block(raw["global"], "plot_limits.global")
    )
    per_plot = {
        key: parse_plot_limit_block(value, f"plot_limits.{key}")
        for key, value in raw.items()
        if key != "global" and isinstance(value, dict)
    }
    return PlotLimitSettings(global_limits, per_plot)


def resolve_recursive_search(config: dict[str, Any]) -> bool:
    value = config.get("recursive_search", False)
    if not isinstance(value, bool):
        raise ValueError("recursive_search must be true or false.")
    return value


def resolve_title_settings(config: dict[str, Any]) -> TitleSettings:
    raw = config.get("title_from_filename", {})
    if not isinstance(raw, dict):
        raise ValueError("title_from_filename must be a JSON object.")
    return TitleSettings(
        enabled=bool(raw.get("enabled", False)),
        format=str(raw.get("format", "{plot_title}")),
    )


def resolve_legend_settings(config: dict[str, Any]) -> LegendSettings:
    raw = config.get("plot_legend", {})
    if not isinstance(raw, dict):
        raise ValueError("plot_legend must be a JSON object.")
    mode = str(raw.get("mode", "off")).strip().lower()
    if mode not in {"off", "full"}:
        raise ValueError("plot_legend.mode must be off or full.")
    return LegendSettings(mode=mode)


def resolve_average_line_settings(raw: Any, default: AverageLineSettings) -> AverageLineSettings:
    if raw is None:
        return default
    if not isinstance(raw, dict):
        raise ValueError("plot_average entries must be JSON objects.")
    return AverageLineSettings(
        enabled=bool(raw.get("enabled", default.enabled)),
        color=str(raw.get("color", default.color)),
        linewidth=float(raw.get("linewidth", default.linewidth)),
        label=str(raw.get("label", default.label)),
    )


def resolve_plot_average_settings(config: dict[str, Any]) -> PlotAverageSettings:
    raw = config.get("plot_average", {})
    if not isinstance(raw, dict):
        raise ValueError("plot_average must be a JSON object.")
    return PlotAverageSettings(
        direct_db=resolve_average_line_settings(
            raw.get("direct_db"),
            AverageLineSettings(False, "#ff0040", 3.0, "Average dB"),
        ),
        linear_magnitude_db=resolve_average_line_settings(
            raw.get("linear_magnitude_db"),
            AverageLineSettings(False, "#00b7ff", 3.0, "Average linear magnitude to dB"),
        ),
    )


def resolve_exclude_keywords(config: dict[str, Any]) -> list[str]:
    raw = config.get("exclude_keywords", ["Golden"])
    if not isinstance(raw, list):
        raise ValueError("exclude_keywords must be a JSON list.")
    return [str(item) for item in raw if str(item).strip()]


def resolve_xtk_plot_settings(config: dict[str, Any]) -> XtkPlotSettings:
    raw = config.get("xtk_plots", {})
    if not isinstance(raw, dict):
        raise ValueError("xtk_plots must be a JSON object.")
    plot_types = raw.get("plot_types", ["NEXT", "FEXT"])
    if not isinstance(plot_types, list):
        raise ValueError("xtk_plots.plot_types must be a JSON list.")
    parameter = normalize_s_parameter(str(raw.get("parameter", "Sdd21")))
    if not is_supported_parameter(parameter):
        raise ValueError("xtk_plots.parameter is unsupported.")
    return XtkPlotSettings(
        enabled=bool(raw.get("enabled", True)),
        folder_name=str(raw.get("folder_name", "XTK")),
        parameter=parameter,
        plot_types=[str(item).upper() for item in plot_types],
    )


def resolve_frequency_trim_settings(config: dict[str, Any]) -> FrequencyTrimSettings:
    raw = config.get("performance", {}).get("trim_frequency", {})
    if not isinstance(raw, dict):
        raw = {}
    max_hz = raw.get("max_hz")
    return FrequencyTrimSettings(
        enabled=bool(raw.get("enabled", True)),
        margin_ratio=float(raw.get("margin_ratio", 1.05)),
        max_hz=parse_frequency(max_hz) if max_hz is not None else None,
    )


def resolve_processing_workers(config: dict[str, Any]) -> int:
    workers = int(config.get("performance", {}).get("processing_workers", 4))
    if workers < 1 or workers > 16:
        raise ValueError("processing_workers must be between 1 and 16.")
    return workers


def resolve_output_prefix(config: dict[str, Any]) -> str:
    raw = config.get("output_naming", {})
    if not isinstance(raw, dict):
        return ""
    return sanitize_filename_prefix(str(raw.get("part_number_prefix", "")))


def resolve_additional_excel_targets(config: dict[str, Any]) -> list[AdditionalExcelTarget]:
    raw_targets = config.get("additional_excel_targets", [])
    if not isinstance(raw_targets, list):
        raise ValueError("additional_excel_targets must be a JSON list.")
    targets = []
    for index, raw in enumerate(raw_targets):
        if not isinstance(raw, dict):
            raise ValueError(f"additional_excel_targets[{index}] must be a JSON object.")
        if not bool(raw.get("include", True)):
            continue
        groups = raw.get("groups", [])
        if not isinstance(groups, list) or not groups:
            raise ValueError(f"additional_excel_targets[{index}].groups must be a non-empty list.")
        normalized_groups = [str(group).strip().upper() for group in groups]
        if any(group not in {"MAIN", "NEXT", "FEXT"} for group in normalized_groups):
            raise ValueError(f"additional_excel_targets[{index}] contains unsupported groups.")
        parameter = normalize_s_parameter(str(raw.get("parameter", "")))
        if not parameter or not is_supported_parameter(parameter):
            raise ValueError(f"additional_excel_targets[{index}].parameter is unsupported.")
        frequencies = raw.get("frequencies", [])
        if not isinstance(frequencies, list) or not frequencies:
            raise ValueError(f"additional_excel_targets[{index}].frequencies must be non-empty.")
        plot_limits = None
        if isinstance(raw.get("plot_limits"), dict):
            plot_limits = parse_plot_limit_block(raw["plot_limits"], f"additional_excel_targets[{index}].plot_limits")
        targets.append(
            AdditionalExcelTarget(
                groups=normalized_groups,
                parameter=parameter,
                frequencies_hz=sorted({parse_frequency(value) for value in frequencies}),
                plot_limits=plot_limits,
            )
        )
    return targets


def sanitize_filename_prefix(value: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value.strip())
    return sanitized.strip(" ._")


def prefixed_output_name(filename: str, output_prefix: str = "") -> str:
    prefix = sanitize_filename_prefix(output_prefix)
    return f"{prefix}_{filename}" if prefix else filename


def setup_logging(output_dir: Path, output_prefix: str = "") -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / prefixed_output_name("processing_log.txt", output_prefix)
    root = logging.getLogger()
    for handler in list(root.handlers):
        root.removeHandler(handler)
    root.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s: %(message)s")
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    root.addHandler(file_handler)
    root.addHandler(stream_handler)
    logging.info("Log file: %s", log_path)


def log_progress(value: int) -> None:
    logging.info("PROGRESS: %d", max(0, min(100, int(value))))


def is_excluded_path(path: Path, exclude_keywords: list[str]) -> bool:
    text = str(path).casefold()
    return any(keyword.casefold() in text for keyword in exclude_keywords)


def discover_s4p_files(input_dir: Path, recursive_search: bool, exclude_keywords: list[str]) -> list[Path]:
    pattern = "**/*.s4p" if recursive_search else "*.s4p"
    return sorted(
        [
            path
            for path in input_dir.glob(pattern)
            if path.is_file() and not is_excluded_path(path, exclude_keywords)
        ],
        key=lambda path: str(path).casefold(),
    )


def find_xtk_folder(input_dir: Path, folder_name: str = "XTK") -> Path | None:
    target = folder_name.casefold()
    for path in input_dir.iterdir():
        if path.is_dir() and path.name.casefold() == target:
            return path
    return None


def discover_xtk_files(input_dir: Path, settings: XtkPlotSettings, exclude_keywords: list[str]) -> dict[str, list[Path]]:
    result = {plot_type: [] for plot_type in (settings.plot_types or ["NEXT", "FEXT"])}
    folder = find_xtk_folder(input_dir, settings.folder_name)
    if folder is None:
        return result
    for path in sorted(folder.glob("*.s4p"), key=lambda item: str(item).casefold()):
        if not path.is_file() or is_excluded_path(path, exclude_keywords):
            continue
        name = path.stem.casefold()
        for plot_type in result:
            if plot_type.casefold() in name:
                result[plot_type].append(path)
    return result


def mixed_mode_transform_matrix() -> np.ndarray:
    scale = 1 / math.sqrt(2)
    return np.asarray(
        [
            [scale, -scale, 0, 0],
            [0, 0, scale, -scale],
            [scale, scale, 0, 0],
            [0, 0, scale, scale],
        ],
        dtype=float,
    )


def convert_single_ended_to_mixed_mode(s_se: np.ndarray) -> np.ndarray:
    if s_se.ndim != 3 or s_se.shape[1:] != (4, 4):
        raise ValueError(f"Expected 4-port S-parameter array, got {s_se.shape}.")
    transform = mixed_mode_transform_matrix()
    return np.einsum("ab,fbc,dc->fad", transform, s_se, transform)


def process_path(path: Path, max_frequency_hz: float | None = None) -> ProcessedNetwork:
    network = rf.Network(str(path))
    frequency_hz = np.asarray(network.frequency.f, dtype=float)
    single_ended = np.asarray(network.s, dtype=complex)
    if max_frequency_hz is not None:
        keep = frequency_hz <= max_frequency_hz
        if np.any(keep):
            frequency_hz = frequency_hz[keep]
            single_ended = single_ended[keep]
    return ProcessedNetwork(
        file_path=path,
        label=path.stem,
        frequency_hz=frequency_hz,
        single_ended_s=single_ended,
        mixed_mode_s=convert_single_ended_to_mixed_mode(single_ended),
    )


def process_paths(paths: list[Path], workers: int, max_frequency_hz: float | None = None) -> tuple[list[ProcessedNetwork], int]:
    processed = []
    failed = 0
    if workers <= 1:
        for path in paths:
            try:
                processed.append(process_path(path, max_frequency_hz))
            except Exception as exc:
                failed += 1
                logging.exception("Failed to process %s: %s", path, exc)
        return processed, failed

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process_path, path, max_frequency_hz): path for path in paths}
        for future in as_completed(futures):
            path = futures[future]
            try:
                processed.append(future.result())
            except Exception as exc:
                failed += 1
                logging.exception("Failed to process %s: %s", path, exc)
    processed.sort(key=lambda item: str(item.file_path).casefold())
    return processed, failed


def parse_single_ended_parameter(parameter: str) -> tuple[int, int] | None:
    match = re.fullmatch(r"S([1-4])([1-4])", parameter.strip(), flags=re.IGNORECASE)
    if not match:
        return None
    row = int(match.group(1)) - 1
    col = int(match.group(2)) - 1
    return row, col


def parse_mixed_mode_parameter(parameter: str) -> tuple[int, int] | None:
    match = re.fullmatch(r"S([dc])([dc])([12])([12])", parameter.strip(), flags=re.IGNORECASE)
    if not match:
        return None
    row_mode, col_mode, row_pair, col_pair = match.groups()
    row = (0 if row_mode.lower() == "d" else 2) + int(row_pair) - 1
    col = (0 if col_mode.lower() == "d" else 2) + int(col_pair) - 1
    return row, col


def normalize_s_parameter(parameter: str) -> str:
    text = parameter.strip()
    single = parse_single_ended_parameter(text)
    if single is not None:
        return f"S{single[0] + 1}{single[1] + 1}"
    mixed = parse_mixed_mode_parameter(text)
    if mixed is None:
        return text
    row, col = mixed
    row_mode = "d" if row < 2 else "c"
    col_mode = "d" if col < 2 else "c"
    return f"S{row_mode}{col_mode}{row % 2 + 1}{col % 2 + 1}"


def is_supported_parameter(parameter: str) -> bool:
    return parse_single_ended_parameter(parameter) is not None or parse_mixed_mode_parameter(parameter) is not None


def parameter_values(item: ProcessedNetwork, parameter: str) -> np.ndarray:
    single = parse_single_ended_parameter(parameter)
    if single is not None:
        return item.single_ended_s[:, single[0], single[1]]
    mixed = parse_mixed_mode_parameter(parameter)
    if mixed is not None:
        return item.mixed_mode_s[:, mixed[0], mixed[1]]
    raise ValueError(f"Unsupported S-parameter: {parameter}")


def magnitude_db(values: np.ndarray | complex) -> np.ndarray | float:
    return 20 * np.log10(np.maximum(np.abs(values), np.finfo(float).tiny))


def linear_magnitude_to_db(values: pd.Series | np.ndarray | float) -> pd.Series | np.ndarray | float:
    return 20 * np.log10(np.maximum(values, np.finfo(float).tiny))


def extract_param_db(item: ProcessedNetwork, parameter: str) -> np.ndarray:
    return magnitude_db(parameter_values(item, parameter))


def interpolate_complex_series(
    source_frequency_hz: np.ndarray,
    source_values: np.ndarray,
    target_frequency_hz: np.ndarray,
) -> np.ndarray:
    order = np.argsort(source_frequency_hz)
    sorted_freq = source_frequency_hz[order]
    sorted_values = source_values[order]
    real = np.interp(target_frequency_hz, sorted_freq, sorted_values.real)
    imag = np.interp(target_frequency_hz, sorted_freq, sorted_values.imag)
    return real + 1j * imag


def estimate_complex_at_frequency(frequency_hz: np.ndarray, values: np.ndarray, target_hz: float) -> FrequencyEstimate:
    order = np.argsort(frequency_hz)
    sorted_freq = frequency_hz[order]
    sorted_values = values[order]
    index = int(np.searchsorted(sorted_freq, target_hz))
    if index <= 0:
        value = sorted_values[0]
        lower = upper = float(sorted_freq[0])
        method = "nearest_below_range"
    elif index >= len(sorted_freq):
        value = sorted_values[-1]
        lower = upper = float(sorted_freq[-1])
        method = "nearest_above_range"
    elif math.isclose(float(sorted_freq[index]), target_hz, rel_tol=0.0, abs_tol=1e-9):
        value = sorted_values[index]
        lower = upper = float(sorted_freq[index])
        method = "exact"
    else:
        lower = float(sorted_freq[index - 1])
        upper = float(sorted_freq[index])
        ratio = (target_hz - lower) / (upper - lower)
        value = sorted_values[index - 1] + ratio * (sorted_values[index] - sorted_values[index - 1])
        method = "linear_complex"
    return FrequencyEstimate(complex(value), method, lower, upper)


def build_result_tables(
    processed: list[ProcessedNetwork],
    target_freqs: list[float],
    parameter_names: Iterable[str] | None = None,
    extra_columns: dict[str, str] | None = None,
    target_freqs_by_parameter: dict[str, list[float]] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    params = list(PARAM_INDEX) if parameter_names is None else list(parameter_names)
    extras = {} if extra_columns is None else dict(extra_columns)
    rows = []
    for item in processed:
        for parameter in params:
            if not is_supported_parameter(parameter):
                continue
            current_freqs = target_freqs if target_freqs_by_parameter is None else target_freqs_by_parameter.get(parameter, [])
            for target_hz in current_freqs:
                estimate = estimate_complex_at_frequency(item.frequency_hz, parameter_values(item, parameter), target_hz)
                linear = float(abs(estimate.value))
                rows.append(
                    {
                        "file_name": item.file_path.name,
                        "file_path": str(item.file_path),
                        "target_frequency_hz": float(target_hz),
                        "result_frequency_hz": float(target_hz),
                        "interpolation_method": estimate.method,
                        "lower_frequency_hz": estimate.lower_frequency_hz,
                        "upper_frequency_hz": estimate.upper_frequency_hz,
                        "parameter": parameter,
                        "linear_magnitude": linear,
                        "magnitude_db": float(magnitude_db(estimate.value)),
                        **extras,
                    }
                )
    per_file_df = pd.DataFrame(rows)
    group_columns = [*extras.keys(), "parameter", "target_frequency_hz"]
    if per_file_df.empty:
        return pd.DataFrame(columns=group_columns), per_file_df
    summary_df = (
        per_file_df.groupby(group_columns, as_index=False)
        .agg(
            average_db=("magnitude_db", "mean"),
            average_linear_magnitude_db=("linear_magnitude", "mean"),
            maximum_db=("magnitude_db", "max"),
            minimum_db=("magnitude_db", "min"),
            files_used=("file_name", "nunique"),
            lower_frequency_hz_min=("lower_frequency_hz", "min"),
            upper_frequency_hz_max=("upper_frequency_hz", "max"),
        )
        .sort_values(group_columns, kind="stable")
    )
    summary_df["average_linear_magnitude_db"] = linear_magnitude_to_db(summary_df["average_linear_magnitude_db"])
    return summary_df, per_file_df


def additional_excel_targets_for_group(
    targets: list[AdditionalExcelTarget],
    group: str,
) -> dict[str, list[float]]:
    grouped: dict[str, list[float]] = {}
    for target in targets:
        if group in target.groups:
            grouped.setdefault(target.parameter, []).extend(target.frequencies_hz)
    return {parameter: sorted(set(freqs)) for parameter, freqs in grouped.items()}


def build_additional_result_tables(
    main_processed: list[ProcessedNetwork],
    xtk_processed_by_group: dict[str, list[ProcessedNetwork]],
    additional_targets: list[AdditionalExcelTarget],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    summary_frames = []
    per_file_frames = []
    groups = {"MAIN": main_processed, **xtk_processed_by_group}
    for group, processed in groups.items():
        targets = additional_excel_targets_for_group(additional_targets, group)
        if not targets or not processed:
            continue
        summary_df, per_file_df = build_result_tables(
            processed,
            [],
            parameter_names=list(targets),
            extra_columns={"source_group": group},
            target_freqs_by_parameter=targets,
        )
        if not summary_df.empty:
            summary_frames.append(format_additional_results(summary_df))
        if not per_file_df.empty:
            per_file_frames.append(format_additional_results(per_file_df))
    return (
        pd.concat(summary_frames, ignore_index=True) if summary_frames else pd.DataFrame(),
        pd.concat(per_file_frames, ignore_index=True) if per_file_frames else pd.DataFrame(),
    )


def format_additional_results(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    formatted = df.copy()
    if "target_frequency_hz" in formatted.columns:
        insert_at = formatted.columns.get_loc("target_frequency_hz")
        formatted.insert(insert_at, "target_frequency_mhz", formatted["target_frequency_hz"].astype(float) / 1e6)
        formatted = formatted.drop(columns=["target_frequency_hz"])
    preferred = [
        "source_group",
        "parameter",
        "target_frequency_mhz",
        "average_db",
        "average_linear_magnitude_db",
        "minimum_db",
        "maximum_db",
        "files_used",
        "file_name",
        "file_path",
        "result_frequency_hz",
        "interpolation_method",
        "lower_frequency_hz",
        "upper_frequency_hz",
        "linear_magnitude",
        "magnitude_db",
    ]
    columns = [column for column in preferred if column in formatted.columns]
    columns.extend(column for column in formatted.columns if column not in columns)
    sort_columns = [column for column in ("source_group", "parameter", "target_frequency_mhz", "file_name") if column in formatted.columns]
    if sort_columns:
        formatted = formatted.sort_values(sort_columns, kind="stable")
    return formatted.loc[:, columns].reset_index(drop=True)


def label_xtk_parameter(df: pd.DataFrame, plot_type: str, parameter: str) -> pd.DataFrame:
    if df.empty or "parameter" not in df:
        return df
    labeled = df.copy()
    labeled["parameter"] = labeled["parameter"].replace({parameter: f"{plot_type}({parameter})"})
    return labeled


def make_xtk_plots(
    input_dir: Path,
    output_dir: Path,
    xtk_settings: XtkPlotSettings,
    exclude_keywords: list[str],
    plot_limits: PlotLimitSettings,
    title_settings: TitleSettings,
    legend_settings: LegendSettings,
    average_settings: PlotAverageSettings,
    limit_database: dict[str, Any] | None,
    limit_settings: LimitOverlaySettings,
    target_freqs: list[float],
    target_freqs_by_parameter: dict[str, list[float]] | None,
    additional_targets: list[AdditionalExcelTarget],
    workers: int,
    max_frequency_hz: float | None,
) -> tuple[int, int, pd.DataFrame, pd.DataFrame, list[Path], list[tuple[str, list[ProcessedNetwork], list[tuple[str, str, list[str]]]]]]:
    if not xtk_settings.enabled:
        return 0, 0, pd.DataFrame(), pd.DataFrame(), [], []
    discovered = discover_xtk_files(input_dir, xtk_settings, exclude_keywords)
    summary_frames = []
    per_file_frames = []
    picture_paths = []
    plot_jobs = []
    success = failed = 0
    for plot_type, paths in discovered.items():
        if not paths:
            continue
        processed, failed_count = process_paths(paths, workers, max_frequency_hz)
        failed += failed_count
        success += len(processed)
        if not processed:
            continue
        data_group_freqs = []
        if target_freqs_by_parameter is not None:
            data_group_freqs = target_freqs_by_parameter.get(plot_type, [])
        freq_map = {xtk_settings.parameter: data_group_freqs} if data_group_freqs else None
        summary_df, per_file_df = build_result_tables(
            processed,
            target_freqs,
            parameter_names=[xtk_settings.parameter],
            extra_columns={"data_group": plot_type},
            target_freqs_by_parameter=freq_map,
        )
        if not summary_df.empty:
            summary_frames.append(label_xtk_parameter(summary_df, plot_type, xtk_settings.parameter))
        if not per_file_df.empty:
            per_file_frames.append(label_xtk_parameter(per_file_df, plot_type, xtk_settings.parameter))
        specs = [(f"{plot_type}_{xtk_settings.parameter}.png", f"{plot_type} {xtk_settings.parameter}", [xtk_settings.parameter])]
        picture_paths.extend(
            make_plots(
                processed,
                output_dir,
                plot_limits,
                title_settings,
                legend_settings,
                average_settings,
                limit_database,
                limit_settings,
                plot_specs=specs,
            )
        )
        plot_jobs.append((plot_type, processed, specs))
    return (
        success,
        failed,
        pd.concat(summary_frames, ignore_index=True) if summary_frames else pd.DataFrame(),
        pd.concat(per_file_frames, ignore_index=True) if per_file_frames else pd.DataFrame(),
        picture_paths,
        plot_jobs,
    )


def resolve_plot_limits_for_spec(filename: str, title: str, params: list[str], settings: PlotLimitSettings) -> PlotLimits:
    keys = [filename.replace(".png", ""), title, *params]
    plot_limits = PlotLimits()
    for key in keys:
        if key in settings.per_plot:
            plot_limits = settings.per_plot[key]
            break
    return PlotLimits(
        x_min_hz=plot_limits.x_min_hz if plot_limits.x_min_hz is not None else settings.global_limits.x_min_hz,
        x_max_hz=plot_limits.x_max_hz if plot_limits.x_max_hz is not None else settings.global_limits.x_max_hz,
        y_min_db=plot_limits.y_min_db if plot_limits.y_min_db is not None else settings.global_limits.y_min_db,
        y_max_db=plot_limits.y_max_db if plot_limits.y_max_db is not None else settings.global_limits.y_max_db,
    )


def apply_plot_limits(ax: plt.Axes, limits: PlotLimits) -> None:
    if limits.x_min_hz is not None or limits.x_max_hz is not None:
        ax.set_xlim(left=limits.x_min_hz, right=limits.x_max_hz)
    if limits.y_min_db is not None or limits.y_max_db is not None:
        ax.set_ylim(bottom=limits.y_min_db, top=limits.y_max_db)


def build_plot_title(title: str, processed: list[ProcessedNetwork], settings: TitleSettings) -> str:
    if not settings.enabled or not processed:
        return title
    return settings.format.format(plot_title=title, file_name=processed[0].file_path.stem)


def plot_average_lines(ax: plt.Axes, processed: list[ProcessedNetwork], parameter: str, settings: PlotAverageSettings, suffix: str) -> None:
    if not settings.direct_db.enabled and not settings.linear_magnitude_db.enabled:
        return
    base_frequency = processed[0].frequency_hz
    aligned = np.vstack([
        interpolate_complex_series(item.frequency_hz, parameter_values(item, parameter), base_frequency)
        for item in processed
    ])
    if settings.direct_db.enabled:
        ax.plot(
            base_frequency,
            np.mean(magnitude_db(aligned), axis=0),
            color=settings.direct_db.color,
            linewidth=settings.direct_db.linewidth,
            label=f"{settings.direct_db.label}{suffix}",
            zorder=10,
        )
    if settings.linear_magnitude_db.enabled:
        ax.plot(
            base_frequency,
            linear_magnitude_to_db(np.mean(np.abs(aligned), axis=0)),
            color=settings.linear_magnitude_db.color,
            linewidth=settings.linear_magnitude_db.linewidth,
            linestyle="--",
            label=f"{settings.linear_magnitude_db.label}{suffix}",
            zorder=11,
        )


def make_plots(
    processed: list[ProcessedNetwork],
    output_dir: Path,
    plot_limit_settings: PlotLimitSettings,
    title_settings: TitleSettings,
    legend_settings: LegendSettings,
    average_settings: PlotAverageSettings,
    limit_database: dict[str, Any] | None,
    limit_settings: LimitOverlaySettings,
    plot_specs: list[tuple[str, str, list[str]]] | None = None,
    **_: Any,
) -> list[Path]:
    specs = PLOT_SPECS if plot_specs is None else plot_specs
    output_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for filename, title, params in specs:
        fig, ax = plt.subplots(figsize=(11, 7), dpi=150)
        limits = resolve_plot_limits_for_spec(filename, title, params, plot_limit_settings)
        for item in processed:
            for parameter in params:
                suffix = f" {parameter}" if len(params) > 1 else ""
                ax.plot(item.frequency_hz, extract_param_db(item, parameter), label=f"{item.label}{suffix}")
        for parameter in params:
            suffix = f" {parameter}" if len(params) > 1 else ""
            plot_average_lines(ax, processed, parameter, average_settings, suffix)
        ax.set_title(build_plot_title(title, processed, title_settings))
        ax.set_xlabel("Frequency (Hz)")
        ax.set_ylabel("Magnitude (dB)")
        ax.xaxis.set_major_formatter(EngFormatter(unit="Hz"))
        apply_plot_limits(ax, limits)
        ax.grid(True, which="both")
        if legend_settings.mode == "full":
            ax.legend(loc="best", fontsize="small")
        fig.tight_layout()
        output_path = output_dir / filename
        fig.savefig(output_path)
        plt.close(fig)
        written.append(output_path)
        logging.info("Wrote plot: %s", output_path)
    return written


def make_additional_plots(
    main_processed: list[ProcessedNetwork],
    xtk_processed_by_group: dict[str, list[ProcessedNetwork]],
    additional_targets: list[AdditionalExcelTarget],
    output_dir: Path,
    title_settings: TitleSettings,
    legend_settings: LegendSettings,
    average_settings: PlotAverageSettings,
) -> list[Path]:
    groups = {"MAIN": main_processed, **xtk_processed_by_group}
    written = []
    for index, target in enumerate(additional_targets, start=1):
        for group in target.groups:
            processed = groups.get(group, [])
            if not processed:
                continue
            plot_limits = PlotLimitSettings(
                global_limits=target.plot_limits or PlotLimits(1e6, 500e6, -50, 0),
                per_plot={},
            )
            filename = sanitize_filename_prefix(f"Additional_{group}_{target.parameter}_{index}") + ".png"
            written.extend(
                make_plots(
                    processed,
                    output_dir,
                    plot_limits,
                    title_settings,
                    legend_settings,
                    average_settings,
                    None,
                    LimitOverlaySettings(False),
                    plot_specs=[(filename, f"Additional {group} {target.parameter}", [target.parameter])],
                )
            )
    return written


def filtered_plot_specs(plot_specs: list[tuple[str, str, list[str]]], enabled_parameters: set[str] | None) -> list[tuple[str, str, list[str]]]:
    if enabled_parameters is None:
        return plot_specs
    filtered = []
    for filename, title, params in plot_specs:
        kept = [parameter for parameter in params if parameter in enabled_parameters]
        if kept:
            filtered.append((filename, title, kept))
    return filtered


SUMMARY_ORDER = {
    "Sdd21": 0,
    "Sdd11": 1,
    "Sdd22": 2,
    "NEXT(Sdd21)": 3,
    "FEXT(Sdd21)": 4,
    "Scc21": 5,
    "Scd22": 6,
    "Sdc22": 7,
}


def sort_summary_results(summary_df: pd.DataFrame) -> pd.DataFrame:
    if summary_df.empty:
        return summary_df
    sorted_df = summary_df.copy()
    sorted_df["_order"] = sorted_df["parameter"].map(SUMMARY_ORDER).fillna(99)
    sort_columns = ["_order"]
    for column in ("parameter", "target_frequency_hz", "data_group"):
        if column in sorted_df.columns:
            sort_columns.append(column)
    return sorted_df.sort_values(sort_columns, kind="stable").drop(columns="_order").reset_index(drop=True)


def order_summary_columns(summary_df: pd.DataFrame) -> pd.DataFrame:
    preferred = [
        "data_group",
        "parameter",
        "target_frequency_hz",
        "average_linear_magnitude_db",
        "average_db",
        "minimum_db",
        "maximum_db",
        "files_used",
        "lower_frequency_hz_min",
        "upper_frequency_hz_max",
    ]
    columns = [column for column in preferred if column in summary_df.columns]
    columns.extend(column for column in summary_df.columns if column not in columns)
    return summary_df.loc[:, columns]


def add_picture_sheet(writer: pd.ExcelWriter, sheet_name: str, picture_paths: list[Path]) -> None:
    if not picture_paths:
        return
    sheet = writer.book.create_sheet(sheet_name)
    anchor_columns = ("A", "H", "O")
    image_width = 390
    rows_per_image = 18
    for column in anchor_columns:
        sheet.column_dimensions[column].width = 18
    for index, picture_path in enumerate(picture_paths):
        if not picture_path.exists():
            continue
        grid_column = index % len(anchor_columns)
        grid_row = index // len(anchor_columns)
        row = grid_row * rows_per_image + 1
        column = anchor_columns[grid_column]
        sheet[f"{column}{row}"] = picture_path.stem
        sheet[f"{column}{row}"].style = "Title"
        image = ExcelImage(str(picture_path))
        original_width = image.width
        image.width = image_width
        image.height = int(image.height * image.width / original_width)
        sheet.add_image(image, f"{column}{row + 1}")


def remove_written_plots(paths: Iterable[Path]) -> None:
    for path in set(Path(item) for item in paths):
        try:
            path.unlink(missing_ok=True)
        except OSError as exc:
            logging.warning("Could not remove plot %s: %s", path, exc)
