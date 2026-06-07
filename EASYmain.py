from __future__ import annotations

import argparse
import logging
import math
from pathlib import Path
from time import perf_counter
from typing import Any

from cycler import cycler
import EASYcore as core
import pandas as pd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Supplier simple S4P batch processor.")
    parser.add_argument("--config", type=Path, required=True, help="Supplier runtime JSON config.")
    return parser.parse_args()


def read_config(path: Path) -> dict[str, Any]:
    config, _config_path = core.load_config(path)
    return config


def apply_supplier_plot_style() -> None:
    island_palette = [
        "#007f8c",
        "#ff7f50",
        "#2e8b57",
        "#f4a261",
        "#4dabf7",
        "#9c6ade",
        "#00a896",
        "#e76f51",
        "#6c9a8b",
        "#f2c14e",
    ]
    core.plt.rcParams.update(
        {
            "font.sans-serif": [
                "Trebuchet MS",
                "Microsoft JhengHei",
                "Segoe UI",
                "Arial",
                "DejaVu Sans",
            ],
            "axes.unicode_minus": False,
            "axes.facecolor": "#fff8ea",
            "figure.facecolor": "#fff3d6",
            "savefig.facecolor": "#fff3d6",
            "axes.edgecolor": "#0b7285",
            "axes.labelcolor": "#17343a",
            "axes.titlecolor": "#075461",
            "axes.titlesize": 13,
            "axes.titleweight": "bold",
            "axes.labelsize": 10,
            "xtick.color": "#42565b",
            "ytick.color": "#42565b",
            "grid.color": "#d8c59d",
            "grid.linestyle": "--",
            "grid.linewidth": 0.75,
            "grid.alpha": 0.55,
            "lines.linewidth": 1.75,
            "axes.prop_cycle": cycler(color=island_palette),
            "legend.frameon": True,
            "legend.facecolor": "#fff8ea",
            "legend.edgecolor": "#d8c59d",
        }
    )


def require_path(config: dict[str, Any], key: str) -> Path:
    value = config.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{key} is required.")
    return Path(value)


def supplier_frequency_map(config: dict[str, Any]) -> dict[str, list[float]]:
    raw_map = config.get("supplier_excel_frequencies")
    if not isinstance(raw_map, dict) or not raw_map:
        raise ValueError("supplier_excel_frequencies must be a non-empty JSON object.")

    parsed: dict[str, list[float]] = {}
    for parameter, values in raw_map.items():
        if not isinstance(parameter, str) or not parameter.strip():
            raise ValueError("supplier_excel_frequencies parameter names must be non-empty strings.")
        normalized = core.normalize_s_parameter(parameter)
        if not core.is_supported_parameter(normalized) and normalized not in {"NEXT", "FEXT"}:
            raise ValueError(f"Unsupported supplier Excel parameter: {parameter}")
        if not isinstance(values, list) or not values:
            raise ValueError(f"supplier_excel_frequencies.{parameter} must be a non-empty list.")
        parsed[normalized] = sorted({core.parse_frequency(value) for value in values})
    return parsed


def supplier_processing_max_frequency_hz(
    config: dict[str, Any],
    plot_limits: core.PlotLimitSettings,
    target_freqs_by_parameter: dict[str, list[float]],
    additional_targets: list[core.AdditionalExcelTarget],
) -> float | None:
    trim_settings = core.resolve_frequency_trim_settings(config)
    if not trim_settings.enabled:
        logging.info("Frequency trimming disabled.")
        return None
    if trim_settings.max_hz is not None:
        logging.info("Frequency trimming uses manual max %.6g Hz.", trim_settings.max_hz)
        return trim_settings.max_hz

    candidates: list[float] = []
    for frequencies in target_freqs_by_parameter.values():
        candidates.extend(frequencies)
    for target in additional_targets:
        candidates.extend(target.frequencies_hz)
        plot_max = target.plot_limits.x_max_hz if target.plot_limits is not None else None
        if plot_max is not None:
            candidates.append(plot_max)
    for limits in [plot_limits.global_limits, *plot_limits.per_plot.values()]:
        if limits.x_max_hz is not None:
            candidates.append(limits.x_max_hz)

    finite = [value for value in candidates if value > 0 and math.isfinite(value)]
    if not finite:
        logging.info("Frequency trimming skipped because no processing max was available.")
        return None

    max_hz = max(finite) * trim_settings.margin_ratio
    logging.info(
        "Frequency trimming auto max %.6g Hz from required max %.6g Hz and margin_ratio %.4g.",
        max_hz,
        max(finite),
        trim_settings.margin_ratio,
    )
    return max_hz


def write_supplier_excel(
    summary_df: pd.DataFrame,
    per_file_df: pd.DataFrame,
    standard_deviation_df: pd.DataFrame,
    output_dir: Path,
    picture_paths: list[Path],
    additional_summary_df: pd.DataFrame,
    additional_per_file_df: pd.DataFrame,
    additional_picture_paths: list[Path],
    output_prefix: str,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = output_dir / core.prefixed_output_name("summary_results.xlsx", output_prefix)
    all_picture_paths = [*picture_paths, *additional_picture_paths]
    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="summary")
        per_file_df.to_excel(writer, index=False, sheet_name="per_file")
        standard_deviation_df.to_excel(
            writer, index=False, sheet_name="standard deviation"
        )
        if not additional_summary_df.empty:
            additional_summary_df.to_excel(writer, index=False, sheet_name="additional results")
        if not additional_per_file_df.empty:
            additional_per_file_df.to_excel(writer, index=False, sheet_name="additional per file")
        core.add_picture_sheet(writer, "pure picture", all_picture_paths)
    logging.info("Wrote Excel: %s", workbook_path)


def build_standard_deviation_table(per_file_df: pd.DataFrame) -> pd.DataFrame:
    group_columns = [
        column
        for column in (
            "data_group",
            "source_group",
            "parameter",
            "target_frequency_hz",
            "target_frequency_mhz",
        )
        if column in per_file_df.columns
    ]
    if per_file_df.empty or not group_columns:
        return pd.DataFrame(
            columns=[
                *group_columns,
                "target_frequency_mhz",
                "sample_count",
                "mean_db",
                "standard_deviation_db",
                "mean_linear_magnitude",
                "standard_deviation_linear_magnitude",
                "minimum_db",
                "maximum_db",
            ]
        )

    stdev_df = (
        per_file_df.groupby(group_columns, as_index=False)
        .agg(
            sample_count=("magnitude_db", "count"),
            mean_db=("magnitude_db", "mean"),
            standard_deviation_db=("magnitude_db", "std"),
            mean_linear_magnitude=("linear_magnitude", "mean"),
            standard_deviation_linear_magnitude=("linear_magnitude", "std"),
            minimum_db=("magnitude_db", "min"),
            maximum_db=("magnitude_db", "max"),
        )
        .sort_values(group_columns, kind="stable")
    )
    if "target_frequency_hz" in stdev_df.columns and "target_frequency_mhz" not in stdev_df.columns:
        stdev_df["target_frequency_mhz"] = stdev_df["target_frequency_hz"] / 1e6

    preferred_columns = [
        column
        for column in (
            "data_group",
            "source_group",
            "parameter",
            "target_frequency_hz",
            "target_frequency_mhz",
            "sample_count",
            "mean_db",
            "standard_deviation_db",
            "mean_linear_magnitude",
            "standard_deviation_linear_magnitude",
            "minimum_db",
            "maximum_db",
        )
        if column in stdev_df.columns
    ]
    return stdev_df.loc[:, preferred_columns].reset_index(drop=True)


def run_supplier(config: dict[str, Any], config_path: Path) -> int:
    apply_supplier_plot_style()
    input_dir = require_path(config, "input_dir")
    output_dir = require_path(config, "output_dir")
    target_freqs_by_parameter = supplier_frequency_map(config)
    target_freqs = sorted({freq for freqs in target_freqs_by_parameter.values() for freq in freqs})
    if not target_freqs:
        raise ValueError("At least one supplier target frequency is required.")

    output_prefix = core.resolve_output_prefix(config)
    plot_limits = core.resolve_plot_limits(config)
    recursive_search = core.resolve_recursive_search(config)
    title_settings = core.resolve_title_settings(config)
    legend_settings = core.resolve_legend_settings(config)
    average_settings = core.resolve_plot_average_settings(config)
    exclude_keywords = core.resolve_exclude_keywords(config)
    xtk_settings = core.resolve_xtk_plot_settings(config)
    additional_targets = core.resolve_additional_excel_targets(config)
    processing_workers = core.resolve_processing_workers(config)

    core.setup_logging(output_dir, output_prefix)
    core.log_progress(0)
    logging.info("Supplier simple processor.")
    logging.info("Config file: %s", config_path)
    logging.info("Input folder: %s", input_dir)
    logging.info("Output folder: %s", output_dir)
    logging.info("Output filename prefix: %s", output_prefix if output_prefix else "(none)")
    logging.info("Specification limit overlays: not used in supplier simple processor.")
    logging.info("Supplier Excel target frequency keys: %s", sorted(target_freqs_by_parameter))
    logging.info("Touchstone processing workers: %d", processing_workers)

    limit_database = None
    limit_settings = core.LimitOverlaySettings(enabled=False)
    max_frequency_hz = supplier_processing_max_frequency_hz(
        config, plot_limits, target_freqs_by_parameter, additional_targets
    )

    s4p_files = core.discover_s4p_files(input_dir, recursive_search, exclude_keywords)
    if not s4p_files:
        raise FileNotFoundError(f"No .s4p files found under: {input_dir}")

    logging.info("Found %d .s4p files.", len(s4p_files))
    core.log_progress(10)

    read_started_at = perf_counter()
    processed, failed_count = core.process_paths(s4p_files, processing_workers, max_frequency_hz)
    logging.info("Main Touchstone read/convert stage took %.2fs.", perf_counter() - read_started_at)
    core.log_progress(35)
    if not processed:
        raise RuntimeError("All .s4p files failed to process. See processing_log.txt.")

    main_enabled_plot_parameters = set(target_freqs_by_parameter)
    main_plot_specs = core.filtered_plot_specs(core.PLOT_SPECS, main_enabled_plot_parameters)

    plot_started_at = perf_counter()
    pure_picture_paths = core.make_plots(
        processed,
        output_dir,
        plot_limits,
        title_settings,
        legend_settings,
        average_settings,
        limit_database,
        limit_settings,
        plot_specs=main_plot_specs,
    )
    logging.info("Pure picture plotting stage took %.2fs.", perf_counter() - plot_started_at)
    core.log_progress(45)

    xtk_started_at = perf_counter()
    (
        xtk_successful_count,
        xtk_failed_count,
        xtk_summary_df,
        xtk_per_file_df,
        xtk_pure_picture_paths,
        xtk_plot_jobs,
    ) = core.make_xtk_plots(
        input_dir,
        output_dir,
        xtk_settings,
        exclude_keywords,
        plot_limits,
        title_settings,
        legend_settings,
        average_settings,
        limit_database,
        limit_settings,
        target_freqs,
        target_freqs_by_parameter,
        additional_targets,
        processing_workers,
        max_frequency_hz,
    )
    logging.info("XTK read/plot/table stage took %.2fs.", perf_counter() - xtk_started_at)
    failed_count += xtk_failed_count
    pure_picture_paths.extend(xtk_pure_picture_paths)
    core.log_progress(60)

    table_started_at = perf_counter()
    xtk_processed_by_group = {
        plot_type: xtk_processed for plot_type, xtk_processed, _plot_specs in xtk_plot_jobs
    }
    additional_summary_df, additional_per_file_df = core.build_additional_result_tables(
        processed,
        xtk_processed_by_group,
        additional_targets,
    )
    main_target_freqs_by_parameter = {
        parameter: frequencies
        for parameter, frequencies in target_freqs_by_parameter.items()
        if core.is_supported_parameter(parameter) and parameter not in {"NEXT", "FEXT"}
    }
    summary_df, per_file_df = core.build_result_tables(
        processed,
        target_freqs,
        parameter_names=list(main_target_freqs_by_parameter),
        extra_columns={"data_group": "MAIN"},
        target_freqs_by_parameter=main_target_freqs_by_parameter,
    )
    if not xtk_summary_df.empty and not xtk_per_file_df.empty:
        summary_df = pd.concat([summary_df, xtk_summary_df], ignore_index=True)
        per_file_df = pd.concat([per_file_df, xtk_per_file_df], ignore_index=True)

    ordered_summary_df = core.order_summary_columns(core.sort_summary_results(summary_df))
    standard_deviation_df = build_standard_deviation_table(per_file_df)
    if not additional_per_file_df.empty:
        standard_deviation_df = pd.concat(
            [
                standard_deviation_df,
                build_standard_deviation_table(additional_per_file_df),
            ],
            ignore_index=True,
        )
    logging.info("Summary table stage took %.2fs.", perf_counter() - table_started_at)
    core.log_progress(75)

    additional_picture_paths = core.make_additional_plots(
        processed,
        xtk_processed_by_group,
        additional_targets,
        output_dir,
        title_settings,
        legend_settings,
        average_settings,
    )

    excel_started_at = perf_counter()
    write_supplier_excel(
        ordered_summary_df,
        per_file_df,
        standard_deviation_df,
        output_dir,
        pure_picture_paths,
        additional_summary_df,
        additional_per_file_df,
        additional_picture_paths,
        output_prefix,
    )
    logging.info("Excel writing stage took %.2fs.", perf_counter() - excel_started_at)
    core.log_progress(95)

    cleanup_started_at = perf_counter()
    core.remove_written_plots([*pure_picture_paths, *additional_picture_paths])
    logging.info("Temporary PNG cleanup stage took %.2fs.", perf_counter() - cleanup_started_at)
    core.log_progress(100)
    logging.info(
        "Done. Successful files: %d, failed files: %d, XTK successful files: %d",
        len(processed),
        failed_count,
        xtk_successful_count,
    )
    return 0 if failed_count == 0 else 2


def main() -> int:
    try:
        args = parse_args()
        config = read_config(args.config)
        return run_supplier(config, args.config)
    except Exception as exc:
        logging.exception("Supplier processing stopped: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
