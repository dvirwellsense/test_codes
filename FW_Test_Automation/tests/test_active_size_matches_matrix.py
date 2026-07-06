import test_config as cfg
from frame_parser import parse_active_range


def run(matrix, metadata, **kwargs):

    if not matrix:
        return False, "Empty matrix"

    active_rows_raw = metadata.get("ActiveRows")
    active_cols_raw = metadata.get("ActiveColumns")

    if active_rows_raw is None or active_cols_raw is None:
        return False, "ActiveRows/ActiveColumns missing from metadata"

    try:
        expected_rows = parse_active_range(active_rows_raw)
        expected_cols = parse_active_range(active_cols_raw)
    except ValueError as e:
        return False, f"Could not parse active size: {e}"

    measured_rows = len(matrix)
    measured_cols = len(matrix[0]) - cfg.NUM_REF_COLS

    problems = []

    if measured_rows != expected_rows:
        problems.append(
            f"measured rows={measured_rows}, "
            f"active size (EEPROM ActiveRows={active_rows_raw})={expected_rows}"
        )

    if measured_cols != expected_cols:
        problems.append(
            f"measured cols={measured_cols}, "
            f"active size (EEPROM ActiveColumns={active_cols_raw})={expected_cols}"
        )

    if problems:
        return False, "; ".join(problems)

    return True, (
        f"Measured matrix ({measured_rows}x{measured_cols}) matches "
        f"active size defined in EEPROM"
    )
