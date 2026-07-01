import test_config as cfg
from frame_parser import get_row_numbers


def run(frame, matrix, metadata, **kwargs):

    problems = []

    # --- Row order / completeness, from the raw frame ---
    row_numbers = get_row_numbers(frame)

    expected = list(range(1, cfg.EXPECTED_NUM_ROWS + 1))

    if row_numbers != expected:

        missing = sorted(set(expected) - set(row_numbers))
        duplicates = sorted({
            n for n in row_numbers if row_numbers.count(n) > 1
        })

        detail = []

        if missing:
            detail.append(f"missing={missing}")

        if duplicates:
            detail.append(f"duplicates={duplicates}")

        if row_numbers != sorted(row_numbers) and not (missing or duplicates):
            detail.append("rows out of order")

        problems.append(
            "Row sequence mismatch" +
            (f" ({', '.join(detail)})" if detail else "")
        )

    # --- Matrix dimensions vs what the FW itself reported ---
    try:
        declared_rows = int(metadata.get("NumOfRows", -1))
        declared_cols = int(metadata.get("NumOfColumns", -1))
    except ValueError:
        declared_rows, declared_cols = -1, -1

    if declared_rows != len(matrix):
        problems.append(
            f"metadata NumOfRows={declared_rows} != "
            f"parsed rows={len(matrix)}"
        )

    if matrix and declared_cols != (len(matrix[0]) - cfg.NUM_REF_COLS):
        problems.append(
            f"metadata NumOfColumns={declared_cols} != "
            f"parsed data columns={len(matrix[0]) - cfg.NUM_REF_COLS}"
        )

    if problems:
        return False, "; ".join(problems)

    return True, "Row order and declared dimensions consistent"
