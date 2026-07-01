import test_config as cfg


def run(matrix, **kwargs):

    if len(matrix) != cfg.EXPECTED_NUM_ROWS:
        return False, f"Expected {cfg.EXPECTED_NUM_ROWS} rows, got {len(matrix)}"

    for i, row in enumerate(matrix, start=1):

        if len(row) != cfg.EXPECTED_TOTAL_COLS:
            return False, (
                f"Row {i}: expected {cfg.EXPECTED_TOTAL_COLS} cols, "
                f"got {len(row)}"
            )

    return True, "Matrix size OK"
