import test_config as cfg
from frame_parser import parse_matrix


def load_golden_matrix():

    with open(
        "golden/empty_matrix.txt",
        "r"
    ) as f:

        frame = [
            line.strip()
            for line in f.readlines()
        ]

    return parse_matrix(frame)


def run(matrix, **kwargs):

    golden = load_golden_matrix()

    if len(matrix) != len(golden):
        return False, (
            f"Row count mismatch: got {len(matrix)}, "
            f"golden has {len(golden)}"
        )

    bad_pixels = 0
    max_delta = 0

    data_cols = cfg.EXPECTED_NUM_DATA_COLS

    for row in range(len(matrix)):

        if len(matrix[row]) < data_cols or len(golden[row]) < data_cols:
            return False, f"Row {row + 1}: not enough columns to compare"

        for col in range(data_cols):  # without reference columns

            delta = abs(
                matrix[row][col] -
                golden[row][col]
            )

            max_delta = max(
                max_delta,
                delta
            )

            if delta > cfg.GOLDEN_TOLERANCE:
                bad_pixels += 1

    if bad_pixels > 0:

        return (
            False,
            f"{bad_pixels} pixels differ "
            f"(max delta={max_delta}, tolerance={cfg.GOLDEN_TOLERANCE})"
        )

    return (
        True,
        f"max delta={max_delta}"
    )
