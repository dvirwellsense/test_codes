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
        return False, "Row count mismatch"

    bad_pixels = 0
    max_delta = 0

    TOLERANCE = 5

    for row in range(len(matrix)):

        for col in range(30):  # ללא רפרנסים

            delta = abs(
                matrix[row][col] -
                golden[row][col]
            )

            max_delta = max(
                max_delta,
                delta
            )

            if delta > TOLERANCE:
                bad_pixels += 1

    if bad_pixels > 0:

        return (
            False,
            f"{bad_pixels} pixels differ "
            f"(max delta={max_delta})"
        )

    return (
        True,
        f"max delta={max_delta}"
    )