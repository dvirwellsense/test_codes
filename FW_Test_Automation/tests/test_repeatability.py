import test_config as cfg
from frame_parser import parse_matrix, parse_metadata


def run(fw, matrix, metadata, **kwargs):

    if fw is None:
        return False, "No FW interface provided, cannot capture 2nd frame"

    frame2 = fw.get_frame()
    matrix2 = parse_matrix(frame2)
    metadata2 = parse_metadata(frame2)

    problems = []

    # --- Frame counter must move forward, otherwise the FW is
    # sending a stale/duplicate frame instead of a live one ---
    try:
        frame_num1 = int(metadata.get("Frame", -1))
        frame_num2 = int(metadata2.get("Frame", -1))

        if frame_num2 <= frame_num1:
            problems.append(
                f"Frame counter did not advance: {frame_num1} -> {frame_num2}"
            )

    except ValueError:
        problems.append("Frame counter not numeric")

    # --- Temperature / humidity should not jump unrealistically
    # between two back-to-back reads (a real sudden jump usually
    # means a parsing/read glitch, not physical reality) ---
    try:
        temp_delta = abs(
            float(metadata2.get("Temperature", "nan")) -
            float(metadata.get("Temperature", "nan"))
        )

        if temp_delta > cfg.MAX_TEMPERATURE_JUMP_C:
            problems.append(f"Temperature jumped by {temp_delta:.1f}C between frames")

    except ValueError:
        problems.append("Temperature not numeric in one of the frames")

    try:
        humidity_delta = abs(
            float(metadata2.get("RelativeHumidity", "nan")) -
            float(metadata.get("RelativeHumidity", "nan"))
        )

        if humidity_delta > cfg.MAX_HUMIDITY_JUMP_PCT:
            problems.append(
                f"Humidity jumped by {humidity_delta:.1f}% between frames"
            )

    except ValueError:
        problems.append("RelativeHumidity not numeric in one of the frames")

    # --- Pixel-level noise between the 2 frames ---
    # Only the data columns are held to REPEATABILITY_TOLERANCE - the 3
    # reference-capacitor columns are a different measurement path and
    # aren't known to share the same noise budget, so they're tracked
    # separately (informational only) instead of being lumped into the
    # same pass/fail check.
    ref_max_delta = 0

    if len(matrix2) != len(matrix):
        problems.append(
            f"Row count changed between frames: "
            f"{len(matrix)} -> {len(matrix2)}"
        )
    else:

        bad_pixels = 0
        max_delta = 0

        data_cols = cfg.EXPECTED_NUM_DATA_COLS

        for row_idx in range(len(matrix)):

            row1 = matrix[row_idx]
            row2 = matrix2[row_idx]

            if len(row1) != len(row2):
                problems.append(f"Row {row_idx + 1} length changed between frames")
                break

            for col_idx in range(data_cols):

                delta = abs(row1[col_idx] - row2[col_idx])
                max_delta = max(max_delta, delta)

                if delta > cfg.REPEATABILITY_TOLERANCE:
                    bad_pixels += 1

            for col_idx in range(data_cols, len(row1)):

                ref_delta = abs(row1[col_idx] - row2[col_idx])
                ref_max_delta = max(ref_max_delta, ref_delta)

        if bad_pixels > cfg.REPEATABILITY_MAX_BAD_PIXELS:
            problems.append(
                f"{bad_pixels} data pixels noisy between consecutive frames "
                f"(max delta={max_delta}, tolerance={cfg.REPEATABILITY_TOLERANCE})"
            )

    if problems:
        return False, "; ".join(problems)

    return True, (
        f"Frame advanced, env. stable, data pixels stable across 2 frames "
        f"(ref cols max delta={ref_max_delta}, not gated on tolerance)"
    )
