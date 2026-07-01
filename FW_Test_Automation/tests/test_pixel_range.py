import test_config as cfg


def run(matrix, **kwargs):

    problems = []

    out_of_range = 0
    saturated = 0
    stuck_rows = []

    data_cols = cfg.EXPECTED_NUM_DATA_COLS

    for row_idx, row in enumerate(matrix, start=1):

        data_values = row[:data_cols]

        for value in data_values:

            if value < cfg.PIXEL_MIN or value > cfg.PIXEL_MAX:
                out_of_range += 1

            if value >= cfg.PIXEL_MAX:
                saturated += 1

        # A whole row reading the exact same value on every sensor
        # pixel almost always means that row's readout channel is
        # stuck rather than a real, uniform surface.
        if len(set(data_values)) < cfg.STUCK_ROW_MIN_UNIQUE:
            stuck_rows.append(row_idx)

    if out_of_range > 0:
        problems.append(
            f"{out_of_range} pixels outside "
            f"[{cfg.PIXEL_MIN},{cfg.PIXEL_MAX}]"
        )

    if saturated > 0:
        problems.append(f"{saturated} pixels saturated at max ADC value")

    if stuck_rows:
        problems.append(f"stuck/uniform rows: {stuck_rows}")

    if problems:
        return False, "; ".join(problems)

    return True, "All pixel values within range, no stuck rows"
