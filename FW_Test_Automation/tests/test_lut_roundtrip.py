import test_config as cfg

INT32_MIN = -(2 ** 31)
INT32_MAX = 2 ** 31 - 1


def _scale(value, scale_factor):
    """Scale a float LUT coefficient to the int32 wire format the FW
    expects, matching the reference LT_GUI upload tool
    (Math.Round(v * scale) clamped to int32 range).
    """

    scaled = round(float(value) * scale_factor)

    return max(min(scaled, INT32_MAX), INT32_MIN)


def load_lut_file(path, scale_factor):
    """Parse an LT_GUI-exported LUT CSV:

        <header row - column labels, skipped>
        Row1,a1,b1,c1,d1,a2,b2,c2,d2,...
        Row2,...
        ...
        RowN,...
        <blank line>
        DEFAULT_LUT
        C3,C2,C1,C0            <- header, not parsed
        <default coefficient values>

    Returns (rows, def_values):
      rows       - list of (label, [scaled int values]) in file order
      def_values - list of 4 scaled ints, in whatever order the file has them
    """

    with open(path, "r") as f:
        lines = f.readlines()

    default_index = next(
        (
            i for i, line in enumerate(lines)
            if line.split(",")[0].strip().lower() == "default_lut"
        ),
        None
    )

    if default_index is None:
        raise ValueError("DEFAULT_LUT section not found in file")

    rows = []

    # lines[0] is the column-label header - matrix data starts at 1
    for line in lines[1:default_index]:

        line = line.strip().rstrip(",")

        if not line:
            continue

        parts = line.split(",")

        label = parts[0].strip()

        if not label.lower().startswith("row"):
            raise ValueError(f"Expected a 'RowN,...' line, got: {line!r}")

        values = [_scale(v, scale_factor) for v in parts[1:] if v != ""]

        rows.append((label, values))

    coeff_line_index = default_index + 2

    if coeff_line_index >= len(lines):
        raise ValueError("DEFAULT_LUT coefficient line missing")

    coeff_line = lines[coeff_line_index].strip().rstrip(",")

    coeff_parts = [v for v in coeff_line.split(",") if v != ""][:4]

    if len(coeff_parts) != 4:
        raise ValueError(
            f"DEFAULT_LUT coefficient line needs 4 values, got: {coeff_line!r}"
        )

    def_values = [_scale(v, scale_factor) for v in coeff_parts]

    return rows, def_values


def _parse_mat_lut_response(lines):
    """Parse ReadMatLUT's "MatLUTn,v1,v2,..." lines into a flat,
    row-order list of values (ignores the "MatLUTn" row label - only
    the value sequence is checked against what was sent).
    """

    values = []

    for line in lines:

        if not line.startswith("MatLUT"):
            continue

        parts = line.split(",")

        values.extend(int(v) for v in parts[1:] if v != "")

    return values


def run(fw, **kwargs):

    if fw is None:
        return False, "No FW interface provided, cannot send LUT commands"

    if not cfg.LUT_FILE_PATH:
        return True, "LUT roundtrip check disabled (test_config.LUT_FILE_PATH not set)"

    try:
        rows, def_values = load_lut_file(cfg.LUT_FILE_PATH, cfg.LUT_SCALE_FACTOR)
    except (OSError, ValueError) as e:
        return False, f"Could not load LUT file '{cfg.LUT_FILE_PATH}': {e}"

    problems = []

    print(
        f"  [LUT] loaded {len(rows)} rows, "
        f"{sum(len(v) for _l, v in rows)} total values, "
        f"scale x{cfg.LUT_SCALE_FACTOR}"
    )

    # --- Per-pixel LUT: WriteMatLUT / ReadMatLUT ---
    expected_values = [v for _label, values in rows for v in values]

    write_response = fw.send_lut_matrix(rows, timeout=30.0)

    if not any("stored successfully" in line for line in write_response):
        problems.append(f"WriteMatLUT: unexpected response {write_response}")

    print("  [LUT] sending ReadMatLUT, waiting for readback...")

    read_response = fw.send_command(
        "ReadMatLUT",
        timeout=15.0,
        terminator="Done reading LUT values."
    )

    print(f"  [LUT] ReadMatLUT returned {len(read_response)} lines")

    actual_values = _parse_mat_lut_response(read_response)

    if actual_values != expected_values:
        mismatches = sum(
            1 for a, b in zip(actual_values, expected_values) if a != b
        )
        problems.append(
            f"MatLUT mismatch: sent {len(expected_values)} values, "
            f"read back {len(actual_values)} values, "
            f"{mismatches} differ in the overlapping range"
        )

    # --- DefLUT: WriteDefLUT / ReadDefLUT (values forwarded in file order,
    # no reordering - matches the reference upload tool) ---
    print(f"  [LUT] sending WriteDefLUT,{def_values}...")

    def_write_response = fw.send_command(
        "WriteDefLUT," + ",".join(str(v) for v in def_values),
        timeout=2.0
    )

    if not any("DefLUT saved" in line for line in def_write_response):
        problems.append(f"WriteDefLUT: unexpected response {def_write_response}")

    print("  [LUT] sending ReadDefLUT...")

    def_read_response = fw.send_command("ReadDefLUT", timeout=2.0)

    def_line = next(
        (line for line in def_read_response if line.startswith("DefaultLUT,")),
        None
    )

    if def_line is None:
        problems.append(f"ReadDefLUT: no DefaultLUT response {def_read_response}")
    else:
        actual_def_values = [int(v) for v in def_line[len("DefaultLUT,"):].split(",")]

        if actual_def_values != def_values:
            problems.append(
                f"DefLUT mismatch: sent {def_values}, read back {actual_def_values}"
            )

    if problems:
        return False, "; ".join(problems)

    return True, (
        f"MatLUT ({len(expected_values)} values, x{cfg.LUT_SCALE_FACTOR} scaled) "
        f"and DefLUT round-tripped correctly"
    )
