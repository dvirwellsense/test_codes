import test_config as cfg
from frame_parser import parse_active_range


def run(metadata, **kwargs):

    problems = []

    # --- Mat connection state ---
    # NOTE: MatConnected only reflects whether a sensing mat is
    # physically plugged into the board right now. It's informational,
    # not a pass/fail condition on its own - a board can legitimately
    # be tested with no mat attached.
    mat_connected = metadata.get("MatConnected", "").strip().lower()

    if mat_connected not in ("true", "false"):
        problems.append(
            f"MatConnected has an unexpected value: "
            f"{metadata.get('MatConnected', 'Unknown')}"
        )

    # --- Boot / runtime errors reported by the FW ---
    error_field = metadata.get("Error")

    if error_field is not None:

        error_code = error_field.split(",")[0].strip()

        if error_code not in cfg.ACCEPTABLE_ERROR_CODES:
            problems.append(f"Error reported by FW: {error_field}")

    # --- FW / HW version match ---
    fw_ver = metadata.get("FWVer", "Unknown")
    hw_ver = metadata.get("HWVer", "Unknown")

    if cfg.EXPECTED_FW_VERSION and fw_ver != cfg.EXPECTED_FW_VERSION:
        problems.append(
            f"FWVer={fw_ver} (expected {cfg.EXPECTED_FW_VERSION})"
        )

    if cfg.EXPECTED_HW_VERSION and hw_ver != cfg.EXPECTED_HW_VERSION:
        problems.append(
            f"HWVer={hw_ver} (expected {cfg.EXPECTED_HW_VERSION})"
        )

    # --- Temperature / humidity sanity ---
    try:
        temperature = float(metadata.get("Temperature", "nan"))

        if not (cfg.TEMPERATURE_MIN_C <= temperature <= cfg.TEMPERATURE_MAX_C):
            problems.append(f"Temperature out of range: {temperature}")

    except ValueError:
        problems.append(f"Temperature not numeric: {metadata.get('Temperature')}")

    try:
        humidity = float(metadata.get("RelativeHumidity", "nan"))

        if not (cfg.HUMIDITY_MIN_PCT <= humidity <= cfg.HUMIDITY_MAX_PCT):
            problems.append(f"RelativeHumidity out of range: {humidity}")

    except ValueError:
        problems.append(
            f"RelativeHumidity not numeric: {metadata.get('RelativeHumidity')}"
        )

    # --- Active vs total rows/cols ---
    # ActiveRows/ActiveColumns are "<start>,<end>" 1-based inclusive range
    # strings (e.g. "1,60"), NOT plain integers - confirmed against the FW's
    # parse_range_fixed() and against live hardware. NumOfRows/NumOfColumns
    # ARE plain integers: the FW reports the active-derived size when a mat
    # is connected, or the full physical grid size when it isn't.
    #
    # ActiveRows/ActiveColumns persist their last-configured range
    # regardless of connection state (the FW does not zero them out on
    # disconnect), so this check only applies while connected.
    try:
        total_rows = int(metadata.get("NumOfRows", -1))
        total_cols = int(metadata.get("NumOfColumns", -1))
    except ValueError:
        total_rows = total_cols = -1
        problems.append("NumOfRows/NumOfColumns not numeric")

    if mat_connected == "true":

        try:
            active_rows = parse_active_range(metadata.get("ActiveRows", ""))
            active_cols = parse_active_range(metadata.get("ActiveColumns", ""))
        except ValueError:
            active_rows = active_cols = -1
            problems.append("ActiveRows/ActiveColumns not a valid <start>,<end> range")

        if not (0 < active_rows <= total_rows):
            problems.append(
                f"MatConnected=true but active rows={active_rows} "
                f"(expected 0 < active rows <= {total_rows})"
            )

        if not (0 < active_cols <= total_cols):
            problems.append(
                f"MatConnected=true but active cols={active_cols} "
                f"(expected 0 < active cols <= {total_cols})"
            )

    if problems:
        return False, "; ".join(problems)

    return True, (
        f"FWVer={fw_ver}, HWVer={hw_ver}, "
        f"MatConnected={mat_connected}, no errors"
    )
