import test_config as cfg


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

    # --- Active vs total rows/cols, consistent WITH the connection state ---
    # (not simply "must equal total" - a disconnected mat is expected
    # to report 0 active rows/cols, that is correct FW behavior)
    try:
        active_rows = int(metadata.get("ActiveRows", -1))
        active_cols = int(metadata.get("ActiveColumns", -1))
        total_rows = int(metadata.get("NumOfRows", -1))
        total_cols = int(metadata.get("NumOfColumns", -1))
    except ValueError:
        active_rows = active_cols = total_rows = total_cols = -1
        problems.append("ActiveRows/ActiveColumns/NumOfRows/NumOfColumns not numeric")

    if mat_connected == "true":

        if not (0 < active_rows <= total_rows):
            problems.append(
                f"MatConnected=true but ActiveRows={active_rows} "
                f"(expected 0 < ActiveRows <= {total_rows})"
            )

        if not (0 < active_cols <= total_cols):
            problems.append(
                f"MatConnected=true but ActiveColumns={active_cols} "
                f"(expected 0 < ActiveColumns <= {total_cols})"
            )

    elif mat_connected == "false":

        if active_rows != 0:
            problems.append(
                f"MatConnected=false but ActiveRows={active_rows} (expected 0)"
            )

        if active_cols != 0:
            problems.append(
                f"MatConnected=false but ActiveColumns={active_cols} (expected 0)"
            )

    if problems:
        return False, "; ".join(problems)

    return True, (
        f"FWVer={fw_ver}, HWVer={hw_ver}, "
        f"MatConnected={mat_connected}, no errors"
    )
