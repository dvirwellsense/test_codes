import test_config as cfg
from frame_parser import parse_metadata


# Maps each field to (write_cmd, get_cmd, get_response_prefix).
# get_response_prefix is what to strip off the FW's Get* reply to
# recover the raw value - the FW is NOT consistent about format
# (some Get* replies are "Key,value", others are "Key saved: value").
_MAT_FIELDS = {
    "MatNum": ("MatNum", "GetMatNum", "MatNum,"),
    "ActiveRows": ("ActiveRows", "GetActiveRows", "ActiveRows saved: "),
    "ActiveColumns": ("ActiveColumns", "GetActiveColumns", "ActiveColumns saved: "),
    "MatLifeTime": ("MatLifeTime", "GetMatLifeTime", "MatLifeTime,"),
    "MatActiveTime": ("MatActiveTime", "GetMatActiveTime", "MatActiveTime,"),
}

# The frame's periodic metadata uses a different key than the write
# command for a couple of these fields (firmware naming is inconsistent).
_FRAME_KEY = {
    "MatNum": "MatNum",
    "ActiveRows": "ActiveRows",
    "ActiveColumns": "ActiveColumns",
    "MatLifeTime": "MatLifetime",   # note lowercase "t", unlike the write/get commands
    "MatActiveTime": "MatActiveTime",
}


def _read_field(fw, field):

    write_cmd, get_cmd, prefix = _MAT_FIELDS[field]

    lines = fw.send_command(get_cmd, timeout=2.0)

    for line in lines:
        if line.startswith(prefix):
            return line[len(prefix):]

    return None


def _write_field(fw, field, value):

    write_cmd = _MAT_FIELDS[field][0]

    fw.send_command(f"{write_cmd},{value}", timeout=1.0)


def run(fw, metadata, **kwargs):

    if fw is None:
        return False, "No FW interface provided, ca nnot send EEPROM commands"

    problems = []

    # --- Snapshot current values so they can be restored afterward ---
    snapshot = {}

    for field in _MAT_FIELDS:
        snapshot[field] = _read_field(fw, field)

    pcba_lifetime_snapshot = metadata.get("PCBALifetime")

    try:
        # --- Write test values ---
        for field, value in cfg.EEPROM_TEST_VALUES.items():
            _write_field(fw, field, value)

        fw.send_command(
            f"Lifetime,{cfg.PCBA_LIFETIME_TEST_VALUE}",
            timeout=1.0
        )

        # --- Verify via a fresh frame ---
        frame2 = fw.get_frame()
        metadata2 = parse_metadata(frame2)

        for field, expected in cfg.EEPROM_TEST_VALUES.items():

            frame_key = _FRAME_KEY[field]
            actual = metadata2.get(frame_key)

            if actual != str(expected):
                problems.append(
                    f"{field}: wrote {expected!r}, "
                    f"frame reports {frame_key}={actual!r}"
                )

        actual_pcba_lifetime = metadata2.get("PCBALifetime")

        if actual_pcba_lifetime != str(cfg.PCBA_LIFETIME_TEST_VALUE):
            problems.append(
                f"PCBA Lifetime: wrote {cfg.PCBA_LIFETIME_TEST_VALUE}, "
                f"frame reports PCBALifetime={actual_pcba_lifetime!r}"
            )

    finally:
        # --- Always restore original values, best-effort ---
        restore_problems = []

        for field, original in snapshot.items():

            if original is None:
                restore_problems.append(
                    f"{field} had no original value to restore "
                    f"(was unset before the test)"
                )
                continue

            try:
                _write_field(fw, field, original)
            except Exception as e:
                restore_problems.append(f"{field} restore failed: {e}")

        if pcba_lifetime_snapshot is not None:
            try:
                fw.send_command(
                    f"Lifetime,{pcba_lifetime_snapshot}",
                    timeout=1.0
                )
            except Exception as e:
                restore_problems.append(f"PCBA Lifetime restore failed: {e}")

        if restore_problems:
            problems.append("RESTORE ISSUES: " + "; ".join(restore_problems))

    if problems:
        return False, "; ".join(problems)

    return True, (
        f"All {len(cfg.EEPROM_TEST_VALUES) + 1} EEPROM/metadata fields "
        f"updated and verified, originals restored"
    )
