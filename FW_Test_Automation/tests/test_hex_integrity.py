import test_config as cfg


def _checksum_ok(line):

    line = line.strip()

    if not line.startswith(":"):
        return False

    data = bytes.fromhex(line[1:])

    return (sum(data) & 0xFF) == 0


def run(**kwargs):

    hex_path = cfg.HEX_FILE_PATH

    if not hex_path:
        return True, "HEX file check disabled (test_config.HEX_FILE_PATH not set)"

    try:
        with open(hex_path, "r") as f:
            lines = f.readlines()
    except OSError as e:
        return False, f"Could not open HEX file '{hex_path}': {e}"

    extended_address = 0
    previous_end = None

    gap_count = 0
    overlap_count = 0
    checksum_errors = 0
    data_records = 0

    for line_number, line in enumerate(lines, start=1):

        line = line.strip()

        if not line:
            continue

        if not _checksum_ok(line):
            checksum_errors += 1
            continue

        length = int(line[1:3], 16)
        address = int(line[3:7], 16)
        record = int(line[7:9], 16)
        payload = line[9:9 + length * 2]

        if record == 0x04:
            extended_address = int(payload, 16) << 16
            continue

        if record != 0x00:
            continue

        real_address = extended_address + address
        end_address = real_address + length

        data_records += 1

        if previous_end is not None:

            if real_address > previous_end:
                gap_count += 1

            elif real_address < previous_end:
                overlap_count += 1

        previous_end = end_address

    if checksum_errors or gap_count or overlap_count:
        return False, (
            f"records={data_records}, "
            f"checksum_errors={checksum_errors}, "
            f"gaps={gap_count}, overlaps={overlap_count}"
        )

    return True, f"HEX contiguous, {data_records} records, checksums OK"
