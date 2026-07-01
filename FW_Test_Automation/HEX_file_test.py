from pathlib import Path

HEX_FILE = r"C:\Users\dvirs\Documents\Atmel Studio\7.0\NT_Bootloader_project\NT_PCBA_80\NT_usb\Debug\Version_02_14.hex"


def checksum_ok(line):
    line = line.strip()

    if not line.startswith(":"):
        return False

    data = bytes.fromhex(line[1:])

    return (sum(data) & 0xFF) == 0


extended_address = 0
previous_end = None

gap_count = 0
overlap_count = 0
checksum_errors = 0
data_records = 0

print("=" * 70)

with open(HEX_FILE, "r") as f:

    for line_number, line in enumerate(f, start=1):

        line = line.strip()

        if not line:
            continue

        if not checksum_ok(line):
            checksum_errors += 1
            print(f"Checksum ERROR line {line_number}")

        length = int(line[1:3], 16)
        address = int(line[3:7], 16)
        record = int(line[7:9], 16)
        payload = line[9:9 + length * 2]

        if record == 0x04:
            extended_address = int(payload, 16) << 16
            print(f"Extended address -> 0x{extended_address:08X}")
            continue

        if record != 0x00:
            continue

        real_address = extended_address + address
        end_address = real_address + length

        data_records += 1

        if previous_end is not None:

            if real_address > previous_end:
                gap = real_address - previous_end
                gap_count += 1
                print(
                    f"GAP at line {line_number}: "
                    f"0x{previous_end:08X} -> 0x{real_address:08X} "
                    f"({gap} bytes)"
                )

            elif real_address < previous_end:
                overlap = previous_end - real_address
                overlap_count += 1
                print(
                    f"OVERLAP at line {line_number}: "
                    f"0x{real_address:08X} overlaps "
                    f"{overlap} bytes"
                )

        previous_end = end_address

print("\n" + "=" * 70)
print(f"DATA records      : {data_records}")
print(f"Checksum errors   : {checksum_errors}")
print(f"Gaps              : {gap_count}")
print(f"Overlaps          : {overlap_count}")

if gap_count == 0 and overlap_count == 0:
    print("\n*** HEX IS CONTIGUOUS ***")
    print("*** Your current bootloader should handle it correctly. ***")
else:
    print("\n*** HEX IS NOT CONTIGUOUS ***")
    print("*** Your current parser WILL write wrong addresses. ***")