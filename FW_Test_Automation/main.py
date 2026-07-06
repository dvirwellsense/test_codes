import test_config as cfg
from fw_interface import FWInterface, FrameTimeoutError
from frame_parser import parse_metadata
from frame_parser import parse_matrix

from tests import test_matrix
from tests import test_reference_caps
from tests import test_golden_sample
from tests import test_metadata
from tests import test_pixel_range
from tests import test_frame_integrity
from tests import test_repeatability
from tests import test_required_fields
from tests import test_ref_caps_config
from tests import test_counters_sanity
from tests import test_hex_integrity
from tests import test_eeprom_metadata_update
from tests import test_active_size_matches_matrix
from tests import test_lut_roundtrip

from datetime import datetime
import os
import sys
import time


def run_test(name, func, **kwargs):

    print(f"Running: {name} ...")

    try:
        result, msg = func(**kwargs)

        print(
            f"{name:20} "
            f"{'PASS' if result else 'FAIL'} "
            f"{msg}"
        )

        return result, msg

    except Exception as e:

        print(
            f"{name:20} "
            f"ERROR "
            f"{str(e)}"
        )

        return False, f"ERROR: {e}"


def save_frame(frame):

    os.makedirs("logs", exist_ok=True)

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    filename = (
        f"logs/frame_{timestamp}.txt"
    )

    with open(filename, "w") as f:

        for line in frame:
            f.write(line + "\n")

    return filename


def save_report(metadata, results, aborted=False, reason=None):

    os.makedirs("reports", exist_ok=True)

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    filename = f"reports/report_{timestamp}.txt"

    with open(filename, "w") as f:

        f.write("FW AUTOMATION TEST REPORT\n")
        f.write("===================================\n")
        f.write(f"Timestamp    : {datetime.now().isoformat()}\n")
        f.write(f"COM Port     : {cfg.COM_PORT}\n")
        f.write(f"FW Version   : {metadata.get('FWVer', 'Unknown')}\n")
        f.write(f"HW Version   : {metadata.get('HWVer', 'Unknown')}\n")
        f.write(f"MatConnected : {metadata.get('MatConnected', 'Unknown')}\n")
        f.write("\n")

        if aborted:
            f.write(f"RUN ABORTED: {reason}\n")
            return filename

        f.write("TESTS\n")
        f.write("-----\n")

        for name, (result, msg) in results.items():
            f.write(
                f"{name:20} {'PASS' if result else 'FAIL'}  {msg}\n"
            )

        f.write("\n")

        passed = sum(1 for result, _ in results.values() if result)
        total = len(results)

        f.write(f"TOTAL: {passed}/{total} PASS\n")
        f.write(
            f"OVERALL RESULT: {'PASS' if passed == total else 'FAIL'}\n"
        )

    return filename


def wait_for_mat_connected(fw, frame, metadata, wait_sec):
    """Poll for MatConnected=="true" for up to wait_sec, re-fetching
    frames. Returns the last (frame, metadata) seen - still
    disconnected if it never came up within wait_sec.
    """

    deadline = time.time() + wait_sec

    while metadata.get("MatConnected", "").strip().lower() != "true":

        if time.time() >= deadline:
            break

        try:
            frame = fw.get_frame()
        except FrameTimeoutError:
            break

        metadata = parse_metadata(frame)

    return frame, metadata


def print_metadata(metadata):

    print(
        f"FW Version : "
        f"{metadata.get('FWVer', 'Unknown')}"
    )

    print(
        f"HW Version : "
        f"{metadata.get('HWVer', 'Unknown')}"
    )

    print(
        f"Temperature: "
        f"{metadata.get('Temperature', 'Unknown')}"
    )

    print(
        f"Humidity   : "
        f"{metadata.get('RelativeHumidity', 'Unknown')}"
    )

    print(
        f"Mat Number : "
        f"{metadata.get('MatNum', 'Unknown')}"
    )

    print()


def main():

    print()
    print("===================================")
    print("      FW AUTOMATION TESTER")
    print("===================================")
    print()

    fw = FWInterface(cfg.COM_PORT)

    print("Waiting for frame...")

    start_time = time.time()

    try:
        frame = fw.get_frame()
    except FrameTimeoutError as e:
        print(f"ERROR: {e}")
        print()
        print("OVERALL RESULT: FAIL (no frame received)")
        print()
        sys.exit(1)

    frame_time = time.time() - start_time

    print(
        f"Frame received in "
        f"{frame_time:.2f} sec"
    )

    print()

    log_file = save_frame(frame)

    print(
        f"Frame saved to: {log_file}"
    )

    print()

    metadata = parse_metadata(frame)

    # --- Mat-connect gate: only run the (mostly mat-dependent) test
    # suite once a mat is actually connected. Poll for a while first,
    # since the operator may still be plugging it in. ---
    if metadata.get("MatConnected", "").strip().lower() != "true":

        print(
            f"No mat connected yet - waiting up to "
            f"{cfg.MAT_CONNECT_WAIT_SEC:.0f}s..."
        )

        frame, metadata = wait_for_mat_connected(
            fw, frame, metadata, cfg.MAT_CONNECT_WAIT_SEC
        )

    if metadata.get("MatConnected", "").strip().lower() != "true":

        reason = "No mat connected within the wait window"

        print(f"ERROR: {reason}")

        report_file = save_report(metadata, {}, aborted=True, reason=reason)

        print(f"Report saved to: {report_file}")
        print()
        print("OVERALL RESULT: FAIL (no mat connected)")
        print()

        sys.exit(1)

    # --- Let ADC readings settle before trusting pixel data. The very
    # first frame right after MatConnected flips true can still reflect
    # a not-yet-averaged/settled reading (observed as all-zero pixels
    # and Ref1=0.0 on real hardware) - discard it and grab a fresh one. ---
    print(
        f"Mat connected - letting readings settle "
        f"({cfg.MAT_SETTLE_SEC:.1f}s)..."
    )

    time.sleep(cfg.MAT_SETTLE_SEC)

    frame = fw.get_frame()
    metadata = parse_metadata(frame)

    matrix = parse_matrix(frame)

    print_metadata(metadata)

    print(
        f"Matrix Size : "
        f"{len(matrix)} x {len(matrix[0])}"
    )

    print()

    tests = [
        ("Required Fields", test_required_fields.run),
        ("Metadata", test_metadata.run),
        ("Counters Sanity", test_counters_sanity.run),
        ("Frame Integrity", test_frame_integrity.run),
        ("Matrix Size", test_matrix.run),
        ("Pixel Range", test_pixel_range.run),
        ("Reference Caps", test_reference_caps.run),
        ("RefCaps Config", test_ref_caps_config.run),
        ("Golden Sample", test_golden_sample.run),
        ("Repeatability", test_repeatability.run),
        ("HEX Integrity", test_hex_integrity.run),
        ("Active Size Match", test_active_size_matches_matrix.run),
        ("EEPROM Metadata Update", test_eeprom_metadata_update.run),
        ("LUT Roundtrip", test_lut_roundtrip.run),
    ]

    results = {}

    print("========== TESTS ==========\n")

    for name, func in tests:

        results[name] = run_test(
            name,
            func,
            matrix=matrix,
            metadata=metadata,
            frame=frame,
            fw=fw
        )

    print()
    print("========== SUMMARY ==========\n")

    passed = sum(1 for result, _ in results.values() if result)
    total = len(results)

    for name, (result, _msg) in results.items():

        print(
            f"{name:20} "
            f"{'PASS' if result else 'FAIL'}"
        )

    print()

    print(
        f"TOTAL: {passed}/{total} PASS"
    )

    print()

    if passed == total:
        print("OVERALL RESULT: PASS")
    else:
        print("OVERALL RESULT: FAIL")

    print()

    report_file = save_report(metadata, results)

    print(f"Report saved to: {report_file}")
    print()


if __name__ == "__main__":
    main()