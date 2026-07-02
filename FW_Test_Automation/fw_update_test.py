"""Opt-in FW update test - NOT part of the default main.py run.

Flashes each hex image listed in test_config.FW_UPDATE_IMAGES over
serial, waits for the board to reboot, and confirms FWVer matches the
expected version for that image. Each run overwrites the board's
firmware for real - only run this when you actually intend to reflash
the unit under test.

Usage:
    python fw_update_test.py
"""

import sys
import time

import test_config as cfg
from fw_interface import FWInterface, FrameTimeoutError
from frame_parser import parse_metadata


def run_one_update(fw, hex_path, expected_version):

    print(f"Flashing {hex_path} ...")

    try:
        response = fw.flash_firmware(hex_path, timeout=cfg.FW_UPDATE_TIMEOUT_SEC)
    except Exception as e:
        return False, f"flash_firmware raised: {e}"

    if not any("Done!" in line for line in response):
        return False, f"FW did not confirm success: {response}"

    print("FW confirmed update, waiting for reboot...")

    try:
        fw.reconnect(timeout=cfg.FW_UPDATE_REBOOT_TIMEOUT_SEC)
    except FrameTimeoutError as e:
        return False, f"Board did not come back after flashing: {e}"

    try:
        frame = fw.get_frame()
    except FrameTimeoutError as e:
        return False, f"No frame after reboot: {e}"

    metadata = parse_metadata(frame)

    actual_version = metadata.get("FWVer", "Unknown")

    if actual_version != expected_version:
        return False, (
            f"FWVer after update = {actual_version} "
            f"(expected {expected_version})"
        )

    return True, f"FWVer confirmed as {actual_version}"


def main():

    if not cfg.FW_UPDATE_IMAGES:
        print("test_config.FW_UPDATE_IMAGES is empty - nothing to do.")
        sys.exit(1)

    print()
    print("===================================")
    print("      FW UPDATE TEST")
    print("===================================")
    print()

    fw = FWInterface(cfg.COM_PORT)

    results = []

    for hex_path, expected_version in cfg.FW_UPDATE_IMAGES:

        start_time = time.time()

        passed, msg = run_one_update(fw, hex_path, expected_version)

        elapsed = time.time() - start_time

        results.append((hex_path, expected_version, passed, msg))

        print(
            f"{'PASS' if passed else 'FAIL'} "
            f"[{elapsed:.1f}s] {hex_path} -> {expected_version}: {msg}"
        )

        print()

    print("========== SUMMARY ==========\n")

    passed_count = sum(1 for _, _, passed, _ in results if passed)
    total = len(results)

    for hex_path, expected_version, passed, msg in results:
        print(f"{'PASS' if passed else 'FAIL'} {hex_path} -> {expected_version}")

    print()
    print(f"TOTAL: {passed_count}/{total} PASS")
    print()

    sys.exit(0 if passed_count == total else 1)


if __name__ == "__main__":
    main()
