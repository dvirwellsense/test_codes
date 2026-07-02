"""Opt-in Timer/WDT test - NOT part of the default main.py run.

Sends the `Timer` command, which makes the FW hang in a deliberate
infinite loop, then waits for the hardware watchdog to reset the
board (~10 minutes, see test_config.WDT_EXPECTED_TIMEOUT_SEC) and
confirms it recovers within tolerance and reports the WDT reset
error code (50) on the frame right after reboot.

This call blocks for roughly 10 minutes by design.

Usage:
    python wdt_timer_test.py
"""

import sys
import time

import serial

import test_config as cfg
from fw_interface import FWInterface, FrameTimeoutError
from frame_parser import parse_metadata


def wait_for_recovery(fw, max_wait, poll_timeout=3.0):

    start_time = time.time()

    while time.time() - start_time < max_wait:

        try:
            frame = fw.get_frame(timeout=poll_timeout)
            return frame, time.time() - start_time

        except FrameTimeoutError:
            continue

        except serial.SerialException:

            try:
                fw.reconnect(timeout=poll_timeout, poll_interval=0.5)
            except FrameTimeoutError:
                pass

    return None, time.time() - start_time


def main():

    print()
    print("===================================")
    print("      WDT / TIMER TEST")
    print("===================================")
    print()
    print(
        f"This will hang the board and wait up to "
        f"~{(cfg.WDT_EXPECTED_TIMEOUT_SEC + cfg.WDT_TIMEOUT_TOLERANCE_SEC) / 60:.0f} "
        f"minutes for the watchdog reset. Do not interrupt."
    )
    print()

    fw = FWInterface(cfg.COM_PORT)

    print("Confirming board is currently responsive...")

    try:
        fw.get_frame(timeout=5.0)
    except FrameTimeoutError as e:
        print(f"ERROR: board not responsive before starting the test: {e}")
        sys.exit(1)

    print("Sending Timer command (board will hang now)...")

    fw.send_command("Timer", timeout=0.5)

    max_wait = (
        cfg.WDT_EXPECTED_TIMEOUT_SEC +
        cfg.WDT_TIMEOUT_TOLERANCE_SEC +
        60.0  # extra buffer for USB re-enumeration lag
    )

    frame, elapsed = wait_for_recovery(fw, max_wait)

    print()

    if frame is None:
        print(
            f"FAIL: board did not recover within {max_wait:.0f}s "
            f"(no WDT reset detected)"
        )
        sys.exit(1)

    metadata = parse_metadata(frame)

    error_field = metadata.get("Error", "")
    wdt_error_seen = error_field.split(",")[0].strip() == "50"

    lo = cfg.WDT_EXPECTED_TIMEOUT_SEC - cfg.WDT_TIMEOUT_TOLERANCE_SEC
    hi = cfg.WDT_EXPECTED_TIMEOUT_SEC + cfg.WDT_TIMEOUT_TOLERANCE_SEC

    problems = []

    if not (lo <= elapsed <= hi):
        problems.append(
            f"recovery took {elapsed:.1f}s, expected {lo:.0f}-{hi:.0f}s"
        )

    if not wdt_error_seen:
        problems.append(
            f"no WDT reset error code (50) on first frame after recovery "
            f"(Error={error_field!r})"
        )

    if problems:
        print(f"FAIL: {'; '.join(problems)}")
        sys.exit(1)

    print(
        f"PASS: board recovered in {elapsed:.1f}s "
        f"(expected {lo:.0f}-{hi:.0f}s), WDT reset error code confirmed"
    )

    sys.exit(0)


if __name__ == "__main__":
    main()
