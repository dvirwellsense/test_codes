from fw_interface import FWInterface
from frame_parser import parse_metadata
from frame_parser import parse_matrix

from tests import test_matrix
from tests import test_reference_caps
from tests import test_golden_sample

from datetime import datetime
import os
import time


def run_test(name, func, **kwargs):

    try:
        result, msg = func(**kwargs)

        print(
            f"{name:20} "
            f"{'PASS' if result else 'FAIL'} "
            f"{msg}"
        )

        return result

    except Exception as e:

        print(
            f"{name:20} "
            f"ERROR "
            f"{str(e)}"
        )

        return False


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

    fw = FWInterface("COM4")

    print("Waiting for frame...")

    start_time = time.time()

    frame = fw.get_frame()

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

    matrix = parse_matrix(frame)

    print_metadata(metadata)

    print(
        f"Matrix Size : "
        f"{len(matrix)} x {len(matrix[0])}"
    )

    print()

    tests = [
        ("Matrix Size", test_matrix.run),
        ("Reference Caps", test_reference_caps.run),
        ("Golden Sample", test_golden_sample.run),
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

    passed = sum(results.values())
    total = len(results)

    for name, result in results.items():

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


if __name__ == "__main__":
    main()