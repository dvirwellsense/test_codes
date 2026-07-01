import test_config as cfg


def run(matrix, **kwargs):

    if not matrix:
        return False, "Empty matrix"

    ref1 = [row[-3] for row in matrix]
    ref2 = [row[-2] for row in matrix]
    ref3 = [row[-1] for row in matrix]

    avg1 = sum(ref1) / len(ref1)
    avg2 = sum(ref2) / len(ref2)
    avg3 = sum(ref3) / len(ref3)

    lo1, hi1 = cfg.REF1_RANGE
    lo2, hi2 = cfg.REF2_RANGE
    lo3, hi3 = cfg.REF3_RANGE

    if not (lo1 < avg1 < hi1):
        return False, f"Ref1={avg1:.1f} (expected {lo1}-{hi1})"

    if not (lo2 < avg2 < hi2):
        return False, f"Ref2={avg2:.1f} (expected {lo2}-{hi2})"

    if not (lo3 < avg3 < hi3):
        return False, f"Ref3={avg3:.1f} (expected {lo3}-{hi3})"

    return True, (
        f"Refs OK "
        f"({avg1:.1f}, {avg2:.1f}, {avg3:.1f})"
    )
