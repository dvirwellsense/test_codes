def run(matrix, **kwargs):

    ref1 = [row[-3] for row in matrix]
    ref2 = [row[-2] for row in matrix]
    ref3 = [row[-1] for row in matrix]

    avg1 = sum(ref1) / len(ref1)
    avg2 = sum(ref2) / len(ref2)
    avg3 = sum(ref3) / len(ref3)

    if not (600 < avg1 < 750):
        return False, f"Ref1={avg1:.1f}"

    if not (1200 < avg2 < 1450):
        return False, f"Ref2={avg2:.1f}"

    if not (1800 < avg3 < 2100):
        return False, f"Ref3={avg3:.1f}"

    return True, (
        f"Refs OK "
        f"({avg1:.1f}, {avg2:.1f}, {avg3:.1f})"
    )