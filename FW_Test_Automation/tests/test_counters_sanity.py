import test_config as cfg


def run(metadata, **kwargs):

    problems = []

    for field in cfg.COUNTER_FIELDS:

        raw = metadata.get(field)

        if raw is None:
            problems.append(f"{field} missing")
            continue

        try:
            value = int(raw)
        except ValueError:
            problems.append(f"{field}={raw} is not an integer")
            continue

        if value < 0:
            problems.append(f"{field}={value} is negative")

    if problems:
        return False, "; ".join(problems)

    return True, "All counters are valid non-negative integers"
