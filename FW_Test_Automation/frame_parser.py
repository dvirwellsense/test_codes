def parse_metadata(frame):

    metadata = {}

    for line in frame:

        if line.startswith("Row"):
            break

        parts = line.split(",")

        if len(parts) >= 2:
            # Keep ALL values after the key, not just the first one.
            # Some lines carry more than one field, e.g.:
            #   "Error,50,Watchdog timeout reset"  -> code + description
            #   "RefCaps,5,10,15"                  -> 3 values
            # Joining them back preserves everything while staying
            # backward compatible for single-value lines.
            values = [p for p in parts[1:] if p != ""]
            metadata[parts[0]] = ",".join(values) if values else ""

    return metadata


def get_row_numbers(frame):
    """Return the list of row indices, in the order they appear in
    the raw frame, e.g. "Row1", "Row2", ... -> [1, 2, ...].
    Used to detect missing / duplicated / out-of-order rows, which
    parse_matrix alone can't tell you (it only returns row content).
    """

    row_numbers = []

    for line in frame:

        if not line.startswith("Row"):
            continue

        header = line.split(",")[0]

        try:
            row_numbers.append(int(header[len("Row"):]))
        except ValueError:
            row_numbers.append(None)

    return row_numbers


def parse_active_range(range_str):
    """Parse a "<start>,<end>" 1-based inclusive range string - the
    format the FW stores for the ActiveRows/ActiveColumns metadata
    fields (e.g. "1,60"), confirmed against firmware's
    parse_range_fixed() - into its element count.

    Raises ValueError if the string isn't a valid 2-int range.
    """

    parts = range_str.split(",")

    if len(parts) != 2:
        raise ValueError(f"not a <start>,<end> range: {range_str!r}")

    start = int(parts[0])
    end = int(parts[1])

    return end - start + 1


def parse_matrix(frame):

    matrix = []

    for line in frame:

        if not line.startswith("Row"):
            continue

        parts = line.split(",")

        row = []

        for value in parts[1:]:

            if value == "":
                continue

            row.append(int(value))

        matrix.append(row)

    return matrix