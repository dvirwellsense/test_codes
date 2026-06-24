def parse_metadata(frame):

    metadata = {}

    for line in frame:

        if line.startswith("Row"):
            break

        parts = line.split(",")

        if len(parts) >= 2:
            metadata[parts[0]] = parts[1]

    return metadata


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