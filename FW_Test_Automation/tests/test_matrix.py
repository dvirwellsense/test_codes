def run(matrix, **kwargs):

    if len(matrix) != 60:
        return False, f"Expected 60 rows, got {len(matrix)}"

    for row in matrix:

        if len(row) != 33:
            return False, f"Expected 33 cols, got {len(row)}"

    return True, "Matrix size OK"