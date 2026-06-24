import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from collections import defaultdict

# =========================
# CONFIG
# =========================

CALIB_A_PATH = Path(r"C:\Users\dvirs\Downloads\0-70\files")
CALIB_B_PATH = Path(r"C:\Users\dvirs\Downloads\0-110\files")

MIN_VALID_VALUE = 1.0
Z_THRESHOLD = 3
POLY_DEGREE = 3

# =========================
# LOAD
# =========================

# def load_matrix(path):
#     return pd.read_csv(path, header=None).values

def load_matrix(path):
    with open(path, 'r') as f:
        lines = f.readlines()

    data = []

    # קח רק את השורות שמתחילות ב-Row
    for l in lines:
        l = l.strip()
        if not l or "---" in l or "Pressures" in l:
            break  # סוף המטריצה הראשונה
        if not l.startswith("Row"):
            continue  # דלג על שורות כותרת

        parts = l.split(',')  # הפרדת הערכים בטאב
        if len(parts) < 2:
            continue

        try:
            # המרת כל הערכים למעט העמודה הראשונה ל-float
            row = [float(x.replace('E','e')) for x in parts[1:-2]]
            data.append(row)
        except ValueError:
            continue  # דלג על שורות עם טקסט לא מספרי

    if not data:
        raise ValueError(f"No numeric data found in file {path}")

    return np.array(data)


def load_calibration(folder):
    data = defaultdict(dict)

    for file in folder.glob("*.csv"):
        name = file.stem  # 10u
        pressure = int(name[:-1])
        direction = name[-1]

        data[pressure][direction] = load_matrix(file)

    return data


# =========================
# MASK
# =========================

def build_global_mask(calibA, calibB):
    all_mats = []

    for calib in [calibA, calibB]:
        for p in calib:
            for d in calib[p]:
                all_mats.append(calib[p][d])

    stack = np.stack(all_mats)

    mean_map = np.mean(stack, axis=0)
    std_map = np.std(stack, axis=0)

    # תנאי 1
    valid_min = mean_map > MIN_VALID_VALUE

    # תנאי 2
    z = np.abs((stack - mean_map) / (std_map + 1e-6))
    valid_outlier = ~np.any(z > Z_THRESHOLD, axis=0)

    # תנאי 3 - שורות/עמודות
    row_valid = np.all(valid_min & valid_outlier, axis=1)
    col_valid = np.all(valid_min & valid_outlier, axis=0)

    row_mask = np.repeat(row_valid[:, None], mean_map.shape[1], axis=1)
    col_mask = np.repeat(col_valid[None, :], mean_map.shape[0], axis=0)

    return valid_min & valid_outlier & row_mask & col_mask


def apply_mask(mat, mask):
    return mat[mask]


# =========================
# BASIC METRICS
# =========================

def calc_hysteresis(data, mask):
    vals = []

    for p in data:
        if 'u' in data[p] and 'd' in data[p]:
            u = apply_mask(data[p]['u'], mask)
            d = apply_mask(data[p]['d'], mask)

            vals.append(np.mean(np.abs(u - d)))

    return np.mean(vals)


def calc_noise(data, mask):
    vals = []

    for p in data:
        for d in data[p]:
            vals.append(np.std(apply_mask(data[p][d], mask)))

    return np.mean(vals)


def calc_sensitivity(data, mask):
    pressures = sorted(data.keys())
    means = []

    for p in pressures:
        mats = list(data[p].values())
        avg = np.mean([apply_mask(m, mask) for m in mats])
        means.append(avg)

    slope = np.gradient(means, pressures)
    return np.mean(np.abs(slope))


# =========================
# POLYNOMIAL FIT
# =========================

def fit_polynomials(calib, mask):
    poly_map = {}
    pressures = sorted(calib.keys())

    for i in range(mask.shape[0]):
        for j in range(mask.shape[1]):
            if not mask[i, j]:
                continue

            values = []
            p_list = []

            for p in pressures:
                mats = list(calib[p].values())
                val = np.mean([m[i, j] for m in mats])

                values.append(val)
                p_list.append(p)

            coeffs = np.polyfit(p_list, values, POLY_DEGREE)
            poly_map[(i, j)] = coeffs

    return poly_map


# =========================
# POLY COMPARISON
# =========================

def poly_rmse(polyA, polyB):
    pressures = np.linspace(0, 110, 50)
    errors = []

    for key in polyA:
        if key not in polyB:
            continue

        y1 = np.polyval(polyA[key], pressures)
        y2 = np.polyval(polyB[key], pressures)

        rmse = np.sqrt(np.mean((y1 - y2) ** 2))
        errors.append(rmse)

    return np.mean(errors), np.array(errors)


def compare_coefficients(polyA, polyB):
    diffs = []

    for key in polyA:
        if key in polyB:
            diffs.append(polyA[key] - polyB[key])

    diffs = np.array(diffs)

    return np.mean(diffs, axis=0), np.std(diffs, axis=0)


def build_error_map(polyA, polyB, shape):
    error_map = np.zeros(shape)
    pressures = np.linspace(0, 110, 50)

    for (i, j), c1 in polyA.items():
        if (i, j) not in polyB:
            continue

        c2 = polyB[(i, j)]

        y1 = np.polyval(c1, pressures)
        y2 = np.polyval(c2, pressures)

        error_map[i, j] = np.sqrt(np.mean((y1 - y2) ** 2))

    return error_map


# =========================
# VISUALIZATION
# =========================

def plot_mask(mask):
    plt.figure()
    plt.imshow(mask)
    plt.title("Valid Pixels Mask")
    plt.colorbar()
    plt.show()


def plot_error_map(error_map):
    plt.figure()
    plt.imshow(error_map)
    plt.title("Polynomial RMSE Map")
    plt.colorbar()
    plt.show()


def plot_coeff_hist(polyA, polyB):
    A = np.array(list(polyA.values()))
    B = np.array(list(polyB.values()))

    labels = ['a', 'b', 'c', 'd']

    for i in range(4):
        plt.figure()
        plt.hist(A[:, i], bins=50, alpha=0.5, label='A')
        plt.hist(B[:, i], bins=50, alpha=0.5, label='B')
        plt.title(f"Coeff {labels[i]}")
        plt.legend()
        plt.show()


# =========================
# DECISION
# =========================

def decide(metrics):
    score_A = 0
    score_B = 0

    if metrics['A_hyst'] < metrics['B_hyst']:
        score_A += 1
    else:
        score_B += 1

    if metrics['A_noise'] < metrics['B_noise']:
        score_A += 1
    else:
        score_B += 1

    if metrics['poly_rmse'] < metrics['rmse_threshold']:
        score_A += 1
    else:
        score_B += 1

    return "Calibration A עדיף" if score_A > score_B else "Calibration B עדיף"


# =========================
# MAIN
# =========================

def main():
    print("Loading data...")
    calibA = load_calibration(CALIB_A_PATH)
    calibB = load_calibration(CALIB_B_PATH)

    print("Building mask...")
    mask = build_global_mask(calibA, calibB)

    print(f"Valid pixels: {np.sum(mask)} / {mask.size}")
    plot_mask(mask)

    print("Basic metrics...")
    A_hyst = calc_hysteresis(calibA, mask)
    B_hyst = calc_hysteresis(calibB, mask)

    A_noise = calc_noise(calibA, mask)
    B_noise = calc_noise(calibB, mask)

    print("Fitting polynomials...")
    polyA = fit_polynomials(calibA, mask)
    polyB = fit_polynomials(calibB, mask)

    print("Comparing polynomials...")
    rmse_mean, rmse_all = poly_rmse(polyA, polyB)

    coeff_mean, coeff_std = compare_coefficients(polyA, polyB)

    error_map = build_error_map(polyA, polyB, mask.shape)

    plot_error_map(error_map)
    plot_coeff_hist(polyA, polyB)

    print("\n===== REPORT =====")
    print(f"Hysteresis A: {A_hyst:.4f}")
    print(f"Hysteresis B: {B_hyst:.4f}")
    print(f"Noise A: {A_noise:.4f}")
    print(f"Noise B: {B_noise:.4f}")
    print(f"Polynomial RMSE: {rmse_mean:.4f}")

    print("\nCoeff diff mean:", coeff_mean)
    print("Coeff diff std:", coeff_std)

    metrics = {
        'A_hyst': A_hyst,
        'B_hyst': B_hyst,
        'A_noise': A_noise,
        'B_noise': B_noise,
        'poly_rmse': rmse_mean,
        'rmse_threshold': 0.05  # אפשר לכוון
    }

    print("\n===== CONCLUSION =====")
    print(decide(metrics))


if __name__ == "__main__":
    main()