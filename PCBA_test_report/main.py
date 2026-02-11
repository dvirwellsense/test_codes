import os
import re
import glob
import unicodedata
import numpy as np
import pandas as pd
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import matplotlib.pyplot as plt
from scipy.stats import norm
from openpyxl.styles import PatternFill, Font

# ================= Settings =================
folder = r"C:\Users\dvirs\Desktop\PCB\LT\PCBA tests"  # <-- Set your folder path
EXPECTED_FRAMES = 5

# RowMap: 60 rows mapped to caps
cap1, cap2, cap3, cap4, cap5, cap6 = 1, 5, 10, 15, 20, 30
RowMap = np.zeros(60, dtype=int)
for i in range(20):
    RowMap[i] = cap1 if i % 2 == 0 else cap6
for i in range(20, 40):
    RowMap[i] = cap2 if i % 2 == 0 else cap5
for i in range(40, 60):
    RowMap[i] = cap3 if i % 2 == 0 else cap4
CAP_LIST = [cap1, cap2, cap3, cap4, cap5, cap6]

# Colors
COLOR_PASS = "D1F2D1"
COLOR_WARN = "FFF4CC"
COLOR_FAIL = "F8D7DA"

# ================= Helpers =================
def strip_combining(s: str) -> str:
    nfkd = unicodedata.normalize('NFKD', s)
    return ''.join(ch for ch in nfkd if not unicodedata.combining(ch))

def parse_filename_key(path: str):
    name = os.path.basename(path)
    name_norm = strip_combining(name)
    m = re.match(r'^LT[_ ]?PCBA_(\d{9})_(\d{8})_(\d{6})\.xlsx$', name_norm, flags=re.IGNORECASE)
    if m:
        serial = int(m.group(1))
        ymd = int(m.group(2))
        hms = int(m.group(3))
        return (0, serial, ymd, hms, name_norm.lower())
    else:
        return (1, 0, 0, 0, name_norm.lower())

# ================= Load Excel Frames =================
def load_frame_cube(file_path: str, expected_frames: int = EXPECTED_FRAMES) -> np.ndarray:
    # print(f"Opening file: {file_path}")
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    frame_sheets = [s for s in xls.sheet_names if s.lower().startswith("frame_")]
    if not frame_sheets:
        raise ValueError(f"No Frame_* sheets found in {os.path.basename(file_path)}.")
    try:
        frame_sheets = sorted(frame_sheets, key=lambda s: int(s.split("_")[1]))
    except Exception:
        frame_sheets = sorted(frame_sheets)
    frames = []
    for s in frame_sheets[:expected_frames]:
        df = pd.read_excel(file_path, sheet_name=s, header=None, engine="openpyxl")
        df_num = df.apply(pd.to_numeric, errors="coerce").dropna(axis=0, how="all").dropna(axis=1, how="all")
        if df_num.shape[0] < 60 or df_num.shape[1] < 30:
            raise ValueError(f"{os.path.basename(file_path)} sheet {s}: shape {df_num.shape}, expected at least (60,30).")
        frame = df_num.iloc[:60, :30].to_numpy()
        frames.append(frame)
    data = np.stack(frames, axis=2)
    if data.shape[2] != expected_frames:
        print(f"Warning: {os.path.basename(file_path)} has {data.shape[2]} frames, expected {expected_frames}.")
    return data

# ================= Analyze File =================
def analyze_file(file_path: str) -> pd.DataFrame:
    data = load_frame_cube(file_path, expected_frames=EXPECTED_FRAMES)
    rows = []
    for cap in CAP_LIST:
        mask_rows = (RowMap == cap)
        frames_masked = data[mask_rows, :, :]
        mean_pixels = np.mean(frames_masked, axis=2)
        std_pixels = np.std(frames_masked, axis=2)
        Std_Frames = float(np.mean(std_pixels))
        Error = float(np.mean(mean_pixels - cap) / cap * 100)
        lower, upper = cap * 0.9, cap * 1.1
        total_pixels = mean_pixels.size
        outlier_count = int(np.sum((mean_pixels < lower) | (mean_pixels > upper)))
        Outliers = outlier_count / total_pixels * 100
        Score = 100 - 2*abs(Error) - 6*Std_Frames - 1*Outliers
        if abs(Error) > 15 or Std_Frames > 1.5 or Outliers > 10:
            Status = "FAIL"
        elif Score < 75:
            Status = "WARN"
        else:
            Status = "PASS"
        rows.append({
            "Cap": cap,
            "Error %": round(Error,2),
            "Std_Frames": round(Std_Frames,2),
            "Outliers %": round(Outliers,2),
            "Score": round(Score,2),
            "Status": Status
        })
    df_result = pd.DataFrame(rows)
    df_result["File"] = os.path.basename(file_path)
    return df_result, data  # return data for later per-cap histograms

# ================= Collect Files =================
patterns = ["*.xlsx"]
all_files = []
for pat in patterns:
    all_files.extend(glob.glob(os.path.join(folder, pat)))
all_files = [f for f in all_files if not os.path.basename(f).startswith("~$")]
all_files = sorted(all_files, key=parse_filename_key)
if not all_files:
    print("No Excel files found.")
    raise SystemExit(1)

# ================= Run Analysis =================
final_report = []
errors = []
all_data = {}  # store all loaded data per file
total = len(all_files)
print(f"Found {total} files.\n")

for i, f in enumerate(all_files, start=1):
    fname = os.path.basename(f)
    print(f"[{i}/{total}] Opening: {fname} ...")
    try:
        df, data = analyze_file(f)
        final_report.append(df)
        all_data[fname] = data
        print(f"[{i}/{total}] ✔ OK: {fname}, added {len(df)} rows.\n")
    except Exception as e:
        errors.append((fname, str(e)))
        print(f"[{i}/{total}] ✖ ERROR in {fname}: {e}\n")

report_df = pd.concat(final_report, ignore_index=True) if final_report else pd.DataFrame()

# ================= Boards Summary =================
summary_rows = []
for board, grp in report_df.groupby("File"):
    overall_score = float(grp["Score"].mean())
    has_fail = any(grp["Status"]=="FAIL")
    has_warn = any(grp["Status"]=="WARN")
    overall_status = "FAIL" if has_fail else ("WARN" if has_warn else "PASS")
    summary_rows.append({
        "File": board,
        "Overall Score": round(overall_score,2),
        "Overall Status": overall_status,
        "PASS Count": int((grp["Status"]=="PASS").sum()),
        "WARN Count": int((grp["Status"]=="WARN").sum()),
        "FAIL Count": int((grp["Status"]=="FAIL").sum())
    })
boards_summary = pd.DataFrame(summary_rows)
excel_path = os.path.join(folder, "Final_Report.xlsx")
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    report_df.to_excel(writer, index=False, sheet_name="Per_Cap")
    boards_summary.to_excel(writer, index=False, sheet_name="Boards_Summary")
print(f"✅ Excel saved: {excel_path}")

# Function to update the title size on each slide
def set_title_size(slide, size_pt=24):
    if slide.shapes.title:
        title_tf = slide.shapes.title.text_frame
        for paragraph in title_tf.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size_pt)
                run.font.bold = True  # אפשר להשאיר כ-bold
                run.font.color.rgb = RGBColor(0,0,0)

# ================= PowerPoint =================
prs = Presentation()

# Title slide
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "PCBA Measurement Quality Report"
slide.placeholders[1].text = f"FW / HW Version: TBD\nDate: {datetime.now().strftime('%Y-%m-%d %H:%M')}"

# ================= Parameters Slide =================
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "PCBA Measurement Parameters"
text = (
    "🔹 Score – Overall measurement quality rating\n"
    "  Formula: 100 - 2×|Error %| - 6×Std Frames - 1×Outliers %\n"
    "  High score = good measurement quality\n\n"
    "🔹 Error % – Accuracy of the measurement\n"
    "  % difference between measured and expected capacitance\n\n"
    "🔹 Std Frames – Stability across frames\n"
    "  Standard deviation of pixel values per capacitance\n\n"
    "🔹 Outliers % – Fraction of extreme values\n"
    "  % of pixels outside ±10% of expected capacitance"
)
txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(4))
tf = txBox.text_frame
tf.clear()
p = tf.add_paragraph()
p.text = text
p.font.size = Pt(16)
set_title_size(slide, 24)

# ================= Per-Board Slides =================
boards = boards_summary["File"].tolist()
offset_down = Inches(0.6)
for board in boards:
    df_board = report_df[report_df["File"]==board].copy()
    df_cap = df_board.groupby('Cap', as_index=False).agg({
        'Std_Frames':'mean',
        'Outliers %':'mean',
        'Score':'mean',
        'Error %':'mean',
        'Status': lambda s: 'FAIL' if (s=='FAIL').any() else ('WARN' if (s=='WARN').any() else 'PASS')
    })
    has_fail = any(df_cap["Status"]=="FAIL")
    has_warn = any(df_cap["Status"]=="WARN")
    overall_status = "FAIL" if has_fail else ("WARN" if has_warn else "PASS")
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f"Board: {board}"
    set_title_size(slide, 24)

    # Table
    rows_table = len(CAP_LIST)+1
    cols_table = 6
    table = slide.shapes.add_table(rows_table, cols_table, Inches(0.5), Inches(1) + offset_down, Inches(5), Inches(2)).table
    headers = ['Cap','Score','Error %','Std Frames','Outliers %','Status']
    for c, h in enumerate(headers):
        table.cell(0,c).text = h

    std_vals, out_vals = [], []
    for r, cap in enumerate(CAP_LIST):
        row = df_cap[df_cap['Cap']==cap]
        table.cell(r+1,0).text = str(cap)
        if row.empty:
            table.cell(r+1,1).text = ""
            table.cell(r+1,2).text = ""
            table.cell(r+1,3).text = ""
            table.cell(r+1,4).text = ""
            table.cell(r+1,5).text = ""
            std_vals.append(0.0)
            out_vals.append(0.0)
        else:
            row_data = row.iloc[0]
            table.cell(r+1,1).text = str(round(float(row_data['Score']),2))
            table.cell(r+1,2).text = str(round(float(row_data['Error %']),2))
            table.cell(r+1,3).text = str(round(float(row_data['Std_Frames']),2))
            table.cell(r+1,4).text = str(round(float(row_data['Outliers %']),2))
            table.cell(r+1,5).text = row_data['Status']
            # Color Status cell
            cell = table.cell(r+1,5)
            cell.fill.solid()
            if row_data['Status']=="PASS":
                cell.fill.fore_color.rgb = RGBColor(0xD1,0xF2,0xD1)
            elif row_data['Status']=="WARN":
                cell.fill.fore_color.rgb = RGBColor(0xFF,0xF4,0xCC)
            else:
                cell.fill.fore_color.rgb = RGBColor(0xF8,0xD7,0xDA)
            std_vals.append(float(row_data['Std_Frames']))
            out_vals.append(float(row_data['Outliers %']))

    # Overall status box
    left, top, width, height = Inches(5.7), Inches(0.9) + offset_down, Inches(4), Inches(0.5)
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = f"Overall Status: {overall_status}"
    run.font.bold = True
    run.font.size = Pt(18)
    p.alignment = PP_ALIGN.CENTER
    fill = box.fill
    fill.solid()
    if overall_status=="PASS":
        fill.fore_color.rgb = RGBColor(0xD1,0xF2,0xD1)
    elif overall_status=="WARN":
        fill.fore_color.rgb = RGBColor(0xFF,0xF4,0xCC)
    else:
        fill.fore_color.rgb = RGBColor(0xF8,0xD7,0xDA)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0,0,0)

    # Charts
    std_path = os.path.join(folder, f"{board}_Std_Frames.png")
    out_path = os.path.join(folder, f"{board}_Outliers.png")
    fig, ax = plt.subplots(figsize=(6,3))
    ax.bar(CAP_LIST, std_vals, color='skyblue')
    ax.set_title("Std_Frames vs Capacitance")
    ax.set_xlabel("Capacitance (pF)")
    ax.set_ylabel("Std_Frames (pF)")
    plt.tight_layout()
    plt.savefig(std_path)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(6,3))
    ax.bar(CAP_LIST, out_vals, color='salmon')
    ax.set_title("Outliers % vs Capacitance")
    ax.set_xlabel("Capacitance (pF)")
    ax.set_ylabel("Outliers %")
    plt.tight_layout()
    plt.savefig(out_path)
    plt.close(fig)

    slide.shapes.add_picture(std_path, Inches(5.7), Inches(1.6) + offset_down, width=Inches(4), height=Inches(2))
    slide.shapes.add_picture(out_path, Inches(5.7), Inches(3.8) + offset_down, width=Inches(4), height=Inches(2))
    for tmp in (std_path, out_path):
        try: os.remove(tmp)
        except: pass

# ================= Per-Cap Slides =================
for cap in CAP_LIST:
    all_pixels = []
    avg_per_board = []

    for board, data in all_data.items():
        mask_rows = (RowMap == cap)
        frames_masked = data[mask_rows, :, :]

        all_pixels.extend(frames_masked.flatten())

        mean_board = float(np.mean(frames_masked))  # ממוצע כולל לכל הלוח
        avg_per_board.append(mean_board)

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f"Cap {cap}pF: All Pixels + Board Means"
    set_title_size(slide, 24)

    # =========================
    # Histogram – all pixels
    # =========================
    fig, ax = plt.subplots(figsize=(6, 3))
    ax.hist(all_pixels, bins=50, edgecolor='black')
    ax.set_title(f"All measurements for {cap}pF")
    ax.set_xlabel("Pixel value")
    ax.set_ylabel("Frequency")
    plt.tight_layout()

    path_all = os.path.join(folder, f"{cap}_all_pixels.png")
    plt.savefig(path_all)
    plt.close(fig)

    slide.shapes.add_picture(
        path_all,
        Inches(0.5),
        Inches(1),
        width=Inches(6),
        height=Inches(2.5)
    )

    try:
        os.remove(path_all)
    except:
        pass

    # =========================
    # Histogram – average per board
    # =========================
    fig, ax = plt.subplots(figsize=(6, 3))
    ax.hist(avg_per_board, bins=50, edgecolor='black')
    ax.set_title(f"Average per board for {cap}pF")
    ax.set_xlabel("Mean Pixel Value per Board")
    ax.set_ylabel("Frequency")
    plt.tight_layout()

    path_avg = os.path.join(folder, f"{cap}_avg_per_board.png")
    plt.savefig(path_avg)
    plt.close(fig)

    slide.shapes.add_picture(
        path_avg,
        Inches(0.5),
        Inches(3.6),
        width=Inches(6),
        height=Inches(2.5)
    )

    try:
        os.remove(path_avg)
    except:
        pass

    # =========================
    # Gaussian distribution + σ lines
    # =========================
    mu = np.mean(all_pixels)
    sigma = np.std(all_pixels)

    fig, ax = plt.subplots(figsize=(6, 3))

    # Normalized histogram
    ax.hist(
        all_pixels,
        bins=50,
        density=True,
        alpha=0.6,
        edgecolor='black'
    )

    # Gaussian curve
    x = np.linspace(min(all_pixels), max(all_pixels), 500)
    pdf = norm.pdf(x, mu, sigma)
    ax.plot(x, pdf, linewidth=2)

    # Std deviation markers
    for k in [1, 2, 3]:
        ax.axvline(mu + k * sigma, linestyle='--', linewidth=1)
        ax.axvline(mu - k * sigma, linestyle='--', linewidth=1)

    ax.set_title(
        f"Gaussian Distribution for {cap}pF\n"
        f"μ={mu:.2f}, σ={sigma:.2f}"
    )
    ax.set_xlabel("Pixel value")
    ax.set_ylabel("Probability Density")

    plt.tight_layout()

    path_gauss = os.path.join(folder, f"{cap}_gaussian.png")
    plt.savefig(path_gauss)
    plt.close(fig)

    slide.shapes.add_picture(
        path_gauss,
        Inches(0.5),
        Inches(6.2),
        width=Inches(6),
        height=Inches(2.5)
    )

    try:
        os.remove(path_gauss)
    except:
        pass

# ================= Summary Pie Chart Slide =================
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "Summary: Overall Status Distribution"
set_title_size(slide, 24)

status_counts = boards_summary["Overall Status"].value_counts()
labels = status_counts.index.tolist()
sizes = status_counts.values.tolist()
colors_hex = {"PASS":"#D1F2D1","WARN":"#FFF4CC","FAIL":"#F8D7DA"}
colors = [colors_hex.get(l,"#CCCCCC") for l in labels]

fig, ax = plt.subplots(figsize=(6,6))
ax.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
ax.set_title("Overall Board Status Distribution")
plt.tight_layout()
pie_path = os.path.join(folder,"Summary_Pie.png")
plt.savefig(pie_path)
plt.close(fig)
slide.shapes.add_picture(pie_path, Inches(2), Inches(1.5), width=Inches(6), height=Inches(6))
try: os.remove(pie_path)
except: pass

# ================= Save PPT =================
ppt_path = os.path.join(folder,"PCBA_Report_Full.pptx")
prs.save(ppt_path)
print(f"🖼️ PowerPoint saved: {ppt_path}")
print("Done.")
