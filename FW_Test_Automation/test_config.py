# Central configuration for FW release tests.
# Keeping thresholds here (instead of hard-coded inside each test)
# makes it easy to tune limits per HW revision without touching test logic.

# ---- Expected identity ----
EXPECTED_FW_VERSION = "02.16"      # set to None to skip the exact-match check
EXPECTED_HW_VERSION = "80.16"      # set to None to skip the exact-match check

# ---- Frame / matrix geometry ----
EXPECTED_NUM_ROWS = 60
EXPECTED_NUM_DATA_COLS = 30        # sensor columns (without the 3 reference columns)
NUM_REF_COLS = 3
EXPECTED_TOTAL_COLS = EXPECTED_NUM_DATA_COLS + NUM_REF_COLS

# ---- Environmental sanity ranges ----
TEMPERATURE_MIN_C = 0.0
TEMPERATURE_MAX_C = 60.0

HUMIDITY_MIN_PCT = 0.0
HUMIDITY_MAX_PCT = 100.0

# ---- Pixel value sanity (raw sensor counts) ----
PIXEL_MIN = 1            # a real reading should never be 0/negative
PIXEL_MAX = 4095         # 12-bit ADC ceiling - adjust to actual FW ADC width
STUCK_ROW_MIN_UNIQUE = 2  # a row with fewer unique values than this is suspect

# ---- Reference capacitor expected ranges (avg over all rows) ----
REF1_RANGE = (600, 750)
REF2_RANGE = (1200, 1450)
REF3_RANGE = (1800, 2100)

# ---- Repeatability (frame-to-frame noise) ----
REPEATABILITY_TOLERANCE = 5       # max allowed |delta| per pixel between 2 consecutive frames
REPEATABILITY_MAX_BAD_PIXELS = 0  # how many pixels are allowed to exceed the tolerance

# ---- Golden sample comparison ----
GOLDEN_TOLERANCE = 5

# ---- Known / acceptable FW error codes on boot (leave empty set to fail on ANY error) ----
# Example: {"0"} would treat error code "0" (no error) as acceptable.
ACCEPTABLE_ERROR_CODES = {"0"}

# ---- Communication timing ----
MAX_FRAME_WAIT_SEC = 5.0   # used as a read timeout inside fw_interface

# ---- Metadata schema: every key the FW protocol is expected to send.
# A release that silently drops/renames one of these fields is a
# protocol regression, even if every value it DOES send looks fine. ----
REQUIRED_METADATA_KEYS = [
    "MatConnected", "Frame", "MatNum",
    "NumOfRows", "NumOfColumns",
    "ActiveRows", "ActiveColumns",
    "RefCaps", "Temperature", "RelativeHumidity",
    "PCBALifetime", "MatLifetime", "MatActiveTime",
    "HWVer", "FWVer",
]

# ---- Expected reference-capacitor configuration (calibration setting,
# not the measured values) e.g. "RefCaps,5,10,15" ----
EXPECTED_REF_CAPS = "5,10,15"      # set to None to skip this check

# ---- Lifetime / usage counters: just sanity that they are valid,
# non-negative integers (protects against garbage/overflow) ----
COUNTER_FIELDS = ["PCBALifetime", "MatLifetime", "MatActiveTime", "MatNum"]

# ---- Frame-to-frame progress sanity ----
MAX_TEMPERATURE_JUMP_C = 5.0     # between 2 back-to-back frames
MAX_HUMIDITY_JUMP_PCT = 10.0     # between 2 back-to-back frames

# ---- Release artifact check: compiled .hex file contiguity/checksum.
# Set to a real path to enable; leave None to skip (e.g. when running
# only live on-device tests without the build output at hand). ----
HEX_FILE_PATH = None

