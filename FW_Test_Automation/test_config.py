# Central configuration for FW release tests.
# Keeping thresholds here (instead of hard-coded inside each test)
# makes it easy to tune limits per HW revision without touching test logic.

# ---- Expected identity ----
EXPECTED_FW_VERSION = "02.17"      # set to None to skip the exact-match check
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
# Deliberately wide for now: there are two board variants, each measuring
# a different drop, so their reference readings differ from each other.
# Tighten these back down per-variant once both are characterized.
REF1_RANGE = (400, 1000)
REF2_RANGE = (900, 1800)
REF3_RANGE = (1500, 2700)

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

# ---- Serial port ----
COM_PORT = "COM4"

# ---- Mat-connect gate (main.py checks this before running the
# mat-dependent test suite; a board can be tested with no mat, but
# main.py should not silently run mat tests against a disconnected
# board and report confusing failures) ----
MAT_CONNECT_WAIT_SEC = 15.0   # how long to poll for MatConnected=true before aborting

# Once connected, the first frame(s) can still reflect a not-yet-settled
# ADC average (observed on real hardware as an all-zero matrix and
# Ref1=0.0) - wait this long and re-fetch a fresh frame before trusting
# pixel data.
MAT_SETTLE_SEC = 2.0

# ---- EEPROM metadata round-trip test values (test_eeprom_metadata_update).
# These OVERWRITE the real board's stored values for the duration of the
# test; the original values are snapshotted first and restored afterward. ----
EEPROM_TEST_VALUES = {
    "MatNum": "TEST-0001",
    "ActiveRows": "01,60",
    "ActiveColumns": "01,30",
    "MatLifeTime": 12345,
    "MatActiveTime": 6789,
}
PCBA_LIFETIME_TEST_VALUE = 111222

# ---- LUT round-trip (test_lut_roundtrip). CSV file exported by the LT_GUI
# tool: a header row, then "RowN,a1,b1,c1,d1,a2,..." per-pixel calibration
# coefficient rows, then a "DEFAULT_LUT" marker line, a header line, and one
# line of 4 default coefficients. Set LUT_FILE_PATH to enable; leave None to
# skip. All float coefficients are scaled by LUT_SCALE_FACTOR and rounded to
# int32 before being sent, matching the reference LT_GUI upload tool. ----
LUT_FILE_PATH = r"C:\Users\dvirs\Documents\LT_GUI_Results\LUT_files\LUT_default_plus_10_1_60_1_30_EB.csv"
LUT_SCALE_FACTOR = 10000

# ---- FW update (fw_update_test.py - separate opt-in script, not part of
# the default main.py run since it reflashes and reboots the board).
# List of (hex_path, expected_FWVer_after_flash). Version strings below are
# inferred from the filenames - confirm/adjust against the actual FWVer
# each image reports. ----
FW_UPDATE_IMAGES = [
    (r"C:\Users\dvirs\Documents\Atmel Studio\7.0\NT_Bootloader_project\NT_PCBA_80\NT_usb\Debug\Version_02_17.hex", "02.17"),
    (r"C:\Users\dvirs\Documents\Atmel Studio\7.0\NT_Bootloader_project\NT_PCBA_80\NT_usb\Debug\Version_02_09.hex", "02.09"),
]
FW_UPDATE_TIMEOUT_SEC = 60.0
FW_UPDATE_REBOOT_TIMEOUT_SEC = 30.0

# ---- WDT / Timer test (wdt_timer_test.py - separate opt-in script, not
# part of the default main.py run since it blocks for ~10 minutes). ----
WDT_EXPECTED_TIMEOUT_SEC = 600.0
WDT_TIMEOUT_TOLERANCE_SEC = 30.0

