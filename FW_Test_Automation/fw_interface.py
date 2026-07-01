import time

import serial

from test_config import MAX_FRAME_WAIT_SEC


class FrameTimeoutError(Exception):
    """Raised when a full frame (ending in the last row) isn't
    received within MAX_FRAME_WAIT_SEC seconds."""
    pass


class FWInterface:

    def __init__(self, port="COM8", baudrate=115200):

        self.ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            timeout=1
        )

    def get_frame(self, last_row_prefix="Row60", timeout=MAX_FRAME_WAIT_SEC):
        """Read lines from the serial port until a line starting with
        `last_row_prefix` is seen, or `timeout` seconds elapse.

        The original implementation looped forever if the FW never
        sent "Row60" (e.g. mat disconnected, FW crash, wrong baud
        rate) -- a single missed frame would hang the whole test run.
        This version fails fast with a clear error instead.
        """

        frame = []

        start_time = time.time()

        while True:

            if time.time() - start_time > timeout:
                raise FrameTimeoutError(
                    f"No '{last_row_prefix}' received within "
                    f"{timeout:.1f}s "
                    f"({len(frame)} lines captured before timeout)"
                )

            line = self.ser.readline().decode(
                errors="ignore"
            ).strip()

            if not line:
                continue

            frame.append(line)

            if line.startswith(last_row_prefix):
                break

        return frame
