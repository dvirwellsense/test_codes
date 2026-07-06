import time

import serial

from test_config import MAX_FRAME_WAIT_SEC


class FrameTimeoutError(Exception):
    """Raised when a full frame (ending in the last row) isn't
    received within MAX_FRAME_WAIT_SEC seconds."""
    pass


class FWInterface:

    def __init__(self, port="COM8", baudrate=115200):

        self.port = port
        self.baudrate = baudrate

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

    def send_command(self, cmd, timeout=2.0, terminator=None):
        """Send a text command (e.g. "MatNum,123") and collect the
        response lines the FW sends back.

        Many commands (plain field writes) send NO response at all on
        success - only on error. So this does not raise on an empty
        response; it simply reads for up to `timeout` seconds of
        inactivity, or until a line equal to `terminator` is seen
        (e.g. "Done reading LUT values." for ReadMatLUT), and returns
        whatever lines arrived (possibly none).
        """

        self.ser.reset_input_buffer()

        self.ser.write((cmd + "\r\n").encode())

        return self._read_lines_until_idle(timeout, terminator)

    def _read_lines_until_idle(self, timeout, terminator=None):
        """Read lines for a FIXED window of `timeout` seconds (or until
        `terminator` is seen).

        The FW continuously streams periodic ADC frame lines
        regardless of any command being processed - it never actually
        goes idle. An earlier version of this method reset its
        deadline on every incoming line, which meant it would loop
        forever (deadline kept getting pushed out by frame noise) and
        never return for any command that has no `terminator`, e.g.
        plain Get*/write commands. Do not "extend on activity" here.
        """

        lines = []

        deadline = time.time() + timeout

        while time.time() < deadline:

            line = self.ser.readline().decode(
                errors="ignore"
            ).strip()

            if not line:
                continue

            lines.append(line)

            if terminator is not None and line == terminator:
                break

        return lines

    def _drain_input(self):
        """Discard any bytes currently waiting in the input buffer,
        without blocking.

        Must be called periodically during any long write-heavy
        operation (streaming LUT rows or a HEX file). The FW keeps
        emitting unsolicited output (periodic frames, diagnostics)
        independent of whatever command is in flight. If that output
        is never read, it backs up in the FW's own USB CDC transmit
        buffer; once that's full, the FW's `usb_printf()` calls block,
        which (being on the same single-threaded main loop) also stops
        it from reading any more incoming bytes - a two-way deadlock
        that looks identical to a Python-side hang, but isn't one.
        """

        try:
            waiting = self.ser.in_waiting
            if waiting:
                self.ser.read(waiting)
        except Exception:
            pass

    def send_lut_matrix(self, rows, timeout=15.0, start_delay=0.1, row_delay=0.02, verbose=True):
        """Send the per-pixel LUT matrix via WriteMatLUT. `rows` is a
        list of (label, values) pairs, e.g. ("Row1", [0, 0, 0, 0, 60000, ...]),
        already scaled to the integer wire format (the FW only parses
        int32 - the caller is responsible for any fixed-point scaling).

        Mirrors the reference LUT-upload tool's exact framing: the
        WriteMatLUT/END commands are sent with a bare "\\n", each row
        line with "\\r\\n". The FW's line tokenizer accepts either, but
        this matches the known-working implementation exactly.

        Confirmed (by sending the exact same file through the
        reference GUI tool to the exact same board, which succeeded)
        that this isn't an FW/EEPROM capacity limitation - the actual
        bug was this method never reading the input buffer while it
        writes. See _drain_input()'s docstring: an unread FW transmit
        buffer can stall the FW's main loop from within the write call
        itself, which then also stops it from reading - a deadlock
        that looks identical to a Python-side hang. The GUI tool likely
        avoids this because the surrounding app has a background
        reader thread always draining input, independent of the send
        method itself. start_delay/row_delay mirror the reference
        tool's `await Task.Yield()` pacing (kept as a harmless belt-
        and-suspenders measure, though the drain is the actual fix).
        """

        self.ser.reset_input_buffer()

        if verbose:
            print(f"  [LUT] sending WriteMatLUT command ({len(rows)} rows)...")

        self.ser.write(b"WriteMatLUT\n")

        time.sleep(start_delay)

        self._drain_input()

        for i, (label, values) in enumerate(rows, start=1):

            line = label + "," + ",".join(str(v) for v in values)

            if verbose:
                print(f"  [LUT] sending {label} ({i}/{len(rows)}, {len(values)} values)...")

            self.ser.write((line + "\r\n").encode())

            time.sleep(row_delay)

            self._drain_input()

        if verbose:
            print("  [LUT] all rows sent, sending END and waiting for FW confirmation...")

        self.ser.write(b"END\n")

        response = self._read_lines_until_idle(timeout)

        if verbose:
            print(f"  [LUT] WriteMatLUT response: {response}")

        return response

    def flash_firmware(self, hex_path, timeout=60.0):
        """Send the `flash` command and stream an Intel-HEX file's
        lines to the FW's serial bootloader. Returns the response
        lines - expect a line containing "Done!" on success, "Fail!"
        otherwise. The board resets itself right after replying, so
        the caller must follow this with reconnect().
        """

        self.ser.reset_input_buffer()

        self.ser.write(b"flash\r\n")

        with open(hex_path, "r") as f:

            for line in f:

                line = line.strip()

                if not line:
                    continue

                self.ser.write((line + "\r\n").encode())

                self._drain_input()

        return self._read_lines_until_idle(timeout)

    def reconnect(self, timeout=30.0, poll_interval=1.0):
        """Close and reopen the serial port, polling until the device
        re-enumerates (used after a `flash` update or a WDT reset).
        """

        try:
            if self.ser.is_open:
                self.ser.close()
        except Exception:
            pass

        deadline = time.time() + timeout

        last_error = None

        while time.time() < deadline:

            try:
                self.ser = serial.Serial(
                    port=self.port,
                    baudrate=self.baudrate,
                    timeout=1
                )
                return

            except serial.SerialException as e:
                last_error = e
                time.sleep(poll_interval)

        raise FrameTimeoutError(
            f"Port '{self.port}' did not come back within "
            f"{timeout:.1f}s ({last_error})"
        )
