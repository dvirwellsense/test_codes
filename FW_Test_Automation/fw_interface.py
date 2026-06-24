import serial

class FWInterface:

    def __init__(self, port="COM8", baudrate=115200):

        self.ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            timeout=1
        )

    def get_frame(self):

        frame = []

        while True:

            line = self.ser.readline().decode(
                errors="ignore"
            ).strip()

            if not line:
                continue

            frame.append(line)

            if line.startswith("Row60"):
                break

        return frame