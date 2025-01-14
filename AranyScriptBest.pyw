import tkinter as tk
from tkinter import messagebox
import threading
import socket
import time
from openpyxl import Workbook
import os
import queue
from datetime import datetime
import logging

# Define the base directory
BASE_DIRECTORY = r"C:\Users\det010\Desktop\pi\version_1"

# Ensure the base directory exists
os.makedirs(BASE_DIRECTORY, exist_ok=True)

# Setup logging to file
log_file_path = os.path.join(BASE_DIRECTORY, "app.log")
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format='[%(asctime)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class SequentialHandler(threading.Thread):
    def __init__(self, excel_lock, app):
        super().__init__()
        self.running = False
        self.excel_lock = excel_lock
        self.app = app
        self.iv_socket = None
        self.daq_socket = None

    def run(self):
        self.running = True

        try:
            # Connect to IV Sensor
            self.iv_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.iv_socket.settimeout(15.0)
            self.iv_socket.connect(("192.168.3.100", 8500))
            logging.info("Connected to IV Sensor.")

            # Send initial format command
            if not self.send_command(self.iv_socket, "OF,01\r", "IVSensor"):
                return

            # Receive response for "OF,01\r"
            response_of = self.receive_response(self.iv_socket, "IVSensor")
            if response_of:
                logging.info(f"Received from IVSensor: '{response_of}'")
            else:
                logging.error("No response received for 'OF,01' command.")
                return

            # Connect to DAQ Device
            self.daq_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.daq_socket.settimeout(15.0)
            self.daq_socket.connect(("192.168.3.150", 25000))
            logging.info("Connected to DAQ Device.")

            # Send start logging command
            if not self.send_command(self.daq_socket, "ST\r", "DAQ"):
                return

            # Receive response for "ST\r"
            response_st = self.receive_response(self.daq_socket, "DAQ")
            if response_st:
                logging.info(f"Received from DAQ: '{response_st}'")
            else:
                logging.error("No response received for 'ST' command.")
                return

            while self.running:
                cycle_start = time.time()

                # 1. Send "T2\r" to IV Sensor
                if not self.send_command(self.iv_socket, "T2\r", "IVSensor"):
                    continue  # Skip to next cycle on failure

                # 2. Receive response for "T2\r"
                response_t2 = self.receive_response(self.iv_socket, "IVSensor")
                if response_t2:
                    self.parse_iv_response(response_t2)
                else:
                    logging.error("Failed to receive response for 'T2' command.")
                    continue  # Skip to next cycle

                # 3. Send "BR,0\r" to IV Sensor to get image data
                if not self.send_command(self.iv_socket, "BR,0\r", "IVSensor"):
                    continue

                # 4. Receive and save image data
                image_data = self.receive_br_response()
                if image_data:
                    self.save_image(image_data)
                else:
                    logging.error("Failed to receive image data.")
                    continue  # Skip to next cycle

                # 5. Send "RM\r" to DAQ Device
                if not self.send_command(self.daq_socket, "RM\r", "DAQ"):
                    continue

                # 6. Receive response for "RM\r"
                response_rm = self.receive_response(self.daq_socket, "DAQ")
                if response_rm:
                    self.parse_daq_response(response_rm)
                else:
                    logging.error("Failed to receive response for 'RM' command.")
                    continue  # Skip to next cycle

                # Calculate remaining time to maintain 5-second cycle
                elapsed = time.time() - cycle_start
                time_to_sleep = max(0, 5.0 - elapsed)
                time.sleep(time_to_sleep)

        except Exception as e:
            logging.error(f"SequentialHandler encountered an error: {e}")
        finally:
            self.cleanup()

    def send_command(self, sock, command, device_name):
        try:
            sock.sendall(command.encode('ascii'))
            logging.info(f"Sent '{command.strip()}' to {device_name}.")
            return True
        except Exception as e:
            logging.error(f"Failed to send '{command.strip()}' to {device_name}: {e}")
            return False

    def receive_response(self, sock, device_name):
        try:
            sock.settimeout(5.0)
            response = sock.recv(4096)
            sock.settimeout(15.0)  # Restore original timeout
            if response:
                response_str = response.decode('ascii', errors='replace').strip()
                logging.info(f"Received from {device_name}: '{response_str}'")
                return response_str
            else:
                logging.warning(f"No response received from {device_name}.")
                return None
        except socket.timeout:
            logging.error(f"Timeout while waiting for response from {device_name}.")
            return None
        except Exception as e:
            logging.error(f"Error receiving response from {device_name}: {e}")
            return None

    def receive_br_response(self):
        try:
            # Read header until third comma to get data length
            header = b''
            commas = 0
            while commas < 3:
                byte = self.iv_socket.recv(1)
                if not byte:
                    break
                header += byte
                if byte == b',':
                    commas += 1

            header_str = header.decode('ascii', errors='replace').strip(',')
            parts = header_str.split(',')
            if len(parts) < 3:
                logging.error("Invalid BR response header.")
                return None

            try:
                data_length = int(parts[2])
                logging.info(f"Expecting {data_length} bytes of image data.")
            except ValueError:
                logging.error("Invalid data length in BR response.")
                return None

            # Receive image data
            image_data = b''
            while len(image_data) < data_length:
                chunk = self.iv_socket.recv(min(4096, data_length - len(image_data)))
                if not chunk:
                    break
                image_data += chunk

            if len(image_data) != data_length:
                logging.error(f"Incomplete image data received: {len(image_data)}/{data_length} bytes.")
                return None

            logging.info(f"Received image data: {len(image_data)} bytes.")
            return image_data

        except socket.timeout:
            logging.error("Timeout while receiving BR response.")
            return None
        except Exception as e:
            logging.error(f"Error receiving BR response: {e}")
            return None

    def parse_iv_response(self, response):
        try:
            parts = response.split(',')
            if len(parts) > 11:
                timestamp = time.time()
                hue = float(parts[9])
                saturation = float(parts[10])
                brightness = float(parts[11])

                # Convert HSV to RGB
                r, g, b = self.hsv_to_rgb(hue, saturation, brightness)

                with self.excel_lock:
                    self.app.iv_data.append([timestamp, hue, saturation, brightness, r, g, b])
                    logging.info(f"Appended IV data: Hue={hue}, Saturation={saturation}, Brightness={brightness}, R={r}, G={g}, B={b}")
            else:
                logging.warning("IV response has insufficient data.")
        except Exception as e:
            logging.error(f"Error parsing IV response: {e}")

    def parse_daq_response(self, response):
        try:
            parts = response.split(',')
            if len(parts) >= 4:
                temperature = parts[3].strip()
                timestamp = time.time()

                with self.excel_lock:
                    self.app.temperature_data.append([timestamp, temperature])
                    logging.info(f"Appended Temperature data: {temperature}")
            else:
                logging.warning("DAQ response has insufficient data.")
        except Exception as e:
            logging.error(f"Error parsing DAQ response: {e}")

    def hsv_to_rgb(self, hue, saturation, brightness):
        try:
            import colorsys
            h = hue / 360.0
            s = saturation / 100.0
            v = brightness / 100.0

            r, g, b = colorsys.hsv_to_rgb(h, s, v)
            return int(r * 255), int(g * 255), int(b * 255)
        except Exception as e:
            logging.error(f"Error converting HSV to RGB: {e}")
            return 0, 0, 0

    def save_image(self, image_data):
        try:
            os.makedirs(BASE_DIRECTORY, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"image_{timestamp}.bmp"
            filepath = os.path.join(BASE_DIRECTORY, filename)

            with open(filepath, 'wb') as f:
                f.write(image_data)

            logging.info(f"Saved image: {filepath}")
        except Exception as e:
            logging.error(f"Error saving image: {e}")

    def cleanup(self):
        # Close IV Sensor socket
        if self.iv_socket:
            try:
                self.iv_socket.shutdown(socket.SHUT_RDWR)
            except:
                pass
            self.iv_socket.close()
            logging.info("Closed IV Sensor connection.")

        # Close DAQ Device socket
        if self.daq_socket:
            try:
                self.daq_socket.shutdown(socket.SHUT_RDWR)
            except:
                pass
            self.daq_socket.close()
            logging.info("Closed DAQ Device connection.")

class TestApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sensor Data Logger")
        self.geometry("500x200")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Initialize buttons
        button_frame = tk.Frame(self)
        button_frame.pack(pady=20)

        self.start_button = tk.Button(button_frame, text="Start", command=self.start_process, width=15)
        self.start_button.grid(row=0, column=0, padx=10)

        self.stop_button = tk.Button(button_frame, text="Stop", command=self.stop_process, width=15, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=10)

        # Status label
        self.status_label = tk.Label(self, text="Status: Stopped", fg="red", font=("Arial", 14))
        self.status_label.pack(pady=10)

        # Data storage
        self.iv_data = []
        self.temperature_data = []

        # Threading
        self.excel_lock = threading.Lock()
        self.handler_thread = None

        # Excel file path
        self.excel_file_path = None

    def start_process(self):
        if not self.handler_thread or not self.handler_thread.is_alive():
            self.handler_thread = SequentialHandler(self.excel_lock, self)
            self.handler_thread.start()
            self.status_label.config(text="Status: Running", fg="green")
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.generate_excel_path()
            logging.info("Process started.")

    def stop_process(self):
        if self.handler_thread and self.handler_thread.is_alive():
            self.handler_thread.running = False
            self.handler_thread.join()
            self.handler_thread = None
            self.status_label.config(text="Status: Stopped", fg="red")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.save_to_excel()
            logging.info("Process stopped.")

    def on_close(self):
        if self.handler_thread and self.handler_thread.is_alive():
            if messagebox.askokcancel("Quit", "Process is running. Do you want to quit?"):
                self.stop_process()
                self.destroy()
        else:
            self.destroy()

    def generate_excel_path(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"sensor_data_{timestamp}.xlsx"
        self.excel_file_path = os.path.join(BASE_DIRECTORY, filename)
        logging.info(f"Excel file will be saved as: {self.excel_file_path}")

    def save_to_excel(self):
        if not self.excel_file_path:
            logging.warning("Excel file path not set. Data not saved.")
            return

        with self.excel_lock:
            try:
                wb = Workbook()
                ws = wb.active
                ws.append(["Timestamp", "Hue", "Saturation", "Brightness", "R", "G", "B", "Temperature"])

                min_length = min(len(self.iv_data), len(self.temperature_data))
                logging.info(f"Saving Excel with {min_length} entries.")
                for i in range(min_length):
                    iv_entry = self.iv_data[i]
                    temp_entry = self.temperature_data[i]
                    row = [
                        time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(iv_entry[0])),
                        iv_entry[1], iv_entry[2], iv_entry[3],
                        iv_entry[4], iv_entry[5], iv_entry[6],
                        temp_entry[1]
                    ]
                    ws.append(row)

                wb.save(self.excel_file_path)
                logging.info(f"Data saved to Excel: {self.excel_file_path}")
            except PermissionError:
                logging.error("PermissionError: Cannot save Excel file. Please close it if open.")
            except Exception as e:
                logging.error(f"Error saving Excel file: {e}")

if __name__ == "__main__":
    app = TestApp()
    app.mainloop()
