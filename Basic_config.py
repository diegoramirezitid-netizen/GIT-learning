import serial
import time
def configure_device(port, baidrate, com, hostname, username, password, domain):
    try:
        # Open serial connection
        ser = serial.Serial(com, baudrate=9600, timeout=1)
        time.sleep(2)  # Wait for the connection to establish

        # Send configuration commands
        ser.write(b"enable\n")
        time.sleep(1)
        ser.write(b"configure terminal\n")
        time.sleep(1)
        ser.write(f"set hostname {hostname}\n".encode())
        time.sleep(1)
        ser.write(f"username {username} privilege level 15 password {password}\n".encode())
        time.sleep(1)
        ser.write(f"".encode())
        time.sleep(1)
        ser.write(f"set password {password}\n".encode())
        time.sleep(1)
        ser.write(f"set domain {domain}\n".encode())
        time.sleep(1)
        ser.write("save\n".encode())
        time.sleep(1)

        print("Configuration successful.")