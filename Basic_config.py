import serial
import time
import pandas as pd
import os
def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')
    

def configure_device(port, baudrate, hostname, username, password, domain):
    try:
        ser = serial.Serial(port, baudrate, timeout=1)
        time.sleep(2)  # Esperar a que la conexión se establezca
        ser.write(b'\n')  # Enviar un salto de línea para iniciar la comunicación
        time.sleep(1)
        ser.write(b'enable\n')
        time.sleep(1)
        ser.write(b'configure terminal\n')
        time.sleep(1)
        ser.write(f'hostname {hostname}\n'.encode())
        time.sleep(1)
        ser.write(f'username {username} privilege 15 password {password}\n'.encode())
        time.sleep(1)
        ser.write(f'ip domain-name {domain}\n'.encode())
        time.sleep(1)
        ser.write(b'crypto key generate rsa\n')
        time.sleep(2)  # Esperar a que se genere la clave
        ser.write(b'1024\n')  # Tamaño de la clave
        time.sleep(2)
        ser.write(b'ip ssh version 2\n')
        time.sleep(1)
        ser.write(b'line console 0\n')
        time.sleep(1)
        ser.write(b'login local\n')
        time.sleep(1)
        ser.write(b'line vty 0 4\n')
        time.sleep(1)
        ser.write(b'login local\n')
        time.sleep(1)
        ser.write(b'transport input ssh\n')
        time.sleep(1)
        ser.write(b'transport output ssh\n')
        time.sleep(1)
        ser.write(b'exit\n')
        time.sleep(1)
        ser.write(b'write memory\n')
        time.sleep(1)
        ser.close()
        print("Configuración completada exitosamente.")
    except Exception as e:
        print(f"Error al configurar el dispositivo: {e}")


#nombre del dispositivo, serie del dispositivo, 
#como cargar un excel y que lea campos
