import serial
import time
import pandas as pd
import os

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def configure_device(port, baudrate, hostname, username, password, domain):
    try:
        ser = serial.Serial(port, baudrate, timeout=1)
        print(f"Conectado a {port}. Configurando el dispositivo: {hostname}")
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
        time.sleep(5) # Aumentamos la espera por si la generación de claves es lenta
        ser.write(b'ip ssh version 2\n')
        time.sleep(1)
        ser.write(b'line console 0\n')
        time.sleep(1)
        ser.write(b'login local\n')
        time.sleep(1)
        ser.write(b'exit\n') # Salir de line console 0
        time.sleep(1)
        ser.write(b'line vty 0 4\n')
        time.sleep(1)
        ser.write(b'login local\n')
        time.sleep(1)
        ser.write(b'transport input ssh\n')
        time.sleep(1)
        # El comando 'transport output ssh' no es estándar y puede dar error, se ha comentado
        # ser.write(b'transport output ssh\n')
        # time.sleep(1)
        ser.write(b'exit\n') # Salir de line vty 0 4
        time.sleep(1)
        ser.write(b'end\n') # Salir del modo de configuración
        time.sleep(1)
        ser.write(b'write memory\n')
        time.sleep(5) # Esperar a que se guarde la configuración
        ser.close()
        print(f"Configuración de '{hostname}' completada exitosamente.")
        return True
    except Exception as e:
        print(f"Error al configurar el dispositivo '{hostname}': {e}")
        return False

# --- INICIO DE LA LÓGICA PARA LEER EXCEL Y EJECUTAR ---

# 1. Solicitar información al usuario
clear_screen()
print("--- Script de Configuración de Dispositivos por Consola ---")
excel_file = input("Introduce la ruta de tu archivo Excel: ")
com_port = input("Introduce el puerto COM a utilizar (ej: COM3): ")
baudrate = 9600

# 2. Leer el archivo Excel y procesar los dispositivos
try:
    df = pd.read_excel(excel_file)
    
    # Validar que las columnas necesarias existan
    required_columns = ['Hostname', 'Username', 'Password', 'Domain']
    if not all(col in df.columns for col in required_columns):
        print("\nError: Asegúrate de que tu archivo Excel tenga las columnas: 'Hostname', 'Username', 'Password', 'Domain'")
    else:
        print(f"\nSe encontraron {len(df)} dispositivos para configurar en el archivo.")
        
        # 3. Iterar sobre cada dispositivo en el archivo
        for index, row in df.iterrows():
            # Extraer los datos de la fila
            hostname = row['Hostname']
            username = row['Username']
            password = row['Password']
            domain = row['Domain']

            # Mostrar información y esperar confirmación
            print("\n" + "="*50)
            print(f"Siguiente dispositivo a configurar ({index + 1}/{len(df)}):")
            print(f"  > Hostname: {hostname}")
            print(f"  > Usuario:  {username}")
            print(f"  > Dominio:  {domain}")
            print("="*50)
            
            # Pausa para que el usuario pueda conectar el cable al siguiente dispositivo
            input("Conecta el cable de consola al dispositivo y presiona Enter para continuar...")

            # Llamar a la función de configuración
            configure_device(com_port, baudrate, hostname, username, str(password), domain)
        
        print("\nProceso finalizado. Todos los dispositivos han sido procesados.")

except FileNotFoundError:
    print(f"\nError: No se pudo encontrar el archivo en la ruta '{excel_file}'")
except Exception as e:
    print(f"\nOcurrió un error inesperado: {e}")