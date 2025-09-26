import serial
import time
import pandas as pd
import re

# ========= FUNCIONES =========

def obtener_modelo_serie(ser):
    """
    Ejecuta 'show inventory' en el dispositivo conectado por puerto serial
    y extrae el modelo (PID) y el número de serie (SN).
    """
    ser.write(b"show inventory\n")
    time.sleep(2)

    salida = ""
    if ser.in_waiting:
        # Lee todos los datos disponibles en el buffer de entrada
        salida = ser.read(ser.in_waiting).decode(errors="ignore")

    # Utiliza expresiones regulares para encontrar el PID y SN
    regex_modelo = re.search(r"PID:\s*([\w\-/]+)", salida)
    regex_serie = re.search(r"SN:\s*([\w\d]+)", salida)

    # Extrae los valores si se encontraron coincidencias
    modelo = regex_modelo.group(1) if regex_modelo else None
    serie = regex_serie.group(1) if regex_serie else None

    return modelo, serie, salida


def configurar_dispositivo(ser, nombre, usuario, contrasena, dominio):
    """
    Envía la secuencia de comandos de configuración básicos al dispositivo.
    """
    print(f"⚙️  Aplicando configuración para el hostname: {nombre}...")
    
    comandos = [
        "configure terminal",
        f"hostname {nombre}",
        f"username {usuario}, privilege 15 password {contrasena}",
        f"ip domain-name {dominio}",
        "crypto key generate rsa",
    ]

    for cmd in comandos:
        ser.write(f"{cmd}\n".encode())
        time.sleep(1)

    # El comando 'crypto key' pide confirmación del tamaño de la clave
    ser.write(b"1024\n")
    time.sleep(2)

    # Comandos adicionales para SSH y acceso
    extra_cmds = [
        "ip ssh version 2",
        "line console 0",
        "login local",
        "line vty 0 4",
        "login local",
        "transport input ssh",
        "transport output ssh",
        "end",
        "write memory"
    ]

    for cmd in extra_cmds:
        ser.write(f"{cmd}\n".encode())
        time.sleep(1)

    print(f"✅ Configuración para '{nombre}' aplicada y guardada correctamente.")


def cargar_y_configurar(ruta_excel):
    """
    Función principal: Se conecta a un dispositivo, lo identifica, y luego
    busca una coincidencia en el Excel para configurarlo.
    """
    try:
        df = pd.read_excel(ruta_excel)
    except FileNotFoundError:
        print(f"❌ ERROR: No se pudo encontrar el archivo Excel en la ruta: {ruta_excel}")
        return
    
    # Validar que el Excel contenga las columnas necesarias
    columnas_requeridas = {"modelo", "serie", "puerto", "baudios", "nombre", "usuario", "contrasena", "dominio"}
    if not columnas_requeridas.issubset(df.columns):
        print(f"❌ ERROR: El archivo Excel debe tener las siguientes columnas: {', '.join(columnas_requeridas)}")
        return
    if df.empty:
        print("⚠️ El archivo Excel está vacío. No hay nada que hacer.")
        return

    # Toma los datos de conexión de la primera fila del Excel
    primera_fila = df.iloc[0]
    puerto = primera_fila["puerto"]
    baudios = int(primera_fila["baudios"])
    ser = None
    
    try:
        print(f"🔌 Intentando conectar al puerto {puerto}...")
        ser = serial.Serial(puerto, baudios, timeout=2)
        time.sleep(2)

        # 1. Detectar el dispositivo primero
        modelo_real, serie_real, _ = obtener_modelo_serie(ser)
        
        if not modelo_real or not serie_real:
            print(f"⚠️ No se pudo obtener la información del dispositivo en {puerto}. Verifique la conexión.")
            return

        print("\n===================== DISPOSITIVO DETECTADO =====================")
        print(f"  > Modelo Router: {modelo_real}")
        print(f"  > Numero de Serie: {serie_real}")
        print("=================================================================\n")
        print("🔎 Buscando coincidencia en el archivo Excel...")
        
        coincidencia_exitosa = False
        
        # 2. Iterar sobre el Excel para buscar una coincidencia
        for indice, fila in df.iterrows():
            nombre_router_excel = fila["nombre"]
            print(f"   - Verificando contra '{nombre_router_excel}'...", end="")

            if modelo_real == fila["modelo"] and serie_real == fila["serie"]:
                print(" ✅ Coincide.")
                configurar_dispositivo(
                    ser,
                    fila["nombre"],
                    fila["usuario"],
                    fila["contrasena"],
                    fila["dominio"]
                )
                coincidencia_exitosa = True
                break
            else:
                print(" ❌ No coincide.")
        
        if not coincidencia_exitosa:
            print("\n===========================================================================")
            print("🚫 Proceso finalizado. El dispositivo detectado no coincide con ninguna")
            print("   entrada en el archivo Excel.")
            print("\n   Se verificó contra los siguientes routers:")
            for nombre in df['nombre']:
                print(f"   - {nombre}")
            print("===========================================================================")

    except serial.SerialException as e:
        print(f"❌ Error de conexión en {puerto}: {e}. ¿Está el cable bien conectado?")
    except Exception as e:
        print(f"❌ Ocurrió un error inesperado: {e}")
    finally:
        if ser and ser.is_open:
            ser.close()
            print(f"\n🔌 Puerto {puerto} cerrado.")


# ========= BLOQUE PRINCIPAL =========

if __name__ == "__main__":
    # IMPORTANTE: Cambia esta ruta por la ubicación real de tu archivo Excel
    ruta_archivo_excel = (r"C:\Users\suri yatziri ortiz r\Documents\Codigos\GIT\Routers_Info.xlsx")
    cargar_y_configurar(ruta_archivo_excel)

