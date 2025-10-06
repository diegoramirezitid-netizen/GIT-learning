import serial
import time
import pandas as pd
import re
import textfsm
from io import StringIO
import os
from datetime import datetime

# ========= CONFIGURACIÓN =========
PUERTO = "COM4"
BAUDIOS = 9600
ARCHIVO_EXCEL = os.path.join(os.path.expanduser("~"), "Desktop", "Inventario_Routers.xlsx")

# ========= PLANTILLA TEXTFSM CORREGIDA =========
PLANTILLA_INTERFACES = r"""
Value Interface (\S+)
Value IP_Address (unassigned|[\d\.]+)
Value OK (\S+)
Value Method (\S+)
Value Status (up|down|administratively down)
Value Protocol (up|down)

Start
  ^${Interface}\s+${IP_Address}\s+${OK}\s+${Method}\s+${Status}\s+${Protocol} -> Record
"""

# ========= FUNCIONES =========

def enviar_comando(ser, comando, tiempo_espera=2):
    """Envía un comando al router y retorna la salida"""
    ser.write(f"{comando}\n".encode())
    time.sleep(tiempo_espera)
    
    salida = ""
    while ser.in_waiting:
        salida += ser.read(ser.in_waiting).decode(errors="ignore")
        time.sleep(0.5)
    
    return salida

def obtener_info_basica(ser):
    """Obtiene modelo, serie, hostname y uptime del router"""
    print("🔍 Obteniendo información básica del router...")
    
    # Obtener modelo y serie
    salida_inventory = enviar_comando(ser, "show inventory")
    modelo = re.search(r"PID:\s*([\w\-/]+)", salida_inventory)
    serie = re.search(r"SN:\s*([\w\d]+)", salida_inventory)
    
    # Obtener hostname - método más robusto
    salida_version = enviar_comando(ser, "show version")
    
    # Intentar diferentes patrones para hostname
    hostname = re.search(r'^(\S+)\s+uptime', salida_version)
    if not hostname:
        hostname = re.search(r'^([a-zA-Z0-9\-_]+)#', salida_version)
    if not hostname:
        # Buscar en el prompt de comandos
        hostname = re.search(r'([a-zA-Z0-9\-_]+)>\s*$', salida_version)
    if not hostname:
        hostname = re.search(r'([a-zA-Z0-9\-_]+)#\s*$', salida_version)
    
    uptime = re.search(r'uptime\s+is\s+(.+)', salida_version)
    
    info_basica = {
        'modelo': modelo.group(1) if modelo else 'N/A',
        'serie': serie.group(1) if serie else 'N/A',
        'hostname': hostname.group(1) if hostname else 'N/A',
        'uptime': uptime.group(1) if uptime else 'N/A',
        'fecha_escaneo': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    # Debug para hostname
    if info_basica['hostname'] == 'N/A':
        print("⚠️  No se pudo detectar el hostname, revisando salida...")
        lineas = salida_version.split('\n')
        for i, linea in enumerate(lineas[-5:]):  # Últimas 5 líneas
            print(f"   Línea {i}: {linea.strip()}")
    
    return info_basica

def obtener_interfaces(ser):
    """Obtiene información de interfaces"""
    print("🔍 Obteniendo información de interfaces...")
    
    # Usar "sshow" como mencionaste que funciona
    salida_interfaces = enviar_comando(ser, "sshow ip interface brief")
    
    # Limpiar la salida
    lineas_limpias = []
    for linea in salida_interfaces.split('\n'):
        if 'Invalid input' not in linea and '^' not in linea and linea.strip():
            if '#' in linea:
                linea = linea.split('#')[-1].strip()
            lineas_limpias.append(linea)
    
    salida_limpia = '\n'.join(lineas_limpias)
    
    if not salida_limpia:
        print("❌ No se recibió salida válida del comando")
        return []
    
    # Intentar parsear con TextFSM primero
    interfaces = parsear_con_textfsm(salida_limpia)
    
    # Si TextFSM no funciona, usar método manual
    if not interfaces:
        interfaces = parsear_interfaces_manual(salida_limpia)
    
    return interfaces

def parsear_con_textfsm(salida):
    """Intenta parsear con TextFSM"""
    try:
        template = textfsm.TextFSM(StringIO(PLANTILLA_INTERFACES))
        resultados = template.ParseText(salida)
        
        print(f"✅ Interfaces parseadas con TextFSM: {len(resultados)}")
        
        interfaces = []
        for resultado in resultados:
            interfaz = {
                'Interface': resultado[0],
                'IP-Address': resultado[1],
                'OK': resultado[2],
                'Method': resultado[3],
                'Status': resultado[4],
                'Protocol': resultado[5]
            }
            interfaces.append(interfaz)
        
        return interfaces
        
    except Exception as e:
        print(f"❌ TextFSM falló: {e}")
        return []

def parsear_interfaces_manual(salida):
    """Método manual para parsear interfaces"""
    print("🔧 Usando parser manual...")
    interfaces = []
    
    lineas = salida.split('\n')
    for linea in lineas:
        # Buscar líneas que contengan interfaces
        if any(tipo in linea for tipo in ['GigabitEthernet', 'FastEthernet', 'Serial', 'Loopback', 'Vlan', 'Tunnel']):
            partes = linea.split()
            if len(partes) >= 6:
                interfaz = {
                    'Interface': partes[0],
                    'IP-Address': partes[1],
                    'OK': partes[2],
                    'Method': partes[3],
                    'Status': partes[4],
                    'Protocol': partes[5]
                }
                interfaces.append(interfaz)
    
    print(f"✅ Interfaces parseadas manualmente: {len(interfaces)}")
    return interfaces

def cargar_excel_existente():
    """Carga el Excel existente o crea uno nuevo"""
    try:
        if os.path.exists(ARCHIVO_EXCEL):
            return pd.read_excel(ARCHIVO_EXCEL)
        else:
            # Crear DataFrame vacío con la estructura correcta
            columnas = [
                'modelo', 'Numero de serie', 'Hostname', 'fecha_escaneo',
                'Interface', 'IP-Address', 'OK', 'Method', 'Status', 'Protocol'
            ]
            return pd.DataFrame(columns=columnas)
    except Exception as e:
        print(f"❌ Error al cargar Excel: {e}")
        # Crear DataFrame vacío
        columnas = [
            'modelo', 'Numero de serie', 'Hostname', 'fecha_escaneo',
            'Interface', 'IP-Address', 'OK', 'Method', 'Status', 'Protocol'
        ]
        return pd.DataFrame(columns=columnas)

def dispositivo_existe(df_existente, serie):
    """Verifica si el dispositivo ya existe en el Excel"""
    if df_existente.empty:
        return False
    return serie in df_existente['Numero de serie'].values

def actualizar_dispositivo(df_existente, info_basica, interfaces):
    """Actualiza la información de un dispositivo existente"""
    # Eliminar registros antiguos del dispositivo
    df_actualizado = df_existente[df_existente['Numero de serie'] != info_basica['serie']].copy()
    
    # Agregar nuevos registros
    nuevos_datos = []
    for interfaz in interfaces:
        nuevo_registro = {
            'modelo': info_basica['modelo'],
            'Numero de serie': info_basica['serie'],
            'Hostname': info_basica['hostname'],
            'fecha_escaneo': info_basica['fecha_escaneo'],
            'Interface': interfaz['Interface'],
            'IP-Address': interfaz['IP-Address'],
            'OK': interfaz['OK'],
            'Method': interfaz['Method'],
            'Status': interfaz['Status'],
            'Protocol': interfaz['Protocol']
        }
        nuevos_datos.append(nuevo_registro)
    
    df_nuevo = pd.DataFrame(nuevos_datos)
    df_final = pd.concat([df_actualizado, df_nuevo], ignore_index=True)
    
    return df_final

def agregar_dispositivo(df_existente, info_basica, interfaces):
    """Agrega un nuevo dispositivo al Excel"""
    nuevos_datos = []
    for interfaz in interfaces:
        nuevo_registro = {
            'modelo': info_basica['modelo'],
            'Numero de serie': info_basica['serie'],
            'Hostname': info_basica['hostname'],
            'fecha_escaneo': info_basica['fecha_escaneo'],
            'Interface': interfaz['Interface'],
            'IP-Address': interfaz['IP-Address'],
            'OK': interfaz['OK'],
            'Method': interfaz['Method'],
            'Status': interfaz['Status'],
            'Protocol': interfaz['Protocol']
        }
        nuevos_datos.append(nuevo_registro)
    
    df_nuevo = pd.DataFrame(nuevos_datos)
    
    if df_existente.empty:
        return df_nuevo
    else:
        return pd.concat([df_existente, df_nuevo], ignore_index=True)

def guardar_en_excel(info_basica, interfaces, df_existente):
    """Guarda la información en el Excel acumulativo"""
    
    # Verificar si el dispositivo ya existe
    if dispositivo_existe(df_existente, info_basica['serie']):
        print(f"🔄 Dispositivo {info_basica['serie']} ya existe, actualizando...")
        df_final = actualizar_dispositivo(df_existente, info_basica, interfaces)
    else:
        print(f"➕ Nuevo dispositivo {info_basica['serie']}, agregando...")
        df_final = agregar_dispositivo(df_existente, info_basica, interfaces)
    
    # Guardar en Excel con formato
    try:
        with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Inventario', index=False)
            
            # Aplicar formato bonito
            workbook = writer.book
            worksheet = workbook['Inventario']
            
            # Ajustar ancho de columnas
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 25)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Congelar paneles (primera fila de encabezados)
            worksheet.freeze_panes = 'A2'
        
        print(f"✅ Información guardada en: {ARCHIVO_EXCEL}")
        print(f"📊 Total de registros en el inventario: {len(df_final)}")
        print(f"🔢 Total de dispositivos únicos: {df_final['Numero de serie'].nunique()}")
        
    except Exception as e:
        print(f"❌ Error al guardar en Excel: {e}")

def main():
    """Función principal - Escanea un router y guarda la información"""
    ser = None
    
    try:
        print("🚀 INICIANDO ESCANEO DE ROUTER")
        print("=" * 50)
        print(f"🔌 Conectando al puerto {PUERTO}...")
        
        ser = serial.Serial(PUERTO, BAUDIOS, timeout=2)
        time.sleep(3)
        
        # Esperar a que el router esté listo
        ser.write(b"\r\n")
        time.sleep(1)
        
        # 1. Obtener información básica
        info_basica = obtener_info_basica(ser)
        
        print("\n📊 INFORMACIÓN BÁSICA DEL ROUTER:")
        print(f"   Modelo: {info_basica['modelo']}")
        print(f"   Número de Serie: {info_basica['serie']}")
        print(f"   Hostname: {info_basica['hostname']}")
        print(f"   Uptime: {info_basica['uptime']}")
        
        # 2. Obtener información de interfaces
        interfaces = obtener_interfaces(ser)
        
        print(f"\n🔌 INTERFACES DETECTADAS: {len(interfaces)}")
        for interfaz in interfaces:
            print(f"   - {interfaz['Interface']}: {interfaz['IP-Address']} ({interfaz['Status']}/{interfaz['Protocol']})")
        
        # 3. Cargar Excel existente y guardar/actualizar información
        df_existente = cargar_excel_existente()
        guardar_en_excel(info_basica, interfaces, df_existente)
        
        print(f"\n✅ ESCANEO COMPLETADO PARA: {info_basica['hostname']}")
        
    except serial.SerialException as e:
        print(f"❌ Error de conexión: {e}")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")
    finally:
        if ser and ser.is_open:
            ser.close()
            print("🔌 Conexión cerrada.")

# ========= EJECUCIÓN =========
if __name__ == "__main__":
    main()