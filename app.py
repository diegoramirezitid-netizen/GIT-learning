import socket
print("Sexo")

hostname = socket.gethostname()
print(f"Hostname: {hostname}")

ipaddress = socket.gethostbyname(hostname)
print(f"IP: {ipaddress}")

print("Nuevo")

print("Creo que ya") 

numero_a = int(input("Dame el primer numero: "))
numero_b = int(input("Dame el segundo numero: "))
print(f"La suma es: {numero_a + numero_b}")
print(f"La resta de los numeros es: {numero_a - numero_b}")

print(f"La multiplicacion de los numeros es: {numero_a * numero_b}")