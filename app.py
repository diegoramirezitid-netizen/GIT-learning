import socket
print("Sexo")

hostname = socket.gethostname()
print(f"Hostname: {hostname}")

ipaddress = socket.gethostbyname(hostname)
print(f"IP: {ipaddress}")

print("Nuevo")

print("Creo que ya")