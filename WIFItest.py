import subprocess
import socket

def get_devices():
    try:
        result = subprocess.check_output(["arp", "-a"], universal_newlines = True)
        lines = result.split("\n")[3:]
        connected_devices = []

        for line in lines:
            parts = line.split()
            if len(parts) == 3:
                ip_address, mac_address, _ = parts
                try:
                    hostname = socket.gethostbyaddr(ip_address)[0]
                except socket.herror:
                    hostname = "None"
                
                connected_devices.append((ip_address, mac_address, hostname))
        return connected_devices
    
    except subprocess.CalledProcessError as error:
        print(f"Error: {error}")
        return []
    
if __name__ == "__main__":
    devices = get_devices()
    if devices:
        print("Device list:")
        for hostname in devices:
            print(f"Hostname: {hostname}")
    else:
        print("No devices")