
import subprocess
import socket

def get_connected_devices():
    try:
        # Run the "arp -a" command to get a list of devices in the ARP table
        result = subprocess.check_output(["arp", "-a"], universal_newlines=True)
        # Split the result into lines and skip the header
        lines = result.split('\n')[3:]
        connected_devices = []

        for line in lines:
            # Split each line to extract IP and MAC addresses
            parts = line.split()
            if len(parts) == 3:
                ip_address, mac_address, _ = parts

                # Try to resolve the hostname using DNS (or mark as "N/A" if not found)
                try:
                    hostname = socket.gethostbyaddr(ip_address)[0]
                except socket.herror:
                    hostname = "N/A"

                # Add the device information to the list
                connected_devices.append((ip_address, mac_address, hostname))

        return connected_devices
    except subprocess.CalledProcessError as e:
        # Handle any errors that may occur during command execution
        print(f"Error: {e}")
        return []

if __name__ == "__main__":
    # Get the list of connected devices
    devices = get_connected_devices()
    if devices:
        print("List of connected devices:")
        for ip, mac, hostname in devices:
            # Print the IP address, MAC address, and hostname (device name)
            print(f"IP Address: {ip}, MAC Address: {mac}, Hostname: {hostname}")
    else:
        print("No devices found.")