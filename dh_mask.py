import subprocess
import socket

def get_devices():
    try:
        # Retrieve ARP table using the 'arp' command in the system shell
        result = subprocess.check_output(["arp", "-a"], universal_newlines=True)

        # Split the result into lines and ignore the first three lines (headers)
        lines = result.split("\n")[3:]
        connected_devices = []

    # Loop through the lines retrieved from the ARP table
        for line in lines:
            parts = line.split()

            # If the line has three parts (IP, MAC, and description), process it
            if len(parts) == 3:
                ip_address, mac_address, _ = parts

                # Try to get the hostname using IP address
                try:
                    hostname = socket.gethostbyaddr(ip_address)[0]
                except socket.herror:
                    hostname = "None"
            
            # Append the details (IP, MAC, hostname) to the list of connected devices
            connected_devices.append((ip_address, mac_address, hostname))

    # Return the list of connected devices
        return connected_devices
    
    # Catch any errors if they occur
    except subprocess.CalledProcessError as error:
        print(f"Error: {error}")
        return []
    
if __name__ == "__main__":
    # Call the get_devices function to get the list of connected devices
    devices = get_devices()
    
    # Print the device list if any devices are found; otherwise, print "No devices"
    if devices:
        print("Device list:")
        for ip, mac, hostname in devices:
            print(f"IP: {ip}, MAC: {mac}, Hostname: {hostname}")
    else:
        print("No devices")