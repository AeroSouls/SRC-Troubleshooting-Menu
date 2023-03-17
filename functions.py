import os
import socket
import subprocess
import winreg

def get_computer_name():
    return os.environ['COMPUTERNAME']

def get_ip_address():
    try:
        hostname = socket.gethostname()
        ip_address = socket.gethostbyname(hostname)
        return ip_address
    except Exception as e:
        return "Unknown"

def get_serial_number():
    output = os.popen("wmic bios get serialnumber")
    for line in output:
        if "SerialNumber" in line or len(line.strip()) == 0:
            continue
        serial_number = line.strip()
        if serial_number == "System Serial Number" or serial_number == "To be filled by O.E.M.":
            return "Custom Build"
        elif "VMware" in serial_number:
            return "VMware"
        else:
            return serial_number
    return "Unknown"

def get_connectwise_id():
    registry_keys = [
        r"HKLM\SOFTWARE\WOW6432Node\LabTech\Service",
        r"HKLM\SOFTWARE\LabTech\Service",
    ]

    for registry_key in registry_keys:
        try:
            output = subprocess.check_output(
                f'reg query "{registry_key}" /v ID', shell=True, text=True, stderr=subprocess.DEVNULL
            )

            for line in output.splitlines():
                if "ID" in line:
                    _, _, connectwise_id = line.strip().partition("ID")
                    connectwise_id = connectwise_id.strip().split()[-1]  # Extract the last element after splitting the remaining string
                    if connectwise_id.startswith("0x"):
                        connectwise_id = int(connectwise_id, 16)  # Convert the hexadecimal value to an integer
                    if connectwise_id:
                        return connectwise_id
        except subprocess.CalledProcessError:
            pass

    return "Not Installed"