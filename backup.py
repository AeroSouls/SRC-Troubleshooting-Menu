import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import webbrowser
import os
import socket
import subprocess
import threading
import queue
import datetime
import shutil
import tempfile
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

ping_output_queue = queue.Queue()

def restart_computer():
    confirm = messagebox.askyesno("Restart Computer", "Are you sure you want to restart the computer?")
    if confirm:
        run_command("shutdown /r /t 5")

def ping_google(ip_address):
    global p, ping_thread

    ping_output_var.set("")  # Clear the ping output variable

    # Start a subprocess to ping the specified IP address
    p = subprocess.Popen(["ping", "-t", ip_address], stdout=subprocess.PIPE, stderr=subprocess.PIPE, bufsize=1, universal_newlines=True)

    # Define a function to continuously update the ping output text
    def update_ping_output_text():
        global p, last_three_lines

        # Loop through each line of output from the ping command
        for line in iter(p.stdout.readline, ""):
            if not line:
                break

            # Add a timestamp to the output line and add it to the queue for logging
            timestamp = datetime.datetime.now().strftime("%m-%d-%Y %I:%M:%S %p")
            line_with_timestamp = f"{timestamp} - {line.strip()}"
            ping_output_queue.put(line_with_timestamp)

            # Update the last three lines of the ping output
            last_three_lines.pop(0)
            last_three_lines.append(line_with_timestamp)

            # Clear the output text widget and add the last three lines to it
            output_text.config(state=tk.NORMAL)
            output_text.delete("1.0", tk.END)
            output_text.tag_configure("success", foreground="green")
            output_text.tag_configure("failure", foreground="red")

            for line in last_three_lines:
                if "Reply from" in line:
                    output_text.insert(tk.END, line + "\n", "success")
                else:
                    output_text.insert(tk.END, line + "\n", "failure")

            # Disable the output text widget and update the scrollbar
            output_text.config(state=tk.DISABLED)
            scrollbar.update()

    # Start a thread to continuously update the ping output text
    ping_thread = threading.Thread(target=update_ping_output_text, daemon=True)
    ping_thread.start()

def stop_ping():
    # Access the global variable 'p'
    global p
    # If 'p' exists, terminate the process and set the ping output variable
    if p:
        p.terminate()
        ping_output_var.set("Ping HAS STOPPED")

def write_log_file():
    while True:
        try:
            # Get the next line from the ping output queue, waiting for 1 second if necessary
            line = ping_output_queue.get(timeout=1)
        except queue.Empty:
            # If no line is available in the queue, continue to the next iteration of the loop
            continue

        # Append the line to the 'ping_log.txt' file
        with open("ping_log.txt", "a") as log_file:
            log_file.write(line + "\n")
            log_file.flush()

# Create a new thread for the log writer and start it
log_writer_thread = threading.Thread(target=write_log_file, daemon=True)
log_writer_thread.start()

def download_ping_log():
    # If the 'ping_log.txt' file exists
    if os.path.exists("ping_log.txt"):
        # Create a temporary directory and copy the log file to it
        temp_dir = tempfile.mkdtemp()
        temp_file = os.path.join(temp_dir, "ping_log.txt")
        shutil.copy("ping_log.txt", temp_file)
        # Open the temporary file in the default web browser
        webbrowser.open(temp_file)
    else:
        # If the log file doesn't exist, set the ping output variable
        ping_output_var.set("No log file found.")

# Create the main window
window = tk.Tk()
window.title("SRC Troubleshooter Menu #252-756-0004")

# Create a StringVar to store the ping output variable and a list to store the last three lines of output
ping_output_var = tk.StringVar()
last_three_lines = ["", "", "", "", "", "", ""]

def run_command(command):
    # Run the specified command and capture its output
    output = os.popen(command).read()
    # Enable the output text widget, clear its contents, and insert the command output
    output_text.config(state=tk.NORMAL)
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, output)
    # Disable the output text widget and update the scrollbar
    output_text.config(state=tk.DISABLED)
    scrollbar.update()

#Fancy buttons
style = ttk.Style()
style.configure("Fancy.TButton",
                foreground="black",
                background="black",
                font=("Helvetica", 12, "bold"),
                padding=10,
                relief="raised",
                borderwidth=2)
style.map("Fancy.TButton",
          background=[("active", "#4f6f8f"), ("!disabled", "black")])

# Create a menu bar
menu_bar = tk.Menu(window)
window.config(menu=menu_bar)

# Add "Open Admin CMD" and "Open Computer Manag" directly to the menu bar
menu_bar.add_command(label="Open Admin CMD", command=lambda: run_command("start cmd"))
menu_bar.add_command(label="Open Computer Management ", command=lambda: run_command("start compmgmt.msc"))
menu_bar.add_command(label="Open Device Manager", command=lambda: run_command("start devmgmt.msc"))


# Add "Tools" menu with "Download Files (Agent Installer)" option
download_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Downloads", menu=download_menu)
download_menu.add_command(label="Agent Installer", command=lambda: webbrowser.open("https://itbysrc.com/agent/Agent_Install.MSI"))

# Modify the button_frame layout
button_frame = ttk.Frame(window)
button_frame.grid(column=0, row=0, padx=20, pady=10)

button1 = ttk.Button(button_frame, text="Restart Computer", command=restart_computer, style="Fancy.TButton") 
button2 = ttk.Button(button_frame, text="Show Temp Directory", command=lambda: run_command("dir %temp%\\*"), style="Fancy.TButton")
button3 = ttk.Button(button_frame, text="Clear Temp Files", command=lambda: run_command("del /s /q /f %temp%\\*"), style="Fancy.TButton")

support_button = ttk.Button(button_frame, text="Remote Support", command=lambda: webbrowser.open("https://itbysrc.com/remote-support/"), style="Fancy.TButton")

output_text = tk.Text(window, height=10, state=tk.DISABLED)
scrollbar = tk.Scrollbar(window, command=output_text.yview)
output_text.config(yscrollcommand=scrollbar.set)

clear_button = ttk.Button(window, text="Clear Output", command=lambda: clear_output(), style="Fancy.TButton")

def clear_output():
    output_text.config(state=tk.NORMAL)
    output_text.delete("1.0", tk.END)
    output_text.config(state=tk.DISABLED)
    scrollbar.update()

button1.grid(row=0, column=0, padx=5, pady=5)
button2.grid(row=1, column=0, padx=5, pady=5)
button3.grid(row=2, column=0, padx=5, pady=5)
support_button.grid(row=3, column=0, padx=5, pady=5)

output_text.grid(column=1, row=0, padx=10, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
scrollbar.grid(column=1, row=0, padx=(0, 10), pady=10, sticky=(tk.NE, tk.SE))  # Adjust the padx to shift the scrollbar to the right

clear_button.grid(column=1, row=1, padx=10, pady=5, sticky=(tk.W, tk.E))

# Ping Frame / Ping Entry
ping_frame = ttk.Frame(window)
ping_frame.grid(column=0, row=2, padx=20, pady=5)

ip_entry = ttk.Entry(ping_frame)
ip_entry.grid(row=0, column=0, padx=5, pady=5)
ip_entry.insert(0, "google.com")  # Set a default value

ping_button = ttk.Button(ping_frame, text="Ping IP", command=lambda: ping_google(ip_entry.get()), style="Fancy.TButton")
ping_button.grid(row=1, column=0, padx=5, pady=5)

download_log_button = ttk.Button(ping_frame, text="Open Ping Log", command=download_ping_log, style="Fancy.TButton")
download_log_button.grid(row=2, column=0, padx=5, pady=5)

stop_button = ttk.Button(window, text="Stop Ping", command=stop_ping, style="Fancy.TButton")
stop_button.grid(column=1, row=2, padx=10, pady=5, sticky=(tk.W, tk.E))

ping_output_label = ttk.Label(window, textvariable=ping_output_var, anchor=tk.W, justify=tk.LEFT)
ping_output_label.grid(column=1, row=3, padx=10, pady=10, sticky=(tk.W, tk.E))

# Create a status panel
status_panel = ttk.Frame(window)
status_panel.grid(column=2, row=0, padx=5, pady=5, rowspan=3, sticky=tk.N)  # Adjust the rowspan to accommodate the scrollbar

# Add labels for computer name, IP address, and serial number
computer_name_label = ttk.Label(status_panel, text="Computer: {}".format(get_computer_name()), anchor=tk.W)
computer_name_label.grid(column=0, row=0, padx=5, pady=5, sticky=tk.W)

ip_address_label = ttk.Label(status_panel, text="IP Address: {}".format(get_ip_address()), anchor=tk.W)
ip_address_label.grid(column=0, row=1, padx=5, pady=5, sticky=tk.W)

serial_number_label = ttk.Label(status_panel, text="Serial Number: {}".format(get_serial_number()), anchor=tk.W)
serial_number_label.grid(column=0, row=2, padx=5, pady=5, sticky=tk.W)

connectwise_id_label = ttk.Label(status_panel, text="ConnectWise ID: {}".format(get_connectwise_id()), anchor=tk.W)
connectwise_id_label.grid(column=0, row=3, padx=5, pady=5, sticky=tk.W)

window.columnconfigure(1, weight=1)
window.rowconfigure(0, weight=1)

window.mainloop()