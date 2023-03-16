import tkinter as tk
from tkinter import ttk
import os
import webbrowser
import subprocess
import threading
import shutil
import tempfile
import queue
import datetime
from tkinter import PhotoImage
import time

ping_output_queue = queue.Queue()

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

def open_multi_ping_window():
    def add_ping():
        ip_address = ip_entry.get()
        if ip_address:
            treeview.insert("", "end", text=ip_address, values=(ip_address, ""))
            ip_entry.delete(0, "end")

    def remove_ping():
        selected_item = treeview.selection()[0]
        treeview.delete(selected_item)

    def start_multi_ping():
        global stop_pinging
        stop_pinging = False
        for item in treeview.get_children():
            ip_address = treeview.item(item)["values"][0]
            ping_thread = threading.Thread(target=update_ping_output, args=(ip_address, item), daemon=True)
            ping_thread.start()

    def stop_multi_ping():
        global stop_pinging
        stop_pinging = True

    def update_ping_output(ip_address, item):
        while not stop_pinging:
            p = subprocess.Popen(["ping", "-n", "1", ip_address], stdout=subprocess.PIPE, stderr=subprocess.PIPE, bufsize=1, universal_newlines=True)
            output, _ = p.communicate()

            ping_output = ""
            for line in output.split("\n"):
                if "time=" in line:
                    ping_output = line.split("time=")[-1]
                    break
                elif "Request timed out" in line:
                    ping_output = "Request timed out"
                    break
                elif "could not find host" in line.lower():
                    ping_output = "Could not find host"
                    break

            treeview.item(item, values=(ip_address, ping_output))

            # Sleep for the selected ping interval
            ping_interval = int(ping_interval_combobox.get())
            time.sleep(ping_interval)

    multi_ping_window = tk.Toplevel(window)
    multi_ping_window.title("Multi-Ping Tool")

    ip_entry = ttk.Entry(multi_ping_window)
    ip_entry.pack(padx=10, pady=10)

    button_frame = ttk.Frame(multi_ping_window)
    button_frame.pack(padx=10, pady=10)

    add_button = ttk.Button(button_frame, text="Add", command=add_ping)
    add_button.grid(row=0, column=0, padx=5, pady=5)

    remove_button = ttk.Button(button_frame, text="Remove", command=remove_ping)
    remove_button.grid(row=0, column=1, padx=5, pady=5)

    start_button = ttk.Button(button_frame, text="Start", command=start_multi_ping)
    start_button.grid(row=0, column=2, padx=5, pady=5)

    stop_button = ttk.Button(button_frame, text="Stop", command=stop_multi_ping)
    stop_button.grid(row=0, column=3, padx=5, pady=5)

    # Add the ping interval dropdown box
    ping_interval_label = ttk.Label(multi_ping_window, text="Ping Interval (seconds):")
    ping_interval_label.pack(padx=10, pady=(10, 0))

    ping_interval_combobox = ttk.Combobox(multi_ping_window, values=list(range(1, 61)), state="readonly")
    ping_interval_combobox.set(5)
    ping_interval_combobox.pack(padx=10, pady=(0, 10))

    treeview = ttk.Treeview(multi_ping_window, columns=("IP", "Ping Output"), show="headings")
    treeview.heading("IP", text="IP Address")
    treeview.heading("Ping Output", text="Ping Output")
    treeview.pack(padx=10, pady=10, fill="both", expand=True)

    scrollbar = ttk.Scrollbar(multi_ping_window, orient="vertical", command=treeview.yview)
    scrollbar.pack(side="right", fill="y")
    treeview.configure(yscrollcommand=scrollbar.set)

# Create a global variable to control the pinging process
stop_pinging = False

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

#Icons to buttons
restart_icon = PhotoImage(file="icons/power.png")
show_temp_icon = PhotoImage(file="icons/temp.png")
clear_temp_icon = PhotoImage(file="icons/delete.png")
remote_support_icon = PhotoImage(file="icons/web.png")
src_tools_icon = PhotoImage(file="icons/web.png")
src_clear_icon = PhotoImage(file="icons/clear.png")
src_ip_icon = PhotoImage(file="icons/ip.png")
src_log_icon = PhotoImage(file="icons/log.png")
stop_icon = PhotoImage(file="icons/stop.png")


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
menu_bar.add_command(label="Open Multi-Ping", command=lambda: open_multi_ping_window())


# Modify the button_frame layout
button_frame = ttk.Frame(window)
button_frame.grid(column=0, row=0, padx=20, pady=10)

button1 = ttk.Button(button_frame, text="Restart Computer", image=restart_icon, compound=tk.LEFT, command=lambda: run_command("shutdown /r /t 5"), style="Fancy.TButton")
button2 = ttk.Button(button_frame, text="Show Temp Directory", image=show_temp_icon, compound=tk.LEFT, command=lambda: run_command("dir %temp%\\*"), style="Fancy.TButton")
button3 = ttk.Button(button_frame, text="Clear Temp Files", image=clear_temp_icon, compound=tk.LEFT, command=lambda: run_command("del /s /q /f %temp%\\*"), style="Fancy.TButton")

support_button = ttk.Button(button_frame, text="Remote Support", image=remote_support_icon, compound=tk.LEFT, command=lambda: webbrowser.open("https://itbysrc.com/remote-support/"), style="Fancy.TButton")
tools_button = ttk.Button(button_frame, text="SRC Tools Download", image=src_tools_icon, compound=tk.LEFT, command=lambda: webbrowser.open("https://itbysrc.com/agent/index.php?dir=%21%20SRC%20Tools/"), style="Fancy.TButton")

output_text = tk.Text(window, height=10, state=tk.DISABLED)
scrollbar = tk.Scrollbar(window, command=output_text.yview)
output_text.config(yscrollcommand=scrollbar.set)

clear_button = ttk.Button(window, text="Clear Output",image=src_clear_icon, command=lambda: clear_output(), style="Fancy.TButton")

def clear_output():
    output_text.config(state=tk.NORMAL)
    output_text.delete("1.0", tk.END)
    output_text.config(state=tk.DISABLED)
    scrollbar.update()

button1.grid(row=0, column=0, padx=5, pady=5)
button2.grid(row=1, column=0, padx=5, pady=5)
button3.grid(row=2, column=0, padx=5, pady=5)
support_button.grid(row=3, column=0, padx=5, pady=5)
tools_button.grid(row=4, column=0, padx=5, pady=5)

output_text.grid(column=1, row=0, padx=10, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
scrollbar.grid(column=2, row=0, padx=10, pady=10, sticky=(tk.N, tk.S))
clear_button.grid(column=1, row=1, padx=10, pady=5, sticky=(tk.W, tk.E))

# Ping Frame / Ping Entry
ping_frame = ttk.Frame(window)
ping_frame.grid(column=0, row=2, padx=20, pady=5)

ip_entry = ttk.Entry(ping_frame)
ip_entry.grid(row=0, column=0, padx=5, pady=5)
ip_entry.insert(0, "google.com")  # Set a default value

ping_button = ttk.Button(ping_frame, text="Ping IP", image=src_ip_icon,compound=tk.LEFT , command=lambda: ping_google(ip_entry.get()), style="Fancy.TButton")
ping_button.grid(row=1, column=0, padx=5, pady=5)

download_log_button = ttk.Button(ping_frame, text="Open Ping Log", image=src_log_icon, compound=tk.LEFT, command=download_ping_log, style="Fancy.TButton")
download_log_button.grid(row=2, column=0, padx=5, pady=5)

stop_button = ttk.Button(window, text="Stop Ping", image=stop_icon, compound=tk.LEFT, command=stop_ping, style="Fancy.TButton")
stop_button.grid(column=1, row=2, padx=10, pady=5, sticky=(tk.W, tk.E))

ping_output_label = ttk.Label(window, textvariable=ping_output_var, anchor=tk.W, justify=tk.LEFT)
ping_output_label.grid(column=1, row=3, padx=10, pady=10, sticky=(tk.W, tk.E))

window.columnconfigure(1, weight=1)
window.rowconfigure(0, weight=1)

window.mainloop()
