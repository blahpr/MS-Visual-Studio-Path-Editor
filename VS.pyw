import tkinter as tk
from tkinter import messagebox
import winreg
import subprocess
import os
import re
from collections import defaultdict
import shutil
import webbrowser
import sys
import ctypes


# Function to check if the script is running as admin
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Utility function to handle the icon path in both development and bundled exe
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    try:
        base_path = sys._MEIPASS  # PyInstaller stores the temp path here
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def update_registry():
    try:
        new_drive_letter = drive_entry.get().upper()  # Get and uppercase the drive letter
        new_installation_path = fr"{new_drive_letter}:\Program Files (x86)\Microsoft Visual Studio\Shared"

        # Open the registry key
        key_path = r"SOFTWARE\Microsoft\VisualStudio\Setup"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE)

        # Set the new value for SharedInstallationPath
        winreg.SetValueEx(key, "SharedInstallationPath", 0, winreg.REG_SZ, new_installation_path)

        # Close the registry key
        winreg.CloseKey(key)

        # Show the changes in the label
        show_changes(new_drive_letter, new_installation_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def show_current_registry():
    try:
        # Open the registry key
        key_path = r"SOFTWARE\Microsoft\VisualStudio\Setup"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ)

        # Get the current value for SharedInstallationPath
        current_installation_path, _ = winreg.QueryValueEx(key, "SharedInstallationPath")

        # Close the registry key
        winreg.CloseKey(key)

        # Extract the current drive letter
        current_drive_letter = current_installation_path[0] if current_installation_path else ""

        # Update the label with the current contents
        current_label.config(text=f"Registry:\nComputer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\VisualStudio\\Setup\n\nDrive Letter: {current_drive_letter}\n {current_installation_path}")
    except FileNotFoundError:
        # Handle the case where the registry key is not found
        messagebox.showwarning("Not Found", 
            "Microsoft Visual Studio not found on this system. Please download Microsoft Visual Studio from the following link:\n\nhttps://visualstudio.microsoft.com/downloads/")
        current_label.config(text="Registry:\nMicrosoft Visual Studio not found.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def show_changes(new_drive_letter, new_path):
    # Show the changes made to the registry
    changes_label.config(text=f"Changes Made: {new_drive_letter}\n {new_path}")

def refresh_registry():
    try:
        # Refresh the displayed current contents of the registry
        show_current_registry()
    except Exception as e:
        messagebox.showerror("Error", str(e))

def open_registry_editor():
    try:
        # Open the Registry Editor and navigate to the specified registry key
        key_path = r"SOFTWARE\Microsoft\VisualStudio\Setup"
        
        # On Windows, use the 'start' command to open the Registry Editor at a specific key
        os.system(f'start regedit')
    except Exception as e:
        messagebox.showerror("Error", str(e))

def open_github():
    webbrowser.open("https://github.com/blahpr/MS-Visual-Studio-Path-Editor")

def show_about():
    # Create a new window for the About dialog
    about_window = tk.Toplevel(root)
    about_window.title("About MS Visual Studio Path Editor")

    # Set the window size and icon
    about_window.geometry("390x245")
    about_window.iconbitmap(resource_path('images/R.ico'))

    # Add a label with the about text
    about_text = "This Application Allows you to Change the Drive Letter for the Shared Installation Path or Visual Studio 2022: Shared components, tools, and SDK's in the Windows Registry for Microsoft Visual Studio 2022.\n\nMade With Python & GPT.\n\nContact: geebob273@gmail.com\n\nhttps://github.com/blahpr/MS-Visual-Studio-Path-Editor\n\n© 2024 Blahp Software"
    about_label = tk.Label(about_window, text=about_text, wraplength=350, justify="left")
    about_label.pack(pady=10)

    # Add a GitHub button that opens the GitHub URL
    github_button = tk.Button(about_window, text="GitHub", command=open_github)
    github_button.pack(pady=0)


# Function to attempt to restart the script with admin privileges
def request_admin_privileges():
    # Try to re-launch the script as an admin
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit()

# Create the main Tkinter window
root = tk.Tk()
root.title("MS Visual Studio Path Editor")

# Set the window icon
root.iconbitmap(resource_path('images/R.ico'))

# Create a menu bar
menubar = tk.Menu(root)
root.config(menu=menubar)

# Create a top-level menu button
file_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="File", menu=file_menu)

# Add an "About" menu item
file_menu.add_command(label="About", command=show_about)

# Create a label and entry for the new drive letter
drive_label = tk.Label(root, text="Change Drive Letter:\n Visual Studio 2022: Shared components, tools, and SDK's:")
drive_label.pack(pady=5)

drive_entry = tk.Entry(root)
drive_entry.pack(pady=5)

# Create a button to update the registry (this will require admin privileges)
def update_registry_with_admin():
    if not is_admin():
        request_admin_privileges()
    update_registry()

update_button = tk.Button(root, text="Add\\Update Registry", command=update_registry_with_admin)
update_button.pack(pady=10)

# Create a button to refresh the registry display
refresh_button = tk.Button(root, text="Refresh", command=refresh_registry)
refresh_button.pack(pady=10)

# Create a button to open the Registry Editor
open_editor_button = tk.Button(root, text="Open Registry Editor", command=open_registry_editor)
open_editor_button.pack(pady=10)

# Create a label to show the current registry information
current_label = tk.Label(root, text="Registry:\n")
current_label.pack(pady=10)

# Create a label to show the changes made
changes_label = tk.Label(root, text="\n")
changes_label.pack(pady=10)

# Show the current registry information when the application starts
show_current_registry()

# Run the Tkinter main loop
root.mainloop()
