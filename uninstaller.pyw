print("Portable Application Uninstaller")
print("Written by Stephen-Hamilton-C. Licensed under GPLv3.")
print("Find the source code at https://github.com/Stephen-Hamilton-C/portable-application-installer")
print("Please note that Portable Application Installer is not fully aware of all files that may have been created by the application.")
print("You may need to manually remove files on your own, such as configuration or data files.")
print()

# Platform check
import platform, sys
if platform.system() != "Windows":
    print("This script will only work on Windows!")
    sys.exit(1)

import os
import winreg
import shutil

CLI_MODE = len(sys.argv) > 1 and sys.argv[1].lower() == "nogui"
APP_NAME = None
INSTALL_MODE = None

def show_no_data_error():
    if CLI_MODE:
        print("CRITICAL! Uninstaller data could not be read! This application cannot be uninstalled by Portable Application Uninstaller.")
        print("You will have to manually remove the files.")
        input("Press enter to exit...")
        sys.exit(1)
    else:
        # show error with gui
        from tkinter import messagebox
        messagebox.showerror("Critical Error", "Uninstaller data could not be read! This application cannot be uninstalled by Portable Application Uninstaller. You will have to manually remove the files.")
        sys.exit(1)

try:
    with open("uninstaller-data", "r") as dataFile:
        # First two lines are a warning
        dataFile.readline()
        dataFile.readline()

        APP_NAME = dataFile.readline().strip()
        INSTALL_MODE = dataFile.readline().strip()
except Exception:
    show_no_data_error()

if APP_NAME == None or APP_NAME == "" or INSTALL_MODE == None or INSTALL_MODE == "":
    show_no_data_error()

def run():
    REGKEY_APPPATHS = "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\"+APP_NAME
    REGKEY_UNINSTALL = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"+APP_NAME
    PROGRAMDATA = os.getenv("PROGRAMDATA")
    APPDATA = os.getenv("APPDATA")
    LOCALAPPDATA = os.getenv("LOCALAPPDATA")
    PROGRAMFILES = os.getenv("PROGRAMFILES")
    START_MENU_APPPATH = os.path.join(PROGRAMDATA if INSTALL_MODE == "a" else APPDATA, "Microsoft", "Windows", "Start Menu", "Programs")

    # Remove registry keys
    print("Removing registry entries...")
    key_type = winreg.HKEY_LOCAL_MACHINE if INSTALL_MODE == "a" else winreg.HKEY_CURRENT_USER
    winreg.DeleteKey(key_type, REGKEY_APPPATHS)
    winreg.DeleteKey(key_type, REGKEY_UNINSTALL)

    # Remove shortcut from start menu
    print("Removing start menu shortcut...")
    os.remove(os.path.join(START_MENU_APPPATH, APP_NAME+".lnk"))

    # Remove files from directory
    print("Removing files from install directory...")
    INSTALL_DIR = PROGRAMFILES if INSTALL_MODE == "a" else LOCALAPPDATA
    try:
        shutil.rmtree(os.path.join(INSTALL_DIR, APP_NAME))
    except:
        pass # Will throw because can't remove directory


if CLI_MODE:
    try:
        run()

        print("Finished uninstalling.")
    except Exception as e:
        print(e)
        print("Failed to uninstall! See above for error stacktrace.")
        print("Perhaps you should submit an issue to https://github.com/Stephen-Hamilton-C/portable-application-installer/issues")

    input("Press enter to exit...")
else:
    # launch gui
    import tkinter as tk
    from tkinter import messagebox

    window = tk.Tk()
    window.title("Portable Application Uninstaller")
    font = "helvetica 12"

    # Header
    title = tk.Label(window, text="Portable Application Uninstaller", font=font)
    title.pack()
    author = tk.Label(window, text="Written by Stephen-Hamilton-C. Licensed under GPLv3.", font=font)
    author.pack()
    source = tk.Label(window, text="Find the source code at https://github.com/Stephen-Hamilton-C/portable-application-installer", fg="blue", cursor="hand2", font=font+" underline")
    source.bind("<Button-1>", lambda e: os.startfile("https://github.com/Stephen-Hamilton-C/portable-application-installer"))
    source.pack()

    # Body
    body = tk.Label(window, font=font, text="This application will attempt to uninstall "+APP_NAME+" from your Windows PC.\nPress the button below to start. Please note that Portable Application Uninstaller\nis not responsible for any left behind files\nsuch as config or data files created by the application.")
    body.pack(pady=24)

    # Uninstall button
    def click_uninstall():
        global uninstall_button
        try:
            uninstall_button.config(state=tk.DISABLED, text="Uninstalling...")
            run()
            messagebox.showinfo("Success!", "Finished uninstalling.")
            sys.exit(0)
        except Exception:
            messagebox.showerror("Error", "An unexpected error occurred while uninstalling. Perhaps you should submit an issue to https://github.com/Stephen-Hamilton-C/portable-application-installer/issues")
            sys.exit(1)
    uninstall_button = tk.Button(window, font=font, text="Uninstall", command=click_uninstall)
    uninstall_button.pack(ipadx=60, pady=12)

    # Launch
    try:
        # Attempt to make text not blurry
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    finally:
        window.mainloop()


