print("Portable Application Installer")
print("Written by Stephen-Hamilton-C. Licensed under GPLv3.")
print("Find the source code at https://github.com/Stephen-Hamilton-C/portable-application-installer")
print()

# Platform check
import platform, sys
if platform.system() != "Windows":
    print("This script will only work on Windows!")
    sys.exit(1)

import winreg
import os
import shutil
import ctypes
import subprocess

CLI_MODE = len(sys.argv) > 1 and sys.argv[1].lower() == "nogui"
LOCAL_APPDATA = os.getenv("LOCALAPPDATA")
PROGRAM_FILES = os.getenv("PROGRAMFILES")

def install_package(pkg: str) -> bool:
    print(pkg+" is not installed, which is required for this script to run.")
    response = input("Would you like to install "+pkg+"? (Y/n): ").lower()
    if len(response) < 1 or response[0] != "n":
        print("Installing "+pkg+" with pip3...")
        try:
            subprocess.run(["pip3", "install", pkg], check=True)
        except:
            print("Error using pip3, trying pip instead...")
            try:
                subprocess.run(["pip", "install", pkg], check=True)
            except:
                print("Could not install "+pkg+".")
                return False
    else:
        print(pkg+" will not be installed.")
        return False
    return True

# Try to import pywin32
try:
    from win32com.client import Dispatch
except ImportError:
    if install_package("pywin32"):
        from win32com.client import Dispatch
    else:
        print("Unable to continue without pywin32. Exiting...")
        sys.exit(1)

def get_path(prompt: str, file: bool) -> str:
    prompt_result = input(prompt).replace("\"", "")

    while not os.path.exists(prompt_result) and ((file and not os.path.isfile(prompt_result)) or (not file and not os.path.isdir(prompt_result))):
        if file:
            print("That file does not exist or is not a file!")
        else:
            print("That directory does not exist or is not a directory!")
        prompt_result = input(prompt).replace("\"", "")

    return prompt_result

def get_app_name_and_install_dir(prompt: str, install_mode: str):
    global PROGRAM_FILES
    global LOCAL_APPDATA

    app_name = input(prompt)
    if install_mode == "a":
        install_dir = os.path.join(PROGRAM_FILES, app_name)
    else:
        install_dir = os.path.join(LOCAL_APPDATA, app_name)
    while os.path.exists(install_dir) and os.listdir(install_dir) != []:
        print("Name already taken! Try a different name.")
        app_name = input(prompt)
        if install_mode == "a":
            install_dir = os.path.join(PROGRAM_FILES, app_name)
        else:
            install_dir = os.path.join(LOCAL_APPDATA, app_name)
    return (app_name, install_dir)

def get_install_mode() -> str:
    while install_mode != "u" and install_mode != "a":
        install_mode = input("Should this be installed for all users (a) or just your current user (u)?: ").lower()[0]
    return install_mode

def run(app_name: str, install_dir: str, exe_path: str, install_mode: str, copy_all: str):
    appdata = os.getenv("APPDATA")
    program_data = os.getenv("PROGRAMDATA")
    exe_dir = os.path.dirname(exe_path)
    exe_name = os.path.basename(exe_path)
    START_MENU_APPPATH = os.path.join(program_data if install_mode == "a" else appdata, "Microsoft", "Windows", "Start Menu", "Programs")
    installed_exe_path = os.path.join(install_dir, exe_name)



    ### Registry keys
    REGKEY_APPPATHS = "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\"+app_name
    REGKEY_UNINSTALL = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"+app_name



    ### Copy files to install directory
    if len(copy_all) < 1 or copy_all[0] == "n":
        print("Copying portable application to \""+install_dir+"\"...")
        os.makedirs(install_dir, exist_ok=True)
        shutil.copy2(exe_path, installed_exe_path)
    else:
        print("Copying all files in \""+exe_dir+"\" to \""+install_dir+"\"...")
        try:
            shutil.rmtree(install_dir)
        except FileNotFoundError:
            pass
        shutil.copytree(exe_dir, install_dir)



    ### Create start menu shortcut
    print("Creating start menu shortcut...")
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(os.path.join(START_MENU_APPPATH, app_name+".lnk"))
    shortcut.Targetpath = installed_exe_path
    shortcut.WorkingDirectory = install_dir
    shortcut.IconLocation = installed_exe_path
    shortcut.save()



    ### Create uninstaller
    print("Copying uninstaller...")
    with open(os.path.join(install_dir, "uninstaller-data"), "w") as dataFile:
        dataFile.write("# DO NOT MODIFY THIS FILE!\n")
        dataFile.write("# ANY MODIFICATION COULD RESULT IN BREAKING WINDOWS REGISTRY!\n")
        dataFile.write(app_name+"\n")
        dataFile.write(install_mode+"\n")
    print("Compiling uninstaller to exe...")
    shutil.copy("uninstaller.pyw", install_dir)
    shutil.copy("uninstaller.bat", install_dir)



    ### Write Registry keys
    # Registry keys to write:
    # REGKEY_APPPATHS "" "INSTALL_DIR\\EXECUTABLE_NAME"
    # REGKEY_UNINSTALL "DisplayName" APP_NAME
    # TODO: Get version from file
    # REGKEY_UNINSTALL "DisplayVersion" "Portable"
    # REGKEY_UNINSTALL "UninstallString" "INSTALL_DIR\\uninstaller.pyw"
    # REGKEY_UNINSTALL "DisplayIcon" "INSTALL_DIR\\EXECUTABLE_NAME"
    print("Writing registry keys...")
    key_type = winreg.HKEY_LOCAL_MACHINE if install_mode == "a" else winreg.HKEY_CURRENT_USER
    try:
        apppath_key = winreg.OpenKey(key_type, REGKEY_APPPATHS, 0, winreg.KEY_WRITE)
    except:
        apppath_key = winreg.CreateKey(key_type, REGKEY_APPPATHS)
    winreg.SetValueEx(apppath_key, "", 0, winreg.REG_SZ, installed_exe_path)
    winreg.SetValueEx(apppath_key, "Path", 0, winreg.REG_SZ, install_dir)
    winreg.CloseKey(apppath_key)
    try:
        uninstall_key = winreg.OpenKey(key_type, REGKEY_UNINSTALL, 0, winreg.KEY_WRITE)
    except:
        uninstall_key = winreg.CreateKey(key_type, REGKEY_UNINSTALL)
    winreg.SetValueEx(uninstall_key, "DisplayName", 0, winreg.REG_SZ, app_name)
    winreg.SetValueEx(uninstall_key, "DisplayVersion", 0, winreg.REG_SZ, "Portable")
    winreg.SetValueEx(uninstall_key, "UninstallString", 0, winreg.REG_SZ, os.path.join(install_dir, "uninstaller.bat"))
    winreg.SetValueEx(uninstall_key, "DisplayIcon", 0, winreg.REG_SZ, installed_exe_path)
    winreg.CloseKey(uninstall_key)



### GUI
if not CLI_MODE:
    # Import tkinter
    try:
        import tkinter as tk
        from tkinter import messagebox
        from tkinter.filedialog import askopenfilename
    except ImportError:
        if install_package("tkinter"):
            import tkinter as tk
            from tkinter import messagebox
            from tkinter.filedialog import askopenfilename
        else:
            print("No GUI available. Run with 'nogui' as an argument to enable CLI mode.")
            sys.exit(1)

    # init window
    window = tk.Tk()
    window.title("Portable Application Installer")
    font = "helvetica 12"

    # Header
    title = tk.Label(window, text="Portable Application Installer", font=font)
    title.grid(column=1, row=0)
    author = tk.Label(window, text="Written by Stephen-Hamilton-C. Licensed under GPLv3.", font=font)
    author.grid(column=1, row=10)
    source = tk.Label(window, text="Find the source code at https://github.com/Stephen-Hamilton-C/portable-application-installer", fg="blue", cursor="hand2", font=font+" underline")
    source.bind("<Button-1>", lambda e: os.startfile("https://github.com/Stephen-Hamilton-C/portable-application-installer"))
    source.grid(column=1, row=20)

    # Path to exe
    def click_exe_browse():
        filename = askopenfilename()
        if filename != "":
            exe_input.delete("0.0", tk.END)
            exe_input.insert("0.0", filename)
    exe_input_label = tk.Label(window, text="Path to Application:", font=font)
    exe_input = tk.Text(window, height=1, font=font)
    exe_input.insert("0.0", __file__)
    exe_input_browse = tk.Button(window, text="...", command = click_exe_browse, font=font)
    exe_input_label.grid(column=0, row=40, padx=6, pady=6)
    exe_input.grid(column=1, row=40, pady=6)
    exe_input_browse.grid(column=2, row=40, padx=6, ipadx=12, pady=6)
    
    # App name
    app_name_label = tk.Label(window, text="Application Name:", font=font)
    app_name_input = tk.Text(window, height=1, font=font)
    app_name_label.grid(column=0, row=50, padx=6, pady=6)
    app_name_input.grid(column=1, row=50, pady=6)

    # Install mode
    install_mode = tk.StringVar(window, "u")
    system_radio_button = tk.Radiobutton(window, text="Install for all users", value="a", variable=install_mode, font=font)
    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        system_radio_button.config(state=tk.DISABLED, text="Install for all users (Run as admin to do this)")
    user_radio_button = tk.Radiobutton(window, text="Install just for this user", value="u", variable=install_mode, font=font)
    system_radio_button.grid(column=0, row=60)
    user_radio_button.grid(column=0, row=70)

    # Copy all
    copy_all = tk.StringVar(window, "y")
    copy_all_checkbox = tk.Checkbutton(window, text="Copy all files in application folder?", variable=copy_all, onvalue="y", offvalue="n", font=font)
    copy_all_checkbox.grid(column=0, row=80)

    # Run button
    def click_install():
        global app_name_input
        global app_name_label
        global exe_input
        global install_button
        global copy_all
        global install_mode

        install_button.config(state=tk.DISABLED, text="Installing...")

        app_name = app_name_input.get(1.0, "end-1c")
        exe_path = exe_input.get(1.0, "end-1c")
        install_root = PROGRAM_FILES if install_mode.get() == "a" else LOCAL_APPDATA
        install_dir = os.path.join(install_root, app_name)

        # Validate inputs
        invalid = False
        if app_name.strip() == "" or (os.path.exists(install_dir) and os.listdir(install_dir) != []):
            app_name_input.config(highlightcolor="red", highlightbackground="red", highlightthickness=1)
            if app_name.strip() != "":
                app_name_label.config(text="Application Name (name already taken!):")
            invalid = True
        else:
            app_name_input.config(highlightthickness=0)
            app_name_label.config(text="Application Name:")

        if not os.path.exists(exe_path):
            exe_input.config(highlightcolor="red", highlightbackground="red", highlightthickness=1)
            invalid = True
        else:
            exe_input.config(highlightthickness=0)

        # Cancel if invalid
        if invalid:
            install_button.config(state=tk.NORMAL, text="Install")
            return

        if os.path.dirname(exe_path).lower().endswith("downloads"):
            copy_all.set("n")

        # Run
        try:
            run(app_name, install_dir, exe_path, install_mode.get(), copy_all.get())
            messagebox.showinfo("Success!", app_name+" finished installing successfully. You may run the application from the start menu and uninstall it like any other Windows application.")
            sys.exit(0)
        except Exception as e:
            print(e)
            messagebox.showerror("Error", "An unknown error occurred while installing! Try running from a command prompt to see the error in more detail.")
            install_button.config(state=tk.NORMAL, text="Install")
    install_button = tk.Button(window, text="Install", font=font, command=click_install)
    install_button.grid(column=1, row=90, pady=12, ipadx=60)
    
    # Launch
    try:
        # Attempt to make text not blurry
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    finally:
        window.mainloop()



### CLI
if CLI_MODE:
    exe_path = get_path("What is the path to the executable you wish to install?: ", True)
    if ctypes.windll.shell32.IsUserAnAdmin() != 0:
        install_mode = get_install_mode()
    else:
        print("No admin privileges detected. This application will be installed to the current user.")
        install_mode = "u"
    app_name_and_install_dir = get_app_name_and_install_dir("What is the name of this application?: ", install_mode)
    app_name = app_name_and_install_dir[0]
    install_dir = app_name_and_install_dir[1]
    if os.path.dirname(exe_path).lower().endswith("downloads"):
        copy_all = "n"
    else:
        copy_all = input("Should all files with the application be installed? (y/N): ").lower()

    try:
        run(app_name, install_dir, exe_path, install_mode, copy_all)
    
        print(app_name+" has been successfully installed! You can run the application from the start menu.")
        print("To uninstall the program, just uninstall it like any other Windows application.")
    except Exception as e:
        print(e)
        print("Failed to install! See above for error stacktrace.")
        print("Perhaps you should submit an issue to https://github.com/Stephen-Hamilton-C/portable-application-installer/issues")
    
    input("Press enter to exit...")
