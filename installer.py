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

# Try to import pywin32
try:
    from win32com.client import Dispatch
except ImportError:
    print("pywin32 is not installed, which is required for this script to run.")
    response = input("Would you like to install pywin32? (Y/n): ").lower()
    if len(response) < 1 or response[0] != "n":
        print("Installing pywin32 with pip3...")
        try:
            subprocess.run(["pip3", "install", "pywin32"], check=True)
        except:
            print("Error using pip3, trying pip instead...")
            try:
                subprocess.run(["pip", "install", "pywin32"], check=True)
            except:
                print("Could not install pywin32. Exiting...")
                sys.exit(1)
        from win32com.client import Dispatch
    else:
        print("pywin32 will not be installed. Exiting...")
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

def get_app_name_and_install_dir(prompt: str):
    global INSTALL_MODE
    global PROGRAMFILES
    global LOCALAPPDATA

    app_name = input(prompt)
    if INSTALL_MODE == "a":
        install_dir = os.path.join(PROGRAMFILES, app_name)
    else:
        install_dir = os.path.join(LOCALAPPDATA, app_name)
    while os.path.exists(install_dir) and os.listdir(install_dir) != []:
        print("Name already taken! Try a different name.")
        app_name = input(prompt)
        if INSTALL_MODE == "a":
            install_dir = os.path.join(PROGRAMFILES, app_name)
        else:
            install_dir = os.path.join(LOCALAPPDATA, app_name)
    return (app_name, install_dir)

def get_install_mode() -> str:
    while install_mode != "u" and install_mode != "a":
        install_mode = input("Should this be installed for all users (a) or just your current user (u)?: ").lower()[0]
    return install_mode

### Environment variables
PROGRAMFILES = os.getenv("PROGRAMFILES")
APPDATA = os.getenv("APPDATA")
LOCALAPPDATA = os.getenv("LOCALAPPDATA")
PROGRAMDATA = os.getenv("PROGRAMDATA")


### User input
if ctypes.windll.shell32.IsUserAnAdmin() != 0:
    INSTALL_MODE = get_install_mode()
else:
    print("No admin privileges detected. This application will be installed to the current user.")
    INSTALL_MODE = "u"
EXECUTABLE_PATH = get_path("What is the path to the executable you wish to install?: ", True)
EXECUTABLE_DIR = os.path.dirname(EXECUTABLE_PATH)
EXECUTABLE_NAME = os.path.basename(EXECUTABLE_PATH)
START_MENU_APPPATH = os.path.join(PROGRAMDATA if INSTALL_MODE == "a" else APPDATA, "Microsoft", "Windows", "Start Menu", "Programs")

APP_NAME_AND_INSTALL_DIR = get_app_name_and_install_dir("What is the name of this application?: ")
APP_NAME = APP_NAME_AND_INSTALL_DIR[0]
INSTALL_DIR = APP_NAME_AND_INSTALL_DIR[1]
INSTALLED_EXE_PATH = os.path.join(INSTALL_DIR, EXECUTABLE_NAME)



### Registry keys
REGKEY_APPPATHS = "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\"+APP_NAME
REGKEY_UNINSTALL = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"+APP_NAME



### Copy files to install directory
if os.path.dirname(EXECUTABLE_DIR).lower() != "downloads":
    copy_all = input("Should all files with the application be installed? (y/N): ").lower()
else:
    copy_all = "n"

if len(copy_all) < 1 or copy_all[0] == "n":
    print("Copying portable application to \""+INSTALL_DIR+"\"...")
    os.makedirs(INSTALL_DIR)
    shutil.copy2(EXECUTABLE_PATH, INSTALLED_EXE_PATH)
else:
    print("Copying all files in \""+EXECUTABLE_DIR+"\" to \""+INSTALL_DIR+"\"...")
    try:
        shutil.rmtree(INSTALL_DIR)
    except FileNotFoundError:
        pass
    shutil.copytree(EXECUTABLE_DIR, INSTALL_DIR)



### Create start menu shortcut
print("Creating start menu shortcut...")
shell = Dispatch("WScript.Shell")
shortcut = shell.CreateShortcut(os.path.join(START_MENU_APPPATH, APP_NAME+".lnk"))
shortcut.Targetpath = INSTALLED_EXE_PATH
shortcut.WorkingDirectory = INSTALL_DIR
shortcut.IconLocation = INSTALLED_EXE_PATH
shortcut.save()



### Create uninstaller
print("Copying uninstaller...")
with open(os.path.join(INSTALL_DIR, "uninstaller-data"), "w") as dataFile:
    dataFile.write("# DO NOT MODIFY THIS FILE!\n")
    dataFile.write("# ANY MODIFICATION COULD RESULT IN BREAKING WINDOWS REGISTRY!\n")
    dataFile.write(APP_NAME+"\n")
    dataFile.write(INSTALL_MODE+"\n")
print("Compiling uninstaller to exe...")
shutil.copy("uninstaller.py", INSTALL_DIR)
shutil.copy("uninstaller.bat", INSTALL_DIR)



### Write Registry keys
# Registry keys to write:
# REGKEY_APPPATHS "" "INSTALL_DIR\\EXECUTABLE_NAME"
# REGKEY_UNINSTALL "DisplayName" APP_NAME
# TODO: Get version from file
# REGKEY_UNINSTALL "DisplayVersion" "Portable"
# REGKEY_UNINSTALL "UninstallString" "INSTALL_DIR\\uninstaller.py"
# REGKEY_UNINSTALL "DisplayIcon" "INSTALL_DIR\\EXECUTABLE_NAME"
print("Writing registry keys...")
key_type = winreg.HKEY_LOCAL_MACHINE if INSTALL_MODE == "a" else winreg.HKEY_CURRENT_USER
try:
    apppath_key = winreg.OpenKey(key_type, REGKEY_APPPATHS, 0, winreg.KEY_WRITE)
except:
    apppath_key = winreg.CreateKey(key_type, REGKEY_APPPATHS)
winreg.SetValueEx(apppath_key, "", 0, winreg.REG_SZ, INSTALLED_EXE_PATH)
winreg.SetValueEx(apppath_key, "Path", 0, winreg.REG_SZ, INSTALL_DIR)
winreg.CloseKey(apppath_key)
try:
    uninstall_key = winreg.OpenKey(key_type, REGKEY_UNINSTALL, 0, winreg.KEY_WRITE)
except:
    uninstall_key = winreg.CreateKey(key_type, REGKEY_UNINSTALL)
winreg.SetValueEx(uninstall_key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
winreg.SetValueEx(uninstall_key, "DisplayVersion", 0, winreg.REG_SZ, "Portable")
winreg.SetValueEx(uninstall_key, "UninstallString", 0, winreg.REG_SZ, os.path.join(INSTALL_DIR, "uninstaller.bat"))
winreg.SetValueEx(uninstall_key, "DisplayIcon", 0, winreg.REG_SZ, INSTALLED_EXE_PATH)
winreg.CloseKey(uninstall_key)

print(APP_NAME+" has been successfully installed! You can run the application from the start menu.")
print("To uninstall the program, just uninstall it like any other Windows application.")
enter = input("Press enter to exit...")
