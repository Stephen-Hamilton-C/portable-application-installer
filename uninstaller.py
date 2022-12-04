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

APP_NAME = None
INSTALL_MODE = None

with open("uninstaller-data", "r") as dataFile:
    # First two lines are a warning
    dataFile.readline()
    dataFile.readline()

    APP_NAME = dataFile.readline().strip()
    INSTALL_MODE = dataFile.readline().strip()

if APP_NAME == None or INSTALL_MODE == None:
    print("CRITICAL! Uninstaller data could not be read! This application cannot be uninstalled by Portable Application Uninstaller.")
    print("You will have to manually remove the files.")
    sys.exit(1)

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

print("Finished uninstalling.")
enter = input("Press enter to exit...")
