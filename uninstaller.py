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


APP_NAME = None
INSTALL_MODE = None

with open("uninstaller-data", "r") as dataFile:
    # First two lines are a warning
    dataFile.readline()
    dataFile.readline()

    APP_NAME = dataFile.readline()
    INSTALL_MODE = dataFile.readline()

if APP_NAME == None or INSTALL_MODE == None:
    print("CRITICAL! Uninstaller data could not be read! This application cannot be uninstalled by Portable Application Uninstaller.")
    print("You will have to manually remove the files.")
    sys.exit(1)

# Get registry keys from APP_NAME
REGKEY_APPPATHS = "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\"+APP_NAME
REGKEY_UNINSTALL = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"+APP_NAME

# TODO: Remove registry keys
# TODO: Remove shortcut from start menu
# TODO: Remove files from directory

