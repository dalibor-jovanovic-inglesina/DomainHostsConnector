# Domain Hosts Connector
An application that permits you to remote connect to Domain endpoints (Servers, Computers and Notebooks) with RDP or UltraVNC for now and also permit to gather information data as Serial Number, Manufacturer or TPM Version that you can use for an invetory system.

# Requirements
This application, to run, requires: pywin32, PyQt6, PyQt6_sip.\
Use the command:\
python -m pip install -r requirements.txt\
to get the needed libraries.

# Important
In the main.py file there is a excluded_devices array variable that, in case you want to exclude some devices from getting scanned, must be populated.

# Execution
You can execute the application with pyhton3 main.py.

# Compile For Windows
You can compile this program with pyinstaller.\
\
Install pyinstaller with pip:\
pip install pyinstaller\
or\
python3 -m pip install pyinstaller\
\
Run this command to compile main.py:\
pyinstaller --onefile --windowed main.py\
(this will create an .exe file under the "dist" folder)
