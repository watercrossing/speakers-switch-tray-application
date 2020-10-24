# Switch IoT Speakers from the task bar

My monitor speakers do not have an accessible on/off switch, so I have them wired to a WiFi smart plug (which I flashed with [Tasmota](https://tasmota.github.io/docs/)). This application creates a tray icon that allows me to switch the speakers on and off. It also monitors the audio output, and if no sound output is detected for a while the speakers will be switched off. 

## Installation

 * create a virtual environment (name it `venv` for the `run-in-background.bat` to work)
 * install the dependencies (`pip -r requirements.txt`)
 * copy `config-default.ini` to `config.ini`, adjusting at least the `BASEURL`
 * run `trayapp.py`, and potentially link `run-in-background.bat` to your `shell:startup` directory.

