# Switch IoT Speakers from the task bar

My monitor speakers do not have an accessible on/off switch, so I have them wired to a WiFi smart plug (which I flashed with [Tasmota](https://tasmota.github.io/docs/)). This application creates a tray icon that allows me to switch the speakers on and off. It also monitors the audio output, and if no sound output is detected for a while the speakers will be switched off. Since the speakers pull in around 100W each, this saves quite a bit of energy!

There is also an option to automatically change Windows' default audio output device when speakers are turned on and automatically turned off. In my case my display has (very poor) inbuild speakers, but they are still better than no audio.

## Installation

 * create and enter the virtual environment (name it `venv` for the `run-in-background.bat` to work)
 * install the dependencies (`pip -r requirements.txt`)
 * copy `config-default.ini` to `config.ini`, adjusting at least the `BASEURL`
 * run `trayapp.py`, and potentially link `run-in-background.bat` to your `shell:startup` directory.

# Automatic turn off on shutdown/standy/hibernation

It turns out it is impossible to receive Windows 10 shutdown notices from [any application that is launched using `pythonw.exe`](https://stackoverflow.com/questions/64522390/). It seems the only solution to capture shutdowns is to use a Windows service. Hence `service.py`  contains a simple service that listens for shutdowns, standy and hibernation, and turns of the power to the speakers. It can also capture logoff and screen lock events, but these are currently not used.

## Installation

* open the virtual environment in an elevated PS
* run `python venv\Scripts\pywin32_postinstall.py -install` (from [SO](https://stackoverflow.com/questions/34696815/))
* run `python service.py --startup auto install`
* open the Windows Services app to
  * if the environment is installed in a home directory, the service needs to be started with that users' permissions
  * start the service (or do this via `sc.exe start AutomaticSpeakerTurnOff`).

It seems that Windows 10 version upgrades remove the service. I had to repeat the four steps above when upgrading to 2004.

## Notes

* [pywin32](https://github.com/mhammond/pywin32) has changed how services are called / installed in a non-backwards compatible manner. Should upgrade at some point! 
* investigate https://github.com/mhammond/pywin32/issues/1450#issuecomment-1119014554 - maybe we don't need the `sys.path.insert` logic
