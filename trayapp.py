from infi.systray import SysTrayIcon
import urllib.request as r
import json, time, sys
import numpy as np
import soundcard as sc
from datetime import datetime, timedelta

import logging, configparser, os



def setupCFG(cfg):
    if os.path.exists('config-default.ini'):
        with open('config-default.ini') as f:
            cfg.read_file(f)
    cfg.read('config.ini')

def setupLogging():
    logging.basicConfig(format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
                        datefmt='%m-%d %H:%M:%S',
                        filename=CFG['LOGS']['LOGFILE'], 
                        level=getattr(logging, CFG['LOGS']['LOGLEVEL'], logging.DEBUG))

    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))


def runCMD(systray: SysTrayIcon, cmd: str):
    try:
        status = json.loads(r.urlopen(CFG['APP']['BASEURL']+cmd).read())
        logging.debug(f"CMD {cmd}: {status}")
    except Exception:
        logging.exception("Something went wrong talking to the plug")
        setTrayIcon(systray, "ERROR")
        return "ERR"
    else:
        setTrayIcon(systray, status['POWER'])
        return status['POWER']
    

def getState(systray: SysTrayIcon):
    return runCMD(systray, 'Power1')

def toggle(systray: SysTrayIcon):
    return runCMD(systray, 'Power1%20TOGGLE')

def turnOff(systray:SysTrayIcon):
    return runCMD(systray, 'Power1%20OFF')


def setTrayIcon(systray: SysTrayIcon, status):
    if status == 'ON':
        systray.update(icon='icons/icon-on.ico')
    elif status == 'OFF':
        systray.update(icon='icons/icon-off.ico')
    else:
        systray.update(icon='icons/icon-unsure.ico')

def window_rms(a, window_size):
  a2 = np.power(a,2)
  window = np.ones(window_size)/float(window_size)
  return np.sqrt(np.convolve(a2, window, 'valid'))


def isThereAudioOutput():
    main_speaker = sc.default_speaker()
    main_speaker_loopback = [x for x in sc.all_microphones(include_loopback=True) if x.name == main_speaker.name and x.isloopback][0]
    df = main_speaker_loopback.record(numframes=int(CFG['AUTOTURNOFF']['LISTEN'])*int(CFG['AUTOTURNOFF']['SAMPLERATE']), 
                                        samplerate=int(CFG['AUTOTURNOFF']['SAMPLERATE']))
    if df.max() == 0:
        ## Well nothing was recorded since nothing is trying to play back anything right now... maybe we should retry a few times?
        return False
    channels = df.shape[1]
    for i in range(channels):
        loudness = window_rms(df[:,i], int(CFG['AUTOTURNOFF']['SECONDS']))
        # lets say something is not noise if the 90 percentile is > 0.01. 
        # I have little evidence if this is any good.
        if np.percentile(loudness, 90) > float(CFG['AUTOTURNOFF']['LEVEL']):
            return True
    return False

def regularAudioChecks(systray: SysTrayIcon):
    audioChecks = []
    while systray._message_loop_thread.is_alive():
        state = getState(systray)
        if state == 'ON' and datetime.now() > disableUntil and bool(CFG['AUTOTURNOFF']['ENABLE']):
            audioChecks.append(isThereAudioOutput())
            logging.debug(audioChecks)
            # Turn off after x intervals of no outputs; this should be configurable.
            if len(audioChecks) > int(CFG['AUTOTURNOFF']['INTERVAL']):
                audioChecks.pop(0)
                if not any(audioChecks):
                    turnOff(systray)
                    audioChecks = []
        time.sleep(int(CFG['AUTOTURNOFF']['WAIT']))
    logging.debug("Quitting.")

def disableAutoOff(systray: SysTrayIcon, hours : int):
    global disableUntil
    disableUntil = datetime.now() + timedelta(hours=hours)
    if hours > 0:
        systray.update(hover_text="Manage Speakers via IoT: auto-off disabled until " + disableUntil.strftime("%H:%M"))
    else:
        systray.update(hover_text="Manage Speakers via IoT")


def main():
    menu_options = (("Toggle power", None, toggle),
                    ("Disable auto-off", None, (('1 hour', None, lambda x: disableAutoOff(x, 1)),
                                                ('2 hour', None, lambda x: disableAutoOff(x, 2)),
                                                ('3 hour', None, lambda x: disableAutoOff(x, 3)),
                                                ('4 hour', None, lambda x: disableAutoOff(x, 4)) )),
                    ("Enable auto-off (default)", None, lambda x: disableAutoOff(x, 0)))
    with SysTrayIcon("icons/icon-unsure.ico", "Manage Speakers via IoT", menu_options) as systray:
        getState(systray)
        regularAudioChecks(systray)

if __name__ == '__main__':
    CFG = configparser.ConfigParser()
    disableUntil = datetime.now()
    setupCFG(CFG)
    setupLogging()
    main()