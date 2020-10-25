import os, logging, sys

## The service gets started in a strange place, rootdir should be the directory of this file.
ROOTDIR = os.sep.join(os.getcwd().split(os.sep)[:-4])

logging.basicConfig(format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
    datefmt='%m-%d %H:%M:%S',
    level = logging.DEBUG, 
    filename = ROOTDIR + '\\SpeakerShutdown-service.log',
)

sys.path.insert(0, ROOTDIR + '\\venv\\lib\\site-packages')
sys.path.insert(0, ROOTDIR + '\\venv\\lib\\site-packages\\win32\\lib')

import win32serviceutil, win32service
import win32event
import servicemanager

import win32gui, win32gui_struct, win32con

import urllib.request as r
import json, configparser

VERSION = 8

class SpeakerShutdown(win32serviceutil.ServiceFramework):

    _svc_name_ = "AutomaticSpeakerTurnOff"
    _svc_display_name_ = "Automatic Speaker Turn Off"
    _svc_description_ = "Automatically turn off speakers on powerdown"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        logging.debug(f"Init was called, version {VERSION}")
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.setupCFG()
    
    # Override the base class so we can accept additional events.
    def GetAcceptedControls(self):
        rc = win32serviceutil.ServiceFramework.GetAcceptedControls(self)
        rc |= win32service.SERVICE_CONTROL_SHUTDOWN \
              | win32service.SERVICE_ACCEPT_SESSIONCHANGE
        return rc
    
    # Override in order to process service shutdown requests.
    def SvcShutdown(self):
        logging.info(f"Shutdown event captured in SvcShutdown")
        self.turnPowerOff()
        self.SvcStop()

    # All extra events are sent via SvcOtherEx (SvcOther remains as a
    # function taking only the first args for backwards compat)
    def SvcOtherEx(self, control, event_type, data):
        # This is only showing a few of the extra events - see the MSDN
        # docs for "HandlerEx callback" for more info.
        if control == win32service.SERVICE_CONTROL_SESSIONCHANGE:
            # data is a single elt tuple, but this could potentially grow
            # in the future if the win32 struct does
            logging.debug(f"Session event: type={event_type}, data={data}")
            if event_type == 6: ## Logoff - could also turn off here, but shutdown is nicer?
                pass
                #self.turnPowerOff()
        else:
            logging.debug("Other event: code=%d, type=%s, data=%s" % (control, event_type, data))

    def turnPowerOff(self):
        try:
            status = json.loads(r.urlopen(self.CFG['APP']['BASEURL'] + 'Power1%20OFF').read())
            logging.debug(f"Send turn-off, returned: {status}")
        except Exception:
            logging.exception("Something went wrong when trying to turn the power off")
    
    def setupCFG(self):
        cfg = configparser.ConfigParser()
        if os.path.exists(os.path.join(ROOTDIR,'config-default.ini')):
            with open(os.path.join(ROOTDIR, 'config-default.ini')) as f:
                cfg.read_file(f)
        cfg.read(os.path.join(ROOTDIR, 'config.ini'))
        self.CFG = cfg
    
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        # do nothing at all - just wait to be stopped
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
        # Write a stop message.
        logging.info("Service stopped.")
        servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STOPPED,
                (self._svc_name_, '')
                )

if __name__=='__main__':
    win32serviceutil.HandleCommandLine(SpeakerShutdown)