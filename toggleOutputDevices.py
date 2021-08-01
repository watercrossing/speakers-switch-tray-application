
import comtypes
import soundcard as sc
import configparser, os

import policyconfig as pc

def switchSpeaker(switchTo):
    policy_config = comtypes.CoCreateInstance(
        pc.CLSID_PolicyConfigClient,
        pc.IPolicyConfig,
        comtypes.CLSCTX_ALL
    )

    policy_config.SetDefaultEndpoint(switchTo.id, pc.ERole.eMultimedia)


def setupCFG(cfg):
    if os.path.exists('config-default.ini'):
        with open('config-default.ini') as f:
            cfg.read_file(f)
    cfg.read('config.ini')


print([x.name for x in sc.all_speakers()])
CFG = configparser.ConfigParser()
setupCFG(CFG)

print(sc.default_speaker())

mainSpeaker = CFG['AUTOTURNOFF']['SPEAKERTOMONITOR']
SPEAKERIFOFF = CFG['AUTOTURNOFF']['SPEAKERIFOFF']

if mainSpeaker not in sc.default_speaker().name:
    switchTo = next(x for x in sc.all_speakers() if mainSpeaker in x.name)
    print(switchTo.id)
    switchSpeaker(switchTo)
else:
    switchTo = next(x for x in sc.all_speakers() if SPEAKERIFOFF in x.name)
    print(switchTo.id)
    switchSpeaker(switchTo)


print(sc.default_speaker())



