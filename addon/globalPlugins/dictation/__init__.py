from ctypes import *
from ctypes.wintypes import *
import os
from win32api import *
from win32con import *
import win32con

import wx

import controlTypes
import eventHandler
from globalPluginHandler import GlobalPlugin as BaseGlobalPlugin
from NVDAObjects.IAccessible import getNVDAObjectFromEvent
from NVDAObjects import NVDAObject
import speech
import windowUtils
import winUser

currentEntry = None
autoFlushTimer = None
requestedWSRShowHideEvents = False
wsrAlternatesPanel = None
wsrPanelHiddenFunction = None

def requestWSRShowHideEvents(fn=None):
    global requestedWSRShowHideEvents, hookId, eventCallback, wsrPanelHiddenFunction
    if fn is None:
        fn = wsrPanelHiddenFunction
    else:
        wsrPanelHiddenFunction = fn
    if requestedWSRShowHideEvents:
        return
    try:
        hwnd = winUser.FindWindow(u"MS:SpeechTopLevel", None)
    except:
        hwnd = None
    if hwnd:
        pid, tid = winUser.getWindowThreadProcessID(hwnd)
        eventHandler.requestEvents(eventName='show', processId=pid, windowClassName='#32770')
        eventCallback = make_callback(fn)
        hookId = winUser.setWinEventHook(win32con.EVENT_OBJECT_HIDE, win32con.EVENT_OBJECT_HIDE, 0, eventCallback, pid, 0, 0)
        requestedWSRShowHideEvents = True

def make_callback(fn):
    @WINFUNCTYPE(None, c_int, c_int, c_int, c_int, c_int, c_int, c_int)
    def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
        fn(hwnd)
    return callback

def flushCurrentEntry():
    global currentEntry, autoFlushTimer
    if autoFlushTimer is not None:
        autoFlushTimer.Stop()
        autoFlushTimer = None
    start, text = currentEntry
    speech.speakText(text)
    currentEntry = None
    requestWSRShowHideEvents()

@WINFUNCTYPE(None, HWND, DWORD, c_wchar_p)
def textInsertedCallback(hwnd, start, text):
    global currentEntry, autoFlushTimer
    if currentEntry is not None:
        prevStart, prevText = currentEntry
        if start < prevStart or start > (prevStart + len(prevText)):
            flushCurrentEntry()
    if currentEntry is not None:
        prevStart, prevText = currentEntry
        currentEntry = (prevStart, prevText[:start - prevStart] + text)
    else:
        currentEntry = (start, text)
    if autoFlushTimer is not None:
        autoFlushTimer.Stop()
        autoFlushTimer = None
    def autoFlush(*args, **kwargs):
        global autoFlushTimer
        autoFlushTimer = None
        flushCurrentEntry()
    autoFlushTimer = wx.CallLater(100, autoFlush)

masterDLL = None

def initialize():
    global masterDLL
    addonRootDir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    dllPath = os.path.join(addonRootDir, "DictationBridgeMaster.dll")
    masterDLL = windll.LoadLibrary(dllPath)
    masterDLL.DBMaster_SetTextInsertedCallback(textInsertedCallback)
    if not masterDLL.DBMaster_Start():
        raise WinError()

def terminate():
    global masterDLL
    if masterDLL is not None:
        masterDLL.DBMaster_Stop()
        masterDLL = None

def speakAndSpellWSRAlternatesPanelItem(obj):
    text = obj.name[2:] # strip symbol 2776 and space
    speech.speakText("%d. %s" % (obj.IAccessibleChildID, text))
    speech.speakSpelling(text)

IDOK = 1
IDCANCEL = 2

class WSRAlternatesPanel(NVDAObject):
    def script_cancel(self, gesture):
        buttonWindowHandle = windll.user32.GetDlgItem(self.windowHandle, IDCANCEL)
        button = getNVDAObjectFromEvent(buttonWindowHandle, winUser.OBJID_CLIENT, 0)
        button.doAction()

    __gestures = {
        'kb:escape': 'cancel',
    }

def isInWSRAlternatesPanel(obj):
    while obj is not None:
        if isinstance(obj, WSRAlternatesPanel):
            return True
        obj = obj.parent
    return False

class GlobalPlugin(BaseGlobalPlugin):
    def __init__(self):
        super(GlobalPlugin, self).__init__()
        initialize()
        requestWSRShowHideEvents(self.wsrPanelHidden)

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        if obj.windowClassName == '#32770' and obj.name == "Alternates panel":
            clsList.insert(0, WSRAlternatesPanel)

    def event_show(self, obj, nextHandler):
        global wsrAlternatesPanel
        if isinstance(obj, WSRAlternatesPanel):
            wsrAlternatesPanel = obj
            speech.cancelSpeech()
            speech.speakText(obj.name)
            for descendant in obj.recursiveDescendants:
                if descendant.role == controlTypes.ROLE_STATICTEXT:
                    speech.speakText(descendant.name)
                elif descendant.role == controlTypes.ROLE_LISTITEM:
                    speakAndSpellWSRAlternatesPanelItem(descendant)
            return
        nextHandler()

    def wsrPanelHidden(self, windowHandle):
        global wsrAlternatesPanel
        if wsrAlternatesPanel is not None and windowHandle == wsrAlternatesPanel.windowHandle:
            speech.speakText("Closed alternates panel")
            wsrAlternatesPanel = None

    def event_selection(self, obj, nextHandler):
        if obj.role == controlTypes.ROLE_LISTITEM and isInWSRAlternatesPanel(obj):
            speech.speakText("selected")
            speakAndSpellWSRAlternatesPanelItem(obj)
            return
        nextHandler()

    def getScript(self, gesture):
        if wsrAlternatesPanel is not None:
            result = wsrAlternatesPanel.getScript(gesture)
            if result is not None:
                return result
        return super(GlobalPlugin, self).getScript(gesture)

    def terminate(self):
        super(GlobalPlugin, self).terminate()
        terminate()
