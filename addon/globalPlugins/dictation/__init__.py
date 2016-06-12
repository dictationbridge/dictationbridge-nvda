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
import speech
import winUser

currentEntry = None
autoFlushTimer = None
requestedWSRShowEvents = False
wsrAlternatesPanelWindowHandle = None
wsrAlternatesPanelHiddenFunction = None

def requestWSRShowEvents(fn=None):
	global requestedWSRShowEvents, hookId, eventCallback, wsrAlternatesPanelHiddenFunction
	if fn is None:
		fn = wsrAlternatesPanelHiddenFunction
	else:
		wsrAlternatesPanelHiddenFunction = fn
	if requestedWSRShowEvents:
		return
	try:
		hwnd = FindWindow("MS:SpeechTopLevel", None)
	except:
		hwnd = None
	if hwnd:
		pid, tid = winUser.getWindowThreadProcessID(hwnd)
		eventHandler.requestEvents(eventName='show', processId=pid, windowClassName='#32770')
		eventHandler.requestEvents(eventName='hide', processId=pid, windowClassName='#32770')
		eventCallback = make_callback(fn)
		hookId = winUser.setWinEventHook(win32con.EVENT_OBJECT_HIDE, win32con.EVENT_OBJECT_HIDE, 0, eventCallback, pid, 0, 0)
		requestedWSRShowEvents = True

def make_callback(fn):
	@WINFUNCTYPE(None, c_int, c_int, c_int, c_int, c_int, c_int, c_int)
	def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
		fn()
	return callback

def flushCurrentEntry():
	global currentEntry, autoFlushTimer
	if autoFlushTimer is not None:
		autoFlushTimer.Stop()
		autoFlushTimer = None
	start, text = currentEntry
	speech.speakText(text)
	currentEntry = None
	requestWSRShowEvents()

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

# HACK
def speakSpellingImmediately(text):
	for _ in speech._speakSpellingGen(text, speech.getCurrentLanguage(), False):
		pass

def speakAndSpellWSRAlternatesPanelItem(obj):
	text = obj.name[2:] # strip symbol 2776 and space
	speech.speakText("%d. %s" % (obj.IAccessibleChildID, text))
	speakSpellingImmediately(text)

def isInWSRAlternatesPanel(obj):
	while obj is not None:
		if obj.name == "Alternates panel" and obj.windowClassName == "#32770":
			return True
		obj = obj.parent
	return False

class GlobalPlugin(BaseGlobalPlugin):
	def __init__(self):
		super(GlobalPlugin, self).__init__()
		initialize()
		requestWSRShowEvents(self.alternates_hidden)

	def event_show(self, obj, nextHandler):
		global wsrAlternatesPanelWindowHandle
		if obj.name == "Alternates panel":
			wsrAlternatesPanelWindowHandle = obj.windowHandle
			speech.speakText(obj.name)
			for descendant in obj.recursiveDescendants:
				if descendant.role == controlTypes.ROLE_STATICTEXT:
					speech.speakText(descendant.name)
				elif descendant.role == controlTypes.ROLE_LISTITEM:
					speakAndSpellWSRAlternatesPanelItem(descendant)

	def alternates_hidden(self):
		global wsrAlternatesPanelWindowHandle
		if wsrAlternatesPanelWindowHandle is not None:
			speech.speakText("Closed alternates panel")
			wsrAlternatesPanelWindowHandle = None

	def event_selection(self, obj, nextHandler):
		if obj.role == controlTypes.ROLE_LISTITEM and isInWSRAlternatesPanel(obj):
			speech.speakText("selected")
			speakAndSpellWSRAlternatesPanelItem(obj)

	def terminate(self):
		super(GlobalPlugin, self).terminate()
		terminate()
