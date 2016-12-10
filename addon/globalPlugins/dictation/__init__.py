from ctypes import *
from ctypes.wintypes import *
import os
from win32api import *
from win32con import *
import win32con

import wx

import config
import controlTypes
import eventHandler
import globalCommands
from globalPluginHandler import GlobalPlugin as BaseGlobalPlugin
import inputCore
import keyboardHandler
from logHandler import log
from NVDAObjects.IAccessible import getNVDAObjectFromEvent
from NVDAObjects import NVDAObject
import speech
import time
import windowUtils
import winInputHook
import winUser

currentEntry = None
autoFlushTimer = None
requestedWSRShowHideEvents = False
wsrAlternatesPanel = None
wsrSpellingPanel = None
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
	text = text.replace("\r\n", "\n")
	text = text.replace("\r", "\n")
	log.info("text %r" % text)
	if text == "\n":
		speech.speakText("new line")
	elif text == "\n\n":
		speech.speakText("new paragraph")
	elif text == "": # new paragraph from Dragon in Word
		speech.speakText("new paragraph")
	else:
		speech.speakText(text)
	currentEntry = None
	requestWSRShowHideEvents()

def textInserted(hwnd, start, text):
	global currentEntry, autoFlushTimer
	if currentEntry is not None:
		prevStart, prevText = currentEntry
		if (not (start == -1 and prevStart == -1)) and (start < prevStart or start > (prevStart + len(prevText))):
			flushCurrentEntry()
	if currentEntry is not None:
		prevStart, prevText = currentEntry
		if prevStart == -1 and start == -1:
			currentEntry = (-1, prevText + text)
		else:
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
cTextInsertedCallback = WINFUNCTYPE(None, HWND, DWORD, c_wchar_p)(textInserted)

def textDeleted(hwnd, start, text):
	speech.speakText("deleted %s" % text)
cTextDeletedCallback = WINFUNCTYPE(None, HWND, DWORD, c_wchar_p)(textDeleted)

def makeKeyName(knm):
	"Function to process key names for sending to executeGesture"
	colonSplit=knm.split(":")
	tk=colonSplit[1]
	nvm="insert"
	if(tk.startswith("NVDA")):
		tKey=tk.replace("NVDA",nvm)
	else:
		tKey=tk

	return tKey

def execCommand(action):
	"""take commands from speech-recognition and send them to NVDA
	The function accepts script names to execute"""
	for key, funcName in globalCommands.commands._GlobalCommands__gestures.items():
		if (key.startswith("kb:") or key.startswith("kb(%s):" % config.conf["keyboard"]["keyboardLayout"])) and action == funcName:
			inputCore.manager.executeGesture(keyboardHandler.KeyboardInputGesture.fromName(makeKeyName(key)))
			break

def commandCallback(command):
	execCommand(command)
cCommandCallback = WINFUNCTYPE(None, c_char_p)(commandCallback)

lastKeyDownTime = None

def patchKeyDownCallback():
	originalCallback = winInputHook.keyDownCallback
	def callback(*args, **kwargs):
		global lastKeyDownTime
		lastKeyDownTime = time.time()
		return originalCallback(*args, **kwargs)
	winInputHook.keyDownCallback = callback

masterDLL = None

def initialize():
	global masterDLL
	addonRootDir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
	dllPath = os.path.join(addonRootDir, "DictationBridgeMaster32.dll")
	masterDLL = windll.LoadLibrary(dllPath)
	masterDLL.DBMaster_SetTextInsertedCallback(cTextInsertedCallback)
	masterDLL.DBMaster_SetTextDeletedCallback(cTextDeletedCallback)
	masterDLL.DBMaster_SetCommandCallback(cCommandCallback)
	if not masterDLL.DBMaster_Start():
		raise WinError()
	patchKeyDownCallback()

def terminate():
	global masterDLL
	if masterDLL is not None:
		masterDLL.DBMaster_Stop()
		masterDLL = None

def getCleanedWSRAlternatesPanelItemName(obj):
	return obj.name[2:] # strip symbol 2776 and space

def speakWSRAlternatesPanelItem(obj):
	text = getCleanedWSRAlternatesPanelItemName(obj)
	speech.speakText(text)

def speakAndSpellWSRAlternatesPanelItem(obj):
	text = getCleanedWSRAlternatesPanelItemName(obj)
	speech.speakText(text)
	speech.speakSpelling(text)

def selectListItem(obj):
	obj.IAccessibleObject.accSelect(2, obj.IAccessibleChildID)

IDOK = 1
IDCANCEL = 2
IDC_SPELLING_WORD = 6304

class WSRAlternatesPanel(NVDAObject):
	def script_ok(self, gesture):
		buttonWindowHandle = windll.user32.GetDlgItem(self.windowHandle, IDOK)
		button = getNVDAObjectFromEvent(buttonWindowHandle, winUser.OBJID_CLIENT, 0)
		button.doAction()

	def script_cancel(self, gesture):
		buttonWindowHandle = windll.user32.GetDlgItem(self.windowHandle, IDCANCEL)
		button = getNVDAObjectFromEvent(buttonWindowHandle, winUser.OBJID_CLIENT, 0)
		button.doAction()

	def script_selectPreviousItem(self, gesture):
		for obj in self.recursiveDescendants:
			if obj.role != controlTypes.ROLE_LISTITEM:
				continue
			if controlTypes.STATE_SELECTED in obj.states:
				if obj.previous is not None and obj.previous.role == controlTypes.ROLE_LISTITEM:
					selectListItem(obj.previous)
				break

	def script_selectNextItem(self, gesture):
		firstListItem = None
		for obj in self.recursiveDescendants:
			if obj.role != controlTypes.ROLE_LISTITEM:
				continue
			if firstListItem is None:
				firstListItem = obj
			if controlTypes.STATE_SELECTED in obj.states:
				if obj.next is not None and obj.next.role == controlTypes.ROLE_LISTITEM:
					selectListItem(obj.next)
				break
		else:
			if firstListItem is not None:
				selectListItem(firstListItem)

	def script_selectFirstItem(self, gesture):
		firstListItem = None
		for obj in self.recursiveDescendants:
			if obj.role != controlTypes.ROLE_LISTITEM:
				continue
			if firstListItem is None:
				firstListItem = obj
				break
		if firstListItem is not None:
			selectListItem(firstListItem)

	def script_selectLastItem(self, gesture):
		lastListItem = None
		for obj in self.recursiveDescendants:
			if obj.role != controlTypes.ROLE_LISTITEM:
				continue
			lastListItem = obj
		if lastListItem is not None:
			selectListItem(lastListItem)

	__gestures = {
		'kb:enter': 'ok',
		'kb:escape': 'cancel',
		'kb:upArrow': 'selectPreviousItem',
		'kb:downArrow': 'selectNextItem',
		'kb:home': 'selectFirstItem',
		'kb:control+home': 'selectFirstItem',
		'kb:end': 'selectLastItem',
		'kb:control+end': 'selectLastItem',
	}

class WSRSpellingPanel(NVDAObject):
	pollTimer = None
	previousWord = None

	def _get_word(self):
		wordWindowHandle = windll.user32.GetDlgItem(self.windowHandle, IDC_SPELLING_WORD)
		wordObject = getNVDAObjectFromEvent(wordWindowHandle, winUser.OBJID_CLIENT, 0)
		if controlTypes.STATE_INVISIBLE in wordObject.states:
			return ""
		return wordObject.name

	def poll(self, *args, **kwargs):
		self.pollTimer = None
		oldWord = self.previousWord or ""
		newWord = self.word or ""
		if newWord != oldWord:
			self.previousWord = newWord
			if len(newWord) > len(oldWord) and newWord[:len(oldWord)] == oldWord:
				speech.speakSpelling(newWord[len(oldWord):])
			elif newWord:
				speech.speakText(newWord)
				speech.speakSpelling(newWord)
			elif oldWord:
				speech.speakText("cleared")
		self.schedulePoll()

	def cancelPoll(self):
		if self.pollTimer is not None:
			self.pollTimer.Stop()
			self.pollTimer = None

	def schedulePoll(self):
		self.cancelPoll()
		self.pollTimer = wx.CallLater(100, self.poll)

	def script_ok(self, gesture):
		buttonWindowHandle = windll.user32.GetDlgItem(self.windowHandle, IDOK)
		button = getNVDAObjectFromEvent(buttonWindowHandle, winUser.OBJID_CLIENT, 0)
		button.doAction()

	def script_cancel(self, gesture):
		buttonWindowHandle = windll.user32.GetDlgItem(self.windowHandle, IDCANCEL)
		button = getNVDAObjectFromEvent(buttonWindowHandle, winUser.OBJID_CLIENT, 0)
		button.doAction()

	__gestures = {
		'kb:enter': 'ok',
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
		elif obj.windowClassName == '#32770' and obj.name == "Spelling panel":
			clsList.insert(0, WSRSpellingPanel)

	def event_show(self, obj, nextHandler):
		global wsrAlternatesPanel, wsrSpellingPanel
		if isinstance(obj, WSRAlternatesPanel):
			wsrAlternatesPanel = obj
			speech.cancelSpeech()
			speech.speakText(obj.name)
			for descendant in obj.recursiveDescendants:
				if controlTypes.STATE_INVISIBLE in descendant.states or controlTypes.STATE_INVISIBLE in descendant.parent.states:
					continue
				if descendant.role == controlTypes.ROLE_STATICTEXT:
					speech.speakText(descendant.name)
				elif descendant.role == controlTypes.ROLE_LINK:
					speech.speakText("Or say")
					speech.speakText(descendant.name)
				elif descendant.role == controlTypes.ROLE_LISTITEM:
					speech.speakText(str(descendant.positionInfo["indexInGroup"]))
					speakWSRAlternatesPanelItem(descendant)
			return
		elif isinstance(obj, WSRSpellingPanel):
			if wsrSpellingPanel is not None:
				wsrSpellingPanel.cancelPoll()
			wsrSpellingPanel = obj
			wsrSpellingPanel.schedulePoll()
			speech.cancelSpeech()
			speech.speakText(obj.name)
			for descendant in obj.recursiveDescendants:
				if controlTypes.STATE_INVISIBLE in descendant.states or controlTypes.STATE_INVISIBLE in descendant.parent.states:
					continue
				if descendant.role == controlTypes.ROLE_STATICTEXT:
					speech.speakText(descendant.name)
				elif descendant.role == controlTypes.ROLE_LINK:
					speech.speakText("Or say")
					speech.speakText(descendant.name)
			return
		nextHandler()

	def wsrPanelHidden(self, windowHandle):
		global wsrAlternatesPanel, wsrSpellingPanel
		if wsrAlternatesPanel is not None and windowHandle == wsrAlternatesPanel.windowHandle:
			if wsrSpellingPanel is None:
				speech.speakText("Closed alternates panel")
			wsrAlternatesPanel = None
		elif wsrSpellingPanel is not None and windowHandle == wsrSpellingPanel.windowHandle:
			wsrSpellingPanel.cancelPoll()
			speech.speakText("Closed spelling panel")
			wsrSpellingPanel = None

	def event_selection(self, obj, nextHandler):
		if obj.role == controlTypes.ROLE_LISTITEM and isInWSRAlternatesPanel(obj):
			speech.speakText(str(obj.positionInfo["indexInGroup"]))
			speakAndSpellWSRAlternatesPanelItem(obj)
			return
		nextHandler()

	def getScript(self, gesture):
		if wsrAlternatesPanel is not None:
			result = wsrAlternatesPanel.getScript(gesture)
			if result is not None:
				return result
		elif wsrSpellingPanel is not None:
			result = wsrSpellingPanel.getScript(gesture)
			if result is not None:
				return result
		return super(GlobalPlugin, self).getScript(gesture)

	def event_typedCharacter(self, obj, nextHandler, ch):
		if lastKeyDownTime is None or (time.time() - lastKeyDownTime) >= 0.5:
			if obj.windowClassName != "ConsoleWindowClass":
				textInserted(obj.windowHandle, -1, ch)
			return
		nextHandler()

	def terminate(self):
		super(GlobalPlugin, self).terminate()
		terminate()
