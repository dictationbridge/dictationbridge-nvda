import os
import sys
import shutil
import subprocess
import threading
import time
from ctypes import *
from ctypes.wintypes import *
import wx
import api
import braille
import config
import ui
import controlTypes
import core
import eventHandler
import globalCommands
import gui
import inputCore
import keyboardHandler
import queueHandler
import speech
import windowUtils
import winInputHook
import winUser
from globalPluginHandler import GlobalPlugin as BaseGlobalPlugin
from logHandler import log
from NVDAObjects import NVDAObject
from NVDAObjects.IAccessible import getNVDAObjectFromEvent
from win32api import *
from dictationGesture import DictationGesture
import inputCore

addonRootDir = os.path.abspath(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
currentEntry = None
autoFlushTimer = None
requestedWSRShowHideEvents = False
wsrAlternatesPanel = None
wsrSpellingPanel = None
wsrPanelHiddenFunction = None

def escape(input):
		input = input.replace("<","&lt;")
		input = input.replace(">","&gt;")
		input = input.replace('"',"&quot;")
		input = input.replace("'","&apos;")
		return input

class HelpCategory(object):
	def __init__(self,categoryName):
		self.categoryName = escape(categoryName)
		self.rows = []

	def addRow(self, command, help):
		self.rows.append((escape(command), escape(help)))

	def html(self):
		html = "<h3>"+self.categoryName+"</h3>"
		html += "<table><tr><th>"
		#Translators: The name of the column header for what the user speaks to activate a command.
		html += escape(_("Say This"))
		html += "</th><th>"
		#Translators: The name of a column header for the help text for what this speech will do.
		html += escape(_("To Do This"))
		html += "</th></tr>"
		for row in self.rows:
			html+="<tr><th>"+row[0]+"</th><td>"+row[1]+"</td></tr>"
		html += "</table>"
		return html

def dbHelp():
	sys.path.append(addonRootDir)
	#Import this late, as we only need this here and it's a large autobuilt blob.
	from NVDA_helpCommands import commands
	sys.path.remove(sys.path[-1])
	html = "<p>"
	#Translators: Description of how to Move to the next topic in the in-built help.
	html += escape(_('To use this help documentation, you can navigate to the next category with h, or by saying "next heading".'))
	#Translators: Description of how to Move  through help tables in the in-built help.
	html += escape(_('To move by column or row, use table navigation, (You can say "Previous row/column", "prev row", "prev column", "next row", "next column" to navigate with speech).'))
	#Translators: Description of how to  find a command  in the in-built help.
	html += escape(_('to find specific text, say "find text", wait for the find dialog to appear, then dictate your text, then say "press enter" or "click ok".'))
	html += "</p><h2>"
	#Translators: The Context sensitive help heading, telling the user what these commands are..
	html +=escape(_("Currently available commands."))
	html += "</h2>"
	categories = {}
	#Translators: The name of a category in Dictationbridge for the commands help.
	miscName = _("Miscellaneous")
	categories[miscName] = HelpCategory(miscName)
	for command in commands:
		if command["identifier_for_NVDA"] in SPECIAL_COMMANDS:
			#All special commands get the miscelaneous category.
			categories[miscName].addRow(command["text"], command["helpText"])
			continue
		#Creating a gesture is not efficient, but helps eliminate code bloat.
		gesture = DictationGesture(command["identifier_for_NVDA"])
		scriptInfo = gesture.script_hacky
		if not scriptInfo:
			#This script is not active right now!
			continue
		doc = getattr(scriptInfo[0], "__doc__", "")
		category = ""
		try:
			category = scriptInfo[0].category
		except AttributeError:
			category = getattr(scriptInfo[1], "scriptCategory", miscName)
		if not categories.get(category):
			categories[cat] = HelpCategory(cat)
		categories[cat].addRow(command["text"], doc)
	for category in categories.values():
		html+=category.html()
	ui.browseableMessage(html,
		#Translators: The title of the context sensitive help for Dictation Bridge NVDA Commands.
		_("Dictation Bridge NVDA Context Sensitive Help"),
		True)

SPECIAL_COMMANDS = {
	"stopTalking" : speech.cancelSpeech,
	"toggleTalking" : lambda:speech.pauseSpeech(not speech.isPaused),
	"dbHelp" : dbHelp,
}

def successDialog(program):
	#Translators: Message shown if the commands were installed into dragon successfully.
	gui.messageBox(_("The {0}  commands were successfully installed. Please restart your {0} profile to proceed. See the manual for details on how to do this.").format(program),
	#Translators: Title of the successfully installed commands dialog
	_("Success!"))

def _onInstallDragonCommands():
	si = subprocess.STARTUPINFO()
	si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
	si.wShowWindow = subprocess.SW_HIDE
	dragonDir = r"C:\Program Files (x86)\Nuance\NaturallySpeaking15\Program"
	#Translators: Title of an error dialog shown in dictation bridge.
	DB_ERROR_TITLE = _("Dictation Bridge Error")
	if not os.path.exists(dragonDir):
		dragonDir.replace(r" (x86)", "")
	if not os.path.exists(dragonDir):
		#Translators: Message given to the user when the addon can't find an installed copy of dragon.
		gui.messageBox(_("Cannot find dragon installed on your machine. Please install dragon and then try this process again."),
			DB_ERROR_TITLE)
		return
	xml2dat = os.path.join(dragonDir, "mycmdsxml2dat.exe")
	nsadmin = os.path.join(dragonDir, "nsadmin.exe")
	
	#Translators: The official name of Dragon in your language, this probably should be left as Dragon.
	thisProgram = _("Dragon")
	if os.path.exists(os.path.join(addonRootDir, "dragon_dictationBridgeCommands.dat")):
			os.remove(os.path.join(addonRootDir, "dragon_dictationBridgeCommands.dat"))
	try:
		subprocess.check_call([
			xml2dat,
			os.path.join(addonRootDir, "dragon_dictationBridgeCommands.dat"),
			os.path.join(addonRootDir, "dragon_dictationBridgeCommands.xml"),
			], startupinfo=si)
		#Fixme: might need to get the users language, and put them there for non-english locales.
		d=config.execElevated(nsadmin,
			["/commands", os.path.join(addonRootDir, "dragon_dictationBridgeCommands.dat"), "/overwrite=yes"],
			wait=True, handleAlreadyElevated=True)

		successDialog(thisProgram)
	except:
		#Translators: Message shown if dragon commands failed to install.
		gui.messageBox(_("There was an error while performing the addition of dragon commands into dragon. Are you running as an administrator? If so, please send the error in your log to the dictation bridge team as a bug report."),
			DB_ERROR_TITLE)
		raise
	finally:
		if os.path.exists(os.path.join(addonRootDir, "dragon_dictationBridgeCommands.dat")):
			os.remove(os.path.join(addonRootDir, "dragon_dictationBridgeCommands.dat"))

def _onInstallMSRCommands():
	MSRPATH = os.path.expanduser(r"~\documents\speech macros")
	commandsFile = os.path.join(addonRootDir, "dictationBridge.WSRMac")
	#Translators: Official Title of Microsoft speech Recognition in your language.
	thisProgram = _("Microsoft Speech Recognition")
	if os.path.exists(MSRPATH):
		shutil.copy(commandsFile, MSRPATH)
		successDialog(thisProgram)
	else:
		#Translators: The user doesn't have microsoft speech recognition profiles, or we can't find them.
		gui.messageBox(_("Failed to locate your Microsoft Speech Macros folder. Please see the troublshooting part of the documentation for more details."),
			#Translators: Title for the microsoft speech recognitioninstalation  error dialog.
			_("Error installing MSR utilities"))


def onInstallMSRCommands(evt):
	dialog = gui.IndeterminateProgressDialog(gui.mainFrame,
		#Translators: Title for a dialog shown when Microsoft speech recognition Commands are being installed!
		_("Microsoft Speech Recognition Command Installation"),
		#Translators: Message shown in the progress dialog for MSR command installation.
		_("Please wait while microsoft speech recognition commands are installed."))
	try:
		gui.ExecAndPump(_onInstallMSRCommands)
	except: #Catch all, because if this fails, bad bad bad.
		log.error("DictationBridge commands failed to install!", exc_info=True)
	finally:
		dialog.done()

def onInstallDragonCommands(evt):
	#Translators: Warning about having custom commands already.
	goAhead = gui.messageBox(_("If you are on a computer with shared commands, and you have multiple users using these commands, this will override them. Please do not proceed unless you are sure you aren't sharing commands over a network. if you are, please read \"Importing Commands Into a Shared System\" in the Dictation Bridge documentation for manual steps.\nDo you want to proceed?"),
		#Translators: Warning dialog title.
		_("Warning: Potentially Dangerous Opperation Will be Performed"),
		wx.YES|wx.NO)
	if goAhead==wx.NO:
		return
	dialog = gui.IndeterminateProgressDialog(gui.mainFrame,
		#Translators: Title for a dialog shown when Dragon  Commands are being installed!
		_("Dragon Command Installation"),
		#Translators: Message shown in the progress dialog for dragon command installation.
		_("Please wait while Dragon commands are installed."))
	try:
		gui.ExecAndPump(_onInstallDragonCommands)
	except: #Catch all, because if this fails, bad bad bad.
		log.error("DictationBridge commands failed to install!", exc_info=True)
	finally:
		dialog.done()

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
		hookId = winUser.setWinEventHook(winUser.EVENT_OBJECT_HIDE, winUser.EVENT_OBJECT_HIDE, 0, eventCallback, pid, 0, 0)
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
	while True:
		i = text.find("\n")
		if i == -1:
			break
		if i > 0:
			speech.speakText(text[:i])
		if text[i:i + 2] == "\n\n":
			# Translators: The text which is spoken when a new paragraph is added.
			speech.speakText(_("new paragraph"))
			text = text[i + 2:]
		else:
			# Translators: The message spoken when a new line is entered.
			speech.speakText(_("new line"))
			text = text[i + 1:]
	if text != "":
		speech.speakText(text)
	braille.handler.handleCaretMove(api.getFocusObject())
	currentEntry = None
	requestWSRShowHideEvents()

def textInserted(hwnd, start, text):
	global currentEntry, autoFlushTimer
	log.debug("textInserted %r" % text)
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
cTextInsertedCallback = WINFUNCTYPE(None, HWND, LONG, c_wchar_p)(textInserted)

def textDeleted(hwnd, start, text):
	# Translators: The message spoken when a piece of text is deleted.
	speech.speakText(_("deleted %s" % text))
cTextDeletedCallback = WINFUNCTYPE(None, HWND, LONG, c_wchar_p)(textDeleted)

def commandCallback(command):
	if command in SPECIAL_COMMANDS:
		queueHandler.queueFunction(queueHandler.eventQueue, SPECIAL_COMMANDS[command])
		return
	inputCore.manager.executeGesture(DictationGesture(command))
cCommandCallback = WINFUNCTYPE(None, c_char_p)(commandCallback)

def debugLogCallback(msg):
	log.debug(msg)
cDebugLogCallback = WINFUNCTYPE(None, c_char_p)(debugLogCallback)

lastKeyDownTime = None

def patchKeyDownCallback():
	originalCallback = winInputHook.keyDownCallback
	def callback(*args, **kwargs):
		global lastKeyDownTime
		lastKeyDownTime = time.time()
		return originalCallback(*args, **kwargs)
	winInputHook.keyDownCallback = callback

masterDLL = None
installMenuItem = None

def initialize():
	global masterDLL, installMenuItem
	path = os.path.join(config.getUserDefaultConfigPath(), ".dbInstall")
	if os.path.exists(path):
		#First time reinstall of old code without the bail if updating code. Remove the temp file.
		#Also, import install tasks, then fake an install to get around the original path bug.
		sys.path.append(addonRootDir)
		import installTasks
		installTasks.onInstall(postPathBug = True)
		sys.path.remove(sys.path[-1])
		os.remove(path)
	dllPath = os.path.join(addonRootDir, "DictationBridgeMaster32.dll")
	masterDLL = windll.LoadLibrary(dllPath)
	masterDLL.DBMaster_SetTextInsertedCallback(cTextInsertedCallback)
	masterDLL.DBMaster_SetTextDeletedCallback(cTextDeletedCallback)
	masterDLL.DBMaster_SetCommandCallback(cCommandCallback)
	masterDLL.DBMaster_SetDebugLogCallback(cDebugLogCallback)
	if not masterDLL.DBMaster_Start():
		raise WinError()
	patchKeyDownCallback()
	toolsMenu = gui.mainFrame.sysTrayIcon.toolsMenu
	installMenu = wx.Menu()
	#Translators: The Install dragon Commands for NVDA  label.
	installDragonItem= installMenu.Append(wx.ID_ANY, _("Install Dragon Commands"))
	toolsMenu.Parent.Bind(wx.EVT_MENU, onInstallDragonCommands, installDragonItem)
	#Translators: The Install Microsoft Speech Recognition Commands for NVDA  label.
	installMSRItem= installMenu.Append(wx.ID_ANY, _("Install Microsoft Speech Recognition Commands"))
	toolsMenu.Parent.Bind(wx.EVT_MENU, onInstallMSRCommands, installMSRItem)
	#Translators: The Install commands submenu label.
	installMenuItem=toolsMenu.AppendSubMenu(installMenu, _("Install commands for Dictation Bridge"))

def terminate():
	global masterDLL
	if masterDLL is not None:
		masterDLL.DBMaster_Stop()
		masterDLL = None
	try:
		gui.mainFrame.sysTrayIcon.toolsMenu.Remove(gui.mainFrame.sysTrayIcon.toolsMenu.MenuItems.index(installMenuItem))
	except:
		pass

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
				# Translators: The text which is spoken when the spelling dialog is cleared.
				speech.speakText(_("cleared"))
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
		#Phrases which need translated in this function:
		# Translators: The text for "or say," Which is telling the user that they can say the next phrase.
		orSay = _("Or say")
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
					speech.speakText(orSay)
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
					speech.speakText(orSay)
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
		if lastKeyDownTime is None or (time.time() - lastKeyDownTime) >= 0.5 and ch != "":
			if obj.windowClassName != "ConsoleWindowClass":
				log.debug("typedCharacter %r %r" % (obj.windowClassName, ch))
				textInserted(obj.windowHandle, -1, ch)
			return
		nextHandler()

	def terminate(self):
		super(GlobalPlugin, self).terminate()
		terminate()
