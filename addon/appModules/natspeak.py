import time
import os
import sys
import appModuleHandler
import speech
import NVDAObjects
import eventHandler
from NVDAObjects import NVDAObject
from NVDAObjects.window import Window
from NVDAObjects.behaviors import ProgressBar
from logHandler import log
import controlTypes
import colors
import textInfos
import ui
from winUser import user32
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from skipTranslation import translate
sys.path.remove(sys.path[-1])


class CustomList(NVDAObjects.NVDAObject):
	"""
	A display model list for the dragon vocabulary editor.
	"""

	columnNumber = 0
	role = controlTypes.ROLE_LIST
	#Translators: The name of the vocabulary items list in the vocabulary editor for Dragon.
	_name = _("Vocabulary Items")
	_addHeaderNextTime = False
	
	def _get_name(self):
		name = self._name
		name += " "
		name += self.columnHeaders[self.columnNumber]
		return name

	def _get_columnHeaders(self):
		ti = self.makeTextInfo(textInfos.POSITION_FIRST)
		ti.expand(textInfos.UNIT_LINE)
		#[Format Field, Header1, format field, column spacer, Format Field, header2.]
		columnInfo = ti.getTextWithFields()
		#second and last item.
		return (columnInfo[1], columnInfo[-1]) 

	def _get_value(self):
		try:
			ti = self.parent.makeTextInfo(textInfos.POSITION_SELECTION)
		except LookupError:
			return
		base = ""
		if self._addHeaderNextTime:
			base = self.columnHeaders[self.columnNumber] + " "
			self._addHeaderNextTime= False
		fields = ti.getTextWithFields()
		#Second, and last. First is a format field.
		fields = fields[1], fields[-1]
		base += fields[self.columnNumber]
		return base

	def script_moved(self, gesture):
		gesture.send()
		#It seems to take time to refresh.
		time.sleep(.05)
		eventHandler.executeEvent("valueChange", self)

	def _movementHelper(self, dir):
		if (
			(dir and self.columnNumber == 1) or
			(not dir and self.columnNumber==0)):
			ui.message(translate("Edge of table"))
			return
		if dir == 0:
			self.columnNumber-=1
		elif dir == 1:
			self.columnNumber+=1
		self._addHeaderNextTime = True
		eventHandler.executeEvent("valueChange", self)

	def script_right(self, gesture):
		self._movementHelper(1)

	def script_left(self, gesture):
		self._movementHelper(0)

	def getScript(self, gesture):
		#let's make quick nav single letter nav work!
		if len(gesture.identifiers[-1].lstrip("kb:" )) == 1:
			return self.script_moved
		return super(CustomList, self).getScript(gesture)

	__gestures = {
		"kb:upArrow" : "moved",
		"kb:downArrow" : "moved",
		"kb:pageUp" : "moved",
		"kb:pageDown" : "moved",
		"kb:home" : "moved",
		"kb:end" : "moved",
		"kb:rightArrow" : "right",
		"kb:leftArrow" : "left",
	}

class AppModule(appModuleHandler.AppModule):
	lastMicText = None


	def handleMicText(self, text):
		if text == self.lastMicText:
			return
		mOff="Dragon\'s microphone is off;"
		mOn="Normal mode: You can dictate and use voice"
		mSleep="The microphone is asleep;"
		if mOn in text:
			self.lastMicText = text
			ui.message("Dragon mic on")
		elif mOff in text:
			self.lastMicText = text
			ui.message("Dragon mic off")
		elif mSleep in text:
			self.lastMicText = text
			ui.message("Dragon sleeping")

	def event_nameChange(self, obj, nextHandler):
		if obj.windowControlID == 61923 and obj.windowClassName == u"Static":
			text = obj.name or ""
			self.handleMicText(text)
		nextHandler()

	def chooseNVDAObjectOverlayClasses (self, obj, clsList):
		if obj.windowClassName == u'CustomListBox':
			clsList.insert(0, CustomList)
		#The setup wizard uses progress bars to indicate mic status. 
		#This is potentially annoying, as the user is trying to speak, and hearing progress bars is distracting.
		try:
			if obj.windowClassName == u'msctls_progress32' and (
				(obj.parent.parent.role == controlTypes.ROLE_LIST) or 
				(obj.windowControlID == 1148)
				):
				clsList.remove(ProgressBar)
		except ValueError:
			pass

	def event_NVDAObject_init(self, obj):
		if obj.role == controlTypes.ROLE_BUTTON and obj.name == "" and obj.windowClassName == u'Button':
			#Turnary statements aren't used because it'll break translation.
			if obj.windowControlID == 202:
				#Translators: Button title for the dictation box settings.
				obj.name = _("Dictation Box Settings")
			elif obj.windowControlID == 9:
				#Translators: The Dictation boxes help button label.
				obj.name = _("Help")
