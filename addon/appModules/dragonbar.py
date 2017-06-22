from appModuleHandler import AppModule
import controlTypes
import ui
import re
from comtypes import COMError

class AppModule(AppModule):
	lastFlashRightText = None

	def flashRightTextChanged(self, obj):
		text = obj.name
		if not text:
			return
		if self.lastFlashRightText == text:
			return
		self.lastFlashRightText = text
		mOff="Dragon\'s microphone is off;"
		mOn="Normal mode: You can dictate and use voice"
		mSleep="The microphone is asleep;"
		if mOn in text:
			ui.message("Dragon mic on")
		elif mOff in text:
			ui.message("Dragon mic off")
		elif mSleep in text:
			ui.message("Dragon sleeping")

	def event_nameChange(self, obj, nextHandler):
		try:
			automationId = obj.UIAElement.currentAutomationID
		except:
			automationId = None
		if automationId == "txtFlashRight":
			self.flashRightTextChanged(obj)
		nextHandler()

	RE_BAD_MENU_ITEMS= re.compile(r"^mi_?("+"|".join([
		"Top",
		"Profile",
		"Tools",
		"Vocabulary",
		"Audio",
		"Help",
		])+")$")
	def event_NVDAObject_init(self, obj):
		try:
			automationId = obj.UIAElement.CachedAutomationID
		except (COMError, AttributeError):
			return
		if automationId == u'cbRecognitionMode':
			#Fix a combobox with no label.
			obj.name = obj.previous.name
		if self.RE_BAD_MENU_ITEMS.match(automationId):
			obj.role = controlTypes.ROLE_MENU
