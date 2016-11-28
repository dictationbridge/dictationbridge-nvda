from appModuleHandler import AppModule
import controlTypes
import speech
import re

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
			speech.speakText("Dragon mic on")
		elif mOff in text:
			speech.speakText("Dragon mic off")
		elif mSleep in text:
			speech.speakText("Dragon sleeping")

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
		automationId = obj.UIAElement.CachedAutomationID
		if automationId == u'cbRecognitionMode':
			obj.name = obj.previous.name
		if self.RE_BAD_MENU_ITEMS.match(automationId):
			obj.role = controlTypes.ROLE_MENU
