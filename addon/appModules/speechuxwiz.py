from appModuleHandler import AppModule
import api
import controlTypes
import tones
import ui
import windowUtils
from NVDAObjects.behaviors import Dialog
from NVDAObjects.UIA import UIA
import NVDAObjects
import winUser
import time
from logHandler import log

class Wizard(Dialog):
	role = controlTypes.ROLE_DIALOG

class AppModule(AppModule):
	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if obj.windowClassName == "NativeHWNDHost" and obj.role == controlTypes.ROLE_PANE:
			clsList.insert(0, Wizard)

	def event_NVDAObject_init(self, obj):
		#When waiting for a second or two on the headset microphone radio button, an unwanted window gains focus.
		if obj.role == controlTypes.ROLE_WINDOW and isinstance(obj, UIA) and obj.UIAElement.cachedClassName == u'CCRadioButton':
			obj.shouldAllowUIAFocusEvent = False
		if obj.role == controlTypes.ROLE_STATICTEXT and obj.description:
			obj.description = None

	def readTrainingText(self):
		window = api.getForegroundObject()
		for descendant in window.recursiveDescendants:
			if not isinstance(descendant, UIA):
				continue
			try:
				automationID = descendant.UIAElement.currentAutomationID
			except:
				continue
			if automationID == "txttrain":
				api.setNavigatorObject(descendant)
				ui.message(descendant.name)
				break

	def script_readTrainingText(self, gesture):
		self.readTrainingText()

	__gestures = {
		"kb:`": "readTrainingText",
	}

	def event_nameChange(self, obj, nextHandler):
		if obj.role == controlTypes.ROLE_STATICTEXT and obj.windowClassName == "DirectUIHWND":
			self.readTrainingText()
		nextHandler()

	def event_valueChange(self, obj, nextHandler):
		if obj.role == controlTypes.ROLE_PROGRESSBAR:
			self.readTrainingText()
		nextHandler()
