import time
from appModuleHandler import AppModule
from NVDAObjects.behaviors import ProgressBar
import api
import controlTypes
import tones
import ui
import windowUtils
from NVDAObjects.UIA import UIA
import NVDAObjects
import speech
import winUser

class AppModule(AppModule):
	lastMicText = None

	def handleMicText(self, text):
		if text == self.lastMicText:
			return
		mOff="Dragon\'s microphone is off;"
		mOn="Normal mode: You can dictate and use voice"
		mSleep="The microphone is asleep;"
		if mOn in text:
			self.lastMicText = text
			speech.speakText("Dragon mic on")
		elif mOff in text:
			self.lastMicText = text
			speech.speakText("Dragon mic off")
		elif mSleep in text:
			self.lastMicText = text
			speech.speakText("Dragon sleeping")

	def event_nameChange(self, obj, nextHandler):
		text = obj.name or ""
		self.handleMicText(text)
		nextHandler()

	def chooseNVDAObjectOverlayClasses (self, obj, clsList):
		#The setup wizard uses progress bars to indicate mic status. 
		#This is potentially annoying, as the user is trying to speak, and hearing progress bars is distracting.
		try:
			if obj.windowClassName == u'msctls_progress32' and (
				(	obj.parent.parent.role == controlTypes.ROLE_LIST) or 
				(obj.windowControlID == 1148)
				):
				clsList.remove(ProgressBar)
		except ValueError:
			pass
