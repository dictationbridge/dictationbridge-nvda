from appModuleHandler import AppModule
import api
import controlTypes
import tones
import ui
import windowUtils
from NVDAObjects.UIA import UIA
import NVDAObjects
import speech
import winUser
import time

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
