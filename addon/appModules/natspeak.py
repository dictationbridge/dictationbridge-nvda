from appModuleHandler import AppModule
from NVDAObjects.behaviors import ProgressBar
from NVDAObjects import NVDAObject
from logHandler import log
import api
from weakref import ref
import controlTypes
import speech

#scriptCategory = db_con.SCRCAT_DB

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
		print "start", obj.windowClassName, obj.windowControlID, obj.name, obj.parent.windowClassName, "done"
		if obj.windowControlID == 61923 and obj.windowClassName == u"Static":
			text = obj.name or ""
			self.handleMicText(text)
		elif obj.windowClassName == u"DgnResultsBoxWindow" and obj.windowControlID == 0 and obj.name:
			speech.speakMessage(obj.name)
		nextHandler()

	def chooseNVDAObjectOverlayClasses (self, obj, clsList):
		#The setup wizard uses progress bars to indicate mic status. 
		#This is potentially annoying, as the user is trying to speak, and hearing progress bars is distracting.
		try:
			if obj.windowClassName == u'msctls_progress32' and (
				(obj.parent.parent.role == controlTypes.ROLE_LIST) or 
				(obj.windowControlID == 1148)
				):
				clsList.remove(ProgressBar)
				clsList.insert(0, ProgressBarValueCacher)
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
		elif obj.windowClassName == u'Button' and obj.windowControlID == 12324:
			#Set focus to the next button to make life easier and because we want them to not mess the mic up after pressing around to find it.
			obj.setFocus()
