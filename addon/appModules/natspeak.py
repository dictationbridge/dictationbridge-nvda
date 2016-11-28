########--------remove me.
import core #remove me
import re
import IAccessibleHandler
from comtypes import COMError
#########-------/remove me.
import time
from appModuleHandler import AppModule
from NVDAObjects.behaviors import ProgressBar
from logHandler import log
import api
import controlTypes
import tones
import ui
import eventHandler
import windowUtils
from NVDAObjects.UIA import UIA
import NVDAObjects
import speech
import winUser

class CorrectionChoiceDialog(NVDAObjects.NVDAObject):

	def initOverlayClass(self):
		#print "first-child-name", d.getChild(0).name
		#print "Name of me", d.name
			ch = self.parent.getChild(6)
		
		children = self.getChildren(self.parent.IAccessibleObject, self.parent.IAccessibleChildID)
		print "initializing a dialog"
		print len(children)
		print children
		print "\n".join([str(i[0].accName(i[1])) for i in children])
		print "done"
		#log.debug("\n".join([str(i.accName()) for i in ch]))

	def getChildren(self, IAObj,IACID,depth=0):
		final = [(IAObj, IACID)]
		try:
			children = IAccessibleHandler.accessibleChildren(IAObj, 0, IAObj.accChildCount)
			for i in children:
				if i[1] == 0:
					final+=self.getChildren(i[0], i[1], depth+1)
				else:
					final.append(i)
		except COMError:
			pass
		return final



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
		#if obj.windowClassName == u'ClsChoiceBoxWindow':
		if obj.role == controlTypes.ROLE_DIALOG and obj.windowClassName == u'DgnResultsBoxWindow':
			clsList.insert(0, CorrectionChoiceDialog)


	def __init__(self, *args, **kwargs):
		super(AppModule, self).__init__(*args, **kwargs)
		eventHandler.requestEvents("show", self.processID, u'DgnResultsBoxWindow')