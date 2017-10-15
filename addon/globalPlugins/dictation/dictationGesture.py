import time
import functools
import api
import globalCommands
import inputCore
import scriptHandler

class DictationGesture(inputCore.InputGesture):
	_internalID = ""
	_scriptCount = 1


	def __init__(self, action):
		super(DictationGesture, self).__init__()
		count = 1
		if "|" in action:
			action,count = action.split("|")
		self._internalID = action
		self._scriptCount = int(count)

	def _get_identifiers(self):
		return ["db:{identifier}".format(identifier=self._internalID)]

	def _get_displayName(self):
		#fix me: Wouldn't it be nice to return the proper speech?
		id = list(self._internalID)
		for index, char in enumerate(id):
			if char == "_":
				id[index] = " "
				continue
				id[index] = char.lower()
		return "".join(id)

	def _getScriptFromObject(self, obj):
			func = getattr(obj, "script_%s" %self._internalID, None)
			return func

	def scriptWrapper(self, script, gesture):
		scriptHandler._lastScriptCount = gesture._scriptCount-1
		script(gesture)

	def _get_script_hacky(self):
		#Workaround until DB 1.1 when I fix NVDA to not need repeated scripts.
		#Mostly based on scriptHandler.findScript, but no globalGestureMapness
		focus = api.getFocusObject()
		if not focus:
			return None

		ti = focus.treeInterceptor
		if ti:
			func = self._getScriptFromObject(ti)
			if func and (not ti.passThrough or getattr(func,"ignoreTreeInterceptorPassThrough",False)):
				return (func, ti)

		# NVDAObject level.
		func = self._getScriptFromObject(focus)
		if func:
			return (func, focus)
		for obj in reversed(api.getFocusAncestors()):
			func = self._getScriptFromObject(obj)
			if func and getattr(func, 'canPropagate', False):
				return (func, obj)

		# Global commands.
		func = self._getScriptFromObject(globalCommands.commands)
		if func:
			return (func, globalCommands.commands)

	def _get_script(self):
		if inputCore.manager.isInputHelpActive:
			#Don't send it through the hack, because there's no useful help message for it.
			return self.script_hacky[0]
		else:
			script = self.script_hacky
			if script is None:
				return
			wrappedScript = functools.partial(self.scriptWrapper, script[0])
			try:
				wrappedScript.resumeSayAllMode = script[0].resumeSayAllMode
			except AttributeError:
				pass
			return wrappedScript