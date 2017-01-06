import os
import sys
import _winreg

import win32con

import addonHandler
import winUser

def onInstall():
	#Add ourself to the path, so that commands when spoken can be queried to us.

	key = _winreg.OpenKeyEx(_winreg.HKEY_CURRENT_USER, "Environment", 0, _winreg.KEY_READ | _winreg.KEY_WRITE)
	try:
		value, typ = _winreg.QueryValueEx(key, "Path")
	except:
		value, typ = None, _winreg.REG_EXPAND_SZ
	if value is None:
		value = ""
	dir = os.path.dirname(__file__)
	if not isinstance(dir, unicode):
		dir = dir.decode(sys.getfilesystemencoding())
	dir = dir.replace(addonHandler.ADDON_PENDINGINSTALL_SUFFIX, "")
	if value.lower().find(dir.lower()) == -1:
		if value != "":
			value += ";"
		value += dir
		_winreg.SetValueEx(key, "Path", None, typ, value)
		winUser.sendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, u"Environment")

def onUninstall():
	key = _winreg.OpenKeyEx(_winreg.HKEY_CURRENT_USER, "Environment", 0, _winreg.KEY_READ | _winreg.KEY_WRITE)
	try:
		value, typ = _winreg.QueryValueEx(key, "Path")
	except:
		return
	if value is None or value == "":
		return
	dir = os.path.dirname(__file__)
	if not isinstance(dir, unicode):
		dir = dir.decode(sys.getfilesystemencoding())
	dir = dir.replace(addonHandler.DELETEDIR_SUFFIX, "")
	if value.find(dir) != -1:
		value = value.replace(";" + dir, "")
		value = value.replace(dir + ";", "")
		value = value.replace(dir, "")
		_winreg.SetValueEx(key, "Path", None, typ, value)
		winUser.sendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, u"Environment")
