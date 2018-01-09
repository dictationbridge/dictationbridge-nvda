# -*- coding: UTF-8 -*-

# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

# Full getext (please don't change)
_ = lambda x : x

# Add-on information variables
addon_info = {
	# for previously unpublished addons, please follow the community guidelines at:
	# https://bitbucket.org/nvdaaddonteam/todo/raw/master/guideLines.txt
	# add-on Name, internal for nvda
	"addon_name" : "DictationBridge",
	# Add-on summary, usually the user visible name of the addon.
	# Translators: Summary for this add-on to be shown on installation and add-on information.
	"addon_summary" : _("DictationBridge"),
	# Add-on description
	# Translators: Long description to be shown for this add-on on add-on information from add-ons manager
	"addon_description" : _("""Add-on to support Windows Speech Recognition and Dragon NaturallySpeaking"""),
	# version
	"addon_version" : "1.0",
	# Author(s)
	"addon_author" : u"The DictationBridge Team <http://dictationbridge.com/>",
	# URL for the add-on documentation support
	"addon_url" : u"http://dictationbridge.com/",
	# Documentation file name
	"addon_docFileName" : "readme.html",
}


import os.path

# Define the python files that are the sources of your add-on.
# You can use glob expressions here, they will be expanded.
pythonSources = []
for dirpath, dirnames, filenames in os.walk("addon"):
	for fi in filenames:
		if fi.endswith(".py"):
			pythonSources.append(os.path.join(dirpath, fi))
# Files that contain strings for translation. Usually your python sources
i18nSources = pythonSources + ["buildVars.py"]

# Files that will be ignored when building the nvda-addon file
# Paths are relative to the addon directory, not to the root directory of your addon sources.
excludedFiles = []
