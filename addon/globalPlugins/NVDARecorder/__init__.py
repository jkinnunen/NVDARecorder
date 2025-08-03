# -*- coding: UTF-8 -*-
# Copyright (C) 2021-2025 Rui Fontes <rui.fontes@tiflotecnia.com> and Ângelo Abrantes <ampa4374@gmail.com>
# Based on the work of 高生旺 <coscell@gmail.com> with the same name
# and some code copied from the work of James Scholes on speechHistory
# This file is covered by the GNU General Public License.

# import the necessary modules.
import globalPluginHandler
import globalVars
import api
import speech
import speechViewer
from eventHandler import FocusLossCancellableSpeechCommand
import wx
import gui
import ui
import os
import ctypes
from scriptHandler import script
import addonHandler

addonHandler.initTranslation()

# Path to the output text file
_NRIniFile = os.path.join(globalVars.appArgs.configPath, "NVDARecord.txt")
# Recording state flag
start = False
# Determine which module exposes the speak function
try:
	# NVDA => 2021
	_smod = speech.speech
except AttributeError:
	# NVDA < 2021
	_smod = speech

# Save the original speak method
oldSpeak = _smod.speak

# Buffer to accumulate spoken text
contents = ""

def getSequenceText(seq):
	"""Extract only string items from a SpeechSequence."""
	seq = [command for command in seq if not isinstance(command, FocusLossCancellableSpeechCommand)]
	return speechViewer.SPEECH_ITEM_SEPARATOR.join(
		[item for item in seq if isinstance(item, str)]
	)

def mySpeak(sequence, *args, **kwargs):
	"""Wrapped speak method that records speech output."""
	global contents
	# Call the original speak method
	oldSpeak(sequence, *args, **kwargs)
	# Extract and append text
	text = getSequenceText(sequence)
	if text:
		if "\n" not in text:
			text += "\n"
		contents += text


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super(GlobalPlugin, self).__init__()
		# Do not activate in secure screens
		if globalVars.appArgs.secure:
			return

	@script(
		# Translators: Message to be announced during Keyboard Help
		description=_("Activate/deactivate recording on NVDARecorder"),
		gesture="kb:alt+numpadplus"
	)
	def script_record(self, gesture):
		"""Toggle recording of all NVDA speech output."""
		global start, contents
		start = not start
		if start:
			# Start recording: patch the speak method
			ui.message(_("Start recording"))
			_smod.speak = mySpeak
		else:
			# Stop recording: restore original speak
			_smod.speak = oldSpeak
			ui.message(_("Recording stopped"))
			# Write accumulated text to file
			with open(_NRIniFile, "w", encoding="utf-8") as f:
				f.write(contents)
			# Display the results dialog
			gui.mainFrame._popupSettingsDialog(ShowResults)
			# Clear buffer for the next session
			contents = ""


class ShowResults(wx.Dialog):
	"""Dialog to display recorded speech text."""
	def __init__(self, *args, **kwargs):
		kwargs["style"] = kwargs.get("style", 0) | wx.DEFAULT_DIALOG_STYLE
		super(ShowResults, self).__init__(*args, **kwargs)
		self.SetTitle(_("NVDA Recorder"))

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# Static text label
		label = wx.StaticText(self, wx.ID_ANY, _("Here is the recorded text:"))
		mainSizer.Add(label, 0, wx.ALL, 5)

		# Read-only multi-line text control
		global contents
		self.text_ctrl = wx.TextCtrl(
			self, wx.ID_ANY, contents,
			size=(550, 350),
			style=wx.TE_MULTILINE | wx.TE_READONLY
		)
		mainSizer.Add(self.text_ctrl, 1, wx.EXPAND | wx.ALL, 5)

		# Standard dialog buttons
		btnSizer = wx.StdDialogButtonSizer()
		btnFolder = wx.Button(self, wx.ID_ANY, _("Open NVDARecord.txt's folder"))
		self.Bind(wx.EVT_BUTTON, self.openFolder, btnFolder)
		btnSizer.AddButton(btnFolder)

		btnFile = wx.Button(self, wx.ID_ANY, _("Open NVDARecord.txt"))
		self.Bind(wx.EVT_BUTTON, self.openTXTFile, btnFile)
		btnSizer.AddButton(btnFile)

		btnCopy = wx.Button(self, wx.ID_ANY, _("Copy to clipboard"))
		self.Bind(wx.EVT_BUTTON, self.copyToClip, btnCopy)
		btnSizer.AddButton(btnCopy)

		btnClose = wx.Button(self, wx.ID_CLOSE, "")
		self.Bind(wx.EVT_BUTTON, self.quit, btnClose)
		btnSizer.AddButton(btnClose)

		btnSizer.Realize()
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

		self.SetSizer(mainSizer)
		mainSizer.Fit(self)
		self.SetEscapeId(btnClose.GetId())
		self.Layout()
		self.CentreOnScreen()

	def openFolder(self, event):
		"""Open the folder containing the text file."""
		self.Destroy()
		os.startfile(globalVars.appArgs.configPath)

	def openTXTFile(self, event):
		"""Open the recorded text file with default application."""
		self.Destroy()
		ctypes.windll.shell32.ShellExecuteW(None, "open", _NRIniFile, None, None, 10)

	def copyToClip(self, event):
		"""Copy the recorded text to the clipboard."""
		self.Destroy()
		api.copyToClip(self.text_ctrl.GetValue())

	def quit(self, event):
		"""Close the dialog."""
		self.Destroy()
