# -*- coding: utf-8 -*-
# Copyright (C) 2024 Ángel Alcantar  <rayoalcantar@gmail.com>
# This file is covered by the GNU General Public License.
#
# import the necessary modules (NVDA)
import gui
import globalPluginHandler
import globalVars
import wx
import scriptHandler
import ui
import ctypes
from ctypes import wintypes
import datetime
import threading
import addonHandler
addonHandler.initTranslation()


# Definiciones para el manejo del portapapeles
OpenClipboard = ctypes.windll.user32.OpenClipboard
EmptyClipboard = ctypes.windll.user32.EmptyClipboard
SetClipboardData = ctypes.windll.user32.SetClipboardData
CloseClipboard = ctypes.windll.user32.CloseClipboard
GlobalAlloc = ctypes.windll.kernel32.GlobalAlloc
GlobalLock = ctypes.windll.kernel32.GlobalLock
GlobalUnlock = ctypes.windll.kernel32.GlobalUnlock
GMEM_MOVEABLE = 0x0002
CF_UNICODETEXT = 13

def set_clipboard_text(text):
    OpenClipboard(None)
    EmptyClipboard()
    hMem = GlobalAlloc(GMEM_MOVEABLE, len(text) * 2 + 2)
    pMem = GlobalLock(hMem)
    ctypes.cdll.msvcrt.wcscpy(ctypes.c_wchar_p(pMem), text)
    GlobalUnlock(hMem)
    SetClipboardData(CF_UNICODETEXT, hMem)
    CloseClipboard()

class TimeCalculatorDialog(wx.Dialog):
    def __init__(self, parent, onCloseCallback):
        # Translators: Title of the Time Calculator dialog
        super().__init__(parent, title=_("Calculadora de tiempo"))
        self.onCloseCallback = onCloseCallback  # Guarda la referencia al callback
        self.InitUI()
        self.SetSize((400, 200))
        self.Centre()
        self.Bind(wx.EVT_CLOSE, self.onClose)  # Vincula el evento de cierre
    def InitUI(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        # Translators: Label for selecting an hour
        self.select_hour= wx.StaticText(panel, label=_("Selecciona una hora entre 0 y 23:"))
        self.hourCombo = wx.ComboBox(panel, choices=[f"{i:02d}" for i in range(24)], style=wx.CB_READONLY)
        # Translators: Label for selecting a minute
        self.select_minute=wx.StaticText(panel, label=_("Selecciona un minuto entre 0 y 59:"))
        
        self.minuteCombo = wx.ComboBox(panel, choices=[f"{i:02d}" for i in range(60)], style=wx.CB_READONLY)
        self.minuteCombo.SetSelection(0)
        self.hourCombo.SetSelection(0)
        hbox1.Add(self.select_hour, flag=wx.RIGHT, border=8)
        hbox1.Add(self.hourCombo, flag=wx.RIGHT, border=8)
        hbox1.Add(self.select_minute, flag=wx.RIGHT, border=8)
        hbox1.Add(self.minuteCombo, flag=wx.RIGHT, border=8)
        
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        # Translators: Label for the calculate button
        calcBtn = wx.Button(panel, label=_('&Calcular'))
        calcBtn.Bind(wx.EVT_BUTTON, self.OnCalculate)
        #Translators: button to close or exit the interface.
        closeBtn = wx.Button(panel, label=_('&Salir'))
        closeBtn.Bind(wx.EVT_BUTTON, lambda e: self.onClose(e))
        hbox2.Add(calcBtn)
        hbox2.Add(closeBtn, flag=wx.LEFT, border=5)
        
        vbox.Add(hbox1, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, border=10)
        vbox.Add(hbox2, flag=wx.ALIGN_CENTER|wx.TOP|wx.BOTTOM, border=10)
        panel.SetSizer(vbox)

    def OnCalculate(self, event):
        now = datetime.datetime.now()
        target_hour = int(self.hourCombo.GetValue())
        target_minute = int(self.minuteCombo.GetValue())
        target = datetime.datetime(now.year, now.month, now.day, target_hour, target_minute)
        if target < now:
            target += datetime.timedelta(days=1)
        diff = target - now
        hours, remainder = divmod(diff.seconds, 3600)
        minutes = remainder // 60
        # Translators: Message displaying the time remaining. {0} is replaced by the number of hours, and {1} by the number of minutes.
        result = _("Faltan {0} horas y {1} minutos.").format(hours, minutes)
        set_clipboard_text(result)
        ui.message(result)
    def onClose(self, event):
        self.onCloseCallback()  # Llama al callback
        self.Destroy()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = "Calculadora de Tiempo"
    dialogOpen = False

    def __init__(self):
        # Verifica si NVDA se ejecuta en un entorno seguro
        if globalVars.appArgs.secure:
            return
        super(GlobalPlugin, self).__init__()
        # Añade un ítem al menú Herramientas correctamente dentro de __init__
        self.menuItem = gui.mainFrame.sysTrayIcon.toolsMenu.Append(wx.ID_ANY,
                                                                    "Calculadora de tiempo",
                                                                    "Abre la calculadora de tiempo")
        gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self.onOpenDialog, self.menuItem)

    def onOpenDialog(self, event):
        # Llama a la función que abre el diálogo directamente, verificando si ya está abierto
        if not self.dialogOpen:
            threading.Thread(target=self.openDialog, daemon=True).start()
        else:
            #Translators: indicates that the time calculator is already avert.
            ui.message(_("La calculadora de tiempo ya está abierta."))

    @scriptHandler.script(description="Abre la calculadora de tiempo", gesture="kb:NVDA+alt+T")
    def script_openTimeCalculator(self, gesture):
        # Este método activa la misma lógica que el ítem del menú
        self.onOpenDialog(None)

    def openDialog(self):
        # Este método prepara la apertura del diálogo en el hilo de la GUI
        wx.CallAfter(self.showDialog)

    def showDialog(self):
        # Este método muestra el diálogo y configura el estado de `dialogOpen`
        self.dialogOpen = True
        dialog = TimeCalculatorDialog(None, self.closeDialog)
        dialog.ShowModal()

    def closeDialog(self):
        # Este método actualiza el estado cuando el diálogo se cierra
        self.dialogOpen = False

    def terminate(self):
        # Este método asegura que el ítem del menú se elimine cuando el complemento se desactive
        if hasattr(self, 'menuItem'):
            gui.mainFrame.sysTrayIcon.toolsMenu.Remove(self.menuItem.GetId())
