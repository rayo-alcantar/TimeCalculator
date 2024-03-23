import globalPluginHandler
import globalVars
import wx
import scriptHandler
import ui
import ctypes
from ctypes import wintypes
import datetime
import threading

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
        super().__init__(parent, title="Calculadora de tiempo")
        self.onCloseCallback = onCloseCallback  # Guarda la referencia al callback
        self.InitUI()
        self.SetSize((400, 200))
        self.Centre()
        self.Bind(wx.EVT_CLOSE, self.onClose)  # Vincula el evento de cierre
    def InitUI(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        self.select_hour= wx.StaticText(panel, label="Selecciona una hora entre 0 y 23:")
        self.hourCombo = wx.ComboBox(panel, choices=[f"{i:02d}" for i in range(24)], style=wx.CB_READONLY)
        self.select_minute=wx.StaticText(panel, label="Selecciona un minuto entre 0 y 59:")
        
        self.minuteCombo = wx.ComboBox(panel, choices=[f"{i:02d}" for i in range(60)], style=wx.CB_READONLY)
        self.minuteCombo.SetSelection(0)
        self.hourCombo.SetSelection(0)
        hbox1.Add(self.select_hour, flag=wx.RIGHT, border=8)
        hbox1.Add(self.hourCombo, flag=wx.RIGHT, border=8)
        hbox1.Add(self.select_minute, flag=wx.RIGHT, border=8)
        hbox1.Add(self.minuteCombo, flag=wx.RIGHT, border=8)
        
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        calcBtn = wx.Button(panel, label='&Calcular')
        calcBtn.Bind(wx.EVT_BUTTON, self.OnCalculate)
        closeBtn = wx.Button(panel, label='&Salir')
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
        result = f"Faltan {hours} horas y {minutes} minutos."
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

    @scriptHandler.script(description="Abre la calculadora de tiempo", gesture="kb:NVDA+alt+T")
    
    def script_openTimeCalculator(self, gesture):
        if not self.dialogOpen:
            threading.Thread(target=self.openDialog, daemon=True).start()
        else:
            ui.message("La calculadora de tiempo ya está abierta.")
    def openDialog(self):
        wx.CallAfter(self.showDialog)
    def showDialog(self):
        self.dialogOpen = True
        dialog = TimeCalculatorDialog(None, self.closeDialog)  # Pasa la función de cierre como callback
        dialog.ShowModal()

    def closeDialog(self):
        self.dialogOpen = False  # Actualiza el estado cuando se cierra el diálogo
    