Attribute VB_Name = "Mod_Initialization"
Option Explicit

Public UseDMM As Boolean ' DMM állapotának tárolása
Public UseCalibrator As Boolean ' Kalibrátor állapotának tárolása
Public ioMgr As VisaComLib.ResourceManager
Public instrAny As VisaComLib.FormattedIO488
Public instrAny2 As VisaComLib.FormattedIO488

Public Sub InitializeDevices()
    ' DMM és Kalibrátor állapotának beolvasása az AA6 és AA7 cellákból
    UseDMM = CBool(Range("AA6").Value)
    UseCalibrator = CBool(Range("AA7").Value)

    ' Ha a cellák üresek vagy nem megfelelõ értéket tartalmaznak, alapértelmezett értékre állítjuk
    If IsEmpty(Range("AA6").Value) Or Not IsNumeric(Range("AA6").Value) Then UseDMM = False
    If IsEmpty(Range("AA7").Value) Or Not IsNumeric(Range("AA7").Value) Then UseCalibrator = False

    ' DMM kapcsolat létrehozása, ha aktív
    If UseDMM Then
        Set ioMgr = New VisaComLib.ResourceManager
        Set instrAny = New VisaComLib.FormattedIO488
        Set instrAny.IO = ioMgr.Open(Range("V6").Value)
        instrAny.WriteString ("*CLS")
        instrAny.WriteString ("*RST")
    End If

    ' Kalibrátor kapcsolat létrehozása, ha aktív
    If UseCalibrator Then
        If ioMgr Is Nothing Then Set ioMgr = New VisaComLib.ResourceManager ' Ha még nem lett létrehozva
        Set instrAny2 = New VisaComLib.FormattedIO488
        Set instrAny2.IO = ioMgr.Open(Range("V7").Value)
        instrAny2.WriteString ("*CLS")
        instrAny2.WriteString ("*RST")
    End If
End Sub


