Option Explicit

Public UseCalibrator As Boolean ' Kalibrátor állapotának tárolása
Public UseDMM As Boolean ' DMM állapotának tárolása
Public ioMgr As VisaComLib.ResourceManager
Public instrAny As VisaComLib.FormattedIO488
Public instrAny2 As VisaComLib.FormattedIO488

Public Sub InitializeDevices()
    ' Kalibrátor és DMM állapotának beolvasása
    UseCalibrator = CBool(Range("AA7").Value)
    UseDMM = CBool(Range("AA6").Value)

    ' Multimeter kapcsolat létrehozása, ha engedélyezve van
    If UseDMM Then
        Set ioMgr = New VisaComLib.ResourceManager
        Set instrAny = New VisaComLib.FormattedIO488
        Set instrAny.IO = ioMgr.Open(Range("V6").Value)
        instrAny.WriteString ("*CLS")
        instrAny.WriteString ("*RST")
    End If

    ' Kalibrátor kapcsolat létrehozása, ha használjuk
    If UseCalibrator Then
        Set instrAny2 = New VisaComLib.FormattedIO488
        Set instrAny2.IO = ioMgr.Open(Range("V7").Value)
        instrAny2.WriteString ("*CLS")
        instrAny2.WriteString ("*RST")
    End If
End Sub

Public Sub StartMeasurement()
    ' Indítás előtt az F oszlop törlése
    Range("F2:F103").Interior.ColorIndex = xlNone 

    ' Eszközök inicializálása
    Call InitializeDevices

    ' Mérés indítása
    Call RunMeasurement
End Sub
