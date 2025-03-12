VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Munka1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub RunMeasurement()
    ' Inicializálás (DMM és Kalibrátor állapotának betöltése)
    Call InitializeDevices

    ' Ha sem a DMM, sem a Kalibrátor nincs bekapcsolva, leáll
    If Not UseDMM And Not UseCalibrator Then
        MsgBox "Nincs aktív mérõeszköz! Engedélyezd a DMM-et (AA6) vagy a Kalibrátort (AA7).", vbCritical, "Mérési hiba"
        Exit Sub
    End If

    ' Indítsd el a teljes mérést
    Call StartMeasurement
End Sub


