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
    ' Inicializ�l�s (DMM �s Kalibr�tor �llapot�nak bet�lt�se)
    Call InitializeDevices

    ' Ha sem a DMM, sem a Kalibr�tor nincs bekapcsolva, le�ll
    If Not UseDMM And Not UseCalibrator Then
        MsgBox "Nincs akt�v m�r�eszk�z! Enged�lyezd a DMM-et (AA6) vagy a Kalibr�tort (AA7).", vbCritical, "M�r�si hiba"
        Exit Sub
    End If

    ' Ind�tsd el a teljes m�r�st
    Call StartMeasurement
End Sub


