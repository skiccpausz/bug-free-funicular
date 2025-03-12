Attribute VB_Name = "Mod_Initialization"
Option Explicit

Public UseDMM As Boolean ' DMM �llapot�nak t�rol�sa
Public UseCalibrator As Boolean ' Kalibr�tor �llapot�nak t�rol�sa
Public ioMgr As VisaComLib.ResourceManager
Public instrAny As VisaComLib.FormattedIO488
Public instrAny2 As VisaComLib.FormattedIO488

Public Sub InitializeDevices()
    ' DMM �s Kalibr�tor �llapot�nak beolvas�sa az AA6 �s AA7 cell�kb�l
    UseDMM = CBool(Range("AA6").Value)
    UseCalibrator = CBool(Range("AA7").Value)

    ' Ha a cell�k �resek vagy nem megfelel� �rt�ket tartalmaznak, alap�rtelmezett �rt�kre �ll�tjuk
    If IsEmpty(Range("AA6").Value) Or Not IsNumeric(Range("AA6").Value) Then UseDMM = False
    If IsEmpty(Range("AA7").Value) Or Not IsNumeric(Range("AA7").Value) Then UseCalibrator = False

    ' DMM kapcsolat l�trehoz�sa, ha akt�v
    If UseDMM Then
        Set ioMgr = New VisaComLib.ResourceManager
        Set instrAny = New VisaComLib.FormattedIO488
        Set instrAny.IO = ioMgr.Open(Range("V6").Value)
        instrAny.WriteString ("*CLS")
        instrAny.WriteString ("*RST")
    End If

    ' Kalibr�tor kapcsolat l�trehoz�sa, ha akt�v
    If UseCalibrator Then
        If ioMgr Is Nothing Then Set ioMgr = New VisaComLib.ResourceManager ' Ha m�g nem lett l�trehozva
        Set instrAny2 = New VisaComLib.FormattedIO488
        Set instrAny2.IO = ioMgr.Open(Range("V7").Value)
        instrAny2.WriteString ("*CLS")
        instrAny2.WriteString ("*RST")
    End If
End Sub


