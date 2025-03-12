Attribute VB_Name = "Mod_Utilities"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Késleltetés megvalósítása
Public Sub Wait(ByVal Seconds As Single)
    Dim lMilliSeconds As Long
    lMilliSeconds = Seconds * 1000
    Sleep (lMilliSeconds)
End Sub

' DMM válasz konvertálása számértékké
Public Function ConvertDMMResponse(ByVal response As String) As Double
    Dim numericValue As Double

    ' Trim szóközök eltávolítása
    response = Trim(response)

    ' Pont-vesszõ csere a helyi beállítások szerint
    response = Replace(response, ".", Application.DecimalSeparator)

    ' Konvertálás lebegõpontos számmá
    On Error Resume Next
    numericValue = CDbl(response)
    If Err.Number <> 0 Then
        MsgBox "Hiba: A DMM nem konvertálható számmá: " & response, vbCritical, "Mérési hiba"
        Err.Clear
        ConvertDMMResponse = 0
    End If
    On Error GoTo 0

    ConvertDMMResponse = numericValue
End Function

