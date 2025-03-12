Attribute VB_Name = "Mod_Utilities"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' K�sleltet�s megval�s�t�sa
Public Sub Wait(ByVal Seconds As Single)
    Dim lMilliSeconds As Long
    lMilliSeconds = Seconds * 1000
    Sleep (lMilliSeconds)
End Sub

' DMM v�lasz konvert�l�sa sz�m�rt�kk�
Public Function ConvertDMMResponse(ByVal response As String) As Double
    Dim numericValue As Double

    ' Trim sz�k�z�k elt�vol�t�sa
    response = Trim(response)

    ' Pont-vessz� csere a helyi be�ll�t�sok szerint
    response = Replace(response, ".", Application.DecimalSeparator)

    ' Konvert�l�s lebeg�pontos sz�mm�
    On Error Resume Next
    numericValue = CDbl(response)
    If Err.Number <> 0 Then
        MsgBox "Hiba: A DMM nem konvert�lhat� sz�mm�: " & response, vbCritical, "M�r�si hiba"
        Err.Clear
        ConvertDMMResponse = 0
    End If
    On Error GoTo 0

    ConvertDMMResponse = numericValue
End Function

