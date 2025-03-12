Option Explicit

' ---------------------- Késleltetés ----------------------
Public Sub Wait(ByVal Seconds As Single)
    Dim lMilliSeconds As Long
    lMilliSeconds = Seconds * 1000
    Sleep (lMilliSeconds)
End Sub

' ---------------------- DMM válasz konvertálása ----------------------
Public Function ConvertDMMResponse(ByVal response As String) As Double
    Dim numericValue As Double

    ' Trim szóközök eltávolítása
    response = Trim(response)

    ' Pont-vessző csere a helyi beállítások szerint
    response = Replace(response, ".", Application.DecimalSeparator)

    ' Próbáljuk konvertálni
    On Error Resume Next
    numericValue = CDbl(response)
    If Err.Number <> 0 Then
        MsgBox "Hiba: A DMM nem konvertálható számmá: " & response, vbCritical, "Mérési hiba"
        Err.Clear
        ConvertDMMResponse = 0
        Exit Function
    End If
    On Error GoTo 0

    ConvertDMMResponse = numericValue
End Function
