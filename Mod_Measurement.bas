Attribute VB_Name = "Mod_Measurement"
Option Explicit

Public Sub StartMeasurement()
    Dim Counter As Integer, UtolsoSor As Integer, maxMeasurements As Integer
    Dim MeresiMod As String, FunkcioMod As String
    Dim avg As Double, stdev As Variant
    Dim sum As Double, sumSq As Double, i As Integer
    Dim measurements() As Double

    ' Utols� sor meghat�roz�sa
    With ActiveSheet
        UtolsoSor = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With

    ' M�r�si sz�m beolvas�sa a V10 cell�b�l
    maxMeasurements = CInt(Range("V10").Value)
    If maxMeasurements < 1 Then
        MsgBox "Hib�s m�r�si sz�m a V10 cell�ban!", vbCritical, "Hiba"
        Exit Sub
    End If

    ' Eszk�z�k inicializ�l�sa
    Call InitializeDevices

    ' M�r�si ciklus
    Counter = 2
    While Counter < UtolsoSor + 1
        MeresiMod = Cells(Counter, 1).Value
        FunkcioMod = Left(MeresiMod, 3)

        ' Ellen�rizni kell, hogy a m�r�si m�d t�mogatott-e
        If FunkcioMod <> "VDC" And FunkcioMod <> "VAC" Then
            MsgBox "Nincs ilyen m�r�si m�d be�ll�tva: " & MeresiMod, vbExclamation
            Exit Sub
        End If

        ' M�r�si parancs kiad�sa
        If FunkcioMod = "VDC" Then
            instrAny.WriteString ("FUNC " + Chr(34) + "VOLT:DC" + Chr(34))
            instrAny.WriteString ("Volt:DC:Range 10")
        ElseIf FunkcioMod = "VAC" Then
            instrAny.WriteString ("FUNC " + Chr(34) + "VOLT:AC" + Chr(34))
            instrAny.WriteString ("Volt:AC:Range 10")
        End If

        Wait (0.5)

        ' Kalibr�tor aktiv�l�s (ha sz�ks�ges)
        If UseCalibrator Then
            instrAny2.WriteString ("OUT 1V")
            instrAny2.WriteString ("OPER")
        End If

        ' T�bbsz�r�s m�r�s �s statisztika sz�m�t�sa
        sum = 0
        sumSq = 0
        ReDim measurements(1 To maxMeasurements)

        For i = 1 To maxMeasurements
            instrAny.WriteString ("Measure:Volt:DC?")
            measurements(i) = ConvertDMMResponse(instrAny.ReadString)
            sum = sum + measurements(i)
            sumSq = sumSq + measurements(i) * measurements(i)
            Wait (0.2) ' K�sleltet�s a m�r�sek k�z�tt
        Next i

        ' �tlag sz�m�t�sa
        avg = sum / maxMeasurements

        ' Sz�r�s sz�m�t�sa
        If maxMeasurements > 1 Then
            stdev = Sqr((sumSq / maxMeasurements) - (avg * avg))
        Else
            stdev = "N/A"
        End If

        ' Eredm�nyek be�r�sa az Excelbe
        Cells(Counter, "F").Value = avg
        Cells(Counter, "J").Value = stdev

        ' Kalibr�tor kikapcsol�sa (ha sz�ks�ges)
        If UseCalibrator Then instrAny2.WriteString ("STBY")

        ' K�vetkez� sor
        Counter = Counter + 1
    Wend

    ' Eszk�z�k lez�r�sa
    instrAny.IO.Close
    If UseCalibrator Then instrAny2.IO.Close

    MsgBox "M�r�s befejez�d�tt.", vbInformation, "K�sz"
End Sub

