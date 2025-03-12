Attribute VB_Name = "Mod_Measurement"
Option Explicit

Public Sub StartMeasurement()
    Dim Counter As Integer, UtolsoSor As Integer, maxMeasurements As Integer
    Dim MeresiMod As String, FunkcioMod As String
    Dim avg As Double, stdev As Variant
    Dim sum As Double, sumSq As Double, i As Integer
    Dim measurements() As Double

    ' Utolsó sor meghatározása
    With ActiveSheet
        UtolsoSor = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With

    ' Mérési szám beolvasása a V10 cellából
    maxMeasurements = CInt(Range("V10").Value)
    If maxMeasurements < 1 Then
        MsgBox "Hibás mérési szám a V10 cellában!", vbCritical, "Hiba"
        Exit Sub
    End If

    ' Eszközök inicializálása
    Call InitializeDevices

    ' Mérési ciklus
    Counter = 2
    While Counter < UtolsoSor + 1
        MeresiMod = Cells(Counter, 1).Value
        FunkcioMod = Left(MeresiMod, 3)

        ' Ellenõrizni kell, hogy a mérési mód támogatott-e
        If FunkcioMod <> "VDC" And FunkcioMod <> "VAC" Then
            MsgBox "Nincs ilyen mérési mód beállítva: " & MeresiMod, vbExclamation
            Exit Sub
        End If

        ' Mérési parancs kiadása
        If FunkcioMod = "VDC" Then
            instrAny.WriteString ("FUNC " + Chr(34) + "VOLT:DC" + Chr(34))
            instrAny.WriteString ("Volt:DC:Range 10")
        ElseIf FunkcioMod = "VAC" Then
            instrAny.WriteString ("FUNC " + Chr(34) + "VOLT:AC" + Chr(34))
            instrAny.WriteString ("Volt:AC:Range 10")
        End If

        Wait (0.5)

        ' Kalibrátor aktiválás (ha szükséges)
        If UseCalibrator Then
            instrAny2.WriteString ("OUT 1V")
            instrAny2.WriteString ("OPER")
        End If

        ' Többszörös mérés és statisztika számítása
        sum = 0
        sumSq = 0
        ReDim measurements(1 To maxMeasurements)

        For i = 1 To maxMeasurements
            instrAny.WriteString ("Measure:Volt:DC?")
            measurements(i) = ConvertDMMResponse(instrAny.ReadString)
            sum = sum + measurements(i)
            sumSq = sumSq + measurements(i) * measurements(i)
            Wait (0.2) ' Késleltetés a mérések között
        Next i

        ' Átlag számítása
        avg = sum / maxMeasurements

        ' Szórás számítása
        If maxMeasurements > 1 Then
            stdev = Sqr((sumSq / maxMeasurements) - (avg * avg))
        Else
            stdev = "N/A"
        End If

        ' Eredmények beírása az Excelbe
        Cells(Counter, "F").Value = avg
        Cells(Counter, "J").Value = stdev

        ' Kalibrátor kikapcsolása (ha szükséges)
        If UseCalibrator Then instrAny2.WriteString ("STBY")

        ' Következõ sor
        Counter = Counter + 1
    Wend

    ' Eszközök lezárása
    instrAny.IO.Close
    If UseCalibrator Then instrAny2.IO.Close

    MsgBox "Mérés befejezõdött.", vbInformation, "Kész"
End Sub

