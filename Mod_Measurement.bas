Option Explicit

Sub RunMeasurement()
    Dim Counter As Integer
    Dim UtolsoSor As Integer
    Dim maxMeasurements As Integer
    Dim measurements() As Double
    Dim sum As Double, sumSq As Double
    Dim avg As Double, stdev As Variant
    Dim i As Integer
    Dim instrQuery As String

    ' Mérési szám beolvasása a V10 cellából
    maxMeasurements = CInt(Range("V10").Value)
    If maxMeasurements < 1 Then
        MsgBox "Hibás mérési szám a V10 cellában!", vbCritical, "Hiba"
        Exit Sub
    End If

    ' Mérési tartomány meghatározása
    With ActiveSheet
        UtolsoSor = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Counter = 2

    ' ------------------------------- Fő mérési ciklus -------------------------------
    While Counter < UtolsoSor + 1
        Dim MeresiMod As String, FunkcioMod As String
        MeresiMod = Cells(Counter, 1).Value
        FunkcioMod = Left(MeresiMod, 3)

        ' Aktuális sor F oszlopának kiemelése (sárga)
        Cells(Counter, "F").Interior.Color = RGB(255, 239, 174)

        ' Többszörös mérés és statisztika számítása
        sum = 0
        sumSq = 0
        ReDim measurements(1 To maxMeasurements)

        If UseDMM Then
            ' Beállítjuk a DMM-et
            If FunkcioMod = "VDC" Then
                instrAny.WriteString ("FUNC " + Chr(34) + "VOLT:DC" + Chr(34))
                instrAny.WriteString ("Volt:DC:Range 10")
            ElseIf FunkcioMod = "VAC" Then
                instrAny.WriteString ("FUNC " + Chr(34) + "VOLT:AC" + Chr(34))
                instrAny.WriteString ("Volt:AC:Range 10")
            End If

            Wait (0.5)

            ' Kalibrátor beállítása, ha engedélyezve van
            If UseCalibrator Then
                instrAny2.WriteString ("OUT 1V")
                instrAny2.WriteString ("OPER")
            End If

            ' Többszörös mérés
            For i = 1 To maxMeasurements
                instrAny.WriteString ("Measure:Volt:DC?")
                instrQuery = Trim(instrAny.ReadString)

                ' DMM válasz konvertálása
                measurements(i) = ConvertDMMResponse(instrQuery)
                sum = sum + measurements(i)
                sumSq = sumSq + measurements(i) * measurements(i)

                Wait (0.2) ' Késleltetés a mérések között
            Next i
        End If

        ' Átlag számítása
        avg = sum / maxMeasurements

        ' Szórás számítása (ha több mint 1 mérés van)
        If maxMeasurements > 1 Then
            stdev = Sqr((sumSq / maxMeasurements) - (avg * avg))
        Else
            stdev = "N/A"
        End If

        ' Eredmények beírása
        Cells(Counter, "F").Value = avg
        Cells(Counter, "J").Value = stdev

        ' Kalibrátor kikapcsolása (ha használjuk)
        If UseCalibrator Then instrAny2.WriteString ("STBY")

        ' Előző sor F oszlopának háttérszín visszaállítása (törlése)
        If Counter > 2 Then
            Cells(Counter - 1, "F").Interior.ColorIndex = xlNone
        End If

        ' Következő sor
        Counter = Counter + 1
    Wend

    ' ---------------------- Fő ciklus vége ----------------------
    If UseDMM Then instrAny.IO.Close ' DMM kapcsolat zárása
    If UseCalibrator Then instrAny2.IO.Close ' Kalibrátor kapcsolat zárása

    MsgBox "Sikeresen lefutott a Program"
End Sub
