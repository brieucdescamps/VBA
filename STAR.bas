Attribute VB_Name = "STAR"
Public Function AnnualizedReturn(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Dim i As Long
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    Dim d As Date

    ' Ensure both ranges are of the same length
    If DateRange.count <> LevelRange.count Then
        AnnualizedReturn = CVErr(xlErrRef): Exit Function
    End If

    ' Find the first date ≥ StartDate and the last date ≤ EndDate
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' Exit if no valid range was found
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        AnnualizedReturn = CVErr(xlErrNA): Exit Function
    End If

    ' Get the level values at the start and end
    Dim LVstart As Double, LVend As Double
    LVstart = LevelRange.Cells(startIndex, 1).Value
    LVend = LevelRange.Cells(endIndex, 1).Value

    ' Validate input levels
    If LVstart <= 0 Or LVend <= 0 Then
        AnnualizedReturn = CVErr(xlErrNum): Exit Function
    End If

    ' Compute time span in years
    Dim totalYears As Double
    totalYears = Application.WorksheetFunction.YearFrac(DateRange.Cells(startIndex, 1), DateRange.Cells(endIndex, 1))
    If totalYears <= 0 Then
        AnnualizedReturn = CVErr(xlErrDiv0): Exit Function
    End If

    ' Calculate and return CAGR
    AnnualizedReturn = (LVend / LVstart) ^ (1 / totalYears) - 1
End Function


Public Function AnnualizedVolatility(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date, Optional Frequency As Integer = 1) As Variant
    Dim i As Long
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    Dim d As Date

    ' Ensure both ranges are of equal size
    If DateRange.count <> LevelRange.count Then
        AnnualizedVolatility = CVErr(xlErrRef): Exit Function
    End If

    ' Find first date ≥ StartDate and last date ≤ EndDate
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' Validate extracted indices
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        AnnualizedVolatility = CVErr(xlErrNA): Exit Function
    End If

    ' Ensure there are enough data points for the given frequency
    If (endIndex - startIndex) < Frequency Then
        AnnualizedVolatility = CVErr(xlErrValue): Exit Function
    End If

    ' Initialize return series
    Dim retCount As Long
    retCount = (endIndex - startIndex) - Frequency + 1
    If retCount < 1 Then
        AnnualizedVolatility = CVErr(xlErrNA): Exit Function
    End If

    Dim returns() As Double: ReDim returns(1 To retCount)
    Dim valStart As Double, valEnd As Double

    ' Calculate periodic returns
    For i = 1 To retCount
        valStart = LevelRange.Cells(startIndex + i - 1, 1).Value
        valEnd = LevelRange.Cells(startIndex + i - 1 + Frequency, 1).Value
        If valStart > 0 And valEnd > 0 Then
            returns(i) = valEnd / valStart - 1
        Else
            returns(i) = 0 ' fallback if data invalid
        End If
    Next i

    ' Compute sample standard deviation of returns
    Dim sum As Double, sumsq As Double, mean As Double, stdev As Double
    For i = 1 To retCount: sum = sum + returns(i): Next i
    mean = sum / retCount
    For i = 1 To retCount: sumsq = sumsq + (returns(i) - mean) ^ 2: Next i

    If retCount > 1 Then
        stdev = Sqr(sumsq / (retCount - 1))
    Else
        AnnualizedVolatility = CVErr(xlErrDiv0): Exit Function
    End If

    ' Annualize the volatility (assuming 252 trading days/year)
    Dim factor As Double
    factor = Sqr(252 / Frequency)
    AnnualizedVolatility = stdev * factor
End Function


Public Function maxDrawdown(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Dim i As Long
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    Dim d As Date

    ' Ensure date/level ranges are aligned
    If DateRange.count <> LevelRange.count Then
        maxDrawdown = CVErr(xlErrRef): Exit Function
    End If

    ' Locate first valid start and end index using date comparison
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' Validate search result
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        maxDrawdown = CVErr(xlErrNA): Exit Function
    End If

    ' Initialize drawdown tracking
    Dim currentPeak As Double
    Dim currentLevel As Double
    Dim currentDrawdown As Double
    Dim maxDd As Double: maxDd = 0
    Dim worstDate As Date

    currentPeak = LevelRange.Cells(startIndex, 1).Value
    worstDate = DateRange.Cells(startIndex, 1).Value

    ' Loop through data and compute drawdowns
    For i = startIndex To endIndex
        currentLevel = LevelRange.Cells(i, 1).Value
        If IsNumeric(currentLevel) Then
            If currentLevel > currentPeak Then currentPeak = currentLevel
            If currentPeak <> 0 Then
                currentDrawdown = (currentLevel - currentPeak) / currentPeak
                If currentDrawdown < maxDd Then
                    maxDd = currentDrawdown
                    worstDate = DateRange.Cells(i, 1).Value
                End If
            End If
        End If
    Next i

    ' Return an array: [max drawdown %, date of max drawdown]
    maxDrawdown = Array(maxDd, worstDate)
End Function


Public Function Worst10Drawdowns(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Const MAX_DRAWDOWNS As Integer = 10
    Dim i As Long

    ' Validate input
    If DateRange.count <> LevelRange.count Then
        Worst10Drawdowns = CVErr(xlErrRef): Exit Function
    End If

    ' Find usable index range
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            Dim d As Date: d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        Worst10Drawdowns = CVErr(xlErrNA): Exit Function
    End If

    ' Store non-overlapping drawdowns
    Dim allDrawdowns() As Variant
    ReDim allDrawdowns(1 To 1000, 1 To 4)
    Dim count As Long: count = 0
    i = startIndex

    Do While i < endIndex
        ' Find peak
        Dim peakIndex As Long: peakIndex = i
        Dim peakValue As Double: peakValue = LevelRange.Cells(i, 1).Value

        ' Look ahead for trough until a new peak is found
        Dim troughIndex As Long: troughIndex = i
        Dim minDD As Double: minDD = 0

        Dim j As Long
        For j = i + 1 To endIndex
            Dim price As Variant: price = LevelRange.Cells(j, 1).Value
            If Not IsNumeric(price) Then Exit For

            If price > peakValue Then
                Exit For ' new peak found
            Else
                Dim dd As Double: dd = (price - peakValue) / peakValue
                If dd < minDD Then
                    minDD = dd
                    troughIndex = j
                End If
            End If
        Next j

        ' Try to find recovery date (back above peak)
        Dim recoveryDate As Variant: recoveryDate = "-"
        If troughIndex > peakIndex Then
            For j = troughIndex + 1 To endIndex
                If IsNumeric(LevelRange.Cells(j, 1).Value) Then
                    If LevelRange.Cells(j, 1).Value >= peakValue Then
                        recoveryDate = DateRange.Cells(j, 1).Value
                        Exit For
                    End If
                End If
            Next j

            count = count + 1
            If count > UBound(allDrawdowns, 1) Then
                ReDim Preserve allDrawdowns(1 To count + 100, 1 To 4)
            End If
            allDrawdowns(count, 1) = minDD
            allDrawdowns(count, 2) = DateRange.Cells(peakIndex, 1).Value
            allDrawdowns(count, 3) = DateRange.Cells(troughIndex, 1).Value
            allDrawdowns(count, 4) = recoveryDate

            i = troughIndex + 1
        Else
            i = i + 1
        End If
    Loop

    ' Sort drawdowns by severity and take top 10
    If count = 0 Then
        Worst10Drawdowns = CVErr(xlErrNA): Exit Function
    End If

    Dim topN As Long: topN = WorksheetFunction.Min(count, MAX_DRAWDOWNS)
    Dim result() As Variant: ReDim result(0 To topN, 1 To 4)

    ' Add headers
    result(0, 1) = "Drawdown"
    result(0, 2) = "Peak Date"
    result(0, 3) = "Trough Date"
    result(0, 4) = "Recovery Date"

    Dim used() As Boolean: ReDim used(1 To count)
    Dim k As Long, bestIndex As Long

    For i = 1 To topN
        Dim bestDD As Double: bestDD = 0
        bestIndex = -1
        For k = 1 To count
            If Not used(k) Then
                If allDrawdowns(k, 1) < bestDD Then
                    bestDD = allDrawdowns(k, 1)
                    bestIndex = k
                End If
            End If
        Next k

        If bestIndex <> -1 Then
            used(bestIndex) = True
            result(i, 1) = allDrawdowns(bestIndex, 1)
            result(i, 2) = allDrawdowns(bestIndex, 2)
            result(i, 3) = allDrawdowns(bestIndex, 3)
            result(i, 4) = allDrawdowns(bestIndex, 4)
        End If
    Next i

    Worst10Drawdowns = result
End Function


