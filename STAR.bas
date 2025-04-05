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

Public Function StressAwareCorrelationMatrix(DateRange As Range, LevelMatrix As Range, _
    FullStartDate As Date, FullEndDate As Date, _
    StressStartDates As Range, StressEndDates As Range) As Variant

    Dim nAssets As Long: nAssets = LevelMatrix.Columns.count
    Dim nObs As Long: nObs = LevelMatrix.Rows.count
    Dim nStress As Long: nStress = StressStartDates.count
    Dim i As Long, j As Long, t As Long, k As Long

    ' Validate input dimensions
    If DateRange.Rows.count <> nObs Or StressEndDates.count <> nStress Then
        StressAwareCorrelationMatrix = CVErr(xlErrRef)
        Exit Function
    End If

    Dim datesArr() As Variant: datesArr = DateRange.Value
    Dim levelsArr() As Variant: levelsArr = LevelMatrix.Value
    Dim stressStarts() As Variant: stressStarts = StressStartDates.Value
    Dim stressEnds() As Variant: stressEnds = StressEndDates.Value

    Dim result() As Variant
    ReDim result(1 To nAssets, 1 To nAssets)

    For i = 1 To nAssets
        For j = 1 To nAssets
            If i = j Then
                result(i, j) = 1#
            ElseIf i < j Then
                ' Full period correlation
                Dim xFull() As Double, yFull() As Double
                ReDim xFull(1 To nObs)
                ReDim yFull(1 To nObs)
                Dim numFullObs As Long: numFullObs = 0
                
                For t = 1 To nObs
                    If IsDate(datesArr(t, 1)) Then
                        Dim currDate As Date: currDate = datesArr(t, 1)
                        If currDate >= FullStartDate And currDate <= FullEndDate Then
                            If IsNumeric(levelsArr(t, i)) And IsNumeric(levelsArr(t, j)) Then
                                numFullObs = numFullObs + 1
                                xFull(numFullObs) = levelsArr(t, i)
                                yFull(numFullObs) = levelsArr(t, j)
                            End If
                        End If
                    End If
                Next t
                
                If numFullObs >= 2 Then
                    ReDim Preserve xFull(1 To numFullObs)
                    ReDim Preserve yFull(1 To numFullObs)
                    result(i, j) = WorksheetFunction.Correl(xFull, yFull)
                Else
                    result(i, j) = "-"
                End If
            Else
                ' Stress period correlation (aggregated over all defined stress periods)
                Dim xStress() As Double, yStress() As Double
                ReDim xStress(1 To nObs * nStress)
                ReDim yStress(1 To nObs * nStress)
                Dim numStressObs As Long: numStressObs = 0
                
                For k = 1 To nStress
                    If IsDate(stressStarts(k, 1)) And IsDate(stressEnds(k, 1)) Then
                        Dim sStart As Date: sStart = stressStarts(k, 1)
                        Dim sEnd As Date: sEnd = stressEnds(k, 1)
                        For t = 1 To nObs
                            If IsDate(datesArr(t, 1)) Then
                                Dim currDate2 As Date: currDate2 = datesArr(t, 1)
                                If currDate2 >= sStart And currDate2 <= sEnd Then
                                    If IsNumeric(levelsArr(t, i)) And IsNumeric(levelsArr(t, j)) Then
                                        numStressObs = numStressObs + 1
                                        xStress(numStressObs) = levelsArr(t, i)
                                        yStress(numStressObs) = levelsArr(t, j)
                                    End If
                                End If
                            End If
                        Next t
                    End If
                Next k
                
                If numStressObs >= 2 Then
                    ReDim Preserve xStress(1 To numStressObs)
                    ReDim Preserve yStress(1 To numStressObs)
                    result(i, j) = WorksheetFunction.Correl(xStress, yStress)
                Else
                    result(i, j) = "-"
                End If
            End If
        Next j
    Next i

    StressAwareCorrelationMatrix = result
End Function



