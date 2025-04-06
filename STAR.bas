' Module: STAR
' This module provides functions for financial analysis: calculating 
' annualized returns, annualized volatility, maximum drawdown, and worst drawdowns,
' as well as constructing a stress-aware correlation matrix...
' ------------------------------------------------------------------------------

Attribute VB_Name = "STAR"

'==============================================================================
' Function: AnnualizedReturn
' Description: Calculates the Compound Annual Growth Rate (CAGR) for a given
'              date and level range based on the start and end date.
' Parameters:
'   DateRange - range of dates
'   LevelRange - range of corresponding levels (prices)
'   StartDate - analysis start date
'   EndDate - analysis end date
' Returns:
'   CAGR value or an Excel error if input is invalid.
'==============================================================================
Public Function AnnualizedReturn(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Dim i As Long
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    Dim d As Date

    ' Ensure both ranges have the same number of elements.
    If DateRange.count <> LevelRange.count Then
        AnnualizedReturn = CVErr(xlErrRef): Exit Function
    End If

    ' Identify the first date greater or equal to StartDate and the last date less or equal to EndDate.
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' Validate found indices.
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        AnnualizedReturn = CVErr(xlErrNA): Exit Function
    End If

    ' Retrieve level values from the calculated startIndex and endIndex.
    Dim LVstart As Double, LVend As Double
    LVstart = LevelRange.Cells(startIndex, 1).Value
    LVend = LevelRange.Cells(endIndex, 1).Value

    ' Validate that levels are positive.
    If LVstart <= 0 Or LVend <= 0 Then
        AnnualizedReturn = CVErr(xlErrNum): Exit Function
    End If

    ' Calculate total time span in years.
    Dim totalYears As Double
    totalYears = Application.WorksheetFunction.YearFrac(DateRange.Cells(startIndex, 1), DateRange.Cells(endIndex, 1))
    If totalYears <= 0 Then
        AnnualizedReturn = CVErr(xlErrDiv0): Exit Function
    End If

    ' Calculate CAGR.
    AnnualizedReturn = (LVend / LVstart) ^ (1 / totalYears) - 1
End Function

'==============================================================================
' Function: AnnualizedVolatility
' Description: Computes the annualized volatility based on periodic returns.
' Parameters:
'   DateRange - range of dates
'   LevelRange - range of corresponding levels (prices)
'   StartDate - analysis start date
'   EndDate - analysis end date
'   Frequency (optional) - frequency of returns calculation (default = 1)
' Returns:
'   Annualized volatility value or an Excel error if input is invalid.
'==============================================================================
Public Function AnnualizedVolatility(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date, Optional Frequency As Integer = 1) As Variant
    Dim i As Long
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    Dim d As Date

    ' Ensure the DateRange and LevelRange have equal number of elements.
    If DateRange.count <> LevelRange.count Then
        AnnualizedVolatility = CVErr(xlErrRef): Exit Function
    End If

    ' Identify the valid index range within the provided dates.
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' Validate the identified indices.
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        AnnualizedVolatility = CVErr(xlErrNA): Exit Function
    End If

    ' Ensure there are enough data points for return calculation given the frequency.
    If (endIndex - startIndex) < Frequency Then
        AnnualizedVolatility = CVErr(xlErrValue): Exit Function
    End If

    ' Determine the number of return observations.
    Dim retCount As Long
    retCount = (endIndex - startIndex) - Frequency + 1
    If retCount < 1 Then
        AnnualizedVolatility = CVErr(xlErrNA): Exit Function
    End If

    ' Calculate periodic returns.
    Dim returns() As Double: ReDim returns(1 To retCount)
    Dim valStart As Double, valEnd As Double
    For i = 1 To retCount
        valStart = LevelRange.Cells(startIndex + i - 1, 1).Value
        valEnd = LevelRange.Cells(startIndex + i - 1 + Frequency, 1).Value
        If valStart > 0 And valEnd > 0 Then
            returns(i) = valEnd / valStart - 1
        Else
            returns(i) = 0 ' Fall back on zero return if data is invalid.
        End If
    Next i

    ' Calculate the sample standard deviation for the returns.
    Dim sum As Double, sumsq As Double, mean As Double, stdev As Double
    For i = 1 To retCount: sum = sum + returns(i): Next i
    mean = sum / retCount
    For i = 1 To retCount: sumsq = sumsq + (returns(i) - mean) ^ 2: Next i

    If retCount > 1 Then
        stdev = Sqr(sumsq / (retCount - 1))
    Else
        AnnualizedVolatility = CVErr(xlErrDiv0): Exit Function
    End If

    ' Annualize volatility (assumes 252 trading days per year).
    Dim factor As Double
    factor = Sqr(252 / Frequency)
    AnnualizedVolatility = stdev * factor
End Function

'==============================================================================
' Function: PeriodicCorrelation
' Description: Computes the correlation between two time series based on periodic
'              returns. The frequency parameter defines the number of periods
'              (e.g., days) between return observations.
' Parameters:
'   DateRange - range of dates corresponding to the series data.
'   Series1 - range of values for the first time series.
'   Series2 - range of values for the second time series.
'   StartDate - analysis start date.
'   EndDate - analysis end date.
'   Frequency (optional) - number of periods to skip when computing returns
'                           (default = 1, meaning daily returns).
' Returns:
'   The correlation coefficient between the computed periodic returns,
'   or an Excel error if input is invalid.
'==============================================================================
Public Function PeriodicCorrelation(DateRange As Range, Series1 As Range, Series2 As Range, _
    StartDate As Date, EndDate As Date, Optional Frequency As Integer = 1) As Variant

    Dim i As Long
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    Dim d As Date
    
    ' Validate that the ranges have matching counts.
    If DateRange.count <> Series1.count Or DateRange.count <> Series2.count Then
        PeriodicCorrelation = CVErr(xlErrRef)
        Exit Function
    End If

    ' Identify the valid indices for analysis based on StartDate and EndDate.
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i
    
    ' Validate date indices.
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        PeriodicCorrelation = CVErr(xlErrNA)
        Exit Function
    End If

    ' Determine the number of periodic return observations.
    Dim retCount As Long
    retCount = (endIndex - startIndex) - Frequency + 1
    If retCount < 1 Then
        PeriodicCorrelation = CVErr(xlErrNA)
        Exit Function
    End If

    ' Calculate periodic returns for both series.
    Dim returns1() As Double, returns2() As Double
    ReDim returns1(1 To retCount)
    ReDim returns2(1 To retCount)
    Dim valStart1 As Double, valEnd1 As Double, valStart2 As Double, valEnd2 As Double

    For i = 1 To retCount
        valStart1 = Series1.Cells(startIndex + i - 1, 1).Value
        valEnd1 = Series1.Cells(startIndex + i - 1 + Frequency, 1).Value
        valStart2 = Series2.Cells(startIndex + i - 1, 1).Value
        valEnd2 = Series2.Cells(startIndex + i - 1 + Frequency, 1).Value
        
        ' Ensure the starting values are greater than zero to avoid division error.
        If valStart1 <= 0 Or valStart2 <= 0 Then
            returns1(i) = 0
            returns2(i) = 0
        Else
            returns1(i) = valEnd1 / valStart1 - 1
            returns2(i) = valEnd2 / valStart2 - 1
        End If
    Next i

    ' Calculate and return the correlation coefficient between the two return series.
    On Error GoTo ErrHandler
    PeriodicCorrelation = WorksheetFunction.Correl(returns1, returns2)
    Exit Function

ErrHandler:
    PeriodicCorrelation = CVErr(xlErrValue)
End Function

'==============================================================================
' Function: MaxDrawdown
' Description: Computes the maximum drawdown for the given range and returns both
'              the drawdown value and the date of the worst drawdown.
' Parameters:
'   DateRange - range of dates
'   LevelRange - range of corresponding levels (prices)
'   StartDate - analysis start date
'   EndDate - analysis end date
' Returns:
'   An array with the maximum drawdown percentage and the corresponding date, or
'   an Excel error if input is invalid.
'==============================================================================
Public Function MaxDrawdown(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Dim i As Long
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    Dim d As Date

    ' Ensure that the DateRange and LevelRange align.
    If DateRange.count <> LevelRange.count Then
        MaxDrawdown = CVErr(xlErrRef): Exit Function
    End If

    ' Determine the first valid and last valid indices based on the input dates.
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' Validate that the indices were found.
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        MaxDrawdown = CVErr(xlErrNA): Exit Function
    End If

    ' Initialize the tracking variables.
    Dim currentPeak As Double
    Dim currentLevel As Double
    Dim currentDrawdown As Double
    Dim maxDd As Double: maxDd = 0
    Dim worstDate As Date

    currentPeak = LevelRange.Cells(startIndex, 1).Value
    worstDate = DateRange.Cells(startIndex, 1).Value

    ' Loop through the LevelRange to find the maximum drawdown.
    For i = startIndex To endIndex
        currentLevel = LevelRange.Cells(i, 1).Value
        If IsNumeric(currentLevel) Then
            ' Update peak if current level is higher.
            If currentLevel > currentPeak Then currentPeak = currentLevel
            ' Calculate the drawdown from the current peak.
            If currentPeak <> 0 Then
                currentDrawdown = (currentLevel - currentPeak) / currentPeak
                ' Update maximum drawdown if current drawdown is lower.
                If currentDrawdown < maxDd Then
                    maxDd = currentDrawdown
                    worstDate = DateRange.Cells(i, 1).Value
                End If
            End If
        End If
    Next i

    ' Return the results as an array: maximum drawdown and the corresponding date.
    MaxDrawdown = Array(maxDd, worstDate)
End Function

'==============================================================================
' Function: Worst10Drawdowns
' Description: Finds up to 10 worst non-overlapping drawdowns in the historical data.
' Parameters:
'   DateRange - range of dates
'   LevelRange - corresponding levels (prices)
'   StartDate - starting date for analysis
'   EndDate - ending date for analysis
' Returns:
'   A result array containing headers and up to 10 drawdown entries (drawdown, peak date,
'   trough date, recovery date), or an Excel error if input is invalid.
'==============================================================================
Public Function Worst10Drawdowns(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Const MAX_DRAWDOWNS As Integer = 10
    Dim i As Long

    ' Validate that the DateRange and LevelRange have matching counts.
    If DateRange.count <> LevelRange.count Then
        Worst10Drawdowns = CVErr(xlErrRef): Exit Function
    End If

    ' Identify the valid indices for the analysis using StartDate and EndDate.
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            Dim d As Date: d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' Validate the identified index range.
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        Worst10Drawdowns = CVErr(xlErrNA): Exit Function
    End If

    ' Array to hold identified drawdowns.
    Dim allDrawdowns() As Variant
    ReDim allDrawdowns(1 To 1000, 1 To 4)
    Dim count As Long: count = 0
    i = startIndex

    ' Loop through the data to find non-overlapping drawdowns.
    Do While i < endIndex
        ' Identify a peak.
        Dim peakIndex As Long: peakIndex = i
        Dim peakValue As Double: peakValue = LevelRange.Cells(i, 1).Value

        ' Initialize trough variables and search for the trough.
        Dim troughIndex As Long: troughIndex = i
        Dim minDD As Double: minDD = 0

        Dim j As Long
        For j = i + 1 To endIndex
            Dim price As Variant: price = LevelRange.Cells(j, 1).Value
            If Not IsNumeric(price) Then Exit For

            If price > peakValue Then
                Exit For ' A new peak is found, ending current drawdown search.
            Else
                Dim dd As Double: dd = (price - peakValue) / peakValue
                If dd < minDD Then
                    minDD = dd
                    troughIndex = j
                End If
            End If
        Next j

        ' Look for the recovery date when the level goes back above the peak.
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

            ' Store this drawdown in the array.
            count = count + 1
            If count > UBound(allDrawdowns, 1) Then
                ' Increase storage if needed.
                ReDim Preserve allDrawdowns(1 To count + 100, 1 To 4)
            End If
            allDrawdowns(count, 1) = minDD
            allDrawdowns(count, 2) = DateRange.Cells(peakIndex, 1).Value
            allDrawdowns(count, 3) = DateRange.Cells(troughIndex, 1).Value
            allDrawdowns(count, 4) = recoveryDate

            ' Advance pointer to avoid overlapping drawdowns.
            i = troughIndex + 1
        Else
            i = i + 1
        End If
    Loop

    ' Check if any drawdowns were identified.
    If count = 0 Then
        Worst10Drawdowns = CVErr(xlErrNA): Exit Function
    End If

    ' Determine the number of worst drawdowns to display.
    Dim topN As Long: topN = WorksheetFunction.Min(count, MAX_DRAWDOWNS)
    Dim result() As Variant: ReDim result(0 To topN, 1 To 4)

    ' Set the header row.
    result(0, 1) = "Drawdown"
    result(0, 2) = "Peak Date"
    result(0, 3) = "Trough Date"
    result(0, 4) = "Recovery Date"

    ' Array to mark which drawdowns have been used.
    Dim used() As Boolean: ReDim used(1 To count)
    Dim k As Long, bestIndex As Long

    ' Select the worst drawdowns by severity.
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

        ' If a drawdown was selected, copy it to the result array.
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

'==============================================================================
' Function: WorstDrawdowns
' Description: A generic version of Worst10Drawdowns that allows specifying the
'              number of worst drawdowns to be retrieved.
' Parameters:
'   DateRange - range of dates
'   LevelRange - corresponding levels (prices)
'   StartDate - analysis start date
'   EndDate - analysis end date
'   numWorst (optional) - number of worst drawdowns to retrieve (default = 10)
' Returns:
'   A result array containing headers and up to numWorst drawdown entries.
'==============================================================================
Public Function WorstDrawdowns(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date, Optional numWorst As Integer = 10) As Variant
    Dim i As Long
    ' Validate that the input ranges match in size.
    If DateRange.count <> LevelRange.count Then
        WorstDrawdowns = CVErr(xlErrRef): Exit Function
    End If

    ' Determine the valid analysis range based on the provided dates.
    Dim startIndex As Long: startIndex = -1
    Dim endIndex As Long: endIndex = -1
    For i = 1 To DateRange.count
        If IsDate(DateRange.Cells(i, 1).Value) Then
            Dim d As Date: d = Int(DateRange.Cells(i, 1).Value)
            If startIndex = -1 And d >= Int(StartDate) Then startIndex = i
            If d <= Int(EndDate) Then endIndex = i
        End If
    Next i

    ' If the indices are invalid, exit with an error.
    If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex Then
        WorstDrawdowns = CVErr(xlErrNA): Exit Function
    End If

    ' Array to store drawdowns.
    Dim allDrawdowns() As Variant
    ReDim allDrawdowns(1 To 1000, 1 To 4)
    Dim count As Long: count = 0
    i = startIndex

    ' Loop through the data to identify drawdowns.
    Do While i < endIndex
        ' Identify a peak and record its index.
        Dim peakIndex As Long: peakIndex = i
        Dim peakValue As Double: peakValue = LevelRange.Cells(i, 1).Value

        ' Prepare to find the trough following the peak.
        Dim troughIndex As Long: troughIndex = i
        Dim minDD As Double: minDD = 0

        Dim j As Long
        For j = i + 1 To endIndex
            Dim price As Variant: price = LevelRange.Cells(j, 1).Value
            If Not IsNumeric(price) Then Exit For

            If price > peakValue Then
                Exit For ' Exit if a new peak is encountered.
            Else
                Dim dd As Double: dd = (price - peakValue) / peakValue
                If dd < minDD Then
                    minDD = dd
                    troughIndex = j
                End If
            End If
        Next j

        ' Find the recovery date (when the price goes back above the peak).
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

            ' Store the drawdown details.
            count = count + 1
            If count > UBound(allDrawdowns, 1) Then
                ReDim Preserve allDrawdowns(1 To count + 100, 1 To 4)
            End If
            allDrawdowns(count, 1) = minDD
            allDrawdowns(count, 2) = DateRange.Cells(peakIndex, 1).Value
            allDrawdowns(count, 3) = DateRange.Cells(troughIndex, 1).Value
            allDrawdowns(count, 4) = recoveryDate

            ' Move to the next section after the trough to avoid overlap.
            i = troughIndex + 1
        Else
            i = i + 1
        End If
    Loop

    ' If no drawdowns were computed, return an error.
    If count = 0 Then
        WorstDrawdowns = CVErr(xlErrNA): Exit Function
    End If

    ' Determine how many drawdowns to display based on input.
    Dim topN As Long: topN = WorksheetFunction.Min(count, numWorst)
    Dim result() As Variant: ReDim result(0 To topN, 1 To 4)

    ' Add header row.
    result(0, 1) = "Drawdown"
    result(0, 2) = "Peak Date"
    result(0, 3) = "Trough Date"
    result(0, 4) = "Recovery Date"

    ' Array to mark processed drawdowns.
    Dim used() As Boolean: ReDim used(1 To count)
    Dim k As Long, bestIndex As Long

    ' Identify the worst drawdowns by selecting the most severe ones.
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

        ' If a worst drawdown is found, copy it into the results.
        If bestIndex <> -1 Then
            used(bestIndex) = True
            result(i, 1) = allDrawdowns(bestIndex, 1)
            result(i, 2) = allDrawdowns(bestIndex, 2)
            result(i, 3) = allDrawdowns(bestIndex, 3)
            result(i, 4) = allDrawdowns(bestIndex, 4)
        End If
    Next i

    WorstDrawdowns = result
End Function

'==============================================================================
' Function: StressAwareCorrelationMatrix
' Description: Computes a correlation matrix for asset returns over both the full
'              sample period and one or more defined stress periods. For the full period,
'              returns are computed based on periodic returns (using the Frequency parameter)
'              from level data; for stress periods, data across all defined stress intervals
'              are aggregated.
' Parameters:
'   DateRange         - range of dates
'   LevelMatrix       - matrix of asset levels (each column corresponds to an asset)
'   FullStartDate     - start date for full period analysis
'   FullEndDate       - end date for full period analysis
'   StressStartDates  - range of stress period start dates
'   StressEndDates    - range of stress period end dates
'   Frequency (opt.)  - number of periods to skip when computing returns (default = 1)
' Returns:
'   A correlation matrix (nAssets x nAssets) or an Excel error if input is invalid.
'==============================================================================
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

