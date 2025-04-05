Attribute VB_Name = "STAR"

Public Function AnnualizedReturn(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Dim i As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim foundStart As Boolean, foundEnd As Boolean
    Dim LevelStart As Double, LevelEnd As Double
    Dim totalYears As Double
    
    ' Ensure the ranges have the same number of cells.
    If DateRange.Count <> LevelRange.Count Then
        AnnualizedReturn = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Loop through the DateRange to find the indices for StartDate and EndDate.
    foundStart = False
    foundEnd = False
    For i = 1 To DateRange.Count
        If Not foundStart Then
            If CDate(DateRange.Cells(i, 1).Value) = StartDate Then
                startIndex = i
                foundStart = True
            End If
        End If
        If Not foundEnd Then
            If CDate(DateRange.Cells(i, 1).Value) = EndDate Then
                endIndex = i
                foundEnd = True
            End If
        End If
        If foundStart And foundEnd Then Exit For
    Next i
    
    ' If either date is not found, return an error.
    If Not foundStart Or Not foundEnd Then
        AnnualizedReturn = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' Retrieve the level values corresponding to the found dates.
    LevelStart = LevelRange.Cells(startIndex, 1).Value
    LevelEnd = LevelRange.Cells(endIndex, 1).Value
    
    ' Calculate the time difference in years using Excel's YearFrac function.
    totalYears = Application.WorksheetFunction.YearFrac(StartDate, EndDate)
    If totalYears <= 0 Then
        AnnualizedReturn = CVErr(xlErrDiv0)
        Exit Function
    End If
    
    ' Calculate the Compound Annual Growth Rate (CAGR).
    AnnualizedReturn = (LevelEnd / LevelStart) ^ (1 / totalYears) - 1
End Function

Public Function AnnualizedVolatility(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date, Optional Frequency As Integer = 1) As Variant
    Dim i As Long, j As Long
    Dim startIndex As Long, endIndex As Long
    Dim foundStart As Boolean, foundEnd As Boolean
    Dim retCount As Long
    Dim returns() As Double
    Dim sum As Double, sumsq As Double, mean As Double
    Dim stdev As Double, factor As Double
    
    ' Validate that DateRange and LevelRange have the same number of cells.
    If DateRange.Count <> LevelRange.Count Then
        AnnualizedVolatility = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Find the indices for StartDate and EndDate.
    foundStart = False: foundEnd = False
    For i = 1 To DateRange.Count
        If Not foundStart Then
            If CDate(DateRange.Cells(i, 1).Value) = StartDate Then
                startIndex = i
                foundStart = True
            End If
        End If
        If Not foundEnd Then
            If CDate(DateRange.Cells(i, 1).Value) = EndDate Then
                endIndex = i
                foundEnd = True
            End If
        End If
        If foundStart And foundEnd Then Exit For
    Next i
    
    If Not foundStart Or Not foundEnd Then
        AnnualizedVolatility = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' Ensure there are enough observations given the Frequency step.
    If (endIndex - startIndex) < Frequency Then
        AnnualizedVolatility = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Calculate the number of returns using the given frequency.
    retCount = (endIndex - startIndex) - Frequency + 1
    ReDim returns(1 To retCount)
    
    ' Compute returns: using return = (Level at (i+Frequency) / Level at i) - 1.
    For i = 1 To retCount
        Dim startVal As Double, endVal As Double
        startVal = LevelRange.Cells(startIndex + i - 1, 1).Value
        endVal = LevelRange.Cells(startIndex + i - 1 + Frequency, 1).Value
        returns(i) = endVal / startVal - 1
    Next i
    
    ' Calculate the sample standard deviation of the returns.
    sum = 0
    For j = 1 To retCount
        sum = sum + returns(j)
    Next j
    mean = sum / retCount
    
    sumsq = 0
    For j = 1 To retCount
        sumsq = sumsq + (returns(j) - mean) ^ 2
    Next j
    
    If retCount > 1 Then
        stdev = Sqr(sumsq / (retCount - 1))
    Else
        AnnualizedVolatility = CVErr(xlErrDiv0)
        Exit Function
    End If
    
    ' Annualize the volatility.
    ' For daily returns (Frequency = 1), there are approximately 252 trading days.
    ' For returns computed every 'Frequency' days, use an effective number of periods per year = 252 / Frequency.
    factor = Sqr(252 / Frequency)
    
    AnnualizedVolatility = stdev * factor
End Function

Public Function MaxDrawdown(DateRange As Range, LevelRange As Range, StartDate As Date, EndDate As Date) As Variant
    Dim i As Long
    Dim startIndex As Long, endIndex As Long
    Dim foundStart As Boolean, foundEnd As Boolean
    Dim currentPeak As Double
    Dim currentLevel As Double
    Dim currentDrawdown As Double
    Dim maxDD As Double           ' ? Renamed
    Dim worstDate As Date

    ' Validate input
    If DateRange.Count <> LevelRange.Count Then
        MaxDrawdown = CVErr(xlErrRef)
        Exit Function
    End If
    If DateRange.Columns.Count > 1 Or LevelRange.Columns.Count > 1 Then
        MaxDrawdown = CVErr(xlErrRef)
        Exit Function
    End If

    ' Find start and end indices (by comparing dates only, ignoring time)
    For i = 1 To DateRange.Count
        If Not foundStart And Int(CDate(DateRange.Cells(i, 1).Value)) = Int(StartDate) Then
            startIndex = i
            foundStart = True
        End If
        If Not foundEnd And Int(CDate(DateRange.Cells(i, 1).Value)) = Int(EndDate) Then
            endIndex = i
            foundEnd = True
        End If
        If foundStart And foundEnd Then Exit For
    Next i

    If Not foundStart Or Not foundEnd Or endIndex <= startIndex Then
        MaxDrawdown = CVErr(xlErrNA)
        Exit Function
    End If

    currentPeak = LevelRange.Cells(startIndex, 1).Value
    maxDD = 0
    worstDate = DateRange.Cells(startIndex, 1).Value

    For i = startIndex To endIndex
        currentLevel = LevelRange.Cells(i, 1).Value
        If IsNumeric(currentLevel) Then
            If currentLevel > currentPeak Then currentPeak = currentLevel
            If currentPeak <> 0 Then
                currentDrawdown = (currentLevel - currentPeak) / currentPeak
                If currentDrawdown < maxDD Then
                    maxDD = currentDrawdown
                    worstDate = DateRange.Cells(i, 1).Value
                End If
            End If
        End If
    Next i

    MaxDrawdown = Array(maxDD, worstDate)
End Function


