Attribute VB_Name = "ComparingMeans"
' FUNCTION: T_TEST_ONE_SAMPLE
' ---------------------------------
' Performs a one-sample t-test to compare the mean of a sample to a hypothesized population mean (mu).
'
' Arguments:
'   dataRange (Range): The range of sample data.
'   mu (Double): The hypothesized population mean to compare the sample against.
'   Alternative (String, Optional): Specifies the type of hypothesis test. Can be:
'       - "unequal": Two-tailed test (default).
'       - "less": Left-tailed test (H0: mean >= mu, Ha: mean < mu).
'       - "greater": Right-tailed test (H0: mean <= mu, Ha: mean > mu).
'   HasHeader (Boolean, Optional): Indicates whether the first row of the range contains a header.
'       If True, the first row will be ignored in the analysis. Defaults to False.
'
' Returns:
'   Variant: A 2D array containing the following results:
'       - t-Statistic: The calculated t-statistic for the sample.
'       - Degrees of Freedom: The degrees of freedom (n - 1), where n is the sample size.
'       - P-value: The p-value corresponding to the t-statistic, based on the chosen alternative hypothesis.
'       - Sample Mean: The mean of the sample data.
'
' Error Handling:
'   - Returns an error message if there is insufficient data (less than 2 observations).
'   - Returns an error message if an invalid value is provided for the "Alternative" argument.
'
' Example Usage:
'   Dim result As Variant
'   result = T_TEST_ONE_SAMPLE(Range("A:A"), 50, "greater", True)
'   ' This example performs a one-sample t-test on the data in column A, compares the sample mean to 50,
'   ' considers the first row as a header, and performs a right-tailed test (greater alternative hypothesis).
'
' Notes:
'   - The function assumes the data in dataRange is numeric.
'   - The standard error (SE) is calculated as stdDev / sqrt(n).
'   - The t-statistic is calculated as (sampleMean - mu) / SE.
'   - For the two-tailed test, the p-value is calculated using T.Dist_2T; for one-tailed tests, T.Dist is used.
'   - The returned result is a 2D array with labeled output for ease of interpretation.
Function T_TEST_ONE_SAMPLE(dataRange As Range, mu As Double, Optional Alternative As String = "unequal", Optional HasHeader As Boolean = False) As Variant
    Dim cleanedData As Variant
    Dim n As Long
    Dim sampleMean As Double, stdDev As Double, se As Double
    Dim tStat As Double
    Dim df As Long
    Dim pValue As Double
    Dim result(1 To 4, 1 To 2) As Variant

    ' Filter the data and handle the header
    cleanedData = FILTER_DATA(dataRange, HasHeader)

    ' Calculate the number of observations (n)
    n = UBound(cleanedData)
    
    ' Ensure there is data to analyze
    If n < 2 Then
        T_TEST_ONE_SAMPLE = "Error: Insufficient data"
        Exit Function
    End If

    ' Calculate the sample mean and standard deviation
    sampleMean = Application.WorksheetFunction.Average(cleanedData)
    stdDev = Application.WorksheetFunction.StDev_S(cleanedData)
    
    ' Calculate the standard error (SE)
    se = stdDev / Sqr(n)

    ' Calculate the t-statistic
    tStat = (sampleMean - mu) / se

    ' Degrees of freedom (n - 1)
    df = n - 1

    ' Calculate the p-value based on the alternative hypothesis
    Select Case Alternative
        Case "unequal"
            ' Two-tailed test (default)
            pValue = Application.WorksheetFunction.T_Dist_2T(Abs(tStat), df)
        Case "less"
            ' One-tailed test (left-tailed)
            pValue = Application.WorksheetFunction.T_Dist(tStat, df, True)
        Case "greater"
            ' One-tailed test (right-tailed)
            pValue = 1 - Application.WorksheetFunction.T_Dist(tStat, df, True)
        Case Else
            T_TEST_ONE_SAMPLE = "Error: Invalid value for 'Alternative'. Must be 'unequal', 'less', or 'greater'."
            Exit Function
    End Select

    ' Store the results in a 2D array
    result(1, 1) = "t-Statistic"
    result(1, 2) = tStat
    result(2, 1) = "Degrees of Freedom"
    result(2, 2) = df
    result(3, 1) = "P-value"
    result(3, 2) = pValue
    result(4, 1) = "Sample Mean"
    result(4, 2) = sampleMean

    ' Return the result array
    T_TEST_ONE_SAMPLE = result
End Function


' T_TEST_INDEPENDENT_SAMPLE
'
' This function performs an independent samples t-test to compare the means of two groups, with an option for either equal or unequal variances.
'
' Arguments:
' - range1: The first range of data (group 1).
' - range2: The second range of data (group 2).
' - Alternative (Optional): Specifies the alternative hypothesis. Options are:
'       * "unequal" (default): Tests if the two means are different.
'       * "greater": Tests if the mean of group 1 is greater than the mean of group 2.
'       * "less": Tests if the mean of group 1 is less than the mean of group 2.
' - EqualVariances (Optional): Boolean flag. If True, the test assumes equal variances. If False (default), the test uses Welch's t-test, which does not assume equal variances.
' - HasHeader (Optional): Boolean flag. If True, the first row of the data in each range is treated as a header and excluded from the analysis.
'
' Returns:
' - A 4x2 array with the following statistics:
'   * t-Statistic: The calculated t-statistic.
'   * Degrees of Freedom: The degrees of freedom used in the test.
'   * P-Value: The p-value for the test based on the alternative hypothesis.
'   * Mean Difference: The difference in means between group 1 and group 2.
'
' Example Usage:
' - T_TEST_INDEPENDENT_SAMPLE(A2:A11, B2:B11, "greater", True, False)
'   Performs a t-test comparing the means of data in A2:A11 and B2:B11, testing if the mean of group 1 is greater than group 2, assuming equal variances, and considering that the data does not contain a header.
'
' Error Handling:
' - Returns an error message if:
'   * One or both groups contain fewer than two data points.
'   * The Alternative argument is not "unequal", "greater", or "less".
'   * Any other internal calculation issue occurs.
Function T_TEST_INDEPENDENT_SAMPLE(range1 As Range, range2 As Range, Optional Alternative As String = "unequal", Optional EqualVariances As Boolean = False, Optional HasHeader As Boolean = False) As Variant
    Dim cleanedData1 As Variant, cleanedData2 As Variant
    Dim n1 As Long, n2 As Long
    Dim mean1 As Double, mean2 As Double
    Dim var1 As Double, var2 As Double
    Dim se As Double, tStat As Double, df As Double, pValue As Double
    Dim result(1 To 4, 1 To 2) As Variant
    
    ' Clean data to remove missing values and account for header if HasHeader = True
    cleanedData1 = FILTER_DATA(range1, HasHeader)
    cleanedData2 = FILTER_DATA(range2, HasHeader)
    
    ' Get the number of samples in each group
    n1 = UBound(cleanedData1)
    n2 = UBound(cleanedData2)
    
    ' Error check: if insufficient data
    If n1 < 2 Or n2 < 2 Then
        T_TEST_INDEPENDENT_SAMPLE = "Error: Not enough data in one or both groups."
        Exit Function
    End If
    
    ' Calculate means for each group
    mean1 = Application.WorksheetFunction.Average(cleanedData1)
    mean2 = Application.WorksheetFunction.Average(cleanedData2)
    
    ' Calculate variances for each group
    var1 = Application.WorksheetFunction.Var_S(cleanedData1)
    var2 = Application.WorksheetFunction.Var_S(cleanedData2)
    
    ' Calculate the standard error and degrees of freedom depending on whether equal variances are assumed
    If EqualVariances Then
        ' Pooled variance (equal variances assumption)
        pooledVar = ((n1 - 1) * var1 + (n2 - 1) * var2) / (n1 + n2 - 2)
        se = Sqr(pooledVar * (1 / n1 + 1 / n2))
        df = n1 + n2 - 2
    Else
        ' Welch's t-test (unequal variances assumption)
        se = Sqr(var1 / n1 + var2 / n2)
        df = (var1 / n1 + var2 / n2) ^ 2 / ((var1 ^ 2 / (n1 ^ 2 * (n1 - 1))) + (var2 ^ 2 / (n2 ^ 2 * (n2 - 1))))
    End If
    
    ' Calculate the t-statistic
    tStat = (mean1 - mean2) / se
    
    ' Calculate the p-value based on the alternative hypothesis
    Select Case Alternative
        Case "unequal"
            pValue = Application.WorksheetFunction.T_Dist_2T(Abs(tStat), df)
        Case "greater"
            pValue = 1 - Application.WorksheetFunction.T_Dist(tStat, df, True)
        Case "less"
            pValue = Application.WorksheetFunction.T_Dist(tStat, df, True)
        Case Else
            T_TEST_INDEPENDENT_SAMPLE = "Error: Invalid Alternative argument. Use 'unequal', 'greater', or 'less'."
            Exit Function
    End Select
    
    ' Prepare result output
    result(1, 1) = "t-Statistic"
    result(1, 2) = tStat
    result(2, 1) = "Degrees of Freedom"
    result(2, 2) = df
    result(3, 1) = "P-Value"
    result(3, 2) = pValue
    result(4, 1) = "Mean Difference"
    result(4, 2) = mean1 - mean2
    
    ' Return the result array
    T_TEST_INDEPENDENT_SAMPLE = result
End Function


  
' Function: T_TEST_PAIRED
' Performs a paired t-test on two ranges of data and returns the t-statistic, degrees of freedom, and p-value.
'
' Parameters:
'   range1 (Range) - The first range of paired data (e.g., pre-treatment values).
'   range2 (Range) - The second range of paired data (e.g., post-treatment values).
'   Optional Alternative (String) - The type of alternative hypothesis being tested:
'       "unequal" (default) - Two-tailed test (tests for any difference between the means).
'       "greater" - One-tailed test (tests if the mean of range1 is greater than the mean of range2).
'       "less" - One-tailed test (tests if the mean of range1 is less than the mean of range2).
'   Optional HasHeader (Boolean) - If True, the first row of the ranges is treated as a header and ignored in calculations. Default is False.
'
' Returns:
'   A 2D array containing:
'       - The t-statistic value
'       - Degrees of freedom
'       - P-value for the specified alternative hypothesis
'
' The function dynamically calculates differences between the paired values in range1 and range2, removing non-numeric and missing values.
' It then calculates the t-statistic, degrees of freedom, and p-value based on the differences between paired values.
'
' Example usage:
'   =T_TEST_PAIRED(A2:A20, B2:B20, "greater", TRUE)
'   This performs a paired t-test between the ranges A2:A20 and B2:B20, assuming that the first row is a header.
'
' Error Handling:
'   - If there are fewer than two valid numeric pairs, the function returns an error message: "Error: Not enough valid data for paired t-test."
'   - If an invalid value is provided for the 'Alternative' argument, the function returns: "Error: Invalid Alternative argument. Use 'unequal', 'greater', or 'less'."
Function T_TEST_PAIRED(range1 As Range, range2 As Range, Optional Alternative As String = "unequal", Optional HasHeader As Boolean = False) As Variant
    Dim diff() As Double
    Dim diffMean As Double, diffVar As Double
    Dim se As Double, tStat As Double, df As Double, pValue As Double
    Dim result(1 To 3, 1 To 2) As Variant
    Dim i As Long, n As Long, count As Long
    Dim row1Val As Variant, row2Val As Variant
    Dim startRow As Long, lastRow As Long
    
    ' Determine starting row based on whether there's a header
    If HasHeader Then
        startRow = 2
    Else
        startRow = 1
    End If
    
    ' Get the last row of data in each range
    lastRow = Application.Min(FIND_LAST_ROW(range1), FIND_LAST_ROW(range2))
    
    ' Resize the differences array to the maximum possible size
    ReDim diff(1 To lastRow - startRow + 1)
    
    count = 0 ' To keep track of how many valid pairs are included
    
    ' Loop over both columns simultaneously and calculate differences for valid numeric pairs
    For i = startRow To lastRow
        row1Val = range1.Cells(i, 1).value
        row2Val = range2.Cells(i, 1).value
        
        ' Only calculate difference if both values are numeric
        If IsNumeric(row1Val) And IsNumeric(row2Val) Then
            count = count + 1
            ReDim Preserve diff(1 To count)
            diff(count) = CDbl(row1Val) - CDbl(row2Val)
        End If
    Next i
    
    ' Error if there's not enough valid data
    If count < 2 Then
        T_TEST_PAIRED = "Error: Not enough valid data for paired t-test."
        Exit Function
    End If
    
    ' Calculate mean of the differences
    diffMean = Application.WorksheetFunction.Average(diff)
    
    ' Calculate variance of the differences
    diffVar = Application.WorksheetFunction.Var_S(diff)
    
    ' Calculate standard error
    se = Sqr(diffVar / count)
    
    ' Calculate the t-statistic
    tStat = diffMean / se
    
    ' Degrees of freedom for paired t-test
    df = count - 1
    
    ' Calculate the p-value based on the alternative hypothesis
    Select Case Alternative
        Case "unequal"
            pValue = Application.WorksheetFunction.T_Dist_2T(Abs(tStat), df)
        Case "greater"
            pValue = 1 - Application.WorksheetFunction.T_Dist(tStat, df, True)
        Case "less"
            pValue = Application.WorksheetFunction.T_Dist(tStat, df, True)
        Case Else
            T_TEST_PAIRED = "Error: Invalid Alternative argument. Use 'unequal', 'greater', or 'less'."
            Exit Function
    End Select
    
    ' Prepare result output
    result(1, 1) = "t-Statistic"
    result(1, 2) = tStat
    result(2, 1) = "Degrees of Freedom"
    result(2, 2) = df
    result(3, 1) = "P-Value"
    result(3, 2) = pValue
    
    ' Return the result array
    T_TEST_PAIRED = result
End Function


Function MYFUNC(ParamArray groups() As Variant) As Variant
    Dim groupValues As Variant
    Dim groupData() As Variant
    Dim groupMeans() As Double
    Dim overallMean As Double
    Dim sumSquaresBetween As Double
    Dim sumSquaresWithin As Double
    Dim totalSumSquares As Double
    Dim dfBetween As Long, dfWithin As Long
    Dim msBetween As Double, msWithin As Double
    Dim FValue As Double, pValue As Double
    Dim groupCount As Long, i As Long, j As Long
    Dim grandTotal As Double, nTotal As Long
    Dim nGroup As Long
    Dim result(1 To 6, 1 To 2) As Variant
    
    ' Initialize variables
    groupCount = UBound(groups) - LBound(groups) + 1
    ReDim groupData(1 To groupCount)
    ReDim groupMeans(1 To groupCount)
    
    ' Step 1: Filter data for each group and calculate group means
    For i = 1 To groupCount
        ' Filter data for the current group
        groupValues = FILTER_DATA(groups(i - 1), False)
        groupData(i) = groupValues
        
        ' Calculate the mean of the current group
        nGroup = UBound(groupValues)
        If nGroup < 2 Then
            ANOVA = "Error: Not enough data in one of the groups."
            Exit Function
        End If
        groupMeans(i) = Application.WorksheetFunction.Average(groupValues)
        
        ' Update the overall mean calculation
        grandTotal = grandTotal + Application.WorksheetFunction.Sum(groupValues)
        nTotal = nTotal + nGroup
    Next i
    
    ' Calculate the overall mean
    overallMean = grandTotal / nTotal
    
    ' Step 2: Calculate sum of squares between and within groups
    For i = 1 To groupCount
        ' Calculate sum of squares between groups
        nGroup = UBound(groupData(i))
        sumSquaresBetween = sumSquaresBetween + nGroup * (groupMeans(i) - overallMean) ^ 2
        
        ' Calculate sum of squares within groups
        For j = 1 To nGroup
            sumSquaresWithin = sumSquaresWithin + (groupData(i)(j) - groupMeans(i)) ^ 2
        Next j
    Next i
    
    ' Total sum of squares
    totalSumSquares = sumSquaresBetween + sumSquaresWithin
    
    ' Step 3: Calculate degrees of freedom
    dfBetween = groupCount - 1
    dfWithin = nTotal - groupCount
    
    ' Step 4: Calculate mean squares
    msBetween = sumSquaresBetween / dfBetween
    msWithin = sumSquaresWithin / dfWithin
    
    ' Step 5: Calculate F-value
    FValue = msBetween / msWithin
    
    ' Step 6: Calculate p-value using the F distribution
    pValue = Application.WorksheetFunction.F_Dist_RT(FValue, dfBetween, dfWithin)
    
    ' Prepare the result array
    result(1, 1) = "F-Statistic"
    result(1, 2) = FValue
    result(2, 1) = "P-Value"
    result(2, 2) = pValue
    result(3, 1) = "Sum of Squares Between"
    result(3, 2) = sumSquaresBetween
    result(4, 1) = "Sum of Squares Within"
    result(4, 2) = sumSquaresWithin
    result(5, 1) = "Degrees of Freedom Between"
    result(5, 2) = dfBetween
    result(6, 1) = "Degrees of Freedom Within"
    result(6, 2) = dfWithin
    
    ' Return the result array
    MYFUNC = result
End Function


