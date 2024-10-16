Attribute VB_Name = "ComparingFrequencies"
' Function: CHISQ_TEST
' Purpose:  Calculates the chi-square test statistic and p-value for two categorical variables.
'           Incorporates Yates' continuity correction for 2x2 tables.
'
' Arguments:
'   - col1 (Range): The first column of categorical data.
'   - col2 (Range): The second column of categorical data.
'   - HasHeader (Optional Boolean): If True, treats the first row of both columns as headers and excludes them from analysis.
'
' Returns:
'   - A 4x2 array containing the following:
'       * Chi-Square: The chi-square test statistic.
'       * P-value: The p-value for the chi-square test statistic.
'       * Degrees of Freedom: The degrees of freedom for the chi-square test.
'       * Warning (if applicable): A warning if any expected values are below 1, or if more than 20% of expected values are below 5.
'
' Example usage:
'   - =CHISQ_TEST(A2:A100, B2:B100, TRUE)
'
' Notes:
'   - Automatically applies Yates' continuity correction when working with 2x2 tables.
'   - Issues warnings when assumptions of the chi-square test are violated due to low expected values.
Function CHISQ_TEST(col1 As Range, col2 As Range, Optional HasHeader As Boolean = False) As Variant
    Dim observed As Variant
    Dim expectedVals() As Double
    Dim chiSq As Double
    Dim total As Double
    Dim rowTotals() As Long
    Dim colTotals() As Long
    Dim pValue As Double
    Dim result(1 To 4, 1 To 2) As Variant
    Dim r As Long, c As Long
    Dim df As Long
    Dim i As Long, j As Long
    Dim useYates As Boolean
    Dim correctionValue As Double
    Dim warning As String
    Dim expectedCountLessThanFive As Long
    Dim expectedCountLessThanOne As Boolean
    Dim totalCells As Long
    
    ' Get the observed counts by calling CROSSTAB
    observed = CROSSTAB(col1, col2, HasHeader)
    
    expectedVals = expected(observed)
    
    ' Get the number of rows and columns from the observed table
    r = UBound(expectedVals, 1)
    c = UBound(expectedVals, 2)
    
    ' Initialize counters for expected value thresholds
    expectedCountLessThanFive = 0
    expectedCountLessThanOne = False
    totalCells = r * c
    
    ' Check if Yates' continuity correction should be applied (only for 2x2 tables)
    useYates = (r = 2 And c = 2)
    If useYates Then
        correctionValue = 0.5
    Else
        correctionValue = 0
    End If
    
    ' Loop over expectedVals to determine continuity correction and assumption violation
    For i = 1 To r
        For j = 1 To c
            If expectedVals(i, j) < 1 Then
                expectedCountLessThanOne = True
            ElseIf expectedVals(i, j) < 5 Then
                expectedCountLessThanFive = expectedCountLessThanFive + 1
            End If
            
            If useYates Then
                If Abs(observed(i + 1, j + 1) - expectedVals(i, j)) < correctionValue Then
                    correctionValue = Abs(observed(i + 1, j + 1) - expectedVals(i, j))
                End If
            End If
        Next j
    Next i
           
    ' Calculate chi-square statistic
    chiSq = 0
    For i = 1 To r
        For j = 1 To c
            chiSq = chiSq + ((Abs(observed(i + 1, j + 1) - expectedVals(i, j)) - correctionValue) ^ 2) / expectedVals(i, j)
        Next j
    Next i
    
    ' Calculate degrees of freedom
    df = (r - 1) * (c - 1)
    
    ' Calculate p-value using ChiSq_Dist_RT
    pValue = Application.WorksheetFunction.ChiSq_Dist_RT(chiSq, df)
    
    ' Add a warning if the expected values do not meet the threshold
    If expectedCountLessThanOne Then
        warning = "Some expected values are less than 1, chi-square test may be invalid."
    ElseIf expectedCountLessThanFive / totalCells > 0.2 Then
        warning = "More than 20% of expected values are less than 5, chi-square test may be invalid."
    End If
    
    ' Format the output array
    result(1, 1) = "Chi-Square"
    result(1, 2) = chiSq
    result(2, 1) = "P-value"
    result(2, 2) = pValue
    result(3, 1) = "Degrees of Freedom"
    result(3, 2) = df
    
    ' Add warning if needed
    If warning <> "" Then
        result(4, 1) = "Warning:"
        result(4, 2) = warning
    Else
        result(4, 1) = ""
        result(4, 2) = ""
    End If
    
    ' Return the result array
    CHISQ_TEST = result
End Function

' Function: EXPECTED
' Purpose:  Calculates the expected frequencies for a contingency table.
'
' Arguments:
'   - observed (Variant): A 2D array containing the observed counts from a contingency table.
'                         The first row and column are treated as headers and ignored in the calculation.
'
' Returns:
'   - A 2D array of the same size as the observed table (minus headers), containing the expected frequencies.
'
' Example usage:
'   - EXPECTED(CROSSTAB(A2:A100, B2:B100, TRUE))
'
' Notes:
'   - This function is used internally by the CHISQ_TEST function to compute expected values.
'   - It calculates expected counts using the formula: E = (row total * column total) / grand total.
Function expected(observed As Variant) As Variant
    Dim expectedVals() As Double
    Dim rowTotals() As Long, colTotals() As Long
    Dim total As Double
    Dim r As Long, c As Long
    Dim i As Long, j As Long
    
    ' Get the number of rows and columns from the observed table
    r = UBound(observed, 1) - 1 ' Subtract 1 for the header row
    c = UBound(observed, 2) - 1 ' Subtract 1 for the header column
    
    ' Initialize expected counts, row totals, and column totals arrays
    ReDim expectedVals(1 To r, 1 To c)
    ReDim rowTotals(1 To r)
    ReDim colTotals(1 To c)
    
    ' Calculate row totals, column totals, and grand total
    total = 0
    For i = 2 To r + 1
        For j = 2 To c + 1
            rowTotals(i - 1) = rowTotals(i - 1) + observed(i, j)
            colTotals(j - 1) = colTotals(j - 1) + observed(i, j)
            total = total + observed(i, j)
        Next j
    Next i
    
    ' Calculate expected counts
    For i = 1 To r
        For j = 1 To c
            expectedVals(i, j) = (rowTotals(i) * colTotals(j)) / total
        Next j
    Next i
    
    ' Return the expected counts array
    expected = expectedVals
End Function


Function FISHER_2x2(col1 As Range, col2 As Range, Optional Alternative As String = "unequal", Optional HasHeader As Boolean = False) As Variant
    Dim observed As Variant
    Dim r As Long, c As Long
    Dim minMarg As Integer
    Dim pValue As Double
    Dim result(1 To 2, 1 To 2) As Variant
    Dim rowTotals() As Long, colTotals() As Long
    Dim total As Long
    Dim i As Long, j As Long

    ' Get the observed counts by calling CROSSTAB
    observed = CROSSTAB(col1, col2, HasHeader)

    ' Get the number of rows and columns from the observed table
    r = UBound(observed, 1) - 1 ' Subtract 1 for the header row
    c = UBound(observed, 2) - 1 ' Subtract 1 for the header column
    
    ' Fisher's Exact Test is typically applied to 2x2 tables
    If r <> 2 Or c <> 2 Then
        FISHER_2x2 = "Error: Fisher's Exact Test is primarily for 2x2 tables"
        Exit Function
    End If

    ' Initialize row totals, column totals, and grand total
    ReDim rowTotals(1 To r)
    ReDim colTotals(1 To c)
    total = 0
    
    ' Calculate row totals, column totals, and grand total
    For i = 2 To r + 1
        For j = 2 To c + 1
            rowTotals(i - 1) = rowTotals(i - 1) + observed(i, j)
            colTotals(j - 1) = colTotals(j - 1) + observed(i, j)
            total = total + observed(i, j)
        Next j
    Next i
    
    
    minMarg = Application.WorksheetFunction.Min(rowTotals(1), colTotals(1))
    
    ' Use HYPGEOM.DIST to calculate Fisher's Exact Test p-value for a 2x2 table
    ' HYPGEOM.DIST(x, N_draws, successes_in_population, population_size, cumulative)
    If Alternative = "greater" Then
        pValue = 1 - Application.WorksheetFunction.HypGeom_Dist(observed(2, 2) - 1, rowTotals(1), colTotals(1), total, True)
    ElseIf Alternative = "less" Then
        pValue = Application.WorksheetFunction.HypGeom_Dist(observed(2, 2), rowTotals(1), colTotals(1), total, True)
    ElseIf Alternative = "unequal" Then
        pValue = Application.WorksheetFunction.Min(1, 1 - Application.WorksheetFunction.HypGeom_Dist(minMarg - observed(2, 2) + 1, rowTotals(1), colTotals(1), total, True) + Application.WorksheetFunction.HypGeom_Dist(observed(2, 2), rowTotals(1), colTotals(1), total, True))
    Else
        FISHER_2x2 = "Error: alternative must be set as unequal, less or greater"
        Exit Function
    End If

    ' Format the result array
    result(1, 1) = "P-value"
    result(1, 2) = pValue
    result(2, 1) = "Note"
    result(2, 2) = "Fisher's Exact Test for 2x2 Table using HYPGEOM.DIST"

    ' Return the result array
    FISHER_2x2 = result
End Function


Function FISHER_2x3(col1 As Range, col2 As Range, Optional HasHeader As Boolean = False) As Variant
    Dim observed As Variant
    Dim r As Long, c As Long
    Dim pA As Double, pB As Double, pC As Double
    Dim pValue As Double, pObs As Double, pTable As Double
    Dim result(1 To 2, 1 To 2) As Variant
    Dim rowTotals() As Long, colTotals() As Long
    Dim total As Long
    Dim i As Long, j As Long
    Dim k As Integer

    ' Get the observed counts by calling CROSSTAB
    observed = CROSSTAB(col1, col2, HasHeader)

    ' Get the number of rows and columns from the observed table
    r = UBound(observed, 1) - 1 ' Subtract 1 for the header row
    c = UBound(observed, 2) - 1 ' Subtract 1 for the header column

    ' Initialize row totals, column totals, and grand total
    ReDim rowTotals(1 To r)
    ReDim colTotals(1 To c)
    total = 0
    
    ' Calculate row totals, column totals, and grand total
    For i = 2 To r + 1
        For j = 2 To c + 1
            rowTotals(i - 1) = rowTotals(i - 1) + observed(i, j)
            colTotals(j - 1) = colTotals(j - 1) + observed(i, j)
            total = total + observed(i, j)
        Next j
    Next i
    
    ' Calculate probability of observed table (under null hypothesis)
    pA = Application.WorksheetFunction.HypGeom_Dist(observed(2, 2), colTotals(1), observed(2, 2) + observed(2, 3), colTotals(1) + colTotals(2), False)
    pC = Application.WorksheetFunction.HypGeom_Dist(observed(2, 4), colTotals(3), rowTotals(1), total, False)
    'pB = Application.WorksheetFunction.HypGeom_Dist(observed(2, 3), colTotals(2), rowTotals(1), total, False)
    'pC = Application.WorksheetFunction.HypGeom_Dist(observed(2, 4), colTotals(3), rowTotals(1), total, False)
    
    pObs = pA * pC
    
    pValue = pObs
    
    
    For i = 0 To Application.WorksheetFunction.Min(colTotals(1), rowTotals(1))
        For j = 0 To Application.WorksheetFunction.Min(colTotals(2), rowTotals(1) - i)
            k = rowTotals(1) - i - j
            
            pA = Application.WorksheetFunction.HypGeom_Dist(i, colTotals(1), i + j, colTotals(1) + colTotals(2), False)
            'pB = Application.WorksheetFunction.HypGeom_Dist(j, colTotals(2), rowTotals(1), total, False)
            pC = Application.WorksheetFunction.HypGeom_Dist(k, colTotals(3), rowTotals(1), total, False)
            
            pTable = pA * pC
            
            If pTable <= pObs Then
                pValue = pValue + pTable
            End If
            
        Next j
    Next i
       
    FISHER_2x3 = pValue
    
End Function

' ********************************************************************
' Function: GOODNESS_OF_FIT
' Description: This function performs a chi-squared goodness of fit test
'              on an array of categorical data. It compares the observed
'              frequencies in `dataRange` with expected frequencies
'              (provided via `expectedRange`, or assumed to be uniformly distributed if not provided).
'
' Arguments:
'   - dataRange (Range): The range containing the observed categorical data.
'   - expectedRange (Range) [Optional]: A range containing the expected frequencies for each category.
'        If not provided, the function assumes uniform distribution across all categories.
'   - HasHeader (Boolean) [Optional]: A flag indicating if the `dataRange` contains a header row.
'        Defaults to False. If True, the first row is treated as the header.
'
' Returns:
'   - A 2D array with the following statistics:
'     - "Chi-Square Statistic": The calculated chi-squared value.
'     - "Degrees of Freedom": The degrees of freedom (number of categories - 1).
'     - "P-Value": The p-value for the test, based on the chi-squared distribution.
'
' Steps:
'   1. Observed frequencies are calculated using the `DESCRIBE_CATEGORICAL` function.
'   2. If `expectedRange` is provided, it is used to calculate the expected frequencies.
'      Otherwise, a uniform distribution of expected frequencies is assumed.
'   3. The expected values are normalized to match the total number of observed values.
'   4. The chi-squared statistic is calculated by comparing observed and expected values.
'   5. The p-value is calculated using the chi-squared distribution with the appropriate degrees of freedom.
'
' Error Handling:
'   - If the number of categories in `expectedRange` does not match the number of categories in `dataRange`, the function returns an error.
'   - If any non-numeric values are found in `expectedRange`, an error is returned.
'
' Example Usage:
'   - =GOODNESS_OF_FIT(A2:A11, B2:B3, TRUE)
'     Performs a chi-squared goodness of fit test on the observed categorical data in A2:A11,
'     comparing it with the expected frequencies in B2:B3. The first row is treated as a header.
'
'   - =GOODNESS_OF_FIT(A2:A11)
'     Performs a chi-squared goodness of fit test on the observed categorical data in A2:A11,
'     assuming a uniform distribution of expected values.
' ********************************************************************
Function GOODNESS_OF_FIT(dataRange As Range, Optional expectedRange As Range = Nothing, Optional HasHeader As Boolean = False) As Variant
    Dim observedFreq As Variant
    Dim expected() As Double
    Dim totalExpected As Double
    Dim chiSq As Double
    Dim pValue As Double
    Dim df As Long
    Dim result(1 To 3, 1 To 2) As Variant
    Dim totalObserved As Double
    Dim groupCount As Long
    Dim i As Long

    ' Use DESCRIBE_CATEGORICAL to generate observed frequency table
    observedFreq = DESCRIBE_CATEGORICAL(dataRange, HasHeader)

    ' Determine the number of categories (groupCount)
    groupCount = UBound(observedFreq) - 1

    ' Calculate the total number of observed values
    totalObserved = 0
    For i = 2 To groupCount + 1
        totalObserved = totalObserved + observedFreq(i, 2)
    Next i

    ' If the expected range is provided, use it; otherwise assume uniform distribution
    If Not expectedRange Is Nothing Then
        ' Ensure that expectedRange contains valid numeric values
        If expectedRange.Cells.count <> groupCount Then
            GOODNESS_OF_FIT = "Error: Expected values must have the same number of categories as the observed data."
            Exit Function
        End If
        
        ' Convert expectedRange to an array of doubles
        ReDim expected(1 To groupCount)
        For i = 1 To groupCount
            If IsNumeric(expectedRange.Cells(i, 1).value) Then
                expected(i) = expectedRange.Cells(i, 1).value
            Else
                GOODNESS_OF_FIT = "Error: Non-numeric value in expected range."
                Exit Function
            End If
        Next i
        
        ' Calculate total expected counts
        totalExpected = Application.WorksheetFunction.Sum(expected)
        
        ' Normalize expected values
        For i = 1 To groupCount
            expected(i) = expected(i) / totalExpected * totalObserved
        Next i
    Else
        ' If no expected frequencies provided, assume uniform distribution
        ReDim expected(1 To groupCount)
        For i = 1 To groupCount
            expected(i) = totalObserved / groupCount
        Next i
    End If

    ' Initialize chi-square statistic
    chiSq = 0

    ' Calculate chi-square statistic
    For i = 2 To groupCount + 1
        If expected(i - 1) > 0 Then
            chiSq = chiSq + ((observedFreq(i, 2) - expected(i - 1)) ^ 2) / expected(i - 1)
        End If
    Next i

    ' Degrees of freedom: (number of categories - 1)
    df = groupCount - 1

    ' Calculate p-value using ChiSq_Dist_RT
    pValue = Application.WorksheetFunction.ChiSq_Dist_RT(chiSq, df)

    ' Prepare result output
    result(1, 1) = "Chi-Square Statistic"
    result(1, 2) = chiSq
    result(2, 1) = "Degrees of Freedom"
    result(2, 2) = df
    result(3, 1) = "P-Value"
    result(3, 2) = pValue

    ' Return the result array
    GOODNESS_OF_FIT = result
End Function

