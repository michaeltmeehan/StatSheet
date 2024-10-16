Attribute VB_Name = "ComparingMedians"
' Function: WILCOXON_SIGNED_RANK
'
' Description:
' Performs the Wilcoxon signed-rank test for paired samples. This is a non-parametric statistical test
' used to compare two related samples to assess whether their population mean ranks differ.
' It is suitable for non-normally distributed paired data.
'
' Arguments:
' - range1 (Range): The first range of data (paired sample 1).
' - range2 (Range): The second range of data (paired sample 2).
' - Optional HasHeader (Boolean, default = False): If True, the first row is treated as a header
'   and excluded from the analysis.
'
' Returns:
' A 2D array containing:
' 1. W-Statistic: The Wilcoxon signed-rank test statistic.
' 2. P-Value: The p-value for the test, calculated using a normal approximation for large sample sizes (>= 10).
' 3. Warning: If the sample size is smaller than 10, a warning is included indicating that an exact test is needed.
'
' Details:
' 1. The function calculates the differences between the paired samples and ranks the absolute values of these differences.
' 2. It then computes the rank sums for positive and negative differences.
' 3. The test statistic W is the minimum of these rank sums.
' 4. For large samples (>= 10), the p-value is approximated using a normal distribution.
Function WILCOXON_SIGNED_RANK(range1 As Range, range2 As Range, Optional HasHeader As Boolean = False) As Variant
    Dim row1Val As Variant, row2Val As Variant
    Dim diffs() As Double
    Dim absDiffs() As Variant
    Dim absDiffsSorted() As Variant
    Dim ranks() As Double
    Dim n As Long, i As Long
    Dim rankSumPos As Double, rankSumNeg As Double
    Dim W As Double, pValue As Variant
    Dim result(1 To 3, 1 To 2) As Variant
    Dim nonZeroDiffs As Long
    Dim lastRow As Long
    Dim mu As Double, sigma As Double, z As Double
    
    ' Get the last row of the shorter range
    lastRow = Application.WorksheetFunction.Min(FIND_LAST_ROW(range1), FIND_LAST_ROW(range2))
    
        ' Adjust for the header if present
    Dim startRow As Long
    If HasHeader Then
        startRow = 2
    Else
        startRow = 1
    End If
    
    ' Initialize arrays to store differences and their absolute values
    ReDim diffs(1 To lastRow - startRow + 1)
    ReDim absDiffs(1 To lastRow - startRow + 1)
    ReDim absDiffsSorted(1 To lastRow - startRow + 1)
    
        ' Calculate differences between the paired samples
    nonZeroDiffs = 0
    For i = startRow To lastRow
        row1Val = range1.Cells(i, 1).value
        row2Val = range2.Cells(i, 1).value
        
        If IsNumeric(row1Val) And IsNumeric(row2Val) Then
            If row1Val - row2Val <> 0 Then
                nonZeroDiffs = nonZeroDiffs + 1
                diffs(nonZeroDiffs) = row1Val - row2Val
                absDiffs(nonZeroDiffs) = Abs(diffs(nonZeroDiffs))
                absDiffsSorted(nonZeroDiffs) = absDiffs(nonZeroDiffs)
            End If
        End If
    Next i
    
    ' Resize the arrays based on the number of non-zero differences
    ReDim Preserve diffs(1 To nonZeroDiffs)
    ReDim Preserve absDiffs(1 To nonZeroDiffs)
    ReDim Preserve absDiffsSorted(1 To nonZeroDiffs)
    
    ' Sort the absolute differences
    QuickSort absDiffsSorted, LBound(absDiffsSorted), UBound(absDiffsSorted)
    'absDiffs = Application.WorksheetFunction.Sort(absDiffs, 1, 1, False)
    
    ' Initialize ranks array and assign ranks
    ReDim ranks(1 To nonZeroDiffs)
    For i = 1 To nonZeroDiffs
        ranks(i) = CalculateRank(absDiffs(i), absDiffsSorted)
    Next i
    
    ' Calculate the rank sums for positive and negative differences
    rankSumPos = 0
    rankSumNeg = 0
    For i = 1 To nonZeroDiffs
        If diffs(i) > 0 Then
            rankSumPos = rankSumPos + ranks(i)
        ElseIf diffs(i) < 0 Then
            rankSumNeg = rankSumNeg + ranks(i)
        End If
    Next i
    
    ' Wilcoxon test statistic (minimum of the positive and negative rank sums)
    W = Application.WorksheetFunction.Min(rankSumPos, rankSumNeg)
    
    ' Approximate p-value calculation (normal approximation for large n)
    If nonZeroDiffs >= 10 Then
        result(3, 1) = ""
        result(3, 2) = ""
    Else
        result(3, 1) = "Warning: "
        result(3, 2) = "Exact method needed for small sample sizes"
    End If
    
    ' Normal approximation for p-value
    mu = nonZeroDiffs * (nonZeroDiffs + 1) / 4
    sigma = Sqr(mu * (2 * nonZeroDiffs + 1) / 6)
    z = (Abs(W - mu) - 0.5) / sigma
    pValue = 2 * (1 - Application.WorksheetFunction.Norm_S_Dist(z, True))
    
    ' Prepare the result array
    result(1, 1) = "W-Statistic"
    result(1, 2) = W
    result(2, 1) = "P-value"
    result(2, 2) = pValue
    
    ' Return the result array
    WILCOXON_SIGNED_RANK = result
End Function



' Function: MANN_WHITNEY_TEST
'
' Description:
' Performs the Mann-Whitney U test (also known as the Wilcoxon rank-sum test), which is a non-parametric test
' used to determine whether there is a difference in the distribution of two independent samples.
'
' Arguments:
' - range1 (Range): The first range of data (sample 1).
' - range2 (Range): The second range of data (sample 2).
' - Optional HasHeader (Boolean, default = False): If True, the first row is treated as a header
'   and excluded from the analysis.
'
' Returns:
' A 2D array containing:
' 1. U-Statistic: The Mann-Whitney U statistic, which represents the number of times a value in the first group
'    is less than a value in the second group.
' 2. P-Value: The p-value for the test, calculated using the normal approximation.
'
' Details:
' 1. The function ranks all observations from both samples combined.
' 2. It calculates the sum of ranks for each sample and computes the U statistics.
' 3. The test statistic U is the smaller of U1 and U2.
' 4. For large samples, the p-value is approximated using a normal distribution, with a continuity correction
'    applied to the z-score.
Function MANN_WHITNEY_TEST(range1 As Range, range2 As Range, Optional HasHeader As Boolean = False) As Variant
    Dim values1() As Variant, values2() As Variant
    Dim valuesCombined() As Variant
    Dim n1 As Long, n2 As Long
    Dim i As Long
    Dim ranks1() As Double, ranks2() As Double
    Dim R1 As Double, R2 As Double
    Dim U1 As Double, U2 As Double, U As Double
    Dim mean As Double, sigma As Double, z As Double
    Dim pValue As Double
    Dim result(1 To 2, 1 To 2) As Variant
    
    ' Filter the data to remove missing values
    values1 = FILTER_DATA(range1, HasHeader)
    values2 = FILTER_DATA(range2, HasHeader)
    
    n1 = UBound(values1)
    n2 = UBound(values2)
    
    ' Combine the two value arrays into one for ranking
    ReDim valuesCombined(1 To n1 + n2)
    For i = 1 To n1
        If IsNumeric(values1(i)) Then
            valuesCombined(i) = CDbl(values1(i))
        Else
            MANN_WHITNEY_TEST = "Error: Non-numeric data in range1."
            Exit Function
        End If
    Next i
    For i = 1 To n2
        If IsNumeric(values2(i)) Then
            valuesCombined(n1 + i) = CDbl(values2(i))
        'Else
        '    MANN_WHITNEY_TEST = "Error: Non-numeric data in range2."
        '    Exit Function
        End If
    Next i
    
    ' Sort the combined array
    QuickSort valuesCombined, LBound(valuesCombined), UBound(valuesCombined)
    
    ' Initialize ranks arrays
    ReDim ranks1(1 To n1)
    ReDim ranks2(1 To n2)
    
    ' Calculate ranks for each value in values1 and values2
    R1 = 0
    R2 = 0
    For i = 1 To n1
        ranks1(i) = CalculateRank(values1(i), valuesCombined)
        R1 = R1 + ranks1(i)
    Next i
    For i = 1 To n2
        ranks2(i) = CalculateRank(values2(i), valuesCombined)
        R2 = R2 + ranks2(i)
    Next i
    
    ' Calculate U statistics
    U1 = n1 * n2 + (n1 * (n1 + 1)) / 2 - R1
    U2 = n1 * n2 + (n2 * (n2 + 1)) / 2 - R2
    U = Application.WorksheetFunction.Min(U1, U2)
    
    ' Calculate the mean, sigma, and z statistic
    mean = n1 * n2 / 2
    sigma = Sqr(n1 * n2 * (n1 + n2 + 1) / 12)
    z = (Abs(U - mean) - 0.5) / sigma  ' Continuity correction
    
    ' Calculate the p-value using the normal approximation
    pValue = 1 - Application.WorksheetFunction.Norm_S_Dist(z, True)
    
    ' Prepare the result array
    result(1, 1) = "U-Statistic"
    result(1, 2) = U
    result(2, 1) = "P-value"
    result(2, 2) = pValue
    
    ' Return the result array
    MANN_WHITNEY_TEST = result
End Function

