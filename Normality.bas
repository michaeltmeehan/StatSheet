Attribute VB_Name = "Normality"
' ********************************************************************
' Function: SHAPIRO_TEST
' Description: This function performs the Shapiro-Wilk test for normality
'              on a dataset provided in `vals`. It tests the null hypothesis
'              that the data is normally distributed.
'
' Arguments:
'   - vals (Range): The range of numeric values for which to perform the Shapiro-Wilk test.
'   - HasHeader (Boolean) [Optional]: A flag indicating whether the range contains a header row.
'        Defaults to False. If True, the first row is ignored in the test.
'
' Returns:
'   - A 2D array with the following statistics:
'     - "W": The W statistic from the Shapiro-Wilk test.
'     - "P-value": The p-value associated with the W statistic, indicating
'                  the probability of obtaining such a value under the null hypothesis.
'
' Steps:
'   1. The data is cleaned using `FILTER_DATA` to remove missing values.
'   2. The data is sorted in ascending order.
'   3. The expected normal order statistics (`mi`) are computed using the inverse normal distribution function.
'   4. The `a` coefficients, used to construct the test statistic W, are calculated based on the sample size.
'   5. The W statistic is computed as the square of the correlation between the `a` values and the sorted data.
'   6. A transformation is applied to W to obtain a z-score, and a p-value is computed from the normal distribution.
'
' Error Handling:
'   - The function assumes that the input range contains numeric data.
'   - If the number of data points (n) is too small, the results may not be valid.
'     Typically, the Shapiro-Wilk test is designed for sample sizes between 3 and 2000.

' Example Usage:
'   - =SHAPIRO_TEST(A:A, TRUE)
'     Performs the Shapiro-Wilk test for normality on the numeric data in column A,
'     assuming the first row is a header.

'   - =SHAPIRO_TEST(A2:A51)
'     Performs the Shapiro-Wilk test for normality on the numeric data in range A2:A51
'     without a header row.
' ********************************************************************
Function SHAPIRO_TEST(vals As Range, Optional HasHeader As Boolean = False) As Variant
    Dim cleanedVals As Variant
    Dim sortedVals As Variant
    Dim n As Integer
    Dim mi() As Double
    Dim m As Double, U As Double, epsilon As Double
    Dim a() As Double
    Dim W As Double
    Dim mu As Double, sigma As Double, z As Double, pValue As Double
    Dim result(1 To 2, 1 To 2) As Variant
    
    cleanedVals = FILTER_DATA(vals, HasHeader)
    sortedVals = Application.WorksheetFunction.Sort(cleanedVals, 1, 1, True)
    
    n = UBound(sortedVals)
    
    m = 0
    
    ReDim mi(1 To n)
    For i = 1 To n
        'mi(i) = 1
        mi(i) = Application.WorksheetFunction.Norm_S_Inv((i - 0.375) / (n + 0.25))
        m = m + mi(i) ^ 2
    Next i
    
    U = 1 / Sqr(n)
    
    ReDim a(1 To n)
    a(n) = -2.706056 * U ^ 5 + 4.434685 * U ^ 4 - 2.07119 * U ^ 3 - 0.147981 * U ^ 2 + 0.221157 * U + mi(n) * m ^ (-0.5)
    a(n - 1) = -3.582633 * U ^ 5 + 5.682633 * U ^ 4 - 1.752461 * U ^ 3 - 0.293762 * U ^ 2 + 0.042981 * U + mi(n - 1) * m ^ (-0.5)
    
    a(1) = -a(n)
    a(2) = -a(n - 1)
    
    epsilon = (m - 2 * mi(n) ^ 2 - 2 * mi(n - 1) ^ 2) / (1 - 2 * a(n) ^ 2 - 2 * a(n - 1) ^ 2)
    
    For i = 3 To n - 2
        a(i) = mi(i) / Sqr(epsilon)
    Next i
    
    W = Application.WorksheetFunction.Correl(a, sortedVals) ^ 2
    
    mu = 0.0038915 * Log(n) ^ 3 - 0.083751 * Log(n) ^ 2 - 0.31082 * Log(n) - 1.5861
    sigma = Exp(0.0030301 * Log(n) ^ 2 - 0.082676 * Log(n) - 0.4803)
    
    z = Application.WorksheetFunction.Standardize(Log(1 - W), mu, sigma)
    
    pValue = 1 - Application.WorksheetFunction.Norm_S_Dist(z, True)
    
    result(1, 1) = "W"
    result(1, 2) = W
    result(2, 1) = "P-value"
    result(2, 2) = pValue
    
    SHAPIRO_TEST = result 'Application.WorksheetFunction.SumSq(a)
End Function


