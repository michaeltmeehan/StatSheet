Attribute VB_Name = "ConfidenceIntervals"
' Function Name: CI_MEAN
' Purpose: Calculates the confidence interval for the mean of a numeric dataset.
'
' Parameters:
'   dataRange (Range) - The range of numeric data to analyze.
'   Level (Optional Double) - The confidence level (default: 0.95).
'   HasHeader (Optional Boolean) - If True, the first row is considered a header and excluded (default: False).
'
' Returns:
'   A 3-element array with:
'   - The mean
'   - The lower bound of the confidence interval
'   - The upper bound of the confidence interval
'
' Example:
'   ci = CI_MEAN(Range("A:A"), 0.95, True)

Function CI_MEAN(dataRange As Range, Optional Level As Double = 0.95, Optional HasHeader As Boolean = False) As Variant
    Dim n As Integer
    Dim mean As Double, se As Double, lower As Double, upper As Double
    Dim result(1 To 3) As Double
    Dim cleanedData As Variant

    ' Clean data
    cleanedData = FILTER_DATA(dataRange, HasHeader)
    
    ' Calculate statistics
    n = UBound(cleanedData)
    If n = 0 Then
        CI_MEAN = "No valid data"
        Exit Function
    End If
    
    mean = Application.WorksheetFunction.Average(cleanedData)
    se = Application.WorksheetFunction.StDev_S(cleanedData) / Sqr(n)
    lower = mean - Application.WorksheetFunction.T_Inv(1 - (1 - Level) / 2, n - 1) * se
    upper = mean + Application.WorksheetFunction.T_Inv(1 - (1 - Level) / 2, n - 1) * se
    
    ' Store and return result
    result(1) = mean
    result(2) = lower
    result(3) = upper
    CI_MEAN = result
End Function


' Function: CI_GROUPED
'
' Purpose:
'   Calculates confidence intervals for the mean of a numeric variable
'   for each subgroup defined by a categorical variable.
'   The function returns a table with the group name, mean, and lower and upper bounds
'   of the confidence interval for each group.
'
' Parameters:
'   - valuesRange (Range):
'       Range containing the numeric values for which the mean and confidence interval are calculated.
'
'   - groupsRange (Range):
'       Range containing the categorical group labels corresponding to each value in valuesRange.
'       Each label defines a group for which the statistics will be calculated.
'
'   - Level (Optional, Double):
'       The confidence level for the confidence interval (default is 0.95 for a 95% confidence interval).
'
'   - HasHeader (Optional, Boolean):
'       Set to True if the first row of the ranges contains headers and should be skipped (default is False).
'
' Returns:
'   - Variant:
'       A 2D array where each row corresponds to a group. The columns represent:
'       1. Group name
'       2. Mean of the group's numeric values
'       3. Lower bound of the confidence interval
'       4. Upper bound of the confidence interval
'
' Notes:
'   - Groups with fewer than 2 numeric values will return "N/A" for the confidence interval.
'   - Only non-empty, numeric values are included in the calculations.
'   - If HasHeader is True, the function skips the first row in both valuesRange and groupsRange.
'
' Example Usage:
'   CI_GROUPED Range("A2:A100"), Range("B2:B100"), 0.95, True
'   This calculates 95% confidence intervals for the means of values in column A,
'   grouped by the labels in column B, with the first row containing headers.
Function CI_GROUPED(valuesRange As Range, groupsRange As Range, Optional Level As Double = 0.95, Optional HasHeader As Boolean = False) As Variant
    Dim groupDict As Object
    Dim groupName As Variant
    Dim n As Integer, i As Long, startRow As Long
    Dim mean As Double, se As Double, lower As Double, upper As Double
    Dim result As Variant
    Dim groupValues As Collection
    Dim groupCount As Long
    Dim tempArr() As Double
    Dim j As Long
    
    ' Determine start row based on whether there's a header
    If HasHeader Then
        startRow = 2 ' Skip the first row if header exists
    Else
        startRow = 1
    End If
    
    ' Create a dictionary to store unique groups
    Set groupDict = CreateObject("Scripting.Dictionary")
    
    ' Find the last row based on the minimum length of the valueRange and groupRange
    lastRow = WorksheetFunction.Min(FIND_LAST_ROW(valuesRange), FIND_LAST_ROW(groupsRange))
    
    ' Loop through the group labels to find unique groups
    For i = startRow To lastRow
        groupName = groupsRange.Cells(i, 1).value
        
        ' Ensure the value is numeric and the group is valid
        If IsNumeric(valuesRange.Cells(i, 1).value) And Not IsEmpty(groupName) Then
            If Not groupDict.Exists(groupName) Then
                groupDict.Add groupName, New Collection
            End If
            ' Add the corresponding value to the group
            groupDict(groupName).Add valuesRange.Cells(i, 1).value
        End If
    Next i
    
    ' Initialize the result array for the output (number of groups by 4 columns: Group, Mean, Lower CI, Upper CI)
    groupCount = groupDict.count
    ReDim result(1 To (groupCount + 1), 1 To 4)
    
    ' Add labels to the result array
    If HasHeader Then
        result(1, 1) = groupsRange.Cells(1, 1).value
    Else
        result(1, 1) = "Group"
    End If
    result(1, 2) = "Mean"
    result(1, 3) = "Lower"
    result(1, 4) = "Upper"
    
    ' Loop through each unique group and calculate statistics
    i = 1
    For Each groupName In groupDict.Keys
        Set groupValues = groupDict(groupName)
        n = groupValues.count
        
        ' Convert the collection of group values to an array for the worksheet functions
        ReDim tempArr(1 To n)
        For j = 1 To n
            tempArr(j) = groupValues(j)
        Next j
        
        If n > 1 Then
            ' Calculate the mean
            mean = Application.WorksheetFunction.Average(tempArr)
            
            ' Calculate the standard error
            se = Application.WorksheetFunction.StDev_S(tempArr) / Sqr(n)
            
            ' Calculate the confidence interval
            lower = mean - Application.WorksheetFunction.T_Inv(1 - (1 - Level) / 2, n - 1) * se
            upper = mean + Application.WorksheetFunction.T_Inv(1 - (1 - Level) / 2, n - 1) * se
            
            ' Store the results for this group
            result(i + 1, 1) = groupName ' Group name
            result(i + 1, 2) = mean      ' Mean
            result(i + 1, 3) = lower     ' Lower CI
            result(i + 1, 4) = upper     ' Upper CI
        Else
            ' If group size is too small for CI, store N/A values
            result(i + 1, 1) = groupName
            result(i + 1, 2) = Application.WorksheetFunction.Average(tempArr)
            result(i + 1, 3) = "N/A"
            result(i + 1, 4) = "N/A"
        End If
        i = i + 1
    Next groupName
    
    ' Return the result array
    CI_GROUPED = result
End Function

