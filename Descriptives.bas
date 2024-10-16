Attribute VB_Name = "Descriptives"
' Function Name: DESCRIBE_CATEGORICAL
' Purpose: Summarizes categorical data, producing a frequency table and proportions.
'
' Parameters:
'   dataRange (Range) - The range of categorical data to analyze.
'   HasHeader (Optional Boolean) - If True, the first row is considered a header and excluded (default: False).
'
' Returns:
'   A 2D array with the following:
'   - Categories
'   - Frequencies
'   - Proportions
'
' Example:
'   result = DESCRIBE_CATEGORICAL(Range("A:A"), True)
Function DESCRIBE_CATEGORICAL(dataRange As Range, Optional HasHeader As Boolean = False) As Variant
    Dim cleaned As Variant
    Dim header As String
    Dim uniqueCategories As Object
    Dim category As Variant
    Dim cell As Range
    Dim countDict As Object
    Dim categoryArray() As Variant
    Dim freqArray() As Long
    Dim totalCount As Long
    Dim i As Long, j As Long
    Dim result() As Variant
    Dim cleanedDataRange As Variant
    Dim cellValue As Variant
    
    cleanedDataRange = FILTER_DATA(dataRange, HasHeader)
    
    ' Create a dictionary to store unique categories and counts
    Set uniqueCategories = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    
    ' Count occurrences of each category
    For i = LBound(cleanedDataRange) To UBound(cleanedDataRange)
        cellValue = cleanedDataRange(i)
        If Not IsEmpty(cellValue) Then
            If uniqueCategories.Exists(cellValue) Then
                countDict(cellValue) = countDict(cellValue) + 1
            Else
                uniqueCategories.Add cellValue, cellValue
                countDict.Add cellValue, 1
            End If
        End If
    Next i
    
    ' Calculate the total count of non-empty cells
    totalCount = Application.WorksheetFunction.CountA(cleanedDataRange)
    
    ' Create arrays to store categories and frequencies
    ReDim categoryArray(1 To uniqueCategories.count)
    ReDim freqArray(1 To uniqueCategories.count)
    
    i = 1
    For Each category In uniqueCategories
        categoryArray(i) = category
        freqArray(i) = countDict(category)
        i = i + 1
    Next category
    
    ' Store the results in a 2D array
    ReDim result(1 To uniqueCategories.count + 1, 1 To 3)
    
    If HasHeader Then
        result(1, 1) = dataRange(1, 1).value
    Else
        result(1, 1) = "Category"
    End If
    result(1, 2) = "Frequency"
    result(1, 3) = "Proportion"
    
    For j = 1 To UBound(categoryArray)
        result(j + 1, 1) = categoryArray(j)
        result(j + 1, 2) = freqArray(j)
        result(j + 1, 3) = freqArray(j) / totalCount
    Next j
    
    ' Return the frequency table with proportions
    DESCRIBE_CATEGORICAL = result
End Function

' Function Name: DESCRIBE_NUMERICAL
' Purpose: Summarizes numerical data by calculating key descriptive statistics such as mean, median, standard deviation, and quartiles.
'
' Parameters:
'   dataRange (Range) - The range of data to analyze (can include a header).
'   HasHeader (Optional Boolean) - If True, the first row is considered a header and excluded from analysis (default: False).
'
' Returns:
'   A 2D array with summary statistics including:
'   - n (Count of values)
'   - Mean
'   - Median
'   - Standard Deviation
'   - Minimum
'   - Maximum
'   - First quartile (Q1)
'   - Third quartile (Q3)
'
' Example:
'   result = DESCRIBE_NUMERICAL(Range("A:A"), True)
'
' Notes:
'   This function uses FILTER_DATA to clean the input data, removing missing values and handling headers if present.
Function DESCRIBE_NUMERICAL(dataRange As Range, Optional HasHeader As Boolean = False) As Variant
    Dim cleanedData As Variant
    Dim n As Integer
    Dim meanVal As Double, medianVal As Double, stdDev As Double
    Dim minVal As Double, maxVal As Double, q1 As Double, q3 As Double, iqr As Double
    Dim result(1 To 9, 1 To 2) As Variant
    
    ' Filter the data and handle the header
    cleanedData = FILTER_DATA(dataRange, HasHeader)
    
    ' Calculate length of array
    n = UBound(cleanedData)
    
    ' Ensure cleanedData is not empty
    If UBound(cleanedData) = 0 Then
        DESCRIBE_NUMERICAL = "No valid data"
        Exit Function
    End If
    
    
    ' Calculate summary statistics
    meanVal = Application.WorksheetFunction.Average(cleanedData)
    medianVal = Application.WorksheetFunction.Median(cleanedData)
    stdDev = Application.WorksheetFunction.StDev_S(cleanedData)
    minVal = Application.WorksheetFunction.Min(cleanedData)
    maxVal = Application.WorksheetFunction.Max(cleanedData)
    q1 = Application.WorksheetFunction.Quartile_Inc(cleanedData, 1)
    q3 = Application.WorksheetFunction.Quartile_Inc(cleanedData, 3)
    iqr = q3 - q1
    
    ' Store results in a 2D array
    result(1, 1) = "Metric"
    result(1, 2) = "Value"
    result(2, 1) = "n"
    result(2, 2) = n
    result(3, 1) = "Mean"
    result(3, 2) = meanVal
    result(4, 1) = "Median"
    result(4, 2) = medianVal
    result(5, 1) = "Standard Deviation"
    result(5, 2) = stdDev
    result(6, 1) = "Min"
    result(6, 2) = minVal
    result(7, 1) = "Max"
    result(7, 2) = maxVal
    result(8, 1) = "Q1"
    result(8, 2) = q1
    result(9, 1) = "Q3"
    result(9, 2) = q3
    
    ' Return the summary statistics
    DESCRIBE_NUMERICAL = result
End Function

' Function Name: DESCRIBE
' Purpose: Automatically detects the data type in a range and calls the appropriate summary function.
'
' Parameters:
'   dataRange (Range) - The range of data to analyze (can include a header).
'   HasHeader (Optional Boolean) - If True, the first row is considered a header and excluded (default: False).
'
' Returns:
'   A summary of the data:
'   - For numerical data, it calls DESCRIBE_NUMERICAL
'   - For categorical data, it calls DESCRIBE_CATEGORICAL
'
' Example:
'   result = DESCRIBE(Range("A:A"), True)
'
' Notes:
'   This function uses DETECT_DATA_TYPE to determine whether the data is numerical or categorical.
Function DESCRIBE(dataRange As Range, Optional HasHeader As Boolean = False) As Variant
    Dim dataType As String
    Dim result As Variant
    
    ' Detect the data type (Numeric, Categorical, or Empty)
    dataType = DETECT_DATA_TYPE(dataRange, HasHeader)
    
    ' Based on the data type, call the appropriate describe function
    Select Case dataType
        Case "Numeric"
            result = DESCRIBE_NUMERICAL(dataRange, HasHeader)
        Case "Categorical"
            result = DESCRIBE_CATEGORICAL(dataRange, HasHeader)
        Case "Empty"
            result = "No valid data"
        Case Else
            result = "Unknown data type"
    End Select
    
    ' Return the result
    DESCRIBE = result
End Function

