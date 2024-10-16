Attribute VB_Name = "HelperFunctions"
' Function Name: FILTER_DATA
' Purpose: Filters out missing values from a data range and optionally removes the header.
'
' Parameters:
'   dataRange (Range) - The range of data to filter.
'   HasHeader (Optional Boolean) - If True, the first row is considered a header and excluded from the result (default: False).
'
' Returns:
'   A variant array containing only non-missing values from the data range.
'   If HasHeader is True, the first row is excluded.
'
' Example:
'   cleanedData = FILTER_DATA(Range("A:A"), True)
'
' Notes:
'   This function handles both categorical and numerical data.
Function FILTER_DATA(dataRange As Range, Optional HasHeader As Boolean = False) As Variant
    Dim validData() As Variant
    Dim filteredData() As Variant
    Dim cell As Range
    Dim count As Long
    Dim lastRow As Long
    Dim result() As Variant
    Dim i As Long
    
    ' Find the last row with data in the range
    lastRow = FIND_LAST_ROW(dataRange)
    
    ' Count non-empty values up to the last row
    count = 0
    For Each cell In dataRange.Cells(1, 1).Resize(lastRow)
        If Not IsEmpty(cell.value) Then
            count = count + 1
        End If
    Next cell
    
    ' Resize the validData array
    ReDim validData(1 To count)
    
    ' Populate the validData array with non-empty values
    count = 0
    For Each cell In dataRange.Cells(1, 1).Resize(lastRow)
        If Not IsEmpty(cell.value) Then
            count = count + 1
            validData(count) = cell.value
        End If
    Next cell
    
    ' If HasHeader is True, exclude the first element
    If HasHeader Then
        ' Resize the array to exclude the first element
        ReDim filteredData(1 To UBound(validData) - 1)
        For i = 2 To UBound(validData)
            filteredData(i - 1) = validData(i)
        Next i
    Else
        ' No header, keep all valid data
        filteredData = validData
    End If
    
    ' Return the filtered data
    FILTER_DATA = filteredData
End Function


' Function Name: FIND_LAST_ROW
' Purpose: Determines the last non-empty row in a given data range.
'
' Parameters:
'   dataRange (Range) - The range of data to search for the last non-empty row.
'
' Returns:
'   A Long integer representing the row number of the last non-empty cell in the range.
'
' Example:
'   lastRow = FIND_LAST_ROW(Range("A:A"))
'
' Notes:
'   This function is useful for dynamic ranges where the number of rows with data is unknown.
'   It searches from the bottom of the range upward to locate the last filled cell.
Function FIND_LAST_ROW(dataRange As Range) As Long
    Dim lastRow As Long
    
    ' Find the last non-empty row in the provided data range
    lastRow = dataRange.Cells(dataRange.Rows.count, 1).End(xlUp).Row
    
    ' If the data range starts at a different row, adjust the result
    lastRow = lastRow - dataRange.Row + 1
    
    FIND_LAST_ROW = lastRow
End Function





' Function Name: DETECT_DATA_TYPE
' Purpose: Determines whether the data in the given range is numeric, categorical, or empty.
'
' Parameters:
'   dataRange (Range) - The range of data to analyze.
'   HasHeader (Optional Boolean) - If True, the function ignores the header and checks the second row (default: False).
'
' Returns:
'   A string indicating the data type:
'   - "Numeric" for numerical data
'   - "Categorical" for text or non-numeric data
'   - "Empty" if no valid data is detected
'
' Example:
'   dataType = DETECT_DATA_TYPE(Range("A:A"), True)
Function DETECT_DATA_TYPE(dataRange As Range, Optional HasHeader As Boolean = False) As String
    Dim firstValue As Variant
    
    ' Check the first value in the range
    If HasHeader Then
        firstValue = dataRange.Cells(2, 1).value
    Else
        firstValue = dataRange.Cells(1, 1).value
    End If
    
    If IsNumeric(firstValue) Then
        DETECT_DATA_TYPE = "Numeric"
    ElseIf IsEmpty(firstValue) Then
        DETECT_DATA_TYPE = "Empty"
    Else
        DETECT_DATA_TYPE = "Categorical"
    End If
End Function

' Function: CROSSTAB
'
' Purpose:
'   Creates a contingency table (cross-tabulation) of counts for pairs of values from two columns.
'   It returns a table with the counts of occurrences of each unique pair (Ai, Bi) from col1 and col2.
'
' Parameters:
'   - col1 (Range):
'       The first range (e.g., column A) containing categorical data for the cross-tabulation.
'
'   - col2 (Range):
'       The second range (e.g., column B) containing categorical data for the cross-tabulation.
'
'   - HasHeader (Optional, Boolean):
'       Set to True if the first row of both col1 and col2 contains headers and should be skipped (default is False).
'
' Returns:
'   - Variant:
'       A 2D array where:
'       - The first row contains the unique values from col2.
'       - The first column contains the unique values from col1.
'       - The interior cells contain the counts of occurrences of each pair (Ai, Bi).
'       If a pair (Ai, Bi) does not exist in the data, the corresponding cell will contain 0.
'
' Notes:
'   - The function ignores rows where either column contains empty values.
'   - The output array includes headers for the unique values from col1 and col2.
'   - The (Ai, Bi) pair counts are displayed in a matrix-like table.
'
' Example Usage:
'   CROSSTAB Range("A2:A100"), Range("B2:B100"), True
'   This creates a cross-tabulation of the data in columns A and B, skipping the first row (headers).
Function CROSSTAB(col1 As Range, col2 As Range, Optional HasHeader As Boolean = False) As Variant
    Dim dict As Object
    Dim i As Long
    Dim key As String
    Dim firstRow As Long, lastRow As Long
    Dim data1 As Variant, data2 As Variant
    Dim uniqueA As Object, uniqueB As Object
    Dim rowIdx As Long, colIdx As Long
    Dim table() As Variant
    Dim numRows As Long, numCols As Long
    
    ' Create dictionary to store (Ai, Bi) tuple counts
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Create dictionaries for unique values in A and B columns
    Set uniqueA = CreateObject("Scripting.Dictionary")
    Set uniqueB = CreateObject("Scripting.Dictionary")
    
    ' Load data into arrays for faster processing
    data1 = col1.value2
    data2 = col2.value2
    
    ' Determine the last row of data in the ranges
    lastRow = FIND_LAST_ROW(col1)
    
    ' Adjust for header
    firstRow = 1
    If HasHeader Then
        firstRow = 2
    End If
    
    ' Loop through the rows and process non-empty (Ai, Bi) pairs
    For i = firstRow To lastRow
        If Not IsEmpty(data1(i, 1)) And Not IsEmpty(data2(i, 1)) Then
            ' Create a unique key for the (Ai, Bi) tuple
            key = CStr(data1(i, 1)) & "|" & CStr(data2(i, 1))
            
            ' Check if the tuple exists in the dictionary
            If dict.Exists(key) Then
                dict(key) = dict(key) + 1 ' Increment the count
            Else
                dict.Add key, 1 ' Add the new tuple with a count of 1
            End If
            
            ' Store unique values of A and B
            If Not uniqueA.Exists(CStr(data1(i, 1))) Then uniqueA.Add CStr(data1(i, 1)), uniqueA.count + 1
            If Not uniqueB.Exists(CStr(data2(i, 1))) Then uniqueB.Add CStr(data2(i, 1)), uniqueB.count + 1
        End If
    Next i
    
    ' Set dimensions for the table (including headers)
    numRows = uniqueA.count + 1
    numCols = uniqueB.count + 1
    ReDim table(1 To numRows, 1 To numCols)
    
    ' Populate the header row (unique B values)
    For colIdx = 1 To uniqueB.count
        table(1, colIdx + 1) = uniqueB.Keys()(colIdx - 1)
    Next colIdx
    
    ' Populate the header column (unique A values)
    For rowIdx = 1 To uniqueA.count
        table(rowIdx + 1, 1) = uniqueA.Keys()(rowIdx - 1)
    Next rowIdx
    
    ' Populate the contingency table with counts
    For rowIdx = 1 To uniqueA.count
        For colIdx = 1 To uniqueB.count
            key = uniqueA.Keys()(rowIdx - 1) & "|" & uniqueB.Keys()(colIdx - 1)
            If dict.Exists(key) Then
                table(rowIdx + 1, colIdx + 1) = dict(key) ' Add the count
            Else
                table(rowIdx + 1, colIdx + 1) = 0 ' Add 0 if the combination doesn't exist
            End If
        Next colIdx
    Next rowIdx
    
    table(1, 1) = ""
    
    ' Return the array (contingency table)
    CROSSTAB = table
End Function


' Function Name: SPLIT_VALUES
'
' Purpose:
' This function splits a column of values (`valRange`) into separate columns based on unique group IDs from another column (`groupRange`). The result is an array where each column contains values corresponding to a unique group. Any empty cells are populated with blanks instead of zeros.
'
' Parameters:
' - valRange (Range): The range of values that will be split based on the group labels.
' - groupRange (Range): The range of group labels used to split the `valRange` values into different columns.
' - HasHeader (Optional, Boolean, default=False):
'   - If set to `True`, the first row of both `valRange` and `groupRange` is assumed to contain headers and will be excluded from the data processing.
'   - If set to `False`, all rows will be included in the data processing.
'
' Returns:
' - Variant: A 2D array where each column represents the values for a specific group, and the first row contains the group names (headers). Empty cells are filled with empty strings instead of zeros.
'
' Example Usage:
' =SPLIT_VALUES(A2:A10, B2:B10, TRUE)
'
' Notes:
' - The `FIND_LAST_ROW` helper function is used to determine the last row with data in both `valRange` and `groupRange`.
' - The function loops through both columns simultaneously, ensuring the values remain aligned based on their corresponding group labels.
' - If a group has fewer values than the maximum group size, the remaining rows in that column are populated with blanks ("").
'
' Example:
' If you have the following input ranges:
'   A (Values)   | B (Groups)
'   10           | Group1
'   12           | Group2
'   11           | Group1
'   15           | Group3
'   14           | Group2
'   13           | Group1
'
' The output will be:
'   Group1  | Group2  | Group3
'   10      | 12      | 15
'   11      | 14      |
'   13      |         |
'
' Helper Function:
' - FIND_LAST_ROW: Finds the last row with data in a range.
Function SPLIT_VALUES(valRange As Range, groupRange As Range, Optional HasHeader As Boolean = False) As Variant
    Dim groupDict As Object
    Dim i As Long, j As Long
    Dim groupName As Variant
    Dim result As Variant
    Dim lastRow As Long
    Dim startRow As Long
    Dim groupIndex As Long
    Dim val As Variant

    ' Determine start row based on whether there is a header
    If HasHeader Then
        startRow = 2
    Else
        startRow = 1
    End If

    ' Find the last row with data
    lastRow = Application.WorksheetFunction.Min(FIND_LAST_ROW(valRange), FIND_LAST_ROW(groupRange))

    ' Create a dictionary to store unique groups and their associated values
    Set groupDict = CreateObject("Scripting.Dictionary")

    ' Loop through both valRange and groupRange simultaneously
    For i = startRow To lastRow
        groupName = groupRange.Cells(i, 1).value
        val = valRange.Cells(i, 1).value
        
        ' Only process non-empty values
        If Not IsEmpty(groupName) And Not IsEmpty(val) Then
            If Not groupDict.Exists(groupName) Then
                ' Initialize a new collection for the group if it doesn't exist
                groupDict.Add groupName, New Collection
            End If
            ' Add the value to the group's collection
            groupDict(groupName).Add val
        End If
    Next i

    ' Create the result array to hold the split values (number of rows and columns)
    Dim maxValues As Long
    maxValues = 0
    For Each groupName In groupDict
        maxValues = Application.WorksheetFunction.Max(maxValues, groupDict(groupName).count)
    Next groupName

    ReDim result(1 To maxValues + 1, 1 To groupDict.count)

    ' Populate the first row with group names (headers)
    j = 1
    For Each groupName In groupDict.Keys
        result(1, j) = groupName
        j = j + 1
    Next groupName

    ' Populate the result array with values for each group
    For j = 1 To groupDict.count
        groupIndex = 2 ' Start from row 2 (under the headers)
        For i = 1 To groupDict(groupDict.Keys()(j - 1)).count
            result(groupIndex, j) = CDbl(groupDict(groupDict.Keys()(j - 1))(i))
            groupIndex = groupIndex + 1
        Next i

        ' Fill remaining rows in the column with empty strings
        For i = groupIndex To maxValues + 1
            result(i, j) = "" ' Populate empty cells with empty string
        Next i
    Next j
    

    ' Return the result array
    SPLIT_VALUES = result
End Function



' Helper function to calculate ranks in an array with early stopping
Function CalculateRank(value As Variant, valuesArray() As Variant) As Double
    Dim i As Long, rank As Double, count As Long
    rank = 1
    count = 0
    
    For i = LBound(valuesArray) To UBound(valuesArray)
        If valuesArray(i) < value Then
            rank = rank + 1
        ElseIf valuesArray(i) = value Then
            count = count + 1
        Else
            Exit For ' Early exit once valuesArray(i) > value
        End If
    Next i
    
    ' Return the average rank in case of ties
    CalculateRank = rank + (count - 1) / 2
End Function


' QuickSort function to sort an array in ascending or descending order
Sub QuickSort(ByRef arr() As Variant, ByVal low As Long, ByVal high As Long, Optional ByVal ascending As Boolean = True)
    Dim pivotIndex As Long
    If low < high Then
        ' Partition the array and get the pivot index
        pivotIndex = Partition(arr, low, high, ascending)
        ' Recursively sort elements before and after partition
        QuickSort arr, low, pivotIndex - 1, ascending
        QuickSort arr, pivotIndex + 1, high, ascending
    End If
End Sub

' Partition function to partition the array based on the pivot element
Private Function Partition(ByRef arr() As Variant, ByVal low As Long, ByVal high As Long, ByVal ascending As Boolean) As Long
    Dim pivot As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    
    ' Choose the pivot element (here, we take the last element as pivot)
    pivot = arr(high)
    i = low - 1
    
    ' Loop through the array and rearrange elements
    For j = low To high - 1
        If (ascending And arr(j) <= pivot) Or (Not ascending And arr(j) >= pivot) Then
            i = i + 1
            ' Swap arr(i) and arr(j)
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
        End If
    Next j
    
    ' Place the pivot element in the correct position
    temp = arr(i + 1)
    arr(i + 1) = arr(high)
    arr(high) = temp
    
    ' Return the index of the pivot element
    Partition = i + 1
End Function

