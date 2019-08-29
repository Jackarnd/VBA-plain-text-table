Function TableToText(Table As Range) As String
'
' Script by Jack Arnoldi
' version: 1.0
' 29 August 2019
'
' This function transform a Range object into a plain text table.
'
'   Input: 'Table' is the Range object to turn into a plain text table
'   Output: 'TableToText' is a string that contains the plain text table
'
    Dim nbColumns As Integer, nbRows As Integer
    Dim i As Integer, j As Integer, s As Integer
    Dim max As Integer, difference As Integer
    Dim arr As Variant '(1 To Table.Rows.Count, 1 To Table.Columns.Count)
    Dim MaxNbCh As Integer
    Dim CarriageReturn As Boolean
    arr = Table.Value ' Range are passed ByRef so we use another array
    
    MaxNbCh = 130 ' defines the maximum number of characters for a line
    
    nbColumns = Table.Columns.Count
    nbRows = Table.Rows.Count
    MaxNbCh = MaxNbCh / nbColumns
    '' cutting too long cells
    i = LBound(arr, 1)

    While i <= UBound(arr, 1) ' swipes through the whole array.
        ' detecting if there is a need of carriage return in the line
        CarriageReturn = False
        For j = LBound(arr, 2) To UBound(arr, 2)
            If Len(arr(i, j)) >= MaxNbCh Then
                CarriageReturn = True
            End If
        Next j
        
        If CarriageReturn = True Then
        ' adding a new line and moving the data down
            arr = AddLine(arr, i + 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                If Len(arr(i, j)) >= MaxNbCh Then
                    ' put the part of the array element that is too long into the next line of the array
                    arr(i + 1, j) = Right(arr(i, j), Len(arr(i, j)) - MaxNbCh)
                    ' removes the moved part of the stri
                    arr(i, j) = Left(arr(i, j), MaxNbCh)
                End If
            Next j
        Else
        ' if no need for carriage we add a new line in the array with the
        '   array bottom lines
            arr = AddLine(arr, i + 1)
            i = i + 1
            arr(i, 1) = "+"
            For j = 1 To UBound(arr, 2)
                For s = 1 To MaxNbCh
                    arr(i, j) = arr(i, j) + "-"
                Next s
                arr(i, j) = arr(i, j) + "+"
            Next j
        End If
    i = i + 1
    Wend


    '' adding the spaces
        ' Adding the spaces and the |
    For j = 1 To UBound(arr, 2)
        For i = 1 To UBound(arr, 1)
            difference = MaxNbCh - Len(arr(i, j)) ' Get the number of spaces needed
            
            ' add | at the beginning of the line
            If j = 1 And Right(arr(i, 1), 1) <> "+" Then
                arr(i, j) = "|" + CStr(arr(i, j))
            End If
            
            ' add the spaces if the string isn't big enough
            If Len(arr(i, j)) < MaxNbCh Then
                For s = 1 To difference
                    arr(i, j) = CStr(arr(i, j)) + " "
                Next s
            End If
            
            ' add | at the end of the line
            If Right(arr(i, 1), 1) <> "+" Then
                arr(i, j) = CStr(arr(i, j)) + "|"
            End If
        Next i
    Next j
    
    '' Creating the plain text table string
    For i = 1 To UBound(arr, 1)
        TableToText = TableToText + Chr(9) ' This tells to Collabnet to format the string as code
        ' Adding each element of the array to the string
        For j = LBound(arr, 2) To UBound(arr, 2)
            TableToText = TableToText + arr(i, j)
        Next j
        TableToText = TableToText & vbCrLf ' this is a carriage return
    Next i
    Debug.Print TableToText
End Function
Private Function AddLine(iArray As Variant, Line As Integer) As Variant
'
' Script by Jack Arnoldi
' version: 1.0
' 29 August 2019
'
' This function adds a line in an array.
'
'   Inputs:  - 'iArray' is array in which you want to add a line
'            - 'Line' is the index of the line you want to add
'   Output: 'AddLine' is the array containing one more line
'
    Dim oArray As Variant
    iArray = ReDimPreserve(iArray, UBound(iArray, 1) + 1, UBound(iArray, 2))
    
    ReDim oArray(UBound(iArray, 1), UBound(iArray, 2))
    
    ' moving the data
    For i = 1 To UBound(iArray, 1) - 1
        For j = LBound(iArray, 2) To UBound(iArray, 2)
            If i < Line Then
                oArray(i, j) = iArray(i, j)
            Else
                oArray(i + 1, j) = iArray(i, j)
            End If
        Next
    Next
    AddLine = oArray
End Function


Private Function ReDimPreserve(MyArray As Variant, nNewFirstUBound As Long, nNewLastUBound As Long) As Variant
    '
    ' Function taken from the website
    '   https://wellsr.com/vba/2016/excel/dynamic-array-with-redim-preserve-vba/
    '
    ' re-dimensionalizes either the first dimension, the last dimension,
    '   or BOTH dimensions of my 2D array at the same time.
    '
    Dim i, j As Long
    Dim nOldFirstUBound, nOldLastUBound, nOldFirstLBound, nOldLastLBound As Long
    Dim TempArray() As Variant 'Change this to "String" or any other data type if want it to work for arrays other than Variants. MsgBox UCase(TypeName(MyArray))
'---------------------------------------------------------------
'COMMENT THIS BLOCK OUT IF YOU CHANGE THE DATA TYPE OF TempArray
    If InStr(1, UCase(TypeName(MyArray)), "VARIANT") = 0 Then
        MsgBox "This function only works if your array is a Variant Data Type." & vbNewLine & _
               "You have two choice:" & vbNewLine & _
               " 1) Change your array to a Variant and try again." & vbNewLine & _
               " 2) Change the DataType of TempArray to match your array and comment the top block out of the function ReDimPreserve" _
                , vbCritical, "Invalid Array Data Type"
        End
    End If
'---------------------------------------------------------------
    ReDimPreserve = False
    'check if its in array first
    If Not IsArray(MyArray) Then MsgBox "You didn't pass the function an array.", vbCritical, "No Array Detected": End
    
    'get old lBound/uBound
    nOldFirstUBound = UBound(MyArray, 1): nOldLastUBound = UBound(MyArray, 2)
    nOldFirstLBound = LBound(MyArray, 1): nOldLastLBound = LBound(MyArray, 2)
    'create new array
    ReDim TempArray(nOldFirstLBound To nNewFirstUBound, nOldLastLBound To nNewLastUBound)
    'loop through first
    For i = LBound(MyArray, 1) To nNewFirstUBound
        For j = LBound(MyArray, 2) To nNewLastUBound
            'if its in range, then append to new array the same way
            If nOldFirstUBound >= i And nOldLastUBound >= j Then
                TempArray(i, j) = MyArray(i, j)
            End If
        Next
    Next
    'return the array redimmed
    If IsArray(TempArray) Then ReDimPreserve = TempArray
End Function

Sub DisplayArray(arr As Variant)
'
' Script from https://stackoverflow.com/questions/14274949/how-to-print-two-dimensional-array-in-immediate-window-in-vba/24037033#24037033
'
' This function displays an array in the immediate window.
'
'   Input: 'arr' is the array object you want to display
'
    Dim iSubA As Long
    Dim jSubA As Long
    Dim rowString As String

    Debug.Print "The array is: "
    For iSubA = 1 To UBound(arr, 1)
        rowString = arr(iSubA, 1)
        For jSubA = 2 To UBound(arr, 2)
            rowString = rowString & "," & arr(iSubA, jSubA)
        Next jSubA
        Debug.Print rowString
    Next iSubA
End Sub