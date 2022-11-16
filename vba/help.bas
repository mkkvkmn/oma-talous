Option Explicit

Sub sort_blanks(tbl As ListObject, cln_str As String, Optional empty_first As Boolean)
Dim rng As Range
Dim blanks As Range

If empty_first Then
    'replace blanks to help with sorting
    On Error Resume Next
    Set blanks = tbl.ListColumns(cln_str).Range.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    If Not blanks Is Nothing Then
        With blanks
            .Value = "^"
            .Interior.Color = RGB(254, 45, 45)
        End With
        'sort category table
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tbl.ListColumns(cln_str).Range, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .apply
        End With
        Set rng = find_all(tbl.DataBodyRange, "^")
        
        If Not rng Is Nothing Then
            rng.Value = ""
        End If
    End If
Else
    'use natural order
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns(cln_str).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .apply
    End With
    'remove background color
    tbl.Range.Interior.Color = xlNone
End If

End Sub

Function get_end_column(start As String, count As Long) As String
'variables
Dim start_len As Long
Dim start_asc As String
Dim arr(2) As Variant
Dim res As String
Dim ascii_chars As Long
'A = 65, Z = 90 (https://www.asciitable.com/)
ascii_chars = Asc("Z") - Asc("A") + 1 '26 chars
start = UCase(start)
count = count - 1 'total amount of columns, not dif
start_len = Len(start)
'init array
For i = 1 To start_len
    'split string if needed
    If start_len > 1 Then
        start_asc = Asc(Mid(StrReverse(start), i, 1))
    Else
        start_asc = Asc(start)
    End If
    'char is between 65-90
    If Not IsEmpty(start_asc) Then
        arr(i - 1) = start_asc - 64 'A / 65 becomes 1, makes addition possible
    Else
        arr(i - 1) = 0
    End If
Next i
'add chars
For i = 1 To count
    'add "ones"
    arr(0) = arr(0) + 1
    If arr(0) > ascii_chars Then 'last char reached, add new char
        arr(0) = 1
        arr(1) = arr(1) + 1
        'add "tens"
        If arr(1) > ascii_chars Then 'last char reached, add new char
            arr(1) = 1
            'add "hundreds"
            arr(2) = arr(2) + 1
        End If
    End If
Next i
'result
For i = UBound(arr) To 0 Step -1
     If Not IsEmpty(arr(i)) Then
         res = res & Chr(arr(i) + 64) '+64 return the actual char in ascii
     End If
Next i
get_end_column = res
End Function

'https://stackoverflow.com/questions/62236965/how-to-automatically-increase-letter-by-one-to-get-the-next-letter-in-excel
Function next_letr(s As String, nth_letter As Long) As String
    Dim i As Long, L As Long, arrL, arrN

    L = Len(s)

    If L = 1 And Asc(s) + nth_letter < 91 Then
        next_letr = Chr(Asc(s) + nth_letter)
        If next_letr = "[" Then
            next_letr = "AA"
        End If
        Exit Function
    End If

    ReDim arrL(1 To L) As String
    ReDim arrN(1 To L) As Long

    For i = L To 1 Step -1
        arrL(i) = Mid(s, i, 1)
        arrN(i) = Asc(arrL(i))
    Next i
    For i = L To 1 Step -1
        arrN(i) = arrN(i) + nth_letter
        If i = 1 Then Exit For
        If arrN(i) < 91 Then Exit For
        arrN(i) = 65
    Next i

    For i = 1 To L
        arrL(i) = Chr(arrN(i))
    Next i

    next_letr = Join(arrL, "")
    If Left(next_letr, 1) = "[" Then
        next_letr = "AA" & Mid(next_letr, 2)
    End If
End Function

Function find_table_next_row_rng(tbl As ListObject) As Range
    On Error GoTo 0
    Dim row_no As Long
    Dim str_first_row As String
    'get row count
    row_no = tbl.Range.Rows.count
    'if no data, set next row as tables first row
    If Len(tbl.Range.Cells(2, 1)) = 0 Then
        row_no = 1
    End If
    'get next row
    Set find_table_next_row_rng = tbl.Range.Offset(row_no).Resize(1, 1)
End Function


'https://stackoverflow.com/questions/19504858/find-all-matches-in-workbook-using-excel-vba
Function find_all(rng As Range, what As Variant, Optional LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlWhole, Optional SearchOrder As XlSearchOrder = xlByColumns, Optional SearchDirection As XlSearchDirection = xlNext, Optional MatchCase As Boolean = False, Optional MatchByte As Boolean = False, Optional SearchFormat As Boolean = False) As Range
    Dim SearchResult As Range
    Dim firstMatch As String
    With rng
        Set SearchResult = .Find(what, , LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
        If Not SearchResult Is Nothing Then
            firstMatch = SearchResult.Address
            Do
                If find_all Is Nothing Then
                    Set find_all = SearchResult
                Else
                    Set find_all = Union(find_all, SearchResult)
                End If
                Set SearchResult = .FindNext(SearchResult)
            Loop While Not SearchResult Is Nothing And SearchResult.Address <> firstMatch
        End If
    End With
End Function

'https://stackoverflow.com/questions/54823466/check-if-sheet-exists-if-not-create-vba
Function sheet_exists(str As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim ws As Worksheet
    'find
    On Error Resume Next
    Set ws = wb.Sheets(str)
    On Error GoTo 0
    'result
    If Not ws Is Nothing Then sheet_exists = True
End Function

Function table_exists(sht_str As String, tbl_str As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    'find
    On Error Resume Next
    Set tbl = wb.Sheets(sht_str).ListObjects(tbl_str)
    On Error GoTo 0
    'result
    If Not tbl Is Nothing Then table_exists = True
End Function

Sub error_handler(m_sub As String, error_vba As String)
    Debug.Print "Called Error Handler from: " & m_sub
    'print sql error from output param
    Debug.Print error_vba
    'clear erros
    p_error = vbNullString
    Debug.Print ""
    Debug.Print ""
    End
End Sub

Function p(thing As Variant)
    Dim dbug As Boolean: dbug = True
    'print
    If dbug Then
        Debug.Print thing
    End If
End Function


