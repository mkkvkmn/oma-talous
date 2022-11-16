Option Explicit
'these helpers depend on module help

Function get_class(account As String, desc As String, info As String, res As String) As Variant
'global variables
Call reset_globals
'variables
Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("luokat")
Dim tbl As ListObject: Set tbl = ws.ListObjects("luokat")
Dim account_cln_rng As Range: Set account_cln_rng = tbl.ListColumns("tili").Range
Dim find_rng As Range
Dim found_rng As Range
Dim res_rng As Range
'log
p ("Find -> account: " & q & account & q & ", description: " & q & desc & q & ", info: " & q & info & q & " from worksheet " & ws.Name & account_cln_rng.Address & " and return " & res)
'account
Set find_rng = account_cln_rng
Set found_rng = Nothing
Set found_rng = find_all(find_rng, account)
If Not found_rng Is Nothing Then
    'description
    Set find_rng = found_rng.Offset(0, 1)
    p ("account " & q & account & q & " found, find desc " & q & desc & q & " from " & find_rng.Address)
    Set found_rng = Nothing
    Set found_rng = find_all(find_rng, desc)
    If Not found_rng Is Nothing Then
    '    'info
        Set find_rng = found_rng.Offset(0, 1)
        p ("description: " & q & desc & q & " found, find info " & q & info & q & " from " & find_rng.Address)
        Set found_rng = Nothing
        Set found_rng = find_all(find_rng, info)
        If Not found_rng Is Nothing Then
            p ("info: " & q & info & q & " found")
             If Not found_rng Is Nothing Then
                'result
                Set res_rng = found_rng.Offset(0, 1).Resize(1, 3)
                'p (res_rng.Address)
            End If
        End If
    End If
End If

'result
get_class = "n/a"
If Not res_rng Is Nothing Then
    If res = "class" Then
        get_class = res_rng.Cells(1, 1)
    ElseIf res = "category" Then
        get_class = res_rng.Cells(1, 2)
    ElseIf res = "subcategory" Then
        get_class = res_rng.Cells(1, 3)
    End If
End If

'p ("get_class result: " & get_class)

End Function

Function new_sheet(str As String, delete_current As Boolean, Optional hide As Boolean) As Worksheet
    'variables
    Dim tbl As ListObject
    Dim last_col As String
    'get current table
    If sheet_exists(str) Then
        'delete current table
        If delete_current Then
            p "Delete current sheet"
            'delete
            Application.DisplayAlerts = False
            On Error Resume Next
            Sheets(str).Delete
            Application.DisplayAlerts = True
            'create sheet
            Sheets.Add.Name = str
            'create table
            p ("Replaced sheet " & str)
        End If
    'create new table
    Else
        'create sheet
        Sheets.Add.Name = str
        'create table
        p ("Created new sheet " & str)
    End If
    'hide
    If hide Then
        Worksheets(str).Visible = False
    End If
    'result
    Set new_sheet = Worksheets(str)
End Function

Function new_table(sht As String, tbl As String, col_first As String, cols As Long, delete_current As Boolean, Optional hide As Boolean) As ListObject
    'variables
    Dim tbl2 As ListObject
    Dim col_last As String: col_last = get_end_column(col_first, cols)
    Dim tbl_rng As Range
    Dim new_sht As Worksheet
   
    'get sheet
    If sheet_exists(sht) Then
        p ("Set new table rng to " & col_first & "$1:$" & col_last & "$2")
        Set tbl_rng = Worksheets(sht).Range(col_first & "$1:$" & col_last & "$2")
        If table_exists(sht, tbl) Then
            If delete_current Then
                'reset
                Worksheets(sht).ListObjects(tbl).Delete
                Set tbl2 = Worksheets(sht).ListObjects.Add(xlSrcRange, tbl_rng, , xlYes)
                p ("Replaced table " & tbl)
            Else
                Set tbl2 = Worksheets(sht).ListObjects(tbl)
                p ("Use current table " & tbl)
            End If
        Else
            Set tbl2 = Worksheets(sht).ListObjects.Add(xlSrcRange, tbl_rng, , xlYes)
            p ("Created table " & tbl)
        End If
    Else
        Set new_sht = Sheets.Add
        With new_sht
            .Name = sht
            .Move After:=Worksheets(Worksheets.count)
        End With
        Set tbl_rng = Worksheets(sht).Range(col_first & "$1:$" & col_last & "$2")
        Set tbl2 = Worksheets(sht).ListObjects.Add(xlSrcRange, tbl_rng, , xlYes)
        p ("Created new sheet " & sht & " and new table " & tbl)
    End If
    'name table
    With tbl2
        .Name = tbl
    End With
     'hide
    If hide Then
        Worksheets(sht).Visible = False
    End If
    'result
    Set new_table = tbl2
End Function
