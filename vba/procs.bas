'Option Explicit
'global variables
Public i As Integer
Public q As String
Public p_error As String
Public ind As String
'account
Public account_id As Integer
Public account_rng As Range
Public account_str As String
'description
Public desc_id As Integer
Public desc_rng As Range
Public desc_str As String
'info
Public info_id As Integer
Public info_rng As Range
Public info_str As String
'class
Public class_id As Integer
Public class_rng As Range
Public class_str As String
'category
Public category_id As Integer
Public category_rng As Range
Public category_str As String
'sub-category
Public subcategory_id As Integer
Public subcategory_rng As Range
Public subcategory_str As String
'id
Public id_id As Integer
Public id_rng As Range
Public id_str As String

Sub all()
    Application.ScreenUpdating = False
    Call get_data
    Call get_uniq
    Call update_categories
    Call apply_rules
    Call apply_categories
    Call notify
    Application.ScreenUpdating = True
End Sub

Sub class()
    Application.ScreenUpdating = False
    Call apply_rules
    Call apply_categories
    Call notify
    Application.ScreenUpdating = True
End Sub

Sub reset_globals()
    i = 0
    q = """"
    p_error = ""
    ind = "    "
    'account
    account_id = 1
    Set account_rng = Nothing
    account_str = "tili"
    'description
    desc_id = 2
    Set desc_rng = Nothing
    desc_str = "selite"
    'info
    info_id = 3
    Set info_rng = Nothing
    info_str = "info"
    'class
    class_id = 4
    Set class_rng = Nothing
    class_str = "luokka"
    'category
    category_id = 5
    Set category_rng = Nothing
    category_str = "kategoria"
    'sub-category
    subcategory_id = 6
    Set subcategory_rng = Nothing
    subcategory_str = "ala-kategoria"
    'id
    id_id = 7
    Set id_rng = Nothing
    id_str = "id"
End Sub

Sub notify()
On Error Resume Next
'global variables
Call reset_globals
'classes
Dim class_sht_str As String: class_sht_str = "luokat"
Dim class_tbl_str As String: class_tbl_str = "luokat"
Dim class_tbl As ListObject: Set class_tbl = ThisWorkbook.Worksheets(class_sht_str).ListObjects(class_tbl_str)
Dim class_rng As Range: Set class_rng = class_tbl.ListColumns(class_str).DataBodyRange.SpecialCells(xlCellTypeBlanks)
'data
Dim data_sht_str As String: data_sht_str = "data"
Dim data_tbl_str As String: data_tbl_str = "data"
Dim data_sheet As Worksheet
Dim data_tbl As ListObject: Set data_tbl = ThisWorkbook.Worksheets(data_sht_str).ListObjects(data_tbl_str)
Dim data_rng As Range: Set data_rng = data_tbl.ListColumns(class_str).DataBodyRange.SpecialCells(xlCellTypeBlanks)
'other
Dim msg As String
Dim msg2 As String: msg2 = vbNewLine & "Siirry luokittelemaan?"
Dim msg_num As Long

'set message
If Not class_rng Is Nothing And Not data_rng Is Nothing Then
    msg = class_rng.Rows.count & " kpl luokiteltavia rivejä taulukossa luokka" & vbNewLine & " ( data taulussa luokittelemattomia rivejä " & data_rng.Rows.count & " kpl)"
    msg_num = MsgBox(msg & msg2, 4)
ElseIf class_rng Is Nothing And Not data_rng Is Nothing Then
    msg = 0 & " kpl luokiteltavia rivejä taulukossa luokka" & vbNewLine & " ( data taulussa luokittelemattomia rivejä " & data_rng.Rows.count & " kpl)"
    msg_num = MsgBox(msg & msg2, 4)
ElseIf data_rng Is Nothing And Not class_rng Is Nothing Then
    msg = class_rng.Rows.count & " kpl luokiteltavia rivejä taulukossa luokka" & vbNewLine & " ( data taulussa luokittelemattomia rivejä " & 0 & " kpl)"
    msg_num = MsgBox(msg & msg2, 4)
Else
    msg = "Kaikki rivit luokiteltu"
    msg_num = MsgBox(msg, 0)
End If
'result
p ("msg num " & msg_num & " msg: " & msg)
If msg_num = 6 Then ' 6 = yes
    Worksheets("luokat").Activate
Else
    Worksheets("data").Activate
End If

End Sub

Sub apply_categories()
On Error GoTo err_handler
    'global variables
    Call reset_globals
    'variables
    Dim sub_name As String: sub_name = "apply_categories"
    'target
    Dim tgt_sht_str As String: tgt_sht_str = "data"
    Dim tgt_tbl_str As String: tgt_tbl_str = "data"
    Dim tgt_sheet As Worksheet
    Dim tgt_tbl As ListObject: Set tgt_tbl = ThisWorkbook.Worksheets(tgt_sht_str).ListObjects(tgt_tbl_str)
    Dim tgt_rng As Range
    Dim tgt_next_row As Range
    'source
    Dim src_sht_str As String: src_sht_str = "luokat"
    Dim src_tbl_str As String: src_tbl_str = "luokat"
    Dim src_tbl As ListObject: Set src_tbl = ThisWorkbook.Worksheets(src_sht_str).ListObjects(src_tbl_str)
    Dim src_rng As Range
    Dim src_rng_cell As Range
    Dim src_rng_row As Range
    Dim src_rows As Long
    'logic
    Dim missing_rng As Range
    Dim zero_rng As Range
    Dim lu_col_str As String: lu_col_str = "haku"
    Dim lu_cols_str As String: lu_cols_str = "[@" & account_str & "]" & "&" & "[@" & desc_str & "]" & "&" & "[@" & info_str & "]"
    Dim lu_src_str As String: lu_src_str = "luokat" & "[#All]"
    'order data
    Call sort_blanks(tgt_tbl, category_str)
    'add lookup column to src table
    For i = 1 To src_tbl.ListColumns.count
        If src_tbl.ListColumns(i) = lu_col_str Then
           src_tbl.ListColumns(lu_col_str).Delete
        End If
    Next i
    With src_tbl
        .ListColumns.Add(1).Name = lu_col_str
        .ListColumns(lu_col_str).DataBodyRange.Formula = "=" & lu_cols_str
    End With
    'find missing classes
    tgt_tbl.ListColumns(class_str).DataBodyRange.ClearContents
    Set missing_rng = tgt_tbl.ListColumns(class_str).DataBodyRange.SpecialCells(xlCellTypeBlanks)
    'insert classes
    If Not missing_rng Is Nothing Then
        p ("Insert classes to missing range: " & missing_rng.Address)
        With missing_rng
            .Formula = "=IFERROR(VLOOKUP(" & lu_cols_str & "," & lu_src_str & "," & class_id + 1 & ", 0),"""")"
            .Offset(, 1).Formula = "=IFERROR(VLOOKUP(" & lu_cols_str & "," & lu_src_str & "," & category_id + 1 & ", 0),"""")"
            .Offset(, 2).Formula = "=IFERROR(VLOOKUP(" & lu_cols_str & "," & lu_src_str & "," & subcategory_id + 1 & ", 0),"""")"
            .Copy: .PasteSpecial Paste:=xlPasteValues
            .Offset(, 1).Copy: .Offset(, 1).PasteSpecial Paste:=xlPasteValues
            .Offset(, 2).Copy: .Offset(, 2).PasteSpecial Paste:=xlPasteValues
        End With
        Set zero_rng = find_all(tgt_tbl.DataBodyRange, 0)
        If Not zero_rng Is Nothing Then
            zero_rng.Value = ""
        End If
    End If
    'remove lookup column from src table
    src_tbl.ListColumns(lu_col_str).Delete
    'order data
    Call sort_blanks(tgt_tbl, category_str, True)
    'finish
    p ("Complete " & sub_name)
    p ("------------------ ")
Exit Sub
err_handler:
    p_error = "Error number: " & str(Err.Number) & ", Source: " & Err.Source & ", Description: " & Err.Description
    Call error_handler(sub_name, p_error)
End Sub

Sub apply_rules()
Application.ScreenUpdating = False
On Error GoTo err_handler
    'global variables
    Call reset_globals
    'variables
    Dim sub_name As String: sub_name = "apply_rules"
    p ("Start " & sub_name)
    Dim rng As Range
    Dim sort_rng As Range
    'target
    Dim tgt_sht_str As String: tgt_sht_str = "luokat"
    Dim tgt_tbl_str As String: tgt_tbl_str = "luokat"
    Dim tgt_sheet As Worksheet
    Dim tgt_tbl As ListObject
    Dim tgt_rng As Range
    Dim tgt_next_row As Range
    'source
    Dim src_sht_str As String: src_sht_str = "luokat"
    Dim src_tbl_str As String: src_tbl_str = "säännöt"
    Dim src_tbl As ListObject
    Dim src_rng As Range
    Dim src_rng_cell As Range
    Dim src_rng_row As Range
    Dim src_rows As Long
    
    Dim found_account_rng As Range
    Dim found_desc_rng As Range
    Dim found_info_rng As Range
    
    Dim cols_before_insert As Long: cols_before_insert = 3
    Dim insert_class_rng As Range
    Dim insert_category_rng As Range
    Dim insert_subcategory_rng As Range
   
    Dim look_desc_rng As Range
    Dim look_info_rng As Range
    
    'find
    Dim find_account As String
    Dim find_desc As String
    Dim find_info As String
    Dim result_class As String
    Dim result_category As String
    Dim result_subcategory As String
   
    'set source
    Set src_tbl = new_table(src_sht_str, src_tbl_str, "A", 7, False)
    With src_tbl
        .ListColumns(account_id).Name = account_str
        .ListColumns(desc_id).Name = desc_str
        .ListColumns(info_id).Name = info_str
        .ListColumns(class_id).Name = class_str
        .ListColumns(category_id).Name = category_str
        .ListColumns(subcategory_id).Name = subcategory_str
        .ListColumns(id_id).Name = id_str
    End With
    Set src_rng = src_tbl.DataBodyRange
    src_rows = src_rng.Rows.count
    'set destination
    Set tgt_sheet = ThisWorkbook.Sheets(tgt_sht_str)
    Set tgt_tbl = tgt_sheet.ListObjects(tgt_tbl_str)
    
    'find works only if blanks are sorted naturally
    Call sort_blanks(tgt_tbl, category_str, False)
    
    'apply rules
    For Each src_rng_cell In src_rng.Resize(src_rows, 1)
        'set source row from range
        Set src_rng_row = src_rng_cell.Resize(1, 6)
        'find destination accounts rng
        Set tgt_rng = tgt_tbl.ListColumns(account_id).Range
        'set values to find
        find_account = src_rng_row(account_id)
        find_desc = src_rng_row(desc_id)
        find_info = src_rng_row(info_id)
        'set results from find table
        result_class = src_rng_row(class_id)
        result_category = src_rng_row(category_id)
        result_subcategory = src_rng_row(subcategory_id)
        p ("Look for account " & q & find_account & q & " ,description " & q & find_desc & q & " ,info " & q & find_info & q & " from " & tgt_tbl.Name & tgt_rng.Address & "->")
        'find accounts
        Set found_account_rng = find_all(tgt_rng, find_account)
        If Not found_account_rng Is Nothing Then
            p ("1. " & q & find_account & q & " found from " & found_account_rng.Address)
            'find descriptions
            Set look_desc_rng = Range(found_account_rng.Address).Offset(0, desc_id - account_id)
            Set found_desc_rng = find_all(tgt_sheet.Range(look_desc_rng.Address), find_desc)
            If Not found_desc_rng Is Nothing Then
                p ("2. " & q & find_desc & q & " found from " & found_desc_rng.Address)
                'info
                Set look_info_rng = Range(found_desc_rng.Address).Offset(0, info_id - desc_id)
                Set found_info_rng = find_all(tgt_sheet.Range(look_info_rng.Address), find_info)
                If Not found_info_rng Is Nothing Then
                    p ("3. " & q & find_info & q & " found from " & found_info_rng.Address)
                    'insert
                    Set class_rng = found_info_rng.Offset(0, class_id - info_id)
                    Set category_rng = found_info_rng.Offset(0, category_id - info_id)
                    Set subcategory_rng = found_info_rng.Offset(0, subcategory_id - info_id)
                    p ("-> insert class " & q & result_class & q & " to " & class_rng.Address & ", category " & q & result_category & q & " to " & category_rng.Address & " and sub category " & q & result_subcategory & q & " to " & subcategory_rng.Address)
                    class_rng.Value = result_class
                    category_rng.Value = result_category
                    subcategory_rng.Value = result_subcategory
                Else
                    p ("No action: info " & q & find_info & q & " not found from " & look_info_rng.Address)
                End If
            Else
                p ("No action: description " & q & find_desc & q & " not found from " & look_desc_rng.Address)
            End If
        Else
            p ("No action: account " & q & find_account & q & " not found from " & tgt_rng.Address)
        End If
    Next src_rng_cell
    'update rules id
    For Each rng In src_tbl.ListColumns(id_id).DataBodyRange
        rng.Value = rng.Row
    Next rng
    'order missing categories to top
    Call sort_blanks(tgt_tbl, category_str, True)
    'finish
    p ("Complete " & sub_name)
    p ("------------------ ")
Application.ScreenUpdating = True
Exit Sub
err_handler:
    p_error = "Error number: " & str(Err.Number) & ", Source: " & Err.Source & ", Description: " & Err.Description
    Call error_handler(sub_name, p_error)
End Sub

Sub update_categories()
On Error GoTo err_handler
    'global variables
    Call reset_globals
    'variables
    Dim sub_name As String: sub_name = "reset_categories"
    p ("Start " & sub_name)
    'target
    Dim tgt_tbl_str As String: tgt_tbl_str = "luokat"
    Dim tgt_sheet As Worksheet
    Dim tgt_tbl As ListObject
    Dim tgt_rng As Range
    Dim tgt_next_row As Range
    'source
    Dim src_tbl_str As String: src_tbl_str = "temp"
    Dim src_rng As Range
    Dim src_rng_cell As Range
    Dim src_rng_row As Range
    Dim src_rows As Long
    'find strings
    Dim find_desc As String
    Dim find_info As String
    Dim find_account As String
    'other
    Dim account_name As String
    Dim not_found As Boolean: not_found = False
    Dim find_rng As Range
    Dim found_rng As Range
    
    'set
    Set src_rng = ThisWorkbook.Sheets(src_tbl_str).ListObjects(src_tbl_str).DataBodyRange
    src_rows = src_rng.Rows.count
    Set tgt_tbl = new_table(tgt_tbl_str, tgt_tbl_str, "K", 7, False)
    With tgt_tbl
        .ListColumns(account_id).Name = account_str
        .ListColumns(desc_id).Name = desc_str
        .ListColumns(info_id).Name = info_str
        .ListColumns(class_id).Name = class_str
        .ListColumns(category_id).Name = category_str
        .ListColumns(subcategory_id).Name = subcategory_str
        .ListColumns(id_id).Name = id_str
    End With
    Set tgt_sheet = ThisWorkbook.Sheets(tgt_tbl_str)
    
    'look for every temp table row from category table
    For Each src_rng_cell In src_rng.Resize(src_rows, 1)
        'set source row from range
        Set src_rng_row = src_rng_cell.Resize(1, 3)
        'find destination next empty row
        Set tgt_rng = tgt_tbl.ListColumns(account_id).Range
        Set tgt_next_row = find_table_next_row_rng(tgt_tbl).Resize(1, 3)
        'set values to find
        find_account = src_rng_row(account_id)
        find_desc = src_rng_row(desc_id)
        find_info = src_rng_row(info_id)
        p ("Look for account " & q & find_account & q & " ,description " & q & find_desc & q & " ,info " & q & find_info & q & " ->")
        'account
        Set find_rng = tgt_rng
        Set found_rng = Nothing
        Set found_rng = find_all(find_rng, find_account)
        If Not found_rng Is Nothing Then
            'description
            Set find_rng = found_rng.Offset(0, 1)
            'p ("account " & q & account & q & " found, find desc " & q & desc & q & " from " & find_rng.Address)
            Set found_rng = Nothing
            Set found_rng = find_all(find_rng, find_desc)
            If Not found_rng Is Nothing Then
                'info
                Set find_rng = found_rng.Offset(0, 1)
                'p ("description: " & q & desc & q & " found, find info " & q & info & q & " from " & find_rng.Address)
                Set found_rng = Nothing
                Set found_rng = find_all(find_rng, find_info)
                If found_rng Is Nothing Then
                    p (q & find_info & q & " not found")
                    not_found = True
                End If
            Else
                p (q & find_desc & q & " not found")
                not_found = True
            End If
        Else
            p (q & find_account & q & " not found")
            not_found = True
        End If
        'insert if not exists
        If not_found Then
            p ("-> insert account " & q & find_account & q & " ,description " & q & find_desc & q & " ,info " & q & find_info & q & " to " & tgt_next_row.Address)
            With tgt_next_row
                .Cells(1, account_id) = find_account
                .Cells(1, desc_id) = find_desc
                .Cells(1, info_id) = find_info
                .Cells(1, id_id) = .Cells(1, id_id).Row
            End With
        End If
    Next src_rng_cell
    p ("Complete " & sub_name)
    p ("------------------ ")
Exit Sub
err_handler:
    p_error = "Error number: " & str(Err.Number) & ", Source: " & Err.Source & ", Description: " & Err.Description
    Call error_handler(sub_name, p_error)
End Sub

'Collect unique account,description and info combos from data
Sub get_uniq()
On Error GoTo err_handler
    'global variables
    Call reset_globals
    'general variables
    Dim sub_name As String: sub_name = "get_uniq"
    p ("Start " & sub_name)
    'source
    Dim src_ws As Worksheet: Set src_ws = ThisWorkbook.Worksheets("data")
    Dim src_tbl As ListObject: Set src_tbl = src_ws.ListObjects("data")
    Dim src_account_rng As Range
    Dim src_desc_rng As Range
    Dim src_info_rng As Range
    Dim src_records_qty As Long
    
    'data tbl
    Dim tgt_tbl_str As String: tgt_tbl_str = "temp"
    Dim tgt_tbl As ListObject
    Dim tgt_tbl_colums_int As Integer
    Dim tgt_tbl_cols As Variant
    
   'reset temp table
    Set tgt_tbl = new_table(tgt_tbl_str, tgt_tbl_str, "A", 3, True, False)
    'set columns
    With tgt_tbl
        .ListColumns(account_id).Name = account_str
        .ListColumns(desc_id).Name = desc_str
        .ListColumns(info_id).Name = info_str
    End With
    
    'collect uniques from data
    p ("Collect data from sheet: " & src_ws.Name & " table: " & src_tbl.Name)
    'get data from source table
    Set src_account_rng = src_tbl.ListColumns(account_str).DataBodyRange
    Set src_desc_rng = src_tbl.ListColumns(desc_str).DataBodyRange
    Set src_info_rng = src_tbl.ListColumns(info_str).DataBodyRange
    'log records
    src_records_qty = src_desc_rng.Rows.count
    p ("Data records: " & src_records_qty & " in sheet: " & src_ws.Name & " table " & src_tbl.Name)
    'get next empty rows first cell for destination tbl and set target ranges
    Set account_rng = find_table_next_row_rng(tgt_tbl).Resize(src_records_qty, 1)
    Set desc_rng = account_rng.Offset(0, desc_id - 1)
    Set info_rng = account_rng.Offset(0, info_id - 1)
    'copy
    src_account_rng.Copy: account_rng.PasteSpecial Paste:=xlPasteValues
    src_desc_rng.Copy: desc_rng.PasteSpecial Paste:=xlPasteValues
    src_info_rng.Copy: info_rng.PasteSpecial Paste:=xlPasteValues
    'remove duplicates
    tgt_tbl_colums_int = tgt_tbl.ListColumns.count
    ReDim tgt_tbl_cols(0 To tgt_tbl_colums_int - 1)
    For i = 0 To UBound(tgt_tbl_cols)
        tgt_tbl_cols(i) = i + 1
    Next i
    p ("Remove duplicates for " & tgt_tbl_colums_int & " columns")
    tgt_tbl.DataBodyRange.RemoveDuplicates Columns:=(tgt_tbl_cols), Header:=xlNo
    p ("Complete " & sub_name)
    p ("------------------ ")
Exit Sub
err_handler:
    p_error = "Error number: " & str(Err.Number) & ", Source: " & Err.Source & ", Description: " & Err.Description
    Call error_handler(sub_name, p_error)
End Sub

Sub get_data()
On Error GoTo err_handler
    'global variables
    Call reset_globals
    'general variables
    Dim sub_name As String: sub_name = "get_data"
    p ("Start " & sub_name)
    'source
    Dim src_tbl_str As String: src_tbl_str = "d_"
    Dim src_ws As Worksheet
    Dim src_tbl As ListObject
    Dim src_date_rng As Range
    Dim src_desc_rng As Range
    Dim src_info_rng As Range
    Dim src_amount_rng As Range
    Dim src_records_qty As Long
    Dim account_name As String
    'data tbl
    Dim tgt_tbl_name As String: tgt_tbl_name = "data"
    Dim tgt_tbl As ListObject
    Dim tgt_tbl_colums_int As Integer
    'backup table
    Dim bu_tbl As ListObject
    Dim bu_sht_str As String
    Dim bu_tbl_str As String
    Dim bu_tbl_next_row As Range
    'columns
    'category override
    Dim category_m_id As Integer: category_m_id = 8
    Dim category_m_rng As Range: Set category_m_rng = Nothing
    Dim category_m_str As String: category_m_str = "kategoria m"
    'sub category override
    Dim subcategory_m_id As Integer: subcategory_m_id = 9
    Dim subcategory_m_rng As Range: Set subcategory_m_rng = Nothing
    Dim subcategory_m_str As String: subcategory_m_str = "ala-kategoria m"
    'date
    Dim date_id As Integer: date_id = 10
    Dim date_rng As Range: Set date_rng = Nothing
    Dim date_str As String: date_str = "pvm"
    'amount
    Dim amount_id As Integer: amount_id = 11
    Dim amount_rng As Range: Set amount_rng = Nothing
    Dim amount_str As String: amount_str = "eur"
    'share
    Dim share_id As Integer: share_id = 12
    Dim share_rng As Range: Set share_rng = Nothing
    Dim share_str As String: share_str = "osuus"
    'logic
    Dim has_data As Boolean: has_data = False
       
    'create data table
    Set tgt_tbl = new_table(tgt_tbl_name, tgt_tbl_name, "A", 12, False) 'Be careful! set to true if you want to reset data table
    'name columns
    With tgt_tbl
        .ListColumns(account_id).Name = account_str
        .ListColumns(desc_id).Name = desc_str
        .ListColumns(info_id).Name = info_str
        .ListColumns(class_id).Name = class_str
        .ListColumns(category_id).Name = category_str
        .ListColumns(subcategory_id).Name = subcategory_str
        .ListColumns(id_id).Name = id_str
        .ListColumns(category_m_id).Name = category_m_str
        .ListColumns(subcategory_m_id).Name = subcategory_m_str
        .ListColumns(id_id).Name = id_str
        .ListColumns(date_id).Name = date_str
        .ListColumns(amount_id).Name = amount_str
        .ListColumns(share_id).Name = share_str
    End With
    
    'collect descriptions
    For Each src_ws In ThisWorkbook.Sheets
        'get tables
        For Each src_tbl In src_ws.ListObjects
            p ("Look sheet " & q & src_ws.Name & q & " and table " & q & src_tbl.Name & q & " for data " & q & src_tbl_str & q & " ->")
            If (src_tbl.Range.Rows.count) > 3 Or (Len(src_tbl.Range.Cells(2, 1)) > 1 And Len(src_tbl.Range.Cells(2, 2)) > 1) Then
                p (ind & "Data found, rows: " & src_tbl.Range.Rows.count & " len cell 2,1: " & Len(src_tbl.Range.Cells(2, 1)) & ", len cell 2,2: " & Len(src_tbl.Range.Cells(2, 2)))
                If InStr(src_tbl.Name, src_tbl_str) > 0 Then
                    p (ind & "Data string found from " & src_ws.Name & " table: " & src_tbl.Name & "")
                    If Not src_tbl.HeaderRowRange.Find(date_str) Is Nothing And Not src_tbl.HeaderRowRange.Find(desc_str) Is Nothing And Not src_tbl.HeaderRowRange.Find(amount_str) Is Nothing Then
                        account_name = src_ws.Name
                        'get data from source table
                        Set src_date_rng = src_tbl.ListColumns(date_str).DataBodyRange
                        Set src_desc_rng = src_tbl.ListColumns(desc_str).DataBodyRange
                        Set src_amount_rng = src_tbl.ListColumns(amount_str).DataBodyRange
                        For i = 1 To src_tbl.ListColumns.count
                            If src_tbl.ListColumns(i) = info_str Then
                                'p ("Info found for account: " & account_str & " in table " & src_tbl.name)
                                Set src_info_rng = src_tbl.ListColumns(info_str).DataBodyRange
                            End If
                        Next i
                        
                        'log records
                        src_records_qty = src_desc_rng.Rows.count
                        p (ind & "Data records: " & src_records_qty & " in sheet: " & q & src_ws.Name & q & " table " & q & src_tbl.Name & q & ", start copy")
                        'copy to next empty row
                        Set account_rng = find_table_next_row_rng(tgt_tbl).Resize(src_records_qty, 1)
                        Set date_rng = account_rng.Offset(0, date_id - account_id)
                        Set desc_rng = account_rng.Offset(0, desc_id - account_id)
                        Set info_rng = account_rng.Offset(0, info_id - account_id)
                        Set amount_rng = account_rng.Offset(0, amount_id - account_id)
                        p ( _
                            "-> " _
                            & ind & "date from " & src_date_rng.Address & " to " & date_rng.Address & vbNewLine _
                            & ind & ind & "account from value to " & account_rng.Address & vbNewLine _
                            & ind & ind & "description from " & src_desc_rng.Address & " to " & desc_rng.Address & vbNewLine _
                            & ind & ind & "amount from " & src_amount_rng.Address & " to " & amount_rng.Address _
                        )
                        src_date_rng.Copy date_rng
                        account_rng.Value = account_name
                        src_desc_rng.Copy: desc_rng.PasteSpecial Paste:=xlPasteValues
                        If Not src_info_rng Is Nothing Then
                           src_info_rng.Copy: info_rng.PasteSpecial Paste:=xlPasteValues
                           p (ind & "info from " & src_info_rng.Address & " to " & info_rng.Address)
                        End If
                        src_amount_rng.Copy: amount_rng.PasteSpecial Paste:=xlPasteValues
                        'save history
                        p (ind & "Save history")
                        bu_sht_str = "h-" & src_ws.Name
                        bu_tbl_str = Replace(src_tbl.Name, src_tbl_str, "h_")
                        Set bu_tbl = new_table(bu_sht_str, bu_tbl_str, "A", src_tbl.ListColumns.count, False, True)
                        Set bu_tbl_next_row = find_table_next_row_rng(bu_tbl)
                        src_tbl.DataBodyRange.Copy: bu_tbl_next_row.PasteSpecial Paste:=xlPasteValues
                        p (ind & "Saved history for " & q & src_ws.Name & " " & src_tbl.Name & q & " to " & q & bu_sht_str & " " & bu_tbl_str & q)
                        'clear insert sheet
                        src_tbl.DataBodyRange.Delete
                    Else
                        p ("No action: one or more columns missing (" & date_str & "," & desc_str & "," & info_str & ")")
                    End If
                Else
                    p ("No action: table name does not contain " & q & src_tbl_str & q)
                End If
            Else
                p ("No action: no data")
            End If
        Next src_tbl
        'reset for next sheet
        Set src_info_rng = Nothing
    Next src_ws
    p ("Complete " & sub_name)
    p ("------------------ ")
Exit Sub
err_handler:
    p_error = "Error number: " & str(Err.Number) & ", Source: " & Err.Source & ", Description: " & Err.Description
    Call error_handler(sub_name, p_error)
End Sub


