Option Explicit

 
Sub ExportByDepartment()
    Dim tb As Workbook
    Dim th As Worksheet, sh As Worksheet, nh As Worksheet
    Dim lastRow As Long
    Dim export_basis As Variant
    Dim departmentList As Object
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim i As Long, j As Integer
    Dim lastCol As Integer
    Dim found As Boolean
    Dim num_of_head As Integer
    Dim export_column
    
    On Error GoTo the_end
    num_of_head = CInt(InputBox("Please enter the number of lines for the title section (digits only): ", "Output line count", 1))
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    Set departmentList = CreateObject("Scripting.Dictionary")
    
    Set tb = ThisWorkbook
    Set th = tb.Sheets("Export Operation Sheet")

    'Set the column name which the user want to export by
    export_column = th.Cells(1, 3)
    
    'Loop sheets, find columns to export of each sheets
    For Each sh In tb.Sheets
        If sh.Name <> th.Name Then
            'Find the column which user want to export by
            For j = 1 To sh.Cells(num_of_head, Columns.Count).End(xlToLeft).Column
                If sh.Cells(num_of_head, j) = export_column Then Exit For
            Next j
            
            'Sort form
            On Error Resume Next
            sh.AutoFilter.Sort.SortFields.Clear
            sh.AutoFilter.Sort.SortFields.Add2 Key:=Range(sh.Cells(num_of_head, j), sh.Cells(num_of_head, j)), SortOn:=xlSortOnValues, _
                Order:=xlAscending, DataOption:=xlSortNormal
            With sh.AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            On Error GoTo 0
            
            'Loop through each export_basis and add to dictionary
            For i = 2 To sh.Cells(Rows.Count, j).End(xlUp).Row
                export_basis = sh.Cells(i, j).Value
                'Clear and add new items into the dictionary
                If Not departmentList.Exists(export_basis) Then
                    departmentList.Add export_basis, i
                End If
            Next i
        End If
    Next sh
            
     'Loop through each export_basis in dictionary and export to new workbook
    For Each export_basis In departmentList.keys
        Set newWorkbook = Workbooks.Add

        For Each sh In tb.Sheets
            If sh.Name <> th.Name Then
                found = False
            
                'Find the column which user want to export by
                For j = 1 To sh.Cells(num_of_head, Columns.Count).End(xlToLeft).Column
                    If sh.Cells(num_of_head, j) = export_column Then Exit For
                Next j
            
                lastRow = sh.Cells(Rows.Count, j).End(xlUp).Row
                lastCol = sh.Cells(num_of_head, Columns.Count).End(xlToLeft).Column
                
                'Add new sheet and rename, then delete the redundant sheet(Sheet1)
                Set newWorksheet = newWorkbook.Sheets.Add(After:=newWorkbook.Sheets(newWorkbook.Sheets.Count))
                newWorksheet.Name = sh.Name
                
                On Error Resume Next
                Application.DisplayAlerts = False
                newWorkbook.Sheets("Sheet1").Delete
                Application.DisplayAlerts = True
                On Error GoTo 0
                
                'Copy header
                sh.Rows(1 & ":" & num_of_head).Copy newWorksheet.Rows(1 & ":" & num_of_head)
                
                'Copy rows for current export_basis
                Dim departmentFirstRow As Long, departmentLastRow As Long
                'Find first export_basis which name is "export_basis"
                For departmentFirstRow = num_of_head + 1 To lastRow
                    If sh.Cells(departmentFirstRow, j) = export_basis Then
                        found = True
                        Exit For
                    End If
                Next departmentFirstRow
                'Find last export_basis which name is "export_basis"
                For departmentLastRow = lastRow To departmentFirstRow Step -1
                    If sh.Cells(departmentLastRow, j) = export_basis Then
                        found = True
                        Exit For
                    End If
                Next departmentLastRow
                    
                If found = True Then
                    sh.Range(sh.Cells(departmentFirstRow, 1), sh.Cells(departmentLastRow, lastCol)).Copy newWorksheet.Range("A" & num_of_head + 1 & "")
                    newWorksheet.Cells.EntireColumn.AutoFit
                Else
                    newWorksheet.Range("A1") = "Export basis Not Found"
                End If
                
                'If there is any sheet that doesn't contain the export basis, then the sheet will not be contained in new workbook
                For Each nh In newWorkbook.Sheets
                    If nh.Range("A1") = "Export basis Not Found" Then
                        If newWorkbook.Sheets.Count > 1 Then
                            Application.DisplayAlerts = False
                            nh.Delete
                            Application.DisplayAlerts = True
                        Else
                            Application.DisplayAlerts = False
                            newWorkbook.Close
                            Application.DisplayAlerts = True
                            GoTo Next_loop
                        End If
                    End If
                Next nh
                
            End If
        Next sh
        '保存新的workbook，并命名为部门名称——Save new workbook with export_basis name
        newWorkbook.SaveAs tb.Path & "\" & export_basis & ".xlsx"
        newWorkbook.Close
        
Next_loop:
    Next export_basis
    
goon:
    Application.ScreenUpdating = True
    
    MsgBox "Export complete."
    
the_end:

End Sub

