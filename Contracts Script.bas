Sub ZoomUnhideUnmerge()
    
    Dim ws As Worksheet
    
    'Disable ScreenUpdating to improve performance
    Application.ScreenUpdating = False
    
    'Start looping through worksheets
    For Each ws In ThisWorkbook.Worksheets
        
        'Make worksheet active "seems required for some steps"
        ws.Activate
        
        'Change zoom level to be equal for all ws for visual uniformity
        ActiveWindow.Zoom = 40
        
        'Unhide all columns and rows
        ws.Cells.EntireColumn.Hidden = False
        ws.Cells.EntireRow.Hidden = False
        
        'Unmerge all cells
        ws.Cells.UnMerge
        
        'Return to cell A1 for visual uniformity
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ws.Range("A1").Select
        
    Next ws
    
    'Reenable ScreenUpdating
    Application.ScreenUpdating = True
    
    'Return to first sheet
    Worksheets(1).Activate

End Sub

Sub CreateTables()
    
    Dim ws As Worksheet
    Dim Cell As Range
    Dim errcnt As Long
    Dim lRow As Long
    
    'Disable ScreenUpdating to improve performance
    Application.ScreenUpdating = False
    
    'Start looping through worksheets
    For Each ws In ThisWorkbook.Worksheets
        
        'Make worksheet active "seems required for some steps"
        ws.Activate
        
        'Delete formulas and keep values
        For Each Cell In ws.UsedRange
            On Error GoTo ErrHandler
            Cell.Value = Cell.Value
        Next Cell
            
        'Delete Column A if it's an extra
        'Check if cell A3 contains a value, if it doesn't, then column A is an extra
        If ws.Range("A3").Value = "" Then

            'Check if the rest of the column contains data and log before deleting for future record
            If WorksheetFunction.CountA(ws.Range("A7:A" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)) <> 0 Then
                Debug.Print "Column A in sheet " & ws.Name & " is an extra but contains values"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "Column A in sheet " & ws.Name & " is an extra but contains values"
                Close #1
            Else
                Debug.Print "Column A in sheet " & ws.Name & " is an extra without values"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "Column A in sheet " & ws.Name & " is an extra without values"
                Close #1
            End If

            'Delete the column
            ws.Columns(1).Delete

            'Log the operation
            Debug.Print "Column A in sheet " & ws.Name & " has been deleted"
            Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
            Print #1, Now & " " & "Column A in sheet " & ws.Name & " has been deleted"
            Close #1
        End If
        
        'Move B4 value to C4 if misplaced
        'Check if B4 contains value
        If ws.Range("B4").Value <> "" Then

            'Check if C4 contains value as well, if it does then log and do nothing for future investigation
            If ws.Range("C4").Value <> "" Then
                Debug.Print "Both B4 and C4 in " & ws.Name & " have values"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "Both B4 and C4 in " & ws.Name & " have values"
                Close #1
            Else
                'Move value to correct cell and log
                ws.Range("C4").Value = ws.Range("B4").Value
                ws.Range("B4").Value = ""
                Debug.Print "B4 value in " & ws.Name & " has been moved to C4"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "B4 value in " & ws.Name & " has been moved to C4"
                Close #1
            End If
        End If
        
        'Convert main data range to a table
        ws.Range("A3:O4").Select
        Application.CutCopyMode = False
        ws.ListObjects.Add(xlSrcRange, Range("$A$3:$O$4"), , xlYes).Name = _
            "Program_" & ws.Range("C4").Value & "_MainData"
        
        'Convert project list range to table
        
        'Clean columns first
        If ws.Range("AL8").Value = "2024" Then
            ws.Columns(37).Delete
        End If
        
        If ws.Range("AK8").Value <> "2024" Then
            ws.Range("AK8").Value = "2024"
        End If
        
        lRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        'ws.Range("A7:O100").Select
        'Selecting to the latest data record instead of selecting a large table
        ws.Range("A7:AK" & lRow).Select
        Application.CutCopyMode = False
        ws.ListObjects.Add(xlSrcRange, Range("$A$7:$AK$" & lRow), , xlYes).Name = _
            "Program_" & ws.Range("C4").Value & "_Contracts"
        
        'Return to cell A1 for visual uniformity
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ws.Range("A1").Select
        
        Debug.Print "Sheet " & ws.Name & " tables have been created"
        Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
        Print #1, Now & " " & "Sheet " & ws.Name & " tables have been created"
        Close #1
    Next ws
    
    'Reenable ScreenUpdating
    Application.ScreenUpdating = True
    
    'Return to first sheet
    Worksheets(1).Activate

    'Log end of operation
    Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
    Print #1, Now & " " & "Macro Done"
    Close #1

    'Display message that the operation has been completed
    If errcnt = 1 then
        
        MsgBox "Operation Completed with " & errcnt & " Error
    
    Else If errcnt > 1 then
        
        MsgBox "Operation Completed with " & errcnt & " Errors
        
    Else
        
        MsgBox "Operation Completed Successfully"
    
    End If
    
    Exit Sub

ErrHandler:
        
        'Log errors if any
        Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
        Print #1, Now & " " & "Error in sheet " & ws.Name & " cell " & Cell.Address & " Error Code : " & Err.Number & ":" & Err.Description
        Close #1
        errcnt = errcnt + 1
        Resume Next
End Sub
