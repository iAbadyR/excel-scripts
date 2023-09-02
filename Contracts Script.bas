Attribute VB_Name = "ContractsDataCleaningPreporation"
Sub ZoomUnhideUnmerge()
    
    Dim ws As Worksheet
    
    'Disable ScreenUpdating to improve performance
    Application.ScreenUpdating = False
    
    'Start loop
    For Each ws In ThisWorkbook.Worksheets
        
        'Make worksheet active "required for some steps"
        ws.Activate
        
        'Change zoom level to be equal for all ws
        ActiveWindow.Zoom = 40
        
        'Unhide all columns and rows
        ws.Cells.EntireColumn.Hidden = False
        ws.Cells.EntireRow.Hidden = False
        
        'Unmerge all cells
        ws.Cells.UnMerge
        
        'Return to cell A1 for visual uniformaty
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
    
    'Disable ScreenUpdating to improve performance
    Application.ScreenUpdating = False
    
    'Start loop
    For Each ws In ThisWorkbook.Worksheets
        
        'Make worksheet active "required for some steps"
        ws.Activate
        
        'Delete formulas and keep values
        For Each Cell In ws.UsedRange
            On Error GoTo ErrHandler
            Cell.Value = Cell.Value
        Next Cell
            
        'Delete Column A if extra
        If ws.Range("A3").Value = "" Then
            
            If WorksheetFunction.CountA(ws.Range("A9:A100")) <> 0 Then
                Debug.Print "Column A in sheet " & ws.Name & " is extra but contains values"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "Column A in sheet " & ws.Name & " is extra but contains values"
                Close #1
            Else
                Debug.Print "Column A in sheet " & ws.Name & " is extra without values"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "Column A in sheet " & ws.Name & " is extra without values"
                Close #1
            End If
            
            ws.Columns(1).Delete
            
            Debug.Print "Column A in sheet " & ws.Name & " has been deleted"
            Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
            Print #1, Now & " " & "Column A in sheet " & ws.Name & " has been deleted"
            Close #1
        End If
        
        'Move B4 value to C4 if misplaced
        If ws.Range("B4").Value <> "" Then
            If ws.Range("C4").Value <> "" Then
                Debug.Print "Both B4 and C4 in " & ws.Name & " have values"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "Both B4 and C4 in " & ws.Name & " have values"
                Close #1
            Else
                ws.Range("C4").Value = ws.Range("B4").Value
                ws.Range("B4").Value = ""
                Debug.Print "B4 value in " & ws.Name & " has been moved to C4"
                Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
                Print #1, Now & " " & "B4 value in " & ws.Name & " has been moved to C4"
                Close #1
            End If
        End If
        
        'Convert main data range to table
        ws.Range("A3:O4").Select
        Application.CutCopyMode = False
        ws.ListObjects.Add(xlSrcRange, Range("$A$3:$O$4"), , xlYes).Name = _
            "Program_" & ws.Range("C4").Value & "_MainData"
        
        'Convert project list range to table
        'ws.Range("A7:O100").Select
        'Selecting to the latest data record instead of selecting large table
        ws.Range("A7:O" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).Select
        Application.CutCopyMode = False
        ws.ListObjects.Add(xlSrcRange, Range("$A$7:$O$306"), , xlYes).Name = _
            "Program_" & ws.Range("C4").Value & "_Contracts"
        
        'Return to cell A1 for visual uniformaty
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
    
    Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
    Print #1, Now & " " & "Macro Done"
    Close #1
    
    Exit Sub

ErrHandler:
        
        Open Application.ActiveWorkbook.Path & "\Log.txt" For Append As #1
        Print #1, Now & " " & "Error in sheet " & ws.Name & " cell " & Cell.Address & " Error Code : " & Err.Number & ":" & Err.Description
        Close #1
        
        Resume Next
End Sub