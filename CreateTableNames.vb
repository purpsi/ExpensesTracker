Sub RenameTableInAllWorksheets()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim regex As Object
    Dim matches As Object
    Dim numberInSheetName As String
    Dim strE7 As String
    Dim newTableName As String
    
    ' Set reference to the workbook
    Set wb = ThisWorkbook
    
    ' Initialize regex to find numbers in the worksheet name
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\d+"  ' Matches one or more digits
    
    ' Iterate over each worksheet
    For Each ws In wb.Worksheets
    
        ' Find the number in the worksheet name
        Set matches = regex.Execute(ws.Name)
        
        If matches.Count > 0 Then
            numberInSheetName = matches(0)
        Else
            MsgBox "No number found in the worksheet name: " & ws.Name & ". Moving to the next sheet."
            GoTo NextSheet
        End If
        
        ' Get the string from E7
        strE7 = ws.Range("E7").Value
        If Len(strE7) < 4 Then
            MsgBox "String in E7 is too short for the worksheet: " & ws.Name & ". Moving to the next sheet."
            GoTo NextSheet
        End If



' RAW

                Sub RenameTableInAllWorksheets()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim regex As Object
    Dim matches As Object
    Dim numberInSheetName As String
    Dim strE7 As String
    Dim newTableName As String
    
    ' Set reference to the workbook
    Set wb = ThisWorkbook
    
    ' Initialize regex to find numbers in the worksheet name
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\d+"  ' Matches one or more digits
    
    ' Iterate over each worksheet
    For Each ws In wb.Worksheets
    
        ' Find the number in the worksheet name
        Set matches = regex.Execute(ws.Name)
        
        If matches.Count > 0 Then
            numberInSheetName = matches(0)
        Else
            MsgBox "No number found in the worksheet name: " & ws.Name & ". Moving to the next sheet."
            GoTo NextSheet
        End If
        
        ' Get the string from E7
        strE7 = ws.Range("E7").Value
        If Len(strE7) < 4 Then
            MsgBox "String in E7 is too short for the worksheet: " & ws.Name & ". Moving to the next sheet."
            GoTo NextSheet
        End If
        
        ' Construct the new table name
        newTableName = "HTETABLE_" & Left(strE7, 2) & Right(strE7, 2) & numberInSheetName
        
        ' Rename the table (assuming there's only one table in the sheet)
        If ws.ListObjects.Count > 0 Then
            Set tbl = ws.ListObjects(1)
            tbl.Name = newTableName
        Else
            MsgBox "No table found in the worksheet: " & ws.Name & ". Moving to the next sheet."
        End If
        
NextSheet:
    Next ws

    MsgBox "Operation completed."

End Sub

        
        ' Construct the new table name
        newTableName = "HTETABLE_" & Left(strE7, 2) & Right(strE7, 2) & numberInSheetName
        
        ' Rename the table (assuming there's only one table in the sheet)
        If ws.ListObjects.Count > 0 Then
            Set tbl = ws.ListObjects(1)
            tbl.Name = newTableName
        Else
            MsgBox "No table found in the worksheet: " & ws.Name & ". Moving to the next sheet."
        End If
        
NextSheet:
    Next ws

    MsgBox "Operation completed."

End Sub
