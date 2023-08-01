Sub CombinedMacro2()
    If DataRefresh Then
        ActiveSheet.Shapes("GroupBeforePress1").Visible = msoTrue
        MsgBox "Expense reviewed and sent for processing.", vbInformation, "Confirmation"
    End If
End Sub

Function DataRefresh() As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim shouldRefresh As Boolean
    shouldRefresh = True

    Dim refreshTimestampCell As Range
    Set refreshTimestampCell = ws.Range("I3")
    
    
    ' Check if there's a table in the worksheet
    If ws.ListObjects.Count > 0 Then
        Dim tbl As ListObject
        Set tbl = ws.ListObjects(1) ' Refer to the first table

        ' Loop through each row in the table
Dim row As ListRow
For Each row In tbl.ListRows
    ' If there's information in column J (9th column in the table)
    If Trim(row.Range(1, 9).Value) <> "" Then
        ' Check the corresponding cell in column T (19th column in the table)
        If Trim(row.Range(1, 19).Value) = "" Then
            ' Prompt the user if information is missing in column T
            MsgBox "There is an expense with no Line Manager confirmation, please ensure that all expenses have been approved in Column Q and retry.", vbExclamation
            shouldRefresh = False
            Exit For
        End If
    End If
Next row
    Else

        MsgBox "No table found in the active worksheet!", vbExclamation
        shouldRefresh = False
    End If

    ' Refresh the data if the conditions were met
    If shouldRefresh Then
        For Each cn In ThisWorkbook.Connections
            cn.Refresh
        Next cn

        refreshTimestampCell.Value = "Refreshed: " & Now()

        With refreshTimestampCell.Font
            .name = "Arial"
            .Size = 11
            .Bold = True
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Color = RGB(208, 206, 206)
        End With

        lastRefresh = Now()
        MsgBox "Data Refreshed!", vbInformation
    End If

    DataRefresh = shouldRefresh
End Function
