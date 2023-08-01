
Sub CombinedMacro2()
    DataRefresh

    ActiveSheet.Shapes("GroupBeforePress1").Visible = msoTrue
    
     MsgBox "Expense reviewed and sent for processing.", vbInformation, "Confirmation"
End Sub
Sub DataRefresh()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Check if there's a table in the worksheet
    If ws.ListObjects.Count > 0 Then
        Dim tbl As ListObject
        Set tbl = ws.ListObjects(1) ' Refer to the first table

        ' Loop through each row in the table
        Dim row As ListRow
        For Each row In tbl.ListRows
            ' If there's information in column C (3rd column in the table)
            If Trim(row.Range(1, 2).Value) <> "" Then
                ' Check the corresponding cell in column Q (17th column in the table)
                If Trim(row.Range(1, 16).Value) = "" Then
                    ' Prompt the user if information is missing in column Q
                    MsgBox "There is information in column C but not in column Q on row " & row.Index + tbl.HeaderRowRange.row - 1 & ". Please input the necessary information in column Q.", vbExclamation
                    Exit Sub
                End If
            End If
        Next row

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
    
        Else
        MsgBox "No table found in the active worksheet!", vbExclamation
    End If

End Sub


NEW CODE 
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
            ' If there's information in column C (2nd column in the table)
            If Trim(row.Range(1, 2).Value) <> "" Then
                ' Check the corresponding cell in column Q (16th column in the table)
                If Trim(row.Range(1, 16).Value) = "" Then
                    ' Prompt the user if information is missing in column Q
                    MsgBox "There is an expense with no Line Manager Confirmation, please ensure that all expenses have been approved in Column Q and retry.", vbExclamation
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