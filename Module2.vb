Sub HideColumnsInWorksheets()
    Dim ws As Worksheet
    Dim wsNames As Variant
    Dim name As Variant

    ' List the names of worksheets you want to hide columns in
    wsNames = Array("Expense_form1,Expense_form2,Expense_form3,Expense_form4,Expense_form5,Expense_form6,Expense_form7,Expense_form8,Expense_form9,Expense_form10, Expense_form11, Expense_form12,Expense_form13, Expense_form14,Expense_form15, Expense_form16,Expense_form17, Expense_form18 ,Expense_form19 , Expense_form20, Expense_form21, Expense_form22, Expense_form23, Expense_form24")

    ' Loop through each worksheet name in the array
    For Each name In wsNames
        ' Check if the worksheet exists in the workbook
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(name)
        On Error GoTo 0

        ' If the worksheet exists, hide columns R to W
        If Not ws Is Nothing Then
            ws.Columns("R:W").Hidden = True
            Set ws = Nothing
        End If
    Next name
End Sub
