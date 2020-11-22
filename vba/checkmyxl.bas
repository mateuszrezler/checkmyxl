Attribute VB_Name = "checkmyxl"
Sub CheckSelection()
    Dim wb As Workbook
    Dim mymodule As String
    Set wb = ActiveWorkbook
    mymodule = Left(wb.Name, (InStrRev(wb.Name, ".", -1, vbTextCompare) - 1))
    SelectedRange = Replace(selection.Address, "$", "")
    RunPython "from " & mymodule & " import main; main(selection='" & SelectedRange & "')"
End Sub

