Attribute VB_Name = "checkmyxl"
Sub CheckSelection()
    SelectedRange = Replace(selection.Address, "$", "")
    RunPython ("from checkmyxl import main; main('" & SelectedRange & "')")
End Sub

