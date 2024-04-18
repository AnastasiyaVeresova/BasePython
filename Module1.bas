Attribute VB_Name = "Module1"
Sub Fill_Blanks()
    For Each cell In Selection
        If IsEmpty(cell) Then cell.Value = cell.Offset(-1, 0).Value
    Next cell
End Sub

