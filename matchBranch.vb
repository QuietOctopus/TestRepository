Private Sub createBranchColumn()

Application.ScreenUpdating = False
    Range("C4").Select 'Property Column
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Offset(3, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Branch"
    ActiveCell.Offset(1, 0).Range("A1").Select
Do Until (ActiveCell.Offset(0, -1) = vbNullString)
If (Left(ActiveCell.Offset(0, -1).Value, 2) = "CV") Then
ActiveCell.Value = "Central Valley"
    ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "IE") Then
    ActiveCell.Value = "Inland Empire"
        ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "LA") Then
        ActiveCell.Value = "Los Angeles"
            ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "OC") Then
            ActiveCell.Value = "Orange County"
                ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "SD") Then
                ActiveCell.Value = "San Diego"
End If
ActiveCell.Offset(1, 0).Activate
Loop
Application.ScreenUpdating = True
End Sub
                
Private Function matchBranch()


If (Left(ActiveCell.Offset(0, -1).Value, 2) = "CV") Then
ActiveCell.Value = "Central Valley"
    ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "IE") Then
        ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "LA") Then
        ActiveCell.Value = "Los Angeles"
            ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "OC") Then
            ActiveCell.Value = "Orange County"
                ElseIf (Left(ActiveCell.Offset(0, -1).Value, 2) = "SD") Then
                ActiveCell.Value = "San Diego"
                
End If

End Function
