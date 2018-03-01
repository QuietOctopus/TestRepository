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
