Sub noMargins()
    With ActiveDocument.PageSetup
        .TopMargin = InchesToPoints(0)
        .BottomMargin = InchesToPoints(0)
        .LeftMargin = InchesToPoints(0)
        .RightMargin = InchesToPoints(0)
    End With
End Sub
