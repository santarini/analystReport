Sub reportStyle()
    With ActiveDocument.PageSetup
        .TopMargin = InchesToPoints(0)
        .BottomMargin = InchesToPoints(0)
        .LeftMargin = InchesToPoints(0)
        .RightMargin = InchesToPoints(0)
    End With
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(1)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
'single space
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("No Spacing")
'define Heading 1
    With ActiveDocument.Styles("Heading 1").Font
        .Name = "+Headings"
        .Size = 36
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = -738148353
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
End Sub
