Sub DeOxiafy()
    '
    ' DeOxiafy Macro for M/S Word - Replace all of each 16 redundant 'oxia' characters with their 'tonos' equivalents
    '
    '
    ' 1 Lowercase alpha
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1F71)
        .Replacement.Text = ChrW(&H3AC)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 2 Lowercase epsilon
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1F73)
        .Replacement.Text = ChrW(&H3AD)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 3 Lowercase eta
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1F75)
        .Replacement.Text = ChrW(&H3AE)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 4 Lowercase iota
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1F77)
        .Replacement.Text = ChrW(&H3AF)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 5 Lowercase omicron
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1F79)
        .Replacement.Text = ChrW(&H3CC)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 6 Lowercase upsilon
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1F7B)
        .Replacement.Text = ChrW(&H3CD)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 7 Lowercase omega
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1F7D)
        .Replacement.Text = ChrW(&H3CE)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 8 Uppercase alpha
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FBB)
        .Replacement.Text = ChrW(&H386)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 9 Uppercase epsilon
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FC9)
        .Replacement.Text = ChrW(&H388)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 10 Uppercase eta
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FCB)
        .Replacement.Text = ChrW(&H389)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 11 Uppercase iota
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FDB)
        .Replacement.Text = ChrW(&H38A)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 12 Uppercase omicron
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FF9)
        .Replacement.Text = ChrW(&H38C)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 13 Uppercase upsilon
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FEB)
        .Replacement.Text = ChrW(&H38E)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 14 Uppercase omega
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FFB)
        .Replacement.Text = ChrW(&H38F)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 15 Diareses iota
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FD3)
        .Replacement.Text = ChrW(&H390)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' 16 Diareses upsilon
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H1FE3)
        .Replacement.Text = ChrW(&H3B0)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

End Sub
