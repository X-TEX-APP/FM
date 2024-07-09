
        def ToggleTeX():
            return str(r'''
        Sub FormatBeforeToggleTeX()
            ActiveDocument.UndoClear

            With Selection.Find
                .Text = "\\\((*)\\\)"
                .Replacement.Text = "$\1$"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False

                Do While .Execute(Replace:=wdReplaceAll)
                Loop
            End With

            ' Dim DollarInt As Integer
            ' For DollarInt = 0 To 5
            '     With Selection.Find
            '         .Text = "\$\{(*?)\}\$"
            '         .Replacement.Text = "$\1$"
            '         .Forward = True
            '         .Wrap = wdFindContinue
            '         .Format = False
            '         .MatchCase = False
            '         .MatchWholeWord = False
            '         .MatchWildcards = True
            '         .MatchSoundsLike = False
            '         .MatchAllWordForms = False
            '         If Len(Selection.Text) <= 3 Then
            '             Do While .Execute(Replace:=WdReplace.wdReplaceAll)
            '             Loop
            '         End If
            '     End With
            ' Next

            Selection.Find.ClearFormatting

            ActiveDocument.UndoClear


            Dim rng As range
            Dim searchText As String
            Dim replacementText As String
            Dim foundRange As range

            Set rng = ActiveDocument.content
            searchText = "$*?$"

            With rng.Find
                .ClearFormatting
                .Text = searchText
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True

                Do While .Execute
                    Set foundRange = rng.Duplicate
                    If Len(foundRange.Text) <= 12 And InStr(foundRange.Text, "{") = 0 Then
                        replacementText = "$" & Mid(foundRange.Text, 2, Len(foundRange.Text) - 2) & "$"
                        foundRange.Text = "${" & Mid(foundRange.Text, 2, Len(foundRange.Text) - 2) & "}$"
                    End If
                    rng.Collapse wdCollapseEnd
                Loop
            End With
        End Sub
        Sub ToggleTeX()
            Application.ScreenUpdating = False

            Selection.Find.ClearFormatting
            With Selection.Find
                .Text = "(\\\[*\\\])"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                Do While .Execute = True
                    Application.Run MacroName:="MTCommand_TeXToggle"
                    With Selection.Find
                        .Text = "(\\\[*\\\])"
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = True
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        Do While .Execute = True
                                Application.Run MacroName:="MTCommand_TeXToggle"
                                ActiveDocument.UndoClear
                        Loop
                    End With
                    ActiveDocument.UndoClear
                Loop
                ActiveDocument.UndoClear
            End With
            Selection.EndKey Unit:=wdStory

            Selection.Find.ClearFormatting

            With Selection.Find
                .Text = "(\$*\$)"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                Do While .Execute = True
                    Application.Run MacroName:="MTCommand_TeXToggle"
                    With Selection.Find
                        .Text = "(\$*\$)"
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = True
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        Do While .Execute = True
                                Application.Run MacroName:="MTCommand_TeXToggle"
                                ActiveDocument.UndoClear
                        Loop
                    End With
                    ActiveDocument.UndoClear
                Loop
                ActiveDocument.UndoClear
            End With
            Selection.EndKey Unit:=wdStory
            
            Application.ScreenUpdating = True
        End Sub
        ''')

        def ClearLineSpace():
            return str(r'''
        Sub ExpandLine()
            Selection.WholeStory
            With Selection.ParagraphFormat
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(1.5)
                .FirstLineIndent = CentimetersToPoints(0)
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
            End With
        End Sub
        Sub ClearLineSpace()
            With Selection.Find
            .Text = "^p^l"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        Do While .Execute
            .Execute Replace:=wdReplaceAll
        Loop
        End With
        With Selection.Find
            .Text = "^l^p"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        Do While .Execute
            .Execute Replace:=wdReplaceAll
        Loop
        End With
        With Selection.Find
            .Text = "^l^l"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        Do While .Execute
            .Execute Replace:=wdReplaceAll
        Loop
        End With
        With Selection.Find
            .Text = "  "
            .Replacement.Text = " "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
            Loop
        End With
        With Selection.Find
            .Text = "^p^p"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
            Loop
        End With
        Call ExpandLine
    End Sub''')
        
        def QuestionFormat():
            return str(r'''
        Sub XuongdongABCD()
            ActiveDocument.range.ListFormat.ConvertNumbersToText
            Selection.WholeStory
            With Selection.ParagraphFormat
                .FirstLineIndent = CentimetersToPoints(0)
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
            End With
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n A."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n A"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n B."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n B"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n C."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n C"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n D."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n D"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                    .Text = "^11"
                    .Replacement.Text = "^13"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                    .Text = "^b"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            
            With Selection.Find
                .Text = "( )([.:,\)])"
                .Replacement.Text = "\2"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            With Selection.Find
                .Text = "( "
                .Replacement.Text = "("
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                    .Text = "([^32^9])([BCD])(.)"
                    .Replacement.Text = "^p" & "\2" & ". "
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
            With Selection.Find
                .Text = "C©u"
                .Replacement.Text = "Câu"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
            With Selection.Find
                .Text = "Caâu"
                .Replacement.Text = "Câu"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
            With Selection.Find
                .ClearFormatting
                .Text = "^p "
                .Replacement.ClearFormatting
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            With Selection.Find
                .Text = "^p^p"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            With Selection.Find
                .Text = "^t "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            With Selection.Find
                .Text = "^t"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .Text = "  "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
            Loop
            End With
        With Selection.Find
                .Text = "^p "
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
                .Text = " ^p"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
                .Text = "^p^p"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
        
            Selection.Find.ClearFormatting
            Selection.Find.Font.Underline = wdUnderlineSingle
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.Underline = wdUnderlineNone
            With Selection.Find
                .Text = "([.:])"
                .Replacement.Text = "\1"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Font.Underline = wdUnderlineSingle
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.Underline = wdUnderlineNone
            With Selection.Find
                .Text = " "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.WholeStory
            Selection.ParagraphFormat.TabStops.ClearAll
            With ActiveDocument.PageSetup
                .LineNumbering.Active = False
                .Orientation = wdOrientPortrait
                .TopMargin = CentimetersToPoints(1.5)
                .BottomMargin = CentimetersToPoints(1.5)
                .LeftMargin = CentimetersToPoints(1.5)
                .RightMargin = CentimetersToPoints(1)
                .Gutter = CentimetersToPoints(0)
                .HeaderDistance = CentimetersToPoints(0.6)
                .FooterDistance = CentimetersToPoints(0.6)
                .PageWidth = CentimetersToPoints(21)
                .PageHeight = CentimetersToPoints(29.7)
            End With
        End Sub
        Sub S_PageSetup()
            With ActiveDocument.PageSetup
                .TopMargin = CentimetersToPoints(0.8)
                .BottomMargin = CentimetersToPoints(1)
                .LeftMargin = CentimetersToPoints(1.9)
                .RightMargin = CentimetersToPoints(0.9)
                .Gutter = CentimetersToPoints(0)
                .HeaderDistance = CentimetersToPoints(0.8)
                .FooterDistance = CentimetersToPoints(0.7)
                .PageWidth = CentimetersToPoints(21)
                .PageHeight = CentimetersToPoints(29.7)
                .VerticalAlignment = wdAlignVerticalTop
            End With
            With Selection.ParagraphFormat
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .SpaceBefore = 0
                .SpaceAfter = 0
                .Alignment = wdAlignParagraphLeft
                .WidowControl = True
                .Hyphenation = True
                .FirstLineIndent = CentimetersToPoints(0)
                .OutlineLevel = wdOutlineLevelBodyText
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .MirrorIndents = True
            End With
        End Sub
        Sub CH_345(ByRef kind As Byte)
            ActiveDocument.range.ListFormat.ConvertNumbersToText
                Dim l1, l2, l3, l4, d_a, i As Byte
                Dim c As Integer
                Dim lmax As String
                Dim Shape1, Shape2, Shape3, Shape4 As Byte
                Dim title2, msg As String
                Dim ktMsg As Byte
                On Error Resume Next
                Dim myRange As range
                
                Call XuongdongABCD
                
                Application.ScreenUpdating = False
                Selection.HomeKey Unit:=wdStory, Extend:=wdMove
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                    With Selection.Find
                        .Text = "Câu "
                        .MatchWildcards = False
                    End With
                Do While Selection.Find.Execute = True
                    Selection.HomeKey Unit:=wdLine
                    Selection.TypeParagraph
                    Exit Do
                Loop
                For i = 1 To ActiveDocument.Tables.Count
                    ActiveDocument.Tables(i).Select
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    Selection.TypeParagraph
                Next i
                c = 1
                
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Text = "([^13^32^9])([AaBbCcDd])(.)"
                    .Replacement.Text = "\1\2\3" & " "
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.WholeStory
                Selection.ParagraphFormat.TabStops.ClearAll
                Selection.HomeKey Unit:=wdStory, Extend:=wdMove
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                    With Selection.Find
                        .Text = "^pCâu "
                        .MatchWildcards = False
                    End With
                    Do While Selection.Find.Execute = True
                        Selection.Collapse Direction:=wdCollapseEnd
                        With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="c" & c & "q"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                        End With
                        c = c + 1
                        ActiveDocument.UndoClear
                    Loop
                Selection.EndKey Unit:=wdStory
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="c" & c & "q"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    choiceA = "([^13])([Aa])(.)(*)([^13])([Bb])(.)"
                    choiceB = "([^13])([Bb])(.)(*)([^13])([Cc])(.)"
                    choiceC = "([Cc])(.)(*)([^13])([Dd])(.)"
                    choiceD = "([Dd])(.)"
                
                For i = 1 To c - 1
                    S_Wait_STD.Stt = i
                    S_Wait_STD.Repaint
                    Selection.Find.ClearFormatting
                    
                    'Danh dau phuong an A
                    Set myRange = ActiveDocument.range( _
                        Start:=ActiveDocument.Bookmarks("c" & i & "q").range.Start, _
                    End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").range.End)
                    myRange.Select
                    myRange.Find.Execute FindText:=choiceA, MatchWildcards:=True
                If myRange.Find.Found = True Then
                    myRange.MoveStart Unit:=wdCharacter, Count:=1
                    myRange.MoveEnd Unit:=wdCharacter, Count:=-2
                    myRange.Select
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="c" & i & "a"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
        
                    l1 = myRange.Characters.Count
                    Select Case myRange.InlineShapes.Count
                        Case 1
                        l1 = l1 + Round(myRange.InlineShapes(1).Width / 5.8)
                        Case 2
                        l1 = l1 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
                    End Select
                    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="s2"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                End If
                    'Danh dau phuong an B
                    Set myRange = ActiveDocument.range( _
                        Start:=ActiveDocument.Bookmarks("s2").range.Start, _
                        End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").range.End)
                    myRange.Find.Execute FindText:=choiceB, MatchWildcards:=True
                If myRange.Find.Found = True Then
                    myRange.MoveStart Unit:=wdCharacter, Count:=1
                    myRange.MoveEnd Unit:=wdCharacter, Count:=-2
                    myRange.Select
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="c" & i & "b"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                
                    l2 = myRange.Characters.Count
                    Select Case myRange.InlineShapes.Count
                        Case 1
                        l2 = l2 + Round(myRange.InlineShapes(1).Width / 6)
                        Case 2
                        l2 = l2 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
                    End Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="s2"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                End If
                'Danh dau phuong an C
                    Set myRange = ActiveDocument.range( _
                        Start:=ActiveDocument.Bookmarks("s2").range.Start, _
                        End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").range.End)
                    myRange.Find.Execute FindText:=choiceC, MatchWildcards:=True
                If myRange.Find.Found = True Then
                    myRange.MoveStart Unit:=wdCharacter, Count:=0
                    myRange.MoveEnd Unit:=wdCharacter, Count:=-2
                    myRange.Select
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="c" & i & "c"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                
                    l3 = myRange.Characters.Count
                    Select Case myRange.InlineShapes.Count
                    Case 1
                        l3 = l3 + Round(myRange.InlineShapes(1).Width / 6)
                        Case 2
                        l3 = l3 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
                    End Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="s2"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    
                End If
                    'Danh dau phuong an D
                    Set myRange = ActiveDocument.range( _
                        Start:=ActiveDocument.Bookmarks("s2").range.Start, _
                        End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").range.End)
                    myRange.Find.Execute FindText:=choiceD, MatchWildcards:=True
                If myRange.Find.Found = True Then
                    myRange.Select
                    Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
                    With ActiveDocument.Bookmarks
                        .Add range:=Selection.range, Name:="c" & i & "d"
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    
                    l4 = myRange.Characters.Count
                    Select Case myRange.InlineShapes.Count
                        Case 1
                        l4 = l4 + Round(myRange.InlineShapes(1).Width / 6)
                        Case 2
                        l4 = l4 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
                    End Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
                    
                End If
                    lmax = l1
                    If Val(lmax) < l2 Then lmax = l2
                    If Val(lmax) < l3 Then lmax = l3
                    If Val(lmax) < l4 Then lmax = l4
                    If Val(lmax) < 10 Then lmax = "0" & lmax
                    If Val(lmax) > 60 Then lmax = 60
                Dim chia1, chia2  As Byte
                Dim tab2, tab3, tab4 As Long
                Select Case kind
                Case 3
                    chia1 = 24
                    chia2 = 45
                    tab1 = 0.5
                    tab2 = 5
                    tab3 = 9.5
                    tab4 = 14
                Case 4
                    chia1 = 16
                    chia2 = 30
                    tab1 = 0.5
                    tab2 = 3.2
                    tab3 = 5.9
                    tab4 = 8.6
                Case 5
                    chia1 = 13
                    chia2 = 25
                    tab1 = 0.1
                    tab2 = 2.3
                    tab3 = 4.5
                    tab4 = 6.7
                End Select
            
                'Selection.WholeStory
                'Selection.ParagraphFormat.TabStops.ClearAll
                If Val(lmax) < chia1 Then
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
                    'Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(0.5)
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeText Text:=vbTab
                    
                    ActiveDocument.DefaultTabStop = CentimetersToPoints(tab1)
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab1) _
                    , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab2), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab3), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab4), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    
                    
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeBackspace
                    Selection.TypeText Text:=vbTab
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeBackspace
                    Selection.TypeText Text:=vbTab
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeBackspace
                    Selection.TypeText Text:=vbTab
                ElseIf Val(lmax) < chia2 Then
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeText Text:=vbTab

                    ActiveDocument.DefaultTabStop = CentimetersToPoints(tab1)
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab1) _
                    , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab3), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeBackspace
                    Selection.TypeText Text:=vbTab
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"

                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeText Text:=vbTab
                    ActiveDocument.DefaultTabStop = CentimetersToPoints(tab1)
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab1) _
                    , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(tab3), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeBackspace
                    Selection.TypeText Text:=vbTab
                    Else
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeText Text:=vbTab
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeText Text:=vbTab
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeText Text:=vbTab
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
                    Selection.MoveLeft Unit:=wdCharacter
                    Selection.TypeText Text:=vbTab
                    
                End If
                ActiveDocument.UndoClear
                
            Next i
                    Selection.Find.ClearFormatting
                    Selection.Find.Font.Underline = wdUnderlineNone
                    Selection.Find.Replacement.ClearFormatting
                    Selection.Find.Replacement.Font.ColorIndex = wdBlue
                    Selection.Find.Replacement.Font.Bold = True
                    With Selection.Find
                    .Text = "([^9])([Aa])(.)"
                    .Replacement.Text = "\1" & "A."
                    .MatchCase = True
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                    End With
                    
                    With Selection.Find
                    .Text = "([^9])([Bb])(.)"
                    .Replacement.Text = "\1" & "B."
                    .MatchCase = True
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                    End With
                    With Selection.Find
                    .Text = "([^9])([Cc])(.)"
                    .Replacement.Text = "\1" & "C."
                    .MatchCase = True
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                    End With
                    With Selection.Find
                    .Text = "([^9])([Dd])(.)"
                    .Replacement.Text = "\1" & "D."
                    .MatchCase = True
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                    End With
                    
                    Selection.Find.ClearFormatting
                    Selection.Find.Replacement.ClearFormatting
                    Selection.Find.Replacement.Font.ColorIndex = wdBlue
                    Selection.Find.Replacement.Font.Bold = True
                    With Selection.Find
                        .Text = "(Câu [0-9]{1,4})"
                        .Replacement.Text = "\1" & "."
                        .Forward = True
                        .Format = True
                        .Wrap = wdFindContinue
                        .MatchCase = True
                        .MatchWildcards = True
                        .Execute Replace:=wdReplaceAll
                    End With
                    Selection.Find.ClearFormatting
                    Selection.Find.Replacement.ClearFormatting
                    
                    With Selection.Find
                        .Text = ".:"
                        .Replacement.Text = "."
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With Selection.Find
                        .Text = ".."
                        .Replacement.Text = "."
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With Selection.Find
                        .Text = "  "
                        .Replacement.Text = " "
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        Do While .Execute
                            .Execute Replace:=wdReplaceAll
                            ActiveDocument.UndoClear
                        Loop
                    End With
                    With Selection.Find
                        .Text = "^p^p"
                        .Replacement.Text = "^p"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        Do While .Execute
                            .Execute Replace:=wdReplaceAll
                            ActiveDocument.UndoClear
                        Loop
                    End With
                    Selection.Find.ClearFormatting
                    Selection.Find.Font.Underline = wdUnderlineSingle
                    Selection.Find.Replacement.ClearFormatting
                    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
                    Selection.Find.Replacement.Font.ColorIndex = wdBlue
                    With Selection.Find
                        .Text = "^t"
                        .Replacement.Text = "^t"
                        .MatchCase = True
                        .Forward = True
                        .Format = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    Selection.Find.ClearFormatting
                    Selection.Find.Replacement.ClearFormatting
                    Selection.Find.Font.Underline = wdUnderlineSingle
                    Selection.Find.Replacement.Font.ColorIndex = wdBlue
                    Selection.Find.Replacement.Font.Bold = True
                    With Selection.Find
                        .Text = "([ABCD])"
                        .Replacement.Text = "\1"
                        .Forward = True
                        .Format = True
                        .Wrap = wdFindContinue
                        .MatchCase = True
                        .MatchWildcards = True
                        .Execute Replace:=wdReplaceAll
                    End With
            Application.ScreenUpdating = True
            
            Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
            If kind = 3 Then
                Call S_PageSetup
                Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
            ElseIf kind = 4 Then
                Selection.WholeStory
                With ActiveDocument.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = wdOrientPortrait
                    .TopMargin = CentimetersToPoints(0.6)
                    .BottomMargin = CentimetersToPoints(0.8)
                    .LeftMargin = CentimetersToPoints(1.5)
                    .RightMargin = CentimetersToPoints(0.86)
                    .Gutter = CentimetersToPoints(0)
                    .HeaderDistance = CentimetersToPoints(0.8)
                    .FooterDistance = CentimetersToPoints(0.7)
                    .PageWidth = CentimetersToPoints(14.8)
                    .PageHeight = CentimetersToPoints(21)
                    .FirstPageTray = wdPrinterDefaultBin
                    .OtherPagesTray = wdPrinterDefaultBin
                    .SectionStart = wdSectionNewPage
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .VerticalAlignment = wdAlignVerticalTop
                    .SuppressEndnotes = False
                    .MirrorMargins = False
                    .TwoPagesOnOne = False
                    .BookFoldPrinting = False
                    .BookFoldRevPrinting = False
                    .BookFoldPrintingSheets = 1
                    .GutterPos = wdGutterPosLeft
                End With
                Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
            Else
                Selection.WholeStory
                With Selection.PageSetup.TextColumns
                    .SetCount numColumns:=2
                    .EvenlySpaced = True
                    .LineBetween = True
                    .Width = CentimetersToPoints(5.97)
                    .Spacing = CentimetersToPoints(0.5)
                End With
                With ActiveDocument.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = wdOrientPortrait
                    .TopMargin = CentimetersToPoints(0.6)
                    .BottomMargin = CentimetersToPoints(0.8)
                    .LeftMargin = CentimetersToPoints(1)
                    .RightMargin = CentimetersToPoints(0.86)
                    .Gutter = CentimetersToPoints(0)
                    .HeaderDistance = CentimetersToPoints(0.8)
                    .FooterDistance = CentimetersToPoints(0.7)
                    .PageWidth = CentimetersToPoints(21)
                    .PageHeight = CentimetersToPoints(29.7)
                    .FirstPageTray = wdPrinterDefaultBin
                    .OtherPagesTray = wdPrinterDefaultBin
                    .SectionStart = wdSectionNewPage
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .VerticalAlignment = wdAlignVerticalTop
                    .SuppressEndnotes = False
                    .MirrorMargins = False
                    .TwoPagesOnOne = False
                    .BookFoldPrinting = False
                    .BookFoldRevPrinting = False
                    .BookFoldPrintingSheets = 1
                    .GutterPos = wdGutterPosLeft
                End With
                Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
            End If
            Selection.HomeKey Unit:=wdStory
        End Sub''' + r'''
        Sub ChuanhoaSualoi()
            ActiveDocument.range.ListFormat.ConvertNumbersToText
            Selection.HomeKey Unit:=wdStory
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n A."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n A"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            ActiveDocument.UndoClear
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n B."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n B"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            ActiveDocument.UndoClear
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n C."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n C"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            ActiveDocument.UndoClear
            With Selection.Find
                    .Text = "Ch" & ChrW(7885) & "n D."
                    .Replacement.Text = "Ch" & ChrW(7885) & "n D"
                    .Forward = True
                    .MatchCase = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            
            Selection.HomeKey Unit:=wdStory
            
            'Selection.Find.ClearFormatting
            'Selection.Find.Replacement.ClearFormatting
            'Selection.Find.Highlight = True
            'Selection.Find.Replacement.Highlight = True
            'Selection.Find.Replacement.Font.ColorIndex = wdBlue
            'With Selection.Find
                '.Text = "([ABCD])(.)"
                '.Replacement.Text = "\1" & "/."
                '.Forward = True
                '.Wrap = wdFindContinue
                '.MatchCase = True
                '.Format = True
                '.MatchWildcards = True
                '.Execute Replace:=wdReplaceAll
            'End With
            'ActiveDocument.UndoClear
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Font.ColorIndex = wdRed
            Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
            Selection.Find.Replacement.Font.ColorIndex = wdBlue
            With Selection.Find
                .Text = "([ABCD])(.)"
                .Replacement.Text = "\1" & "/."
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = True
                .Format = True
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Font.Underline = wdUnderlineSingle
            Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
            With Selection.Find
                .Text = "([ABCD])(.)"
                .Replacement.Text = "\1" & "/."
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = True
                .Format = True
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
        'Exit Sub
            With Selection.Find
                .Text = "  "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
            
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            With Selection.Find
                .Text = "^p "
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            'Do While .Execute
                .Execute Replace:=wdReplaceAll
            'Loop
            End With
            ActiveDocument.UndoClear
            With Selection.Find
                .Text = "^p^t"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .ClearFormatting
                .Text = "^l"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
            With Selection.Find
                .Text = "^p^p"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            With Selection.Find
                .Text = "( )([.:,\)])"
                .Replacement.Text = "\2"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchWildcards = True
                .Format = True
            Do While .Execute
                .Execute Replace:=wdReplaceAll
                ActiveDocument.UndoClear
            Loop
            End With
            'Selection.Font.Underline = wdUnderlineNone
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "([^32^9])([ABCD])([\/.])"
                .Replacement.Text = "#" & "\2\3"
                '.Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchWildcards = True
                '.Format = True
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
        'Exit Sub
            With Selection.Find
                .Text = "C©u"
                .Replacement.Text = "Câu"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
            With Selection.Find
                .Text = "Caâu"
                .Replacement.Text = "Câu"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
            With Selection.Find
                .Text = "^t"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            ActiveDocument.UndoClear
        'Exit Sub
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            'Selection.Find.Replacement.Font.ColorIndex = wdBlue
            'Selection.Find.Replacement.Font.Bold = True
            With Selection.Find
                .Text = "(\#)([BCD])([\/.])"
                .Replacement.Text = "^9" & "\2\3"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
            End With
                Selection.Find.Execute Replace:=wdReplaceAll
        ActiveDocument.UndoClear
            With Selection.Find
                .Text = "/."
                .Replacement.Text = "."
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            ActiveDocument.UndoClear
            Selection.EndKey Unit:=wdStory
            Selection.EndKey Unit:=wdLine, Extend:=wdExtend
            If Len(Selection) = 1 Then Selection.TypeBackspace
            
        End Sub
        Sub CH_2()
            If Selection.paragraphs.Count > 1 Then
                Dim idx As String
                idx = 1
                Selection.Cut
                ActiveDocument.range.ListFormat.ConvertNumbersToText
                Dim DocGoc As Document
                Dim DocTmp As Document
                Set DocGoc = ActiveDocument
                Set DocTmp = Documents.Add
                Selection.Paste
                Application.ScreenUpdating = False
                Call ChuanhoaSualoi
                Dim tt As Integer
                    Selection.WholeStory
                    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
                    'Selection.Style = ActiveDocument.Styles("Normal")

                    With Selection.ParagraphFormat
                        .FirstLineIndent = CentimetersToPoints(0)
                        .LeftIndent = CentimetersToPoints(1.5)
                        .RightIndent = CentimetersToPoints(0)
                    End With
                    Selection.ParagraphFormat.TabStops.ClearAll
                    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.75)
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(1.75) _
                    , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(6), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(10), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(14), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
            
                With Selection.Find
                    .Text = "(Câu [0-9]{1,4}.)"
                    .Replacement.Text = "##"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = True
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                End With
                ActiveDocument.UndoClear
                With Selection.Find
                    .Text = "(Câu [0-9]{1,4}:)"
                    .Replacement.Text = "##"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = True
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                End With
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Text = "## "
                    .Replacement.Text = "##"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = False
                    .Execute Replace:=wdReplaceAll
                End With
                Set danhsach = ActiveDocument.content
                tt = 0
        Tiep2:
                danhsach.Find.Execute FindText:="##", Forward:=True
                If danhsach.Find.Found = True Then
                    danhsach.Select
                    Selection.ParagraphFormat.TabStops.ClearAll
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(1.5)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphJustify
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(-1.75)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                        .MirrorIndents = False
                        .TextboxTightWrap = wdTightNone
                    End With
                            
                    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
                        .NumberFormat = "Câu " & "%1."
                        .TrailingCharacter = wdTrailingTab
                        .NumberStyle = wdListNumberStyleArabic
                        .NumberPosition = CentimetersToPoints(0)
                        .Alignment = wdListLevelAlignLeft
                        .TextPosition = CentimetersToPoints(0)
                        .TabPosition = wdUndefined
                        .ResetOnHigher = 0
                        .StartAt = Val(idx)
                        .LinkedStyle = ""
                        .Font.Bold = True
                        .Font.color = wdColorBlue
                        .Font.Italic = False
                    End With
                    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
                    Selection.range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                    ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
                    True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                    wdWord10ListBehavior
                    Selection.Delete Unit:=wdCharacter, Count:=1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(1.5)
                        .RightIndent = CentimetersToPoints(0)
                        .FirstLineIndent = CentimetersToPoints(-1.75)
                    End With
                    ActiveDocument.UndoClear
                        tt = tt + 1
                    GoTo Tiep2
                
                End If
                Selection.HomeKey Unit:=wdStory
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                Selection.Find.Replacement.Font.ColorIndex = wdBlue
                With Selection.Find
                    .Text = "([^13^32^9])([ABCD].)"
                    .Replacement.Text = "\1\2"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=wdReplaceAll
                End With
                Selection.WholeStory
                With ActiveDocument.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = wdOrientPortrait
                    .TopMargin = CentimetersToPoints(1)
                    .BottomMargin = CentimetersToPoints(1)
                    .LeftMargin = CentimetersToPoints(2)
                    .RightMargin = CentimetersToPoints(1)
                    .Gutter = CentimetersToPoints(0)
                    .HeaderDistance = CentimetersToPoints(0.6)
                    .FooterDistance = CentimetersToPoints(0.6)
                    .PageWidth = CentimetersToPoints(21)
                    .PageHeight = CentimetersToPoints(29.7)
                End With
                Selection.Find.ClearFormatting
                Selection.Find.Font.Underline = wdUnderlineSingle
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Underline = wdUnderlineNone
                With Selection.Find
                    .Text = ". "
                    .Replacement.Text = ". "
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = True
                    .Format = True
                    .Execute Replace:=wdReplaceAll
                End With
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Text = "  "
                    .Replacement.Text = " "
                    .Forward = True
                    .Wrap = wdFindContinue
                    
                    .MatchWildcards = False
                'Do While .Execute
                    .Execute Replace:=wdReplaceAll
                'Loop
                End With
                Selection.Find.ClearFormatting
                    Selection.Find.Replacement.ClearFormatting
                    With Selection.Find
                        .Text = "^9^32"
                        .Replacement.Text = "^9"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With Selection.Find
                        .Text = "^32^9"
                        .Replacement.Text = "^9"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "(\[)([DGHLS])([STHYO])(*)([abcd])(\])"
                            .Replacement.Text = "\1\2\3\4\5\6"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "( )(\[)([DGHLS])([STHYO])(*)([abcd])(\])"
                            .Replacement.Text = "\2\3\4\5\6\7"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "( )(\[)([012])([DH])(*)([1234])(\])"
                            .Replacement.Text = "\2\3\4\5\6\7"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "(\[)([012])([DH])(*)([1234])(\])"
                            .Replacement.Text = "\1\2\3\4\5\6"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdBlue
                        Selection.Find.Replacement.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With Selection.Find
                            .Text = ChrW(72) & ChrW(432) & ChrW(7899) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(100) & ChrW(7851) & ChrW(110) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Replacement.Text = ChrW(76) & ChrW(7901) & ChrW(105) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = False
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdBlue
                        Selection.Find.Replacement.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With Selection.Find
                            .Text = ChrW(76) & ChrW(7901) & ChrW(105) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Replacement.Text = ChrW(76) & ChrW(7901) & ChrW(105) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = False
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                    Application.ScreenUpdating = True
                    Application.Visible = True
                    Selection.WholeStory
                    Selection.Copy
                    DocTmp.Close (False)
                    DocGoc.Activate
                    Selection.Paste
                        
            Else
                    Call ChuanhoaSualoi
                    Selection.WholeStory
                    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
                    'Selection.Style = ActiveDocument.Styles("Normal")
                    With Selection.ParagraphFormat
                        .FirstLineIndent = CentimetersToPoints(0)
                        .LeftIndent = CentimetersToPoints(1.5)
                        .RightIndent = CentimetersToPoints(0)
                    End With
                    Selection.ParagraphFormat.TabStops.ClearAll
                    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.75)
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(1.75) _
                    , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(6), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(10), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(14), _
                    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
            
                With Selection.Find
                    .Text = "(Câu [0-9]{1,4}.)"
                    .Replacement.Text = "##"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = True
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                End With
                ActiveDocument.UndoClear
                With Selection.Find
                    .Text = "(Câu [0-9]{1,4}:)"
                    .Replacement.Text = "##"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = True
                    .MatchWildcards = True
                    .Execute Replace:=wdReplaceAll
                End With
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                    With Selection.Find
                        .Text = "## "
                        .Replacement.Text = "##"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                Set danhsach = ActiveDocument.content
                'Dim tt As Integer
                tt = 0
        Tiep:
                danhsach.Find.Execute FindText:="##", Forward:=True
                If danhsach.Find.Found = True Then
                    danhsach.Select
                    Selection.ParagraphFormat.TabStops.ClearAll
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(1.5)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphJustify
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(-1.75)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                        .MirrorIndents = False
                        .TextboxTightWrap = wdTightNone
                    End With
                            
                    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
                        .NumberFormat = "Câu " & "%1."
                        .TrailingCharacter = wdTrailingTab
                        .NumberStyle = wdListNumberStyleArabic
                        .NumberPosition = CentimetersToPoints(0)
                        .Alignment = wdListLevelAlignLeft
                        .TextPosition = CentimetersToPoints(0)
                        .TabPosition = wdUndefined
                        .ResetOnHigher = 0
                        .StartAt = 1
                        .LinkedStyle = ""
                        .Font.Bold = True
                        .Font.color = wdColorBlue
                        .Font.Italic = False
                    End With
                    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
                    Selection.range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                    ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
                    True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                    wdWord10ListBehavior
                    Selection.Delete Unit:=wdCharacter, Count:=1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(1.5)
                        .RightIndent = CentimetersToPoints(0)
                        .FirstLineIndent = CentimetersToPoints(-1.75)
                    End With
                    ActiveDocument.UndoClear
                        tt = tt + 1
                    GoTo Tiep
                
                End If
                Selection.HomeKey Unit:=wdStory
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                Selection.Find.Replacement.Font.ColorIndex = wdBlue
                With Selection.Find
                    .Text = "([^13^32^9])([ABCD].)"
                    .Replacement.Text = "\1\2"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=wdReplaceAll
                End With
                Selection.WholeStory
                With ActiveDocument.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = wdOrientPortrait
                    .TopMargin = CentimetersToPoints(1)
                    .BottomMargin = CentimetersToPoints(1)
                    .LeftMargin = CentimetersToPoints(2)
                    .RightMargin = CentimetersToPoints(1)
                    .Gutter = CentimetersToPoints(0)
                    .HeaderDistance = CentimetersToPoints(0.6)
                    .FooterDistance = CentimetersToPoints(0.6)
                    .PageWidth = CentimetersToPoints(21)
                    .PageHeight = CentimetersToPoints(29.7)
                End With
                Selection.Find.ClearFormatting
                Selection.Find.Font.Underline = wdUnderlineSingle
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Underline = wdUnderlineNone
                With Selection.Find
                    .Text = ". "
                    .Replacement.Text = ". "
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = True
                    .Format = True
                    .Execute Replace:=wdReplaceAll
                End With
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Text = "  "
                    .Replacement.Text = " "
                    .Forward = True
                    .Wrap = wdFindContinue
                    
                    .MatchWildcards = False
                'Do While .Execute
                    .Execute Replace:=wdReplaceAll
                'Loop
                End With
                Selection.Find.ClearFormatting
                    Selection.Find.Replacement.ClearFormatting
                    With Selection.Find
                        .Text = "^9^32"
                        .Replacement.Text = "^9"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With Selection.Find
                        .Text = "^32^9"
                        .Replacement.Text = "^9"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "(\[)([DGHLS])([STHYO])(*)([abcd])(\])"
                            .Replacement.Text = "\1\2\3\4\5\6"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "( )(\[)([DGHLS])([STHYOI])(*)([abcd])(\])"
                            .Replacement.Text = "\2\3\4\5\6\7"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "( )(\[)([012])([DH])(*)([1234])(\])"
                            .Replacement.Text = "\2\3\4\5\6\7"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdPink
                        With Selection.Find
                            .Text = "(\[)([012])([DH])(*)([1234])(\])"
                            .Replacement.Text = "\1\2\3\4\5\6"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdBlue
                        Selection.Find.Replacement.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With Selection.Find
                            .Text = ChrW(72) & ChrW(432) & ChrW(7899) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(100) & ChrW(7851) & ChrW(110) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Replacement.Text = ChrW(76) & ChrW(7901) & ChrW(105) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = False
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                        ActiveDocument.UndoClear
                        
                        Selection.Find.ClearFormatting
                        Selection.Find.Replacement.ClearFormatting
                        Selection.Find.Replacement.Font.Bold = True
                        Selection.Find.Replacement.Font.ColorIndex = wdBlue
                        Selection.Find.Replacement.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With Selection.Find
                            .Text = ChrW(76) & ChrW(7901) & ChrW(105) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Replacement.Text = ChrW(76) & ChrW(7901) & ChrW(105) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(7843) & ChrW(105)
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = False
                        End With
                        Selection.Find.Execute Replace:=wdReplaceAll
                Application.ScreenUpdating = True
            End If
        End Sub''' + r'''
        Sub RenumberAuto()
            Application.ScreenUpdating = False
            ActiveDocument.Range.ListFormat.ConvertNumbersToText
            Selection.HomeKey Unit:=wdStory
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(Câu [0-9]{1,4}[.:])"
                .Replacement.text = "#"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = True
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
                .text = "(^13)([0-9]{1,4}[/.:)])"
                .Replacement.text = "\1" & "#"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = True
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = "#"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchWildcards = False
                If Selection.Find.Execute = False Then Exit Sub
            End With
            With Selection.Find
                .text = "#^t"
                .Replacement.text = "#"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                Do While .Execute
                .Execute Replace:=wdReplaceAll
                Loop
            End With
            With Selection.Find
                .text = "# "
                .Replacement.text = "#"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                Do While .Execute
                .Execute Replace:=wdReplaceAll
                Loop
            End With
            Set danhsach = ActiveDocument.Content
            Tiep:
            danhsach.Find.Execute findtext:="#", Forward:=True
            If danhsach.Find.Found = True Then
                danhsach.Select
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.ParagraphFormat.TabStops.ClearAll
                ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
                    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(1.75) _
                    , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
                    .NumberFormat = "Câu " & "%1."
                    .TrailingCharacter = wdTrailingTab
                    .NumberStyle = wdListNumberStyleArabic
                    .NumberPosition = CentimetersToPoints(0)
                    .Alignment = wdListLevelAlignLeft
                    .TextPosition = CentimetersToPoints(1.75)
                    .TabPosition = wdUndefined
                    .ResetOnHigher = 0
                    .StartAt = 1
                    .LinkedStyle = ""
                    .Font.Bold = True
                    .Font.Color = wdColorBlue
                End With
                ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
                Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
                True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                wdWord10ListBehavior
                Selection.Delete Unit:=wdCharacter, Count:=1
                GoTo Tiep
                Else
                Selection.HomeKey Unit:=wdStory
                Application.ScreenUpdating = True
            End If
        End Sub
''')

        def FormatSpecialText():
            return str(r'''
Sub FormatSpecialText()
    Application.ScreenUpdating = False
    Dim doc As Document
    Dim findRange As range
    Dim subRange As range
    Dim re As Object
    Set doc = ActiveDocument
    Set findRange = doc.content
    Set re = CreateObject("VBScript.RegExp")
    Dim matches As Object
    Dim match As Object

    re.Global = True
    re.IgnoreCase = False

    re.pattern = "\\textbf\{((\{[^{}]*\}|[^{}])*)\}"
    Set matches = re.Execute(findRange.Text)
    For Each match In matches
        Set subRange = doc.content
        With subRange.Find
            .ClearFormatting
            .Text = match.Value
            .Replacement.Text = match.SubMatches(0)
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWildcards = False
        End With
        If subRange.Find.Execute Then
            subRange.Text = match.SubMatches(0)
            subRange.Font.Bold = True
        End If
    Next match

    re.pattern = "\\textit\{((\{[^{}]*\}|[^{}])*)\}"
    Set matches = re.Execute(findRange.Text)
    For Each match In matches
        Set subRange = doc.content
        With subRange.Find
            .ClearFormatting
            .Text = match.Value
            .Replacement.Text = match.SubMatches(0)
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWildcards = False
        End With
        If subRange.Find.Execute Then
            subRange.Text = match.SubMatches(0)
            subRange.Font.Italic = True
        End If
    Next match

    re.pattern = "\\underline\{((\{[^{}]*\}|[^{}])*)\}"
    Set matches = re.Execute(findRange.Text)
    For Each match In matches
        Set subRange = doc.content
        With subRange.Find
            .ClearFormatting
            .Text = match.Value
            .Replacement.Text = match.SubMatches(0)
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWildcards = False
        End With
        If subRange.Find.Execute Then
            subRange.Text = match.SubMatches(0)
            subRange.Font.Underline = wdUnderlineSingle
        End If
    Next match

    Application.ScreenUpdating = True
End Sub''')
        
        def ConvertImageToCenter():
            return str(r'''
        Sub ConvertImageToCenter()
            Dim shape As shape
            Dim inlineShape As inlineShape
            Dim para As Paragraph

            For Each shape In ActiveDocument.Shapes
                If shape.Type = msoPicture Then
                    Set inlineShape = shape.ConvertToInlineShape
                    Set para = inlineShape.range.paragraphs(1)
                    para.Alignment = wdAlignParagraphCenter
                End If
            Next shape

            For Each inlineShape In ActiveDocument.InlineShapes
                If inlineShape.Type = wdInlineShapePicture Then
                    Set para = inlineShape.range.paragraphs(1)
                    para.Alignment = wdAlignParagraphCenter
                End If
            Next inlineShape
        End Sub
        ''')
