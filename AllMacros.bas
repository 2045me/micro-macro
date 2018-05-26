Sub Eqn_Italic()

    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).ConvertToMathText

' Caution
' =======
'   Word will **crash** if the equation box is deleted using the BACKSPACE key in paragraph mode.
'     (For example: click the mouse three times to select the entire equation box,
'       then press the backspace key twice)
' Solution
' --------
'   Always use the DELETE key only.
'
' Other Bug
' =========
'   Run-time error '5941':
'     The requested member of the collection does not exist.
'   --It means that the equation boxes conflict.
'     This problem occurs when the equation has been cleared
'       but the empty selection region remains in the equation environment,
'       and at the same time you didn't pay attention and added a new one there.
' Solution
' --------
'   Directly continue typing your equation,
'     or arbitrarily press arrow keys 2 times.

End Sub

'

Sub Eqn_MathML_Correction()

' Introduction
' ============
'   This macro is originally mainly used to fix the following issues:
'     MS Word may mishandle some symbols when converting MathML to MS formulas.
'   And then a few decorative features were added.
'
' Usage
' =====
'   Select the entire equation box and run the macro.
'
' Note
' ====
'   It's better used only for interline formulas, not for inline formulas.
'
' Bug
' ===
'   Incorrectly exclude or include the numerator and denominator.
'   --This bug is caused by subroutines:
'       `Insert thin space` and
'       `Integrals and Large Operators`.


    Selection.OMaths.Linearize
    
    ' ClearFormatting required!
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With Selection.Find
        .MatchWildcards = False
    End With
    
    
        ' Hat circumflex
        With Selection.Find
            .text = ChrW(9524) & "^"
            .Replacement.text = ChrW(770)
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=False
        End With
        
        
        ' Absolute value
        With Selection.Find
            .text = ChrW(8739)
            .Replacement.text = "|"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=False
        End With
        
        
        ' Vector arrow
        With Selection.Find
            .text = ChrW(9524) & ChrW(8594)
            .Replacement.text = ChrW(8407)
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=False
        End With
    
    
    With Selection.Find
        .MatchWildcards = True
    End With
    
    
    ' Font marking - Regular
    ' ======================
    With Selection.Find.Font
        .Bold = 0
        .Italic = 0
    End With
    
    Dim text_identifier__R As String
    text_identifier__R = "~%%~"
    
    With Selection.Find
        .text = "([A-Za-z0-9" & _
                  ChrW(8711) & _
                "]{1,})" ' Including mathematical symbols of Roman typefaces
        .Replacement.text = text_identifier__R & "\1" & text_identifier__R
        .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
    End With
    ' ====================
    
    
    ' Font marking - Bold & Italic
    ' ============================
    With Selection.Find.Font
        .Bold = 1
        .Italic = 1
    End With
    
    Dim text_identifier__B_I As String
    text_identifier__B_I = "~%``%~"
    
    With Selection.Find
        .text = "([! ])"
        .Replacement.text = text_identifier__B_I & "\1" & text_identifier__B_I
        .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
    End With
    ' ====================
    
    
    ' Font marking - Bold & Regular
    ' =============================
    With Selection.Find.Font
        .Bold = 1
        .Italic = 0
    End With
    
    Dim text_identifier__B_R As String
    text_identifier__B_R = "~%`%~"
    
    With Selection.Find
        .text = "(?)"
        .Replacement.text = text_identifier__B_R & "\1" & text_identifier__B_R
        .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
    End With
    ' ====================
    
    
    Selection.Font.Italic = 0 ' Note: Otherwise, the alphabetic characters in the equation environment cannot be found.
    
    ' ClearFormatting required!
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    
        ' Must NOT remove spacing here
        
        
        ' Remove erroneous and redundant placeholders next to the subscript
        With Selection.Find
            .text = ChrW(12310) & "_\(" & "(?@)" & "\)" & ChrW(12311)
            .Replacement.text = "\1"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
        End With
        
        
        ' Remove erroneous and redundant placeholders
        With Selection.Find
            .text = "[" & ChrW(12310) & "]" & "(?@)" & "[" & ChrW(12311) & "]"
            .Replacement.text = "\1"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
        End With
        
        
        ' Integrals and Large Operators
        With Selection.Find
            ' With super- and sub-script
            .text = "([" & _
                      ChrW(8719) & ChrW(8721) & ChrW(8747) & ChrW(8748) & ChrW(8749) & _
                    "]_*?*^^*?*)([ ]{1,2})([!\(\)])"
            .Replacement.text = ChrW(8201) & "\1" & ChrW(9618) & "\3"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            .text = "([" & _
                      ChrW(8719) & ChrW(8721) & ChrW(8747) & ChrW(8748) & ChrW(8749) & _
                    "]_*?@^^*?@[ 0-9]\))([ " & ChrW(8201) & "])"
            .Replacement.text = ChrW(8201) & "\1" & ChrW(9618)
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            ' Remove some incorrect extra additions
            .text = "([\)])" & ChrW(9618) & " ([\(])"
            .Replacement.text = "\1 \2"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            
            ' Without super- or sub-script
            .text = "([" & _
                      ChrW(8719) & ChrW(8721) & ChrW(8747) & ChrW(8748) & ChrW(8749) & _
                    "])([!_])"
            .Replacement.text = ChrW(8201) & "\1" & ChrW(9618) & "\2"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            
            ' With sub-script only
            .text = "([" & _
                      ChrW(8719) & ChrW(8721) & _
                    "]_[! ]@ )"
            .Replacement.text = ChrW(8201) & "\1" & ChrW(9618)
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            ' Remove some incorrect extra additions
            .text = ChrW(9618) & "(\)" & ChrW(9618) & ")"
            .Replacement.text = "\1"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            
            ' Summations or Products of Integrals
            .text = "([" & _
                      ChrW(8719) & ChrW(8721) & _
                        "][! ]*)([" & _
                      ChrW(8747) & ChrW(8748) & ChrW(8749) & _
                    "])"
            .Replacement.text = ChrW(8201) & "\1" & ChrW(9618) & "\2"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            ' Remove some incorrect extra additions
            .text = ChrW(8201) & "(" & ChrW(9618) & ")"
            .Replacement.text = "\1"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            .text = "(" & ChrW(9618) & "[!" & _
                      ChrW(8719) & ChrW(8721) & _
                      ChrW(8747) & ChrW(8748) & ChrW(8749) & _
                    "]* )" & ChrW(9618)
            .Replacement.text = "\1"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
            
            .text = "(" & ChrW(9618) & "{2,})"
            .Replacement.text = ChrW(9618)
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
        End With
        
        
        ' Insert thin space
        With Selection.Find
            ' Principle:
            '   Would rather not add, do not mistakenly add.
            '   Therefore `[!A-z0-9...]`
            .text = "([!A-z0-9_/\(~" & ChrW(34) & "][ \)A-Za-ce-z" & ChrW(8201) & "])([A-Za-z])"
            .Replacement.text = "\1" & ChrW(8201) & "\2"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
        End With
    
    
    Selection.Font.Italic = 1
    
    
    ' Font restoring
    ' ==============
    ' Regular
    ' -------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.text = "\1"
        .MatchWildcards = True
        Do While .Execute(FindText:=text_identifier__R & "([!^^" & text_identifier__R & "]@)" & text_identifier__R)
            If InStr(Selection, " ") = False Then
                Selection.Font.Italic = False
                If Selection Like text_identifier__R & "d" & text_identifier__R Then
                    .Replacement.text = ChrW(8201) & "\1"
                End If
            Else
                Exit Do
            End If
        Loop
        
        ' Reselect the equation block region
        Selection.MoveUp Unit:=wdParagraph
        Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
        
        .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
    End With
    
    ' Bold & Regular
    ' --------------
    With Selection.Find
        .Replacement.text = "\1"
        .MatchWildcards = True
    
        Do While .Execute(FindText:=text_identifier__B_R & "([!^^" & text_identifier__B_R & "]@)" & text_identifier__B_R, _
                    Wrap:=wdFindStop)
                Selection.Font.Bold = True
                Selection.Font.Italic = False
        Loop
        
        ' Reselect the equation block region
        Selection.MoveUp Unit:=wdParagraph
        Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
        
        .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
    End With
    
    ' Bold & Italic
    ' -------------
    With Selection.Find
        .Replacement.text = "\1"
        .MatchWildcards = True
    
        Do While .Execute(FindText:=text_identifier__B_I & "([!^^" & text_identifier__B_I & "]@)" & text_identifier__B_I, _
                    Wrap:=wdFindStop)
                Selection.Font.Bold = True
                Selection.Font.Italic = True
        Loop
        
        ' Reselect the equation block region
        Selection.MoveUp Unit:=wdParagraph
        Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
        
        .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
    End With
    ' ====================
    
    
        ' Linear fraction bar should be used in the superscript
        With Selection.Find
            .text = "(^^\([! " & ChrW(8201) & "]@)/([! \)" & ChrW(8201) & "]@\))"
            .Execute Wrap:=wdFindStop, MatchWildcards:=True
                If Selection Like "^(?*/?*)" = True Then
                    Selection.MoveRight
                    Selection.MoveLeft
                    Selection.MoveLeft Unit:=wdWord, Count:=3
                    Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
                    Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionFrac). _
                        Frac.Type = wdOMathFracLin
                    Selection.MoveEndUntil "(", wdBackward
                    Selection.MoveRight Unit:=wdWord, Count:=3
                    Selection.TypeBackspace
                    Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
                    Selection.Cut
                    Selection.MoveLeft
                    Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
                End If
            ' Reselect the equation block region
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
        End With
    
    
    ' Reselect the equation block region
    Selection.MoveUp Unit:=wdParagraph
    Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
    
    
        ' Fix the latent bug that the equation array does not lie in the same column.
        '   --caused by the preceding `Font marking - Regular` -> `mathematical symbols of Roman typefaces`.
        If Selection.Find.Execute(FindText:="[!^13]" & ChrW(9632) & "\(") Then
            Selection.MoveRight
            Selection.MoveLeft Count:=2, Extend:=wdExtend
            Selection.Cut
            Selection.MoveUp Unit:=wdParagraph
            Selection.Paste
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
        End If
        
        
        ' Remove erroneous and redundant placeholders again
        With Selection.Find
            .text = "[" & ChrW(12310) & "]" & "(?@)" & "[" & ChrW(12311) & "]"
            .Replacement.text = "\1"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
        End With
        
        
        ' Remove redundant spacing
        With Selection.Find
            .text = "(?)[ " & ChrW(8201) & "]{2,}(?)"
            .Replacement.text = "\1" & ChrW(8201) & "\2"
            .Execute Replace:=wdWord, Wrap:=wdFindStop, MatchWildcards:=True
        End With
    
    
    Selection.EndKey
    Selection.OMaths.BuildUp

End Sub

'

Sub Eqn_Num()

' Note
' ====
'   1. Only applies to Microsoft Word 2016 and later,
'        because `#(...)` is used to create an equation array.
'   2. Release your hand as soon as possible after pressing the shortcut key.
'
' Usage
' =====
' STEP 1.
'   Put the cursor in the last position inside the equation box.
'
' STEP 2.
'   Run the macro.


    With CaptionLabels("Equation")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = Selection.Paragraphs(1).Range.ListFormat.ListLevelNumber + 1
        .separator = wdSeparatorHyphen
        'separator = wdSeparatorHyphen "-" || wdSeparatorPeriod "."
    End With
    
    Selection.TypeText text:="#("
    Selection.InsertCaption Label:="Equation", ExcludeLabel:=1
    Selection.TypeText text:=")"
    
    Selection.MoveUp Unit:=wdParagraph, Extend:=wdExtend
    Selection.Font.Size = ActiveDocument.Styles("Normal").Font.Size
    Selection.Font.Color = Automatic
    Selection.EndKey
    
    SendKeys "~{BS}{RIGHT}"
    Selection.Font.Italic = 0

End Sub

'

Sub Eqn_Bookmark()

' Usage
' =====
'   Assume that the Eqn numbering has been generated by
'     another macro `Eqn_Num` in the following format:
'     (1.1-1)
'
' STEP opt.
'   Add a custom label in the following format:
'     (1.1-1{Eqn__Name})
'
' STEP 1.
'   Put the cursor closely before right parenthesis:
'     (1.1-1|)
'           ^
'     (1.1-1{Eqn__Name}|)
'                      ^
' STEP 2.
'   Run the macro.
'     (Then the bookmark object is automatically copied to the clipboard,
'       and you can paste it directly to your target location)


    ActiveWindow.View.ShowFieldCodes = 1
    Selection.MoveRight Extend:=wdExtend
    Selection.Font.Italic = 0
    
    If InStr(Selection, "Eqn__") Then
        
        bookmark_name = RegexExtract(Selection, _
            "(Eqn__.*(?![\\}])\w)")
        ' Equivalent to
        '   `(Eqn__.*\w)(?<![\\}])`
        ' Explanation:
        '   Above shows how to use the "Negative Lookahead" (the former) to simulate
        '     the "Negative Lookbehind" (the latter) in VBA regex.
        '   In this example, the capture group matches until the last `\w` in the string,
        '     regardless of whether there is any `\W` in the group,
        '     and ignores any `\W` after the last `\w`.
        ''MsgBox "Bookmark name:" + Chr(13) + bookmark_name
        
        If bookmark_name Like "Eqn__*[!0-9A-Z_a-z]*" Then
            MsgBox "Bookmark name may only contain alphanumeric characters or underscores."
            Exit Sub
        End If
        
        If ActiveDocument.Bookmarks.Exists(bookmark_name) Then
            ' Generate bookmark REF
            Selection.EndKey
            Selection.MoveEndUntil Chr(21), wdBackward
            Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
                "REF " + bookmark_name + " \h", PreserveFormatting:=False
            Selection.MoveLeft Extend:=wdExtend
            Selection.Cut
            
            ' Override the field to ensure bookmark identifier is consistent with REF
            Selection.EndKey
            Selection.MoveEndUntil Chr(21), wdBackward
            Selection.MoveLeft Extend:=wdExtend
            Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
                bookmark_name & "\", PreserveFormatting:=False
        
        Else
            ' Delete the field of the copied custom new bookmark name
            Selection.MoveRight
            Selection.MoveLeft
            Selection.MoveStartUntil "{", wdBackward
            Selection.MoveLeft Count:=2
            Selection.EndKey Extend:=wdExtend
            Selection.Delete
            Selection.HomeKey Extend:=wdExtend
            
            ' Add bookmark
            With ActiveDocument.Bookmarks
                .Add Range:=Selection.Range, Name:=bookmark_name
            End With
            
            ' Generate bookmark REF
            Selection.EndKey
            Selection.InsertCrossReference ReferenceType:="Bookmark", ReferenceKind:= _
                wdContentText, ReferenceItem:=bookmark_name, InsertAsHyperlink:=True
            Selection.MoveLeft Extend:=wdExtend
            Selection.Cut
            
            ' Add bookmark identifier (hidden)
            Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
                bookmark_name & "\", PreserveFormatting:=False
        End If
    
    Else
        bookmark_name = "Eqn__" + Format(Now, "yyyyMMddHHmmss")
        
        ' Add bookmark
        Selection.MoveRight
        Selection.MoveLeft
        Selection.HomeKey Extend:=wdExtend
        With ActiveDocument.Bookmarks
            .Add Range:=Selection.Range, Name:=bookmark_name
        End With
        Selection.MoveRight
        
        ' Add bookmark identifier (hidden)
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
            bookmark_name & "\", PreserveFormatting:=False
        
        ' Generate bookmark REF and copy it to the clipboard
        Selection.InsertCrossReference ReferenceType:="Bookmark", ReferenceKind:= _
            wdContentText, ReferenceItem:=bookmark_name, InsertAsHyperlink:=True
        Selection.MoveLeft Extend:=wdExtend
        Selection.Cut
    End If

    ActiveWindow.View.ShowFieldCodes = 0

End Sub

'

' Function RegexExtract
'   by @aevanko
'     Link to: https://stackoverflow.com/a/7087145

Function RegexExtract(ByVal text As String, _
                      ByVal extract_what As String, _
                      Optional separator As String = ", ") As String
    
    Dim allMatches As Object
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
    Dim i As Long, j As Long
    Dim result As String
    
    RE.Pattern = extract_what
    RE.Global = True
    Set allMatches = RE.Execute(text)
    
    For i = 0 To allMatches.Count - 1
        For j = 0 To allMatches.Item(i).submatches.Count - 1
            result = result & (separator & allMatches.Item(i).submatches.Item(j))
        Next
    Next
    
    If Len(result) <> 0 Then
        result = Right$(result, Len(result) - Len(separator))
    End If
    
    RegexExtract = result

End Function

'

Sub Shortcut_Keys_Preset_Assignment()

' You can run `Shortcut_Keys_Customization_Pane`
'   to see and modify easily in the visible pane.
'
' Current preset:
'   Alt + B       <=    Eqn_Bookmark
'   Alt + =       <=    Eqn_Italic
'   Alt + 1       <=    Eqn_Num
'   Alt + S, C    <=    Eqn_MathML_Correction


    CustomizationContext = NormalTemplate
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryCommand, Command:="Eqn_Bookmark"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyEquals, wdKeyAlt), KeyCategory _
        :=wdKeyCategoryCommand, Command:="Eqn_Italic"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey1, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryCommand, Command:="Eqn_Num"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, wdKeyAlt), KeyCode2:= _
        BuildKeyCode(wdKeyC), KeyCategory:=wdKeyCategoryCommand, Command:="Eqn_MathML_Correction"

End Sub

'

Sub Shortcut_Keys_Customization_Pane()

    SendKeys "%"
    SendKeys "+{F10}"
    SendKeys "{UP 2}~{TAB 3}~{END}{UP 4}{TAB}"

End Sub
