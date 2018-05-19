Sub Eqn_Ins_in_Italic()

    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).ConvertToMathText

End Sub

'

Sub Eqn_Sym_Correction()

' This macro is used to fix the problem:
'   MS Word may mishandle some symbols when converting MathML to MS formulas.

    Selection.OMaths.Linearize

        ' Hat Circumflex
        With Selection.Find
        .ClearFormatting
        .text = "┴^"
        .Replacement.ClearFormatting
        .Replacement.text = ChrW(770)
        .Execute Replace:=wdWord, Forward:=True, _
            Wrap:=wdFindContinue
        End With
        
        ' Absolute Value
        With Selection.Find
        .ClearFormatting
        .text = "∣"
        .Replacement.ClearFormatting
        .Replacement.text = "|"
        .Execute Replace:=wdWord, Forward:=True, _
            Wrap:=wdFindContinue
        End With

    Selection.OMaths.BuildUp

End Sub

'

Sub Eqn_Num_Automatically()

' Applies To: Microsoft Word 2016 or later.
'
' Usage
' =====
' STEP 1.
'   Put the cursor in the last position inside the formula box.
'
' STEP 2.
'   Apply macro.

    With CaptionLabels("Equation")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = Selection.Paragraphs(1).Range.ListFormat.ListLevelNumber + 1
        .separator = wdSeparatorHyphen
    End With

    Selection.TypeText text:="#("

    Selection.InsertCaption Label:="Equation" _
        , Title:="", Position:=wdCaptionPositionAbove, ExcludeLabel:=1

    Selection.TypeText text:=")"

    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Size = ActiveDocument.Styles("Normal").Font.Size
    Selection.Font.Color = Automatic
    Selection.EndKey Unit:=wdLine

    SendKeys "~"

End Sub

'

Sub Eqn_Bookmark()

' Usage
' =====
'   Assume that the Eqn numbering has been generated by
'     another macro `Sub Eqn_Num_Automatically()`
'     in the following format:
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
'   Apply macro.
'     (Then the bookmark object is automatically copied to the clipboard,
'     and you can paste it directly to your target location)

    ActiveWindow.View.ShowFieldCodes = 1
    Selection.MoveRight Extend:=wdExtend
    Selection.Font.Italic = 0

    If InStr(Selection, "Eqn__") Then

        bookmark_name = RegexExtract(Selection, _
            "(Eqn__.*(?![\\}])\w)")
        ' Equivalent to
        '   `(Eqn__.*\w)(?<![\\}])`
        ' Explanation:
        '   The above is how to use the "Negative Lookahead" (the former) to simulate
        '     the "Negative Lookbehind" (the latter) in VBA regex.
        '   The capture group matches until the last `\w` in the string,
        '     regardless of if any `\W` in the group,
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
                bookmark_name + "\", PreserveFormatting:=False

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
                bookmark_name + "\", PreserveFormatting:=False
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
            bookmark_name + "\", PreserveFormatting:=False

        ' Generate bookmark REF and copy it to the clipboard
        Selection.InsertCrossReference ReferenceType:="Bookmark", ReferenceKind:= _
            wdContentText, ReferenceItem:=bookmark_name, InsertAsHyperlink:=True
        Selection.MoveLeft Extend:=wdExtend
        Selection.Cut
    End If

    ActiveWindow.View.ShowFieldCodes = 0

End Sub

'

' by @aevanko :  https://stackoverflow.com/a/7087145
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

' Run `Sub Shortcut_Keys_Customization()` to see in the visible pane

    CustomizationContext = NormalTemplate

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryCommand, Command:="Eqn_Bookmark"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyEquals, wdKeyAlt), KeyCategory _
        :=wdKeyCategoryCommand, Command:="Eqn_Ins_in_Italic"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyN, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryCommand, Command:="Eqn_Num_Automatically"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, wdKeyAlt), KeyCode2:= _
        BuildKeyCode(wdKeyC), KeyCategory:=wdKeyCategoryCommand, Command:="Eqn_Sym_Correction"

End Sub

'

Sub Shortcut_Keys_Customization()

    SendKeys "%"
    SendKeys "F"
    SendKeys "T"
    SendKeys "{DOWN 7}" & "{TAB 3}" & "~" & "{END}" & "{UP 4}"

End Sub
