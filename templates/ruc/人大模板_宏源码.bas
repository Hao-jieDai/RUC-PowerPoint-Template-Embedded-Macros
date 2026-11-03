Attribute VB_Name = "RUCNavigation"
Option Explicit

' Use UTF-16 strings directly; the Windows ANSI code page is never involved.
#If VBA7 Then
Private Declare PtrSafe Function NativeMessageBoxW Lib "user32" Alias "MessageBoxW" ( _
    ByVal owner As LongPtr, ByVal textPtr As LongPtr, ByVal titlePtr As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function NativeActiveWindow Lib "user32" Alias "GetActiveWindow" () As LongPtr
#Else
Private Declare Function NativeMessageBoxW Lib "user32" Alias "MessageBoxW" ( _
    ByVal owner As Long, ByVal textPtr As Long, ByVal titlePtr As Long, ByVal flags As Long) As Long
Private Declare Function NativeActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
#End If

' RUC PowerPoint Template (Embedded Macros), VBA edition 2.2.1. No network, shell or auto-run.
' Navigation and editable contents are deliberately maintained separately.
Private Const NAV As String = "RUC_NAV_"
Private Const ENTRY As String = "RUC_TOC_ENTRY_"

Public Sub RUC_RefreshNavigation()
    On Error GoTo Failed
    Dim p As Presentation
    Set p = ActivePresentation
    RequireTemplate p
    Application.StartNewUndoEntry
    RefreshDeck p
    Exit Sub
Failed:
    UnicodeMessage U("5237 65B0 672A 5B8C 6210 FF1A") & Err.Description, vbExclamation, U("4EBA 5927 6A21 677F")
End Sub

Public Sub RUC_OnRefresh(control As IRibbonControl)
    RUC_RefreshNavigation
End Sub

Public Sub RUC_OnContents(control As IRibbonControl)
    RUC_RebuildContents
End Sub

Public Sub RUC_OnHelp(control As IRibbonControl)
    RUC_Help
End Sub

Public Sub RUC_Help()
    UnicodeMessage U("0031 002E 0020 5728 5DE6 4FA7 7F29 7565 56FE 4E2D FF0C 53F3 952E 8282 540D 79F0 FF0C 91CD 547D 540D 6216 65B0 589E 8282 3002") & vbCrLf & _
        U("0032 002E 0020 589E 5220 3001 79FB 52A8 9875 9762 540E FF0C 70B9 51FB 3010 5237 65B0 5BFC 822A 548C 9875 7801 3011 3002") & vbCrLf & _
        U("0033 002E 0020 76EE 5F55 6587 5B57 53EF 76F4 63A5 4FEE 6539 FF0C 666E 901A 5237 65B0 4F1A 4FDD 7559 6587 5B57 548C 4F4D 7F6E 3002") & vbCrLf & _
        U("0034 002E 0020 9700 8981 6309 8282 91CD 65B0 751F 6210 76EE 5F55 65F6 FF0C 5355 72EC 70B9 51FB 3010 4ECE 8282 91CD 65B0 751F 6210 76EE 5F55 3011 3002") & vbCrLf & _
        U("0035 002E 0020 7F16 8F91 65F6 4FDD 5B58 4E3A 0020 002E 0070 0070 0074 006D FF1B 653E 6620 524D 5237 65B0 5E76 53E6 5B58 4E3A 0020 002E 0070 0070 0074 0078 3002") & vbCrLf & _
        U("666E 901A 0020 002E 0070 0070 0074 0078 0020 7684 8DF3 8F6C 548C 5DF2 4FDD 5B58 7684 9875 7801 65E0 9700 5B8F 5373 53EF 653E 6620 3002") & vbCrLf & vbCrLf & _
        U("7EDF 4E00 4FEE 6539 9875 811A FF1A 7F16 8F91 7B2C 4E00 9875 5C01 9762 4E2D 7684 9898 76EE 3001 59D3 540D 548C 5B66 9662 FF0C 518D 5237 65B0 3002") & vbCrLf & _
        U("5B8C 6574 4F7F 7528 65B9 6CD5 8BF7 53C2 9605 968F 9644 7684 0020 0052 0045 0041 0044 004D 0045 002E 0074 0078 0074 3002"), vbInformation, U("4EBA 5927 6A21 677F 4F7F 7528 8BF4 660E")
End Sub

Private Sub RequireTemplate(p As Presentation)
    If p.Tags("RUC_MACRO_TEMPLATE") <> "1" Then Err.Raise vbObjectError + 101, , U("8BF7 5148 5207 6362 5230 4EBA 5927 5185 7F6E 5B8F 6A21 677F 521B 5EFA 7684 6F14 793A 6587 7A3F 3002")
    If p.ReadOnly Then Err.Raise vbObjectError + 102, , U("5F53 524D 6587 7A3F 53EA 8BFB FF0C 8BF7 53E6 5B58 4E3A 53EF 7F16 8F91 7684 0020 002E 0070 0070 0074 006D 0020 6587 4EF6 3002")
End Sub

Private Function MinVal(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinVal = a Else MinVal = b
End Function

Private Function MaxVal(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxVal = a Else MaxVal = b
End Function

Private Function Meta(p As Presentation, key As String, fallback As String) As String
    Dim sh As Shape
    Meta = fallback
    If p.Slides.Count = 0 Then Exit Function
    For Each sh In p.Slides(1).Shapes
        If sh.Name = "RUC_META_" & key And sh.HasTextFrame Then
            Meta = Replace(Replace(sh.TextFrame.TextRange.Text, vbCr, " "), vbLf, " ")
            Exit Function
        End If
    Next sh
End Function

Private Function GetSections(p As Presentation) As Collection
    Dim result As New Collection, j As Long, k As Long, first As Long, count As Long
    Dim label As String, excluded As String, indices As Collection, record As Collection
    excluded = "|" & p.Tags("RUC_NAV_EXCLUDE_SECTIONS") & "|"
    If p.SectionProperties.Count = 0 Then
        Set indices = New Collection
        For k = 1 To p.Slides.Count: indices.Add k: Next k
        Set record = New Collection
        record.Add U("62A5 544A 5185 5BB9"), "name": record.Add indices, "slides": record.Add "fallback", "id"
        If indices.Count > 0 Then result.Add record
    Else
        For j = 1 To p.SectionProperties.Count
            label = p.SectionProperties.Name(j)
            first = p.SectionProperties.FirstSlide(j)
            count = p.SectionProperties.SlidesCount(j)
            If count > 0 And first > 0 And InStr(1, excluded, "|" & label & "|", vbTextCompare) = 0 Then
                Set indices = New Collection
                For k = first To MinVal(first + count - 1, p.Slides.Count): indices.Add k: Next k
                Set record = New Collection
                record.Add label, "name": record.Add indices, "slides"
                record.Add p.SectionProperties.SectionID(j), "id"
                result.Add record
            End If
        Next j
    End If
    Set GetSections = result
End Function

Private Function AddText(owner As Object, nm As String, value As String, x As Single, y As Single, _
    w As Single, h As Single, size As Single, color As Long, align As Long) As Shape
    Dim sh As Shape, normalHeight As Single, spacedHeight As Single
    Set sh = owner.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    sh.Name = nm
    With sh.TextFrame
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse: .AutoSize = ppAutoSizeNone
        .TextRange.Text = value
        .TextRange.Font.Name = "Times New Roman": .TextRange.Font.NameFarEast = "KaiTi"
        .TextRange.Font.Size = Int(size + 0.5): .TextRange.Font.Color.RGB = color
        .TextRange.ParagraphFormat.Alignment = align
        .TextRange.ParagraphFormat.Bullet.Visible = msoFalse
        With .TextRange.ParagraphFormat
            .LineRuleWithin = msoTrue: .SpaceWithin = 1.5
            .LineRuleBefore = msoFalse: .SpaceBefore = 0
            .LineRuleAfter = msoFalse: .SpaceAfter = 0
        End With
    End With
    With sh.TextFrame2.TextRange.ParagraphFormat
        .LeftIndent = 0: .RightIndent = 0: .FirstLineIndent = 0
    End With
    sh.TextFrame2.AutoSize = msoAutoSizeNone
    sh.Width = w
    sh.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1
    normalHeight = sh.TextFrame.TextRange.BoundHeight
    sh.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.5
    spacedHeight = sh.TextFrame.TextRange.BoundHeight
    sh.Height = MaxVal(h, spacedHeight + 1)
    sh.Top = MaxVal(0, y + normalHeight - spacedHeight)
    Set AddText = sh
End Function

Private Sub LinkTo(sh As Shape, target As Slide, label As String)
    With sh.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.Address = ""
        .Hyperlink.SubAddress = CStr(target.SlideID) & "," & CStr(target.SlideIndex) & "," & Replace(label, ",", U("FF0C"))
    End With
    sh.AlternativeText = label
End Sub

Private Function EmptyPlaceholder(sh As Shape) As Boolean
    On Error GoTo KeepShape
    If sh.Type <> msoPlaceholder Then Exit Function
    If sh.HasTable Or sh.HasChart Then Exit Function
    Select Case sh.PlaceholderFormat.Type
        Case ppPlaceholderSlideNumber, ppPlaceholderDate, ppPlaceholderFooter, ppPlaceholderHeader
            Exit Function
    End Select
    ' A populated picture is still a placeholder: never remove it.
    If sh.PlaceholderFormat.ContainedType <> msoAutoShape And sh.PlaceholderFormat.ContainedType <> msoTextBox Then Exit Function
    If sh.HasTextFrame Then
        If Len(Trim(Replace(Replace(sh.TextFrame.TextRange.Text, vbCr, ""), vbLf, ""))) <> 0 Then Exit Function
    End If
    EmptyPlaceholder = True
KeepShape:
End Function

Public Sub RUC_CleanEmptyPlaceholders()
    Dim p As Presentation
    Set p = ActivePresentation
    RequireTemplate p
    Application.StartNewUndoEntry
    CleanPlaceholders p
End Sub

Private Sub CleanPlaceholders(p As Presentation)
    Dim s As Slide, j As Long, removed As Long
    For Each s In p.Slides
        For j = s.Shapes.Count To 1 Step -1
            If EmptyPlaceholder(s.Shapes(j)) Then s.Shapes(j).Delete: removed = removed + 1
        Next j
    Next s
    p.Tags.Add "RUC_EMPTY_REMOVED", CStr(removed)
End Sub

' Duplicate a populated text placeholder on its own slide. PowerPoint resolves
' inherited formatting into the resulting ordinary, independently editable box.
Private Sub NormalizeTextPlaceholders(p As Presentation)
    Dim s As Slide, sh As Shape, clone As Shape, copies As ShapeRange
    Dim j As Long, pos As Long, count As Long, nm As String
    Dim x As Single, y As Single, w As Single, h As Single
    For Each s In p.Slides
        For j = s.Shapes.Count To 1 Step -1
            Set sh = s.Shapes(j)
            If sh.Type = msoPlaceholder And sh.HasTextFrame Then
                If sh.PlaceholderFormat.ContainedType = msoAutoShape Or sh.PlaceholderFormat.ContainedType = msoTextBox Then
                    If Len(Trim(sh.TextFrame.TextRange.Text)) > 0 Then
                        nm = sh.Name: x = sh.Left: y = sh.Top: w = sh.Width: h = sh.Height: pos = sh.ZOrderPosition
                        Set copies = sh.Duplicate
                        Set clone = copies(1)
                        If clone.Type = msoPlaceholder Then Err.Raise vbObjectError + 104, , U("672A 80FD 5C06 6587 5B57 5360 4F4D 6846 8F6C 6362 4E3A 666E 901A 6587 672C 6846 3002")
                        clone.Left = x: clone.Top = y: clone.Width = w: clone.Height = h
                        sh.Delete
                        clone.Name = nm
                        Do While clone.ZOrderPosition > pos
                            clone.ZOrder msoSendBackward
                        Loop
                        count = count + 1
                    End If
                End If
            End If
        Next j
    Next s
    p.Tags.Add "RUC_TEXT_CONVERTED", CStr(count)
End Sub

' This routine intentionally does not write, recreate, or delete contents text.
Public Sub RefreshDeck(p As Presentation)
    RequireTemplate p
    Dim n As Long, a As Long, k As Long, i As Long, j As Long, idx As Long, d As Long
    Dim sections As Collection, sec As Collection, ids As Collection
    Dim w As Single, h As Single, sx As Single, sy As Single, weight() As Double, totalWeight As Double
    Dim hx() As Single, hw() As Single, hf() As Single, dx() As Single, dd() As Single, active() As Long
    Dim base As Single, gw As Single, gap As Single, diameter As Single
    Dim mast As Master, s As Slide, sh As Shape, used() As Boolean, muted As Long, white As Long
    n = p.Slides.Count
    If n = 0 Then Exit Sub
    p.Tags.Add "RUC_NAV_UPDATING", "1"
    w = p.PageSetup.SlideWidth: h = p.PageSetup.SlideHeight: sx = w / 720: sy = h / 540
    Set sections = GetSections(p)
    ReDim weight(0 To sections.Count): ReDim hx(0 To sections.Count): ReDim hw(0 To sections.Count): ReDim hf(0 To sections.Count)
    ReDim dx(1 To n): ReDim dd(1 To n): ReDim active(1 To n): ReDim used(1 To p.Designs.Count)
    muted = RGB(184, 128, 137): white = RGB(255, 255, 255)
    For a = 1 To sections.Count
        Set sec = sections(a): Set ids = sec("slides")
        weight(a) = MaxVal(68, MaxVal(Len(sec("name")) * 11 + 10, ids.Count * 7.5 + 10))
        totalWeight = totalWeight + weight(a)
    Next a
    base = 17 * sx
    For a = 1 To sections.Count
        Set sec = sections(a): Set ids = sec("slides")
        gw = (w - 34 * sx) * weight(a) / MaxVal(1, totalWeight)
        gap = MinVal(7.6 * sx, (gw - 6 * sx) / MaxVal(1, ids.Count))
        diameter = MaxVal(1.3 * sx, MinVal(5 * sx, gap * 0.64))
        hx(a) = base: hw(a) = gw: hf(a) = MinVal(10.8 * sx, gw / MaxVal(1, Len(sec("name"))) * 0.94)
        For k = 1 To ids.Count
            idx = ids(k): dx(idx) = base + (k - 1) * gap: dd(idx) = diameter: active(idx) = a
        Next k
        base = base + gw
    Next a
    For Each s In p.Slides: used(s.Design.Index) = True: Next s
    For d = 1 To p.Designs.Count
        If used(d) Then
            Set mast = p.Designs(d).SlideMaster
            For j = mast.Shapes.Count To 1 Step -1
                If Left(mast.Shapes(j).Name, Len(NAV)) = NAV Then mast.Shapes(j).Delete
            Next j
            For a = 1 To sections.Count
                Set sec = sections(a): Set ids = sec("slides")
                Set sh = AddText(mast, NAV & "SECTION_" & CStr(a - 1), CStr(sec("name")), hx(a), 1.4 * sy, hw(a) - 5 * sx, 15 * sy, hf(a), muted, ppAlignLeft)
                LinkTo sh, p.Slides(ids(1)), CStr(sec("name"))
                For k = 1 To ids.Count
                    idx = ids(k)
                    Set sh = mast.Shapes.AddShape(msoShapeOval, dx(idx), 20 * sy, dd(idx), dd(idx))
                    sh.Name = NAV & "DOT_" & p.Slides(idx).SlideID
                    sh.Line.ForeColor.RGB = muted: sh.Line.Weight = 0.65 * sx: sh.Fill.Visible = msoFalse
                    LinkTo sh, p.Slides(idx), CStr(sec("name")) & " " & k & "/" & ids.Count & U("0020 00B7 0020 7B2C") & idx & U("9875")
                Next k
            Next a
            Set sh = AddText(mast, NAV & "AUTHOR", Meta(p, "AUTHOR", U("6C47 62A5 4EBA")), 17 * sx, h - 34 * sy, w * 0.36, 17 * sy, 10.8 * sx, white, ppAlignLeft)
            Set sh = AddText(mast, NAV & "INSTITUTE", Meta(p, "INSTITUTE", U("4E2D 56FD 4EBA 6C11 5927 5B66")), w * 0.38, h - 34 * sy, w * 0.596, 17 * sy, 10.8 * sx, white, ppAlignRight)
            Set sh = AddText(mast, NAV & "TITLE", Meta(p, "TITLE", U("62A5 544A 6807 9898")), 17 * sx, h - 15.2 * sy, w * 0.79, 15 * sy, 10.3 * sx, white, ppAlignLeft)
            Set sh = AddText(mast, NAV & "PAGE", "1", w - 90 * sx, h - 15.2 * sy, 28 * sx, 15 * sy, 10.5 * sx, white, ppAlignRight)
            sh.TextFrame.TextRange.Text = ""
            sh.TextFrame.TextRange.InsertSlideNumber
            sh.TextFrame.TextRange.Font.Name = "Times New Roman": sh.TextFrame.TextRange.Font.Size = Int(10.5 * sx + 0.5): sh.TextFrame.TextRange.Font.Color.RGB = white
            sh.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
            Set sh = AddText(mast, NAV & "TOTAL", " / " & n, w - 62 * sx, h - 15.2 * sy, 45 * sx, 15 * sy, 10.5 * sx, white, ppAlignLeft)
        End If
    Next d
    For i = 1 To n
        Set s = p.Slides(i)
        For j = s.Shapes.Count To 1 Step -1
            If s.Shapes(j).Name = NAV & "ACTIVE_SECTION" Or s.Shapes(j).Name = NAV & "CURRENT" Then s.Shapes(j).Delete
        Next j
        a = active(i)
        If a > 0 Then
            Set sec = sections(a): Set ids = sec("slides")
            Set sh = AddText(s, NAV & "ACTIVE_SECTION", CStr(sec("name")), hx(a), 1.4 * sy, hw(a) - 5 * sx, 15 * sy, hf(a), white, ppAlignLeft)
            LinkTo sh, p.Slides(ids(1)), CStr(sec("name"))
            Set sh = s.Shapes.AddShape(msoShapeOval, dx(i), 20 * sy, dd(i), dd(i))
            sh.Name = NAV & "CURRENT": sh.Line.ForeColor.RGB = white: sh.Line.Weight = 0.65 * sx
            sh.Fill.Solid: sh.Fill.ForeColor.RGB = white
            LinkTo sh, s, CStr(sec("name")) & U("0020 00B7 0020 5F53 524D 9875")
        End If
        s.Tags.Add "RUC_NAV_OWNER", CStr(s.SlideID)
    Next i
    UpdateContentsLinks p, sections
    CleanPlaceholders p
    NormalizeTextPlaceholders p
    ' Deleting a populated placeholder can create a new empty layout prompt.
    ' Remove those replacement prompts after the text has been materialized.
    CleanPlaceholders p
    p.Tags.Add "RUC_NAV_SAVED_COUNT", CStr(n)
    p.Tags.Add "RUC_NAV_VERSION", "2.2.1-VBA"
    p.Tags.Add "RUC_NAV_UPDATED_AT", Format(Now, "yyyy-mm-dd hh:nn:ss")
    p.Tags.Add "RUC_NAV_UPDATING", "0"
End Sub

Private Sub UpdateContentsLinks(p As Presentation, sections As Collection)
    Dim s As Slide, sh As Shape, sec As Collection, ids As Collection, a As Long, sectionId As String, found As Boolean
    For Each s In p.Slides
        For Each sh In s.Shapes
            If Left(sh.Name, Len(ENTRY)) = ENTRY Then
                sectionId = sh.Tags("RUC_TOC_SECTION_ID"): found = False
                For a = 1 To sections.Count
                    Set sec = sections(a)
                    If StrComp(sectionId, CStr(sec("id")), vbTextCompare) = 0 Then
                        Set ids = sec("slides")
                        LinkTo sh, p.Slides(ids(1)), sh.TextFrame.TextRange.Text
                        found = True: Exit For
                    End If
                Next a
                If Not found Then
                    sh.ActionSettings(ppMouseClick).Action = ppActionNone
                End If
            End If
        Next sh
    Next s
End Sub

Public Sub RUC_RebuildContents()
    On Error GoTo Failed
    Dim p As Presentation
    Set p = ActivePresentation
    RequireTemplate p
    If UnicodeMessage(U("5C06 6309 5DE6 4FA7 7684 8282 540D 79F0 91CD 65B0 751F 6210 6240 6709 76EE 5F55 9875 3002") & vbCrLf & _
        U("6A21 677F 76EE 5F55 6761 76EE 7684 624B 52A8 6587 5B57 548C 4F4D 7F6E 4F1A 88AB 66FF 6362 FF1B 4F60 53E6 52A0 7684 6587 672C 6846 4F1A 4FDD 7559 3002") & vbCrLf & _
        U("53EA 60F3 66F4 65B0 9876 90E8 5706 70B9 548C 9875 7801 FF0C 8BF7 53D6 6D88 5E76 9009 62E9 3010 5237 65B0 5BFC 822A 548C 9875 7801 3011 3002") & vbCrLf & vbCrLf & _
        U("786E 8BA4 91CD 65B0 751F 6210 76EE 5F55 FF1F"), vbYesNo + vbQuestion + vbDefaultButton2, U("91CD 65B0 751F 6210 76EE 5F55")) <> vbYes Then Exit Sub
    Application.StartNewUndoEntry
    RebuildContents p
    RefreshDeck p
    Exit Sub
Failed:
    UnicodeMessage U("76EE 5F55 672A 80FD 66F4 65B0 FF1A") & Err.Description, vbExclamation, U("4EBA 5927 6A21 677F")
End Sub

Public Sub RebuildContents(p As Presentation)
    RequireTemplate p
    Dim sections As Collection, sec As Collection, ids As Collection, s As Slide, sh As Shape
    Dim a As Long, j As Long, count As Long, sx As Single, sy As Single, pitch As Single
    Dim x As Single, y As Single, w As Single, h As Single
    Set sections = GetSections(p)
    sx = p.PageSetup.SlideWidth / 720: sy = p.PageSetup.SlideHeight / 540
    For Each s In p.Slides
        If s.Tags("RUC_TOC_SLIDE") = "1" Then
            count = count + 1
            x = Val(s.Tags("RUC_TOC_X")): y = Val(s.Tags("RUC_TOC_Y"))
            w = Val(s.Tags("RUC_TOC_W")): h = Val(s.Tags("RUC_TOC_H"))
            If w <= 0 Or h <= 0 Then x = 95 * sx: y = 170 * sy: w = 520 * sx: h = 260 * sy
            For j = s.Shapes.Count To 1 Step -1
                If Left(s.Shapes(j).Name, Len(ENTRY)) = ENTRY Then s.Shapes(j).Delete
            Next j
            pitch = MinVal(52 * sy, h / MaxVal(1, sections.Count))
            For a = 1 To sections.Count
                Set sec = sections(a): Set ids = sec("slides")
                Set sh = AddText(s, ENTRY & a, a & "  " & CStr(sec("name")), x, y + (a - 1) * pitch, w, MinVal(34 * sy, pitch), 21 * sx, RGB(17, 17, 17), ppAlignLeft)
                If s.SlideIndex >= ids(1) And s.SlideIndex <= ids(ids.Count) Then
                    sh.TextFrame.TextRange.Font.Color.RGB = RGB(151, 31, 48)
                    sh.TextFrame.TextRange.Font.Bold = msoTrue
                End If
                sh.Tags.Add "RUC_TOC_SECTION_ID", CStr(sec("id"))
                LinkTo sh, p.Slides(ids(1)), CStr(sec("name"))
            Next a
        End If
    Next s
    If count = 0 Then Err.Raise vbObjectError + 103, , U("672A 627E 5230 6A21 677F 76EE 5F55 9875 FF1B 8BF7 4ECE 793A 4F8B 590D 5236 4E00 5F20 76EE 5F55 9875 540E 91CD 8BD5 3002")
End Sub


' Decode UTF-16 hexadecimal code units from ASCII-only source literals.
Private Function U(ByVal codeUnits As String) As String
    Dim units() As String, i As Long, value As Long, result As String
    units = Split(codeUnits, " ")
    For i = LBound(units) To UBound(units)
        value = CLng("&H" & units(i))
        If value > 32767 Then value = value - 65536
        result = result & ChrW(value)
    Next i
    U = result
End Function

Private Function UnicodeMessage(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle, _
    ByVal title As String) As VbMsgBoxResult
    UnicodeMessage = NativeMessageBoxW(NativeActiveWindow(), StrPtr(prompt), StrPtr(title), CLng(buttons))
End Function
