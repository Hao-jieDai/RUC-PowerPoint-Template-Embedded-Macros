Attribute VB_Name = "RUCNavigation"
Option Explicit

' RUC Ecology PowerPoint Template (Embedded Macros), VBA edition 2.2. No network, shell or auto-run.
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
    MsgBox "刷新未完成：" & Err.Description, vbExclamation, "人大环院模板"
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
    MsgBox "1. 在左侧缩略图中，右键节名称，重命名或新增节。" & vbCrLf & _
        "2. 增删、移动页面后，点击【刷新导航和页码】。" & vbCrLf & _
        "3. 目录文字可直接修改，普通刷新会保留文字和位置。" & vbCrLf & _
        "4. 需要按节重新生成目录时，单独点击【从节重新生成目录】。" & vbCrLf & _
        "5. 编辑时保存为 .pptm；放映前刷新并另存为 .pptx。" & vbCrLf & _
        "普通 .pptx 的跳转和已保存的页码无需宏即可放映。" & vbCrLf & vbCrLf & _
        "统一修改页脚：编辑第一页封面中的题目、姓名和学院，再刷新。" & vbCrLf & _
        "完整使用方法请参阅随附的 README.txt。", vbInformation, "人大环院模板使用说明"
End Sub

Private Sub RequireTemplate(p As Presentation)
    If p.Tags("RUC_MACRO_TEMPLATE") <> "1" Then Err.Raise vbObjectError + 101, , "请先切换到人大环院内置宏模板创建的演示文稿。"
    If p.ReadOnly Then Err.Raise vbObjectError + 102, , "当前文稿只读，请另存为可编辑的 .pptm 文件。"
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
        record.Add "报告内容", "name": record.Add indices, "slides": record.Add "fallback", "id"
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
        .Hyperlink.SubAddress = CStr(target.SlideID) & "," & CStr(target.SlideIndex) & "," & Replace(label, ",", "，")
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
                        If clone.Type = msoPlaceholder Then Err.Raise vbObjectError + 104, , "未能将文字占位框转换为普通文本框。"
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
    muted = RGB(169, 201, 181): white = RGB(255, 255, 255)
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
                    LinkTo sh, p.Slides(idx), CStr(sec("name")) & " " & k & "/" & ids.Count & " · 第" & idx & "页"
                Next k
            Next a
            Set sh = AddText(mast, NAV & "AUTHOR", Meta(p, "AUTHOR", "汇报人"), 17 * sx, h - 34 * sy, w * 0.36, 17 * sy, 10.8 * sx, white, ppAlignLeft)
            Set sh = AddText(mast, NAV & "INSTITUTE", Meta(p, "INSTITUTE", "中国人民大学生态环境学院"), w * 0.38, h - 34 * sy, w * 0.596, 17 * sy, 10.8 * sx, white, ppAlignRight)
            Set sh = AddText(mast, NAV & "TITLE", Meta(p, "TITLE", "报告标题"), 17 * sx, h - 15.2 * sy, w * 0.79, 15 * sy, 10.3 * sx, white, ppAlignLeft)
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
            LinkTo sh, s, CStr(sec("name")) & " · 当前页"
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
    p.Tags.Add "RUC_NAV_VERSION", "2.2.0-SEE-VBA"
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
    If MsgBox("将按左侧的节名称重新生成所有目录页。" & vbCrLf & _
        "模板目录条目的手动文字和位置会被替换；你另加的文本框会保留。" & vbCrLf & _
        "只想更新顶部圆点和页码，请取消并选择【刷新导航和页码】。" & vbCrLf & vbCrLf & _
        "确认重新生成目录？", vbYesNo + vbQuestion + vbDefaultButton2, "重新生成目录") <> vbYes Then Exit Sub
    Application.StartNewUndoEntry
    RebuildContents p
    RefreshDeck p
    Exit Sub
Failed:
    MsgBox "目录未能更新：" & Err.Description, vbExclamation, "人大环院模板"
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
                    sh.TextFrame.TextRange.Font.Color.RGB = RGB(12, 95, 51)
                    sh.TextFrame.TextRange.Font.Bold = msoTrue
                End If
                sh.Tags.Add "RUC_TOC_SECTION_ID", CStr(sec("id"))
                LinkTo sh, p.Slides(ids(1)), CStr(sec("name"))
            Next a
        End If
    Next s
    If count = 0 Then Err.Raise vbObjectError + 103, , "未找到模板目录页；请从示例复制一张目录页后重试。"
End Sub
