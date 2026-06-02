' =====================================================================
'  RoadmapExport  -  Excel VBA -> PowerPoint roadmap generator
'  Proof-of-concept companion to roadmap.html.
'
'  WHAT IT DOES
'  Reads the same data the web tool reads -- the named tables LOEs,
'  Themes, Activities, Milestones, Benefits, Risks, plus the Lookups
'  sheet -- and draws a single PowerPoint slide of NATIVE, editable
'  shapes: an LOE gutter (vertical text), Theme header blocks, Activity
'  bars positioned/scaled to the timeline, and Milestone diamonds /
'  Benefit stars / Risk warning-triangles coloured from Lookups.
'
'  This is a deliberate VERTICAL SLICE. By default it renders Level-1
'  activities only (set "Max Level" in the settings sheet to go deeper).
'  It mirrors the pure-data-layer logic of roadmap.html: dash-delimited
'  Activity-ID hierarchy, the Lookups colour system, and the
'  month-proportional date->x mapping.
'
'  HOW TO USE
'  1. Open the data workbook in Excel (Windows desktop Office).
'  2. Alt+F11 -> File -> Import File... -> this .bas.
'  3. Optionally add a sheet named "Roadmap Settings" (see README.md).
'  4. Run the macro RoadmapExport.GenerateRoadmap (Alt+F8).
'  PowerPoint opens with the generated slide. No references needed --
'  PowerPoint and Scripting.Dictionary are bound late via CreateObject.
'
'  NOTE: there is no build step and no external dependency, matching the
'  project's conventions. Tested by inspection; run it on Windows Office
'  to verify visual output.
' =====================================================================
Option Explicit

' ---- PowerPoint / Office enum values (late binding has no constants) ----
Private Const ppLayoutBlank As Long = 12
Private Const msoFalse As Long = 0
Private Const msoTrue As Long = -1
Private Const msoShapeRectangle As Long = 1
Private Const msoShapeRoundedRectangle As Long = 5
Private Const msoShapeDiamond As Long = 4
Private Const msoShapeIsoscelesTriangle As Long = 7
Private Const msoShape5pointStar As Long = 92
Private Const msoTextOrientationHorizontal As Long = 1
Private Const msoTextOrientationUpward As Long = 2
Private Const ppAlignLeft As Long = 1
Private Const ppAlignCenter As Long = 2
Private Const ppAlignRight As Long = 3
Private Const msoAnchorMiddle As Long = 3

' ---- Layout constants, in POINTS (1 inch = 72pt). 16:9 slide = 960x540.
'      Direct analogues of the LAYOUT object in roadmap.html. ----
Private Const SLIDE_W As Single = 960
Private Const SLIDE_H As Single = 540
Private Const LOE_GUTTER As Single = 22      ' rotated LOE-title column
Private Const LABEL_WIDTH As Single = 168    ' Activity-label column (~50 chars)
Private Const MARGIN_RIGHT As Single = 18
Private Const TITLE_H As Single = 30
Private Const AXIS_H As Single = 26
Private Const BODY_PAD_TOP As Single = 4
Private Const BODY_PAD_BOTTOM As Single = 10
Private Const ROW_H As Single = 22           ' depth-1 row height
Private Const BAR_H As Single = 11
Private Const MARKER As Single = 5           ' diamond half-size
Private Const MIN_ROW_H As Single = 12       ' floor before we just overflow

' ---- Brand palette (hex -> RGB long). Mirrors BRAND_COLOURS. ----
Private mBrand As Object          ' name -> RGB long
Private mDefaultPalette() As Long ' default colour cycle (see NEUTRAL_RGB for the fallback)

' ---- Parsed Lookups: key "Table|Column" -> Dictionary(value -> RGB long).
'      A parallel dict holds value ORDER if ever needed. ----
Private mLookupColour As Object   ' "Table|Column" -> Dictionary(valueLower -> RGB long)

' ---- Data-model containers ----
Private Type TActivity
    uid As String
    actId As String
    themeId As String
    title As String
    hasStart As Boolean
    sDate As Date
    hasEnd As Boolean
    eDate As Date
    resourcing As String
    depth As Long
End Type

Private Type TMilestone
    activityUid As String
    title As String
    hasDate As Boolean
    mDate As Date
    confidence As String
    hasBenefit As Boolean
    hasRisk As Boolean
    riskRag As String
End Type

Private mLOEtitle As Object       ' loeId -> title (ordered via mLOEorder)
Private mLOEorder As Collection   ' loeId in table order
Private mThemeTitle As Object     ' themeId -> title
Private mThemesByLoe As Object    ' loeId -> Collection of themeId (ordered)
Private mActs() As TActivity
Private mActCount As Long
Private mActsByTheme As Object    ' themeId -> Collection of index into mActs
Private mMsts() As TMilestone
Private mMstCount As Long
Private mMstByActivity As Object  ' activityUid -> Collection of index into mMsts

' =====================================================================
'  ENTRY POINT
' =====================================================================
Public Sub GenerateRoadmap()
    On Error GoTo Fail

    InitPalette
    ParseLookupsSheet
    BuildModel

    ' --- Settings (sheet "Roadmap Settings" or sensible defaults) ---
    Dim title As String, maxLevel As Long, showMstLabels As Boolean
    Dim dStart As Date, dEnd As Date
    ReadSettings title, dStart, dEnd, maxLevel, showMstLabels

    ' --- Drive PowerPoint via late binding (no reference required) ---
    Dim ppt As Object, pres As Object, sld As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    On Error GoTo Fail
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = msoTrue

    ' 16:9 widescreen so our point maths matches the slide.
    Set pres = ppt.Presentations.Add
    pres.PageSetup.SlideWidth = SLIDE_W
    pres.PageSetup.SlideHeight = SLIDE_H
    Set sld = pres.Slides.Add(1, ppLayoutBlank)

    ' Navy background (#071D49) to match branding.
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.Solid
    sld.Background.Fill.ForeColor.RGB = RGBlong(7, 29, 73)

    DrawSlide sld, title, dStart, dEnd, maxLevel, showMstLabels

    MsgBox "Roadmap slide generated in PowerPoint.", vbInformation
    Exit Sub
Fail:
    MsgBox "Roadmap export failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' =====================================================================
'  DRAWING
' =====================================================================
Private Sub DrawSlide(sld As Object, title As String, dStart As Date, dEnd As Date, _
                      maxLevel As Long, showMstLabels As Boolean)
    Dim minDate As Date, maxDate As Date, totalMonths As Long
    minDate = FirstOfMonth(dStart)
    maxDate = AddMonths(FirstOfMonth(dEnd), 1)        ' exclusive boundary
    totalMonths = MonthsBetween(minDate, maxDate)
    If totalMonths < 1 Then totalMonths = 1

    Dim plotLeft As Single, plotRight As Single, plotW As Single, monthW As Single
    plotLeft = LOE_GUTTER + LABEL_WIDTH
    plotRight = SLIDE_W - MARGIN_RIGHT
    plotW = plotRight - plotLeft
    monthW = plotW / totalMonths

    Dim bodyTop As Single
    bodyTop = TITLE_H + AXIS_H + BODY_PAD_TOP

    ' ---- Title ----
    AddText sld, 0, 0, SLIDE_W, TITLE_H, title, 18, True, RGBlong(255, 255, 255), _
            ppAlignLeft, LOE_GUTTER + 6

    ' ---- Build the visible vertical sequence (theme headers + activity rows),
    '      mirroring roadmap.html: LOEs in order -> Themes -> Activities. ----
    Dim seqType() As String, seqId() As String, seqIdx() As Long, seqLoe() As String
    Dim n As Long: n = 0
    ReDim seqType(1 To 1): ReDim seqId(1 To 1): ReDim seqIdx(1 To 1): ReDim seqLoe(1 To 1)

    Dim li As Long, loeId As String
    For li = 1 To mLOEorder.Count
        loeId = mLOEorder(li)
        If mThemesByLoe.Exists(loeId) Then
            Dim themes As Collection: Set themes = mThemesByLoe(loeId)
            Dim ti As Long, themeId As String
            For ti = 1 To themes.Count
                themeId = themes(ti)
                ' Does this theme have any visible activity? Collect them first.
                Dim actIdxs As Collection: Set actIdxs = New Collection
                If mActsByTheme.Exists(themeId) Then
                    Dim c As Collection: Set c = mActsByTheme(themeId)
                    Dim ai As Long
                    For ai = 1 To c.Count
                        If mActs(c(ai)).depth <= maxLevel Then actIdxs.Add c(ai)
                    Next ai
                End If
                If actIdxs.Count > 0 Then
                    n = n + 1: GrowSeq seqType, seqId, seqIdx, seqLoe, n
                    seqType(n) = "theme": seqId(n) = themeId: seqLoe(n) = loeId
                    For ai = 1 To actIdxs.Count
                        n = n + 1: GrowSeq seqType, seqId, seqIdx, seqLoe, n
                        seqType(n) = "act": seqIdx(n) = actIdxs(ai): seqLoe(n) = loeId
                    Next ai
                End If
            Next ti
        End If
    Next li

    If n = 0 Then
        AddText sld, plotLeft, bodyTop, plotW, 40, "No activities to display for the chosen settings.", _
                14, False, RGBlong(255, 255, 255), ppAlignLeft, 0
        DrawAxis sld, TITLE_H, minDate, totalMonths, monthW, plotLeft
        Exit Sub
    End If

    ' ---- Auto-fit vertical scale: shrink rows so everything fits the slide
    '      (PowerPoint can't scroll), with a floor before we just overflow. ----
    Dim themeCount As Long, actCount As Long, i As Long
    For i = 1 To n
        If seqType(i) = "theme" Then themeCount = themeCount + 1 Else actCount = actCount + 1
    Next i
    Dim themeHdrH As Single: themeHdrH = 18
    Dim availH As Single: availH = SLIDE_H - bodyTop - BODY_PAD_BOTTOM
    Dim natural As Single: natural = themeCount * themeHdrH + actCount * ROW_H
    Dim sc As Single: sc = 1
    If natural > availH Then sc = availH / natural
    Dim rowH As Single, thH As Single
    rowH = ROW_H * sc: If rowH < MIN_ROW_H Then rowH = MIN_ROW_H
    thH = themeHdrH * sc: If thH < 14 Then thH = 14
    Dim barH As Single: barH = BAR_H * sc: If barH < 7 Then barH = 7

    ' ---- Month gridlines + axis header ----
    DrawGrid sld, bodyTop, availH, minDate, totalMonths, monthW, plotLeft
    DrawAxis sld, TITLE_H, minDate, totalMonths, monthW, plotLeft

    ' ---- Lay out rows, tracking LOE band extents for the vertical gutter text.
    Dim y As Single: y = bodyTop
    Dim bandLoe As String: bandLoe = ""
    Dim bandY0 As Single: bandY0 = y
    For i = 1 To n
        ' LOE band boundary -> emit the vertical title for the band just ended.
        If seqLoe(i) <> bandLoe Then
            If bandLoe <> "" Then DrawLoeGutter sld, bandLoe, bandY0, y
            bandLoe = seqLoe(i): bandY0 = y
        End If

        If seqType(i) = "theme" Then
            DrawThemeHeader sld, seqId(i), plotLeft - LABEL_WIDTH, y, LABEL_WIDTH + (monthW * totalMonths), thH
            y = y + thH
        Else
            DrawActivityRow sld, seqIdx(i), y, rowH, barH, minDate, maxDate, monthW, plotLeft, _
                            LOE_GUTTER, LABEL_WIDTH, showMstLabels
            y = y + rowH
        End If
    Next i
    If bandLoe <> "" Then DrawLoeGutter sld, bandLoe, bandY0, y
End Sub

Private Sub DrawAxis(sld As Object, top As Single, minDate As Date, totalMonths As Long, _
                     monthW As Single, plotLeft As Single)
    Dim i As Long, d As Date, x As Single
    Dim lastYear As Long: lastYear = 0
    For i = 0 To totalMonths - 1
        d = AddMonths(minDate, i)
        x = plotLeft + i * monthW
        ' Month label (e.g. "Jan")
        AddText sld, x, top + 11, monthW, 13, MonthAbbr(Month(d)), 8, False, _
                RGBlong(200, 215, 240), ppAlignCenter, 0
        ' Year label at each January or the first column.
        If Year(d) <> lastYear Then
            AddText sld, x, top, monthW * 3, 12, CStr(Year(d)), 10, True, _
                    RGBlong(255, 255, 255), ppAlignLeft, 2
            lastYear = Year(d)
        End If
    Next i
End Sub

Private Sub DrawGrid(sld As Object, bodyTop As Single, bodyH As Single, minDate As Date, _
                     totalMonths As Long, monthW As Single, plotLeft As Single)
    Dim i As Long, x As Single, ln As Object
    For i = 0 To totalMonths
        x = plotLeft + i * monthW
        Set ln = sld.Shapes.AddLine(x, bodyTop, x, bodyTop + bodyH)
        ln.Line.ForeColor.RGB = RGBlong(40, 64, 112)
        ln.Line.Weight = 0.5
    Next i
End Sub

Private Sub DrawLoeGutter(sld As Object, loeId As String, y0 As Single, y1 As Single)
    Dim ttl As String
    ttl = IIf(mLOEtitle.Exists(loeId), mLOEtitle(loeId), loeId)
    Dim h As Single: h = y1 - y0
    If h < 8 Then Exit Sub
    Dim box As Object
    Set box = sld.Shapes.AddTextbox(msoTextOrientationUpward, 2, y0, LOE_GUTTER - 4, h)
    With box.TextFrame
        .TextRange.text = ttl
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Color.RGB = RGBlong(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
End Sub

Private Sub DrawThemeHeader(sld As Object, themeId As String, left As Single, top As Single, _
                            width As Single, height As Single)
    Dim ttl As String
    ttl = IIf(mThemeTitle.Exists(themeId), mThemeTitle(themeId), themeId)
    Dim sh As Object
    Set sh = sld.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
    sh.Fill.ForeColor.RGB = RGBlong(13, 41, 92)
    sh.Line.Visible = msoFalse
    With sh.TextFrame
        .TextRange.text = "  " & ttl
        .TextRange.Font.Size = IIf(height < 16, 9, 11)
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Color.RGB = RGBlong(89, 172, 218)   ' accent blue
        .TextRange.ParagraphFormat.Alignment = ppAlignLeft
        .VerticalAnchor = msoAnchorMiddle
        .MarginTop = 0: .MarginBottom = 0
    End With
End Sub

Private Sub DrawActivityRow(sld As Object, idx As Long, top As Single, rowH As Single, barH As Single, _
                            minDate As Date, maxDate As Date, monthW As Single, plotLeft As Single, _
                            loeGutter As Single, labelWidth As Single, showMstLabels As Boolean)
    Dim a As TActivity: a = mActs(idx)
    Dim cy As Single: cy = top + rowH / 2

    ' ---- Label (indented by depth), left of the plot area ----
    Dim indent As Single: indent = (a.depth - 1) * 10
    AddText sld, loeGutter + 4 + indent, top, labelWidth - 6 - indent, rowH, a.title, _
            IIf(rowH < 16, 8, 9.5), (a.depth = 1), RGBlong(255, 255, 255), ppAlignLeft, 0

    ' ---- Activity bar, clamped to the visible window, coloured by Resourcing ----
    If a.hasStart And a.hasEnd Then
        Dim s As Date, e As Date
        s = ClampDate(a.sDate, minDate, maxDate)
        e = ClampDate(a.eDate, minDate, maxDate)
        Dim x1 As Single, x2 As Single
        x1 = XForDate(s, minDate, monthW, plotLeft)
        x2 = XForDate(e, minDate, monthW, plotLeft)
        If x2 < x1 + 2 Then x2 = x1 + 2
        Dim bar As Object
        Set bar = sld.Shapes.AddShape(msoShapeRoundedRectangle, x1, cy - barH / 2, x2 - x1, barH)
        bar.Fill.ForeColor.RGB = ColourFor("Activities", "Resourcing", a.resourcing)
        bar.Line.Visible = msoFalse
    End If

    ' ---- Milestones: diamond, or star if it carries a Benefit; warning
    '      triangle adjacent if it carries a Risk. ----
    If Not mMstByActivity.Exists(a.uid) Then Exit Sub
    Dim coll As Collection: Set coll = mMstByActivity(a.uid)
    Dim k As Long
    For k = 1 To coll.Count
        Dim m As TMilestone: m = mMsts(coll(k))
        If Not m.hasDate Then GoTo NextM
        Dim cd As Date: cd = ClampDate(m.mDate, minDate, maxDate)
        Dim cx As Single: cx = XForDate(cd, minDate, monthW, plotLeft)
        Dim mc As Long: mc = ColourFor("Milestones", "Delivery Confidence", m.confidence)
        Dim sz As Single: sz = MARKER
        If barH < 9 Then sz = MARKER * 0.85

        Dim mk As Object
        If m.hasBenefit Then
            ' 5-point star sized a touch larger than the diamond.
            Set mk = sld.Shapes.AddShape(msoShape5pointStar, cx - sz * 1.3, cy - sz * 1.3, sz * 2.6, sz * 2.6)
        Else
            Set mk = sld.Shapes.AddShape(msoShapeDiamond, cx - sz, cy - sz, sz * 2, sz * 2)
        End If
        mk.Fill.ForeColor.RGB = mc
        mk.Line.ForeColor.RGB = RGBlong(7, 29, 73)
        mk.Line.Weight = 0.75

        If m.hasRisk Then
            ' Small warning triangle ~45deg down-and-right of the marker.
            Dim rr As Single: rr = sz * 0.9
            Dim rx As Single: rx = cx + sz * 0.9
            Dim ry As Single: ry = cy + sz * 0.5
            Dim tri As Object
            Set tri = sld.Shapes.AddShape(msoShapeIsoscelesTriangle, rx - rr, ry - rr, rr * 2, rr * 2)
            tri.Fill.ForeColor.RGB = ColourFor("Risks", "RAG", m.riskRag)
            tri.Line.ForeColor.RGB = RGBlong(255, 255, 255)
            tri.Line.Weight = 0.5
        End If

        If showMstLabels Then
            AddText sld, cx - 40, cy + sz + 1, 80, 12, m.title, 7, False, _
                    RGBlong(220, 230, 245), ppAlignCenter, 0
        End If
NextM:
    Next k
End Sub

' =====================================================================
'  MODEL BUILDING  (mirrors assembleModel / buildHierarchy in roadmap.html)
' =====================================================================
Private Sub BuildModel()
    Set mLOEtitle = NewDict
    Set mLOEorder = New Collection
    Set mThemeTitle = NewDict
    Set mThemesByLoe = NewDict
    Set mActsByTheme = NewDict
    Set mMstByActivity = NewDict

    ' ---- LOEs ----
    Dim lo As Object: Set lo = FindTable("LOEs")
    If Not lo Is Nothing Then
        Dim cId As Long, cTitle As Long
        cId = ColIdx(lo, Array("loe id"))
        cTitle = ColIdx(lo, Array("title", "name"))
        Dim r As Long, v As String, t As String
        For r = 1 To RowCount(lo)
            v = CellStr(lo, r, cId)
            If v <> "" Then
                t = CellStr(lo, r, cTitle): If t = "" Then t = v
                If Not mLOEtitle.Exists(v) Then
                    mLOEtitle.Add v, t
                    mLOEorder.Add v
                End If
            End If
        Next r
    End If

    ' ---- Themes ----
    Set lo = FindTable("Themes")
    If Not lo Is Nothing Then
        Dim cLoe As Long, cTid As Long, cTt As Long
        cLoe = ColIdx(lo, Array("loe id"))
        cTid = ColIdx(lo, Array("theme id"))
        cTt = ColIdx(lo, Array("title", "name"))
        For r = 1 To RowCount(lo)
            Dim loeId As String, themeId As String, ttl As String
            loeId = CellStr(lo, r, cLoe)
            themeId = CellStr(lo, r, cTid)
            If loeId <> "" And themeId <> "" Then     ' skip trailing blank rows
                ttl = CellStr(lo, r, cTt): If ttl = "" Then ttl = themeId
                If Not mThemeTitle.Exists(themeId) Then mThemeTitle.Add themeId, ttl
                If Not mThemesByLoe.Exists(loeId) Then mThemesByLoe.Add loeId, New Collection
                mThemesByLoe(loeId).Add themeId
            End If
        Next r
    End If

    ' ---- Activities ----
    mActCount = 0: ReDim mActs(1 To 1)
    Dim byActId As Object: Set byActId = NewDict   ' actId -> uid (for hierarchy depth)
    Set lo = FindTable("Activities")
    If Not lo Is Nothing Then
        Dim cTh As Long, cAct As Long, cUid As Long, cAt As Long, cS As Long, cE As Long, cRes As Long
        cTh = ColIdx(lo, Array("theme id"))
        cAct = ColIdx(lo, Array("activity id"))
        cUid = ColIdx(lo, Array("uniqueact id", "unique act id", "uniqueactid"))
        cAt = ColIdx(lo, Array("title", "name"))
        cS = ColIdx(lo, Array("start date", "start"))
        cE = ColIdx(lo, Array("end date", "end"))
        cRes = ColIdx(lo, Array("resourcing"))
        For r = 1 To RowCount(lo)
            Dim uid As String, thId As String
            uid = CellStr(lo, r, cUid)
            thId = CellStr(lo, r, cTh)
            If thId <> "" And uid <> "" Then       ' parent col present => real row
                mActCount = mActCount + 1
                ReDim Preserve mActs(1 To mActCount)
                With mActs(mActCount)
                    .uid = uid
                    .actId = CellStr(lo, r, cAct)
                    .themeId = thId
                    .title = CellStr(lo, r, cAt): If .title = "" Then .title = uid
                    .resourcing = CellStr(lo, r, cRes)
                    GetCellDate lo, r, cS, .hasStart, .sDate
                    GetCellDate lo, r, cE, .hasEnd, .eDate
                    ' depth = number of dash-delimited segments in Activity ID.
                    .depth = SegmentCount(.actId)
                End With
                If Not mActsByTheme.Exists(thId) Then mActsByTheme.Add thId, New Collection
                mActsByTheme(thId).Add mActCount
            End If
        Next r
    End If

    ' ---- Milestones (+ benefit/risk flags) ----
    Dim benByMst As Object: Set benByMst = BuildBenefitSet()
    Dim riskByMst As Object: Set riskByMst = BuildRiskMap()

    mMstCount = 0: ReDim mMsts(1 To 1)
    Set lo = FindTable("Milestones")
    If Not lo Is Nothing Then
        Dim cMa As Long, cMu As Long, cMt As Long, cMd As Long, cMc As Long
        cMa = ColIdx(lo, Array("activity id"))
        cMu = ColIdx(lo, Array("uniquemst id", "unique mst id", "uniquemstid"))
        cMt = ColIdx(lo, Array("milestone title", "title"))
        cMd = ColIdx(lo, Array("date"))
        cMc = ColIdx(lo, Array("delivery confidence", "confidence"))
        For r = 1 To RowCount(lo)
            Dim aUid As String, mUid As String
            aUid = CellStr(lo, r, cMa)
            mUid = CellStr(lo, r, cMu)
            If aUid <> "" Then
                mMstCount = mMstCount + 1
                ReDim Preserve mMsts(1 To mMstCount)
                With mMsts(mMstCount)
                    .activityUid = aUid
                    .title = CellStr(lo, r, cMt): If .title = "" Then .title = mUid
                    .confidence = CellStr(lo, r, cMc)
                    GetCellDate lo, r, cMd, .hasDate, .mDate
                    .hasBenefit = benByMst.Exists(LCase$(mUid))
                    .hasRisk = riskByMst.Exists(LCase$(mUid))
                    If .hasRisk Then .riskRag = riskByMst(LCase$(mUid))
                End With
                If Not mMstByActivity.Exists(aUid) Then mMstByActivity.Add aUid, New Collection
                mMstByActivity(aUid).Add mMstCount
            End If
        Next r
    End If
End Sub

' Set of milestone UIDs (lower-cased) that have at least one real benefit.
Private Function BuildBenefitSet() As Object
    Dim d As Object: Set d = NewDict
    Dim lo As Object: Set lo = FindTable("Benefits")
    If lo Is Nothing Then Set BuildBenefitSet = d: Exit Function
    Dim cM As Long, cT As Long, cC As Long, cB As Long, cI As Long
    cM = ColIdx(lo, Array("milestone id"))
    cT = ColIdx(lo, Array("benefit title", "title"))
    cC = ColIdx(lo, Array("category"))
    cB = ColIdx(lo, Array("beneficiary"))
    cI = ColIdx(lo, Array("impact"))
    Dim r As Long, mid As String
    For r = 1 To RowCount(lo)
        mid = CellStr(lo, r, cM)
        ' Drop trailing blank rows whose only value is the auto-filled key.
        If mid <> "" And (CellStr(lo, r, cT) <> "" Or CellStr(lo, r, cC) <> "" _
            Or CellStr(lo, r, cB) <> "" Or CellStr(lo, r, cI) <> "") Then
            If Not d.Exists(LCase$(mid)) Then d.Add LCase$(mid), True
        End If
    Next r
    Set BuildBenefitSet = d
End Function

' Map milestone UID (lower-cased) -> first Risk's RAG value.
Private Function BuildRiskMap() As Object
    Dim d As Object: Set d = NewDict
    Dim lo As Object: Set lo = FindTable("Risks")
    If lo Is Nothing Then Set BuildRiskMap = d: Exit Function
    Dim cM As Long, cT As Long, cO As Long, cR As Long
    cM = ColIdx(lo, Array("milestone id"))
    cT = ColIdx(lo, Array("risk title", "title"))
    cO = ColIdx(lo, Array("risk owner", "owner"))
    cR = ColIdx(lo, Array("rag"))
    Dim r As Long, mid As String
    For r = 1 To RowCount(lo)
        mid = CellStr(lo, r, cM)
        If mid <> "" And (CellStr(lo, r, cT) <> "" Or CellStr(lo, r, cO) <> "" Or CellStr(lo, r, cR) <> "") Then
            If Not d.Exists(LCase$(mid)) Then d.Add LCase$(mid), CellStr(lo, r, cR)
        End If
    Next r
    Set BuildRiskMap = d
End Function

' =====================================================================
'  LOOKUPS  (mirrors parseLookups: table anchored at F1 = column 6)
' =====================================================================
Private Sub ParseLookupsSheet()
    Set mLookupColour = NewDict
    Dim ws As Object
    Set ws = FindSheet(Array("lookups", "lookup"))
    If ws Is Nothing Then Exit Sub

    Dim used As Object: Set used = ws.UsedRange
    Dim maxCol As Long, maxRow As Long
    maxCol = used.Column + used.Columns.Count - 1
    maxRow = used.Row + used.Rows.Count - 1

    Dim c As Long
    For c = 6 To maxCol                      ' column F = 6
        Dim tbl As String, col As String
        tbl = Trim$(SheetStr(ws, 1, c))
        col = Trim$(SheetStr(ws, 2, c))
        If tbl <> "" And col <> "" Then
            Dim spec As String: spec = Trim$(SheetStr(ws, 3, c))
            Dim colourNames() As String, nNames As Long
            nNames = SplitSemis(spec, colourNames)

            Dim valDict As Object: Set valDict = NewDict   ' valueLower -> RGB long
            Dim r As Long, idx As Long, valTxt As String
            idx = 0
            For r = 4 To maxRow
                valTxt = Trim$(SheetStr(ws, r, c))
                If valTxt = "" Then Exit For                ' list ends at first blank
                Dim clr As Long
                If nNames > 0 Then
                    clr = ResolveColour(colourNames(idx Mod nNames))
                Else
                    clr = mDefaultPalette(idx Mod (UBound(mDefaultPalette) + 1))
                End If
                If Not valDict.Exists(LCase$(valTxt)) Then valDict.Add LCase$(valTxt), clr
                idx = idx + 1
            Next r
            mLookupColour(tbl & "|" & col) = valDict
        End If
    Next c
End Sub

' Colour for a list-select value (case-insensitive), neutral fallback.
Private Function ColourFor(tbl As String, col As String, value As String) As Long
    ColourFor = NEUTRAL_RGB()
    If value = "" Then Exit Function
    Dim key As String: key = tbl & "|" & col
    If Not mLookupColour.Exists(key) Then Exit Function
    Dim d As Object: Set d = mLookupColour(key)
    If d.Exists(LCase$(value)) Then ColourFor = d(LCase$(value))
End Function

' =====================================================================
'  PALETTE
' =====================================================================
Private Sub InitPalette()
    Set mBrand = NewDict
    mBrand.Add "blue", RGBlong(89, 172, 218)
    mBrand.Add "orange", RGBlong(224, 141, 72)
    mBrand.Add "amber", RGBlong(224, 141, 72)
    mBrand.Add "yellow", RGBlong(244, 199, 72)
    mBrand.Add "gold", RGBlong(244, 199, 72)
    mBrand.Add "pink", RGBlong(207, 141, 183)
    mBrand.Add "green", RGBlong(121, 198, 138)
    mBrand.Add "purple", RGBlong(164, 150, 215)
    mBrand.Add "violet", RGBlong(164, 150, 215)
    mBrand.Add "red", RGBlong(212, 31, 65)
    mBrand.Add "navy", RGBlong(7, 29, 73)
    ' A few literal basics so "grey"/"white"/"black" still work.
    mBrand.Add "grey", RGBlong(128, 128, 128)
    mBrand.Add "gray", RGBlong(128, 128, 128)
    mBrand.Add "white", RGBlong(255, 255, 255)
    mBrand.Add "black", RGBlong(0, 0, 0)

    ReDim mDefaultPalette(0 To 5)
    mDefaultPalette(0) = RGBlong(89, 172, 218)
    mDefaultPalette(1) = RGBlong(121, 198, 138)
    mDefaultPalette(2) = RGBlong(244, 199, 72)
    mDefaultPalette(3) = RGBlong(224, 141, 72)
    mDefaultPalette(4) = RGBlong(207, 141, 183)
    mDefaultPalette(5) = RGBlong(164, 150, 215)
End Sub

Private Function ResolveColour(name As String) As Long
    Dim k As String: k = LCase$(Trim$(name))
    If k = "" Then ResolveColour = NEUTRAL_RGB(): Exit Function
    If mBrand.Exists(k) Then ResolveColour = mBrand(k) Else ResolveColour = NEUTRAL_RGB()
End Function

Private Function NEUTRAL_RGB() As Long
    NEUTRAL_RGB = RGBlong(138, 160, 204)
End Function

' =====================================================================
'  SETTINGS
' =====================================================================
Private Sub ReadSettings(ByRef title As String, ByRef dStart As Date, ByRef dEnd As Date, _
                         ByRef maxLevel As Long, ByRef showMstLabels As Boolean)
    ' Defaults: current month -> +12 months; Level 1 only; labels off.
    Dim now As Date: now = Date
    dStart = DateSerial(Year(now), Month(now), 1)
    dEnd = DateSerial(Year(now), Month(now) + 12, 1)
    title = "Programme Roadmap"
    maxLevel = 1
    showMstLabels = False

    Dim ws As Object
    Set ws = FindSheet(Array("roadmap settings", "settings"))
    If ws Is Nothing Then GoTo ClampToData

    ' Labels in column A, values in column B.
    Dim r As Long, lbl As String, val As Variant
    For r = 1 To 50
        lbl = LCase$(Trim$(CStr(GetSheetVal(ws, r, 1))))
        val = GetSheetVal(ws, r, 2)
        Select Case lbl
            Case "title": If Trim$(CStr(val)) <> "" Then title = CStr(val)
            Case "start date": If IsDate(val) Then dStart = CDate(val)
            Case "end date": If IsDate(val) Then dEnd = CDate(val)
            Case "max level": If IsNumeric(val) Then maxLevel = CLng(val)
            Case "show milestone labels": showMstLabels = AsBool(val)
        End Select
    Next r

ClampToData:
    If maxLevel < 1 Then maxLevel = 1
    ' Clamp the window to the data extent (mirrors clampMonth).
    Dim mn As Date, mx As Date, has As Boolean
    DataDateExtent mn, mx, has
    If has Then
        If dStart < FirstOfMonth(mn) Then dStart = FirstOfMonth(mn)
        If dEnd > FirstOfMonth(mx) Then dEnd = FirstOfMonth(mx)
        If dEnd < dStart Then dEnd = dStart
    End If
End Sub

Private Sub DataDateExtent(ByRef mn As Date, ByRef mx As Date, ByRef has As Boolean)
    has = False
    Dim i As Long
    For i = 1 To mActCount
        If mActs(i).hasStart Then Span mActs(i).sDate, mn, mx, has
        If mActs(i).hasEnd Then Span mActs(i).eDate, mn, mx, has
    Next i
    For i = 1 To mMstCount
        If mMsts(i).hasDate Then Span mMsts(i).mDate, mn, mx, has
    Next i
End Sub

Private Sub Span(d As Date, ByRef mn As Date, ByRef mx As Date, ByRef has As Boolean)
    If Not has Then mn = d: mx = d: has = True: Exit Sub
    If d < mn Then mn = d
    If d > mx Then mx = d
End Sub

' =====================================================================
'  WORKBOOK READING HELPERS
' =====================================================================
' Find a named ListObject (table) anywhere in the workbook, case-insensitive.
Private Function FindTable(name As String) As Object
    Dim ws As Object, lo As Object
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.name) = LCase$(name) Then Set FindTable = lo: Exit Function
        Next lo
    Next ws
    Set FindTable = Nothing
End Function

Private Function FindSheet(names As Variant) As Object
    Dim ws As Object, i As Long
    For i = LBound(names) To UBound(names)
        For Each ws In ThisWorkbook.Worksheets
            If LCase$(Trim$(ws.name)) = LCase$(names(i)) Then Set FindSheet = ws: Exit Function
        Next ws
    Next i
    Set FindSheet = Nothing
End Function

' First table column (1-based offset) whose header matches any alias; 0 if none.
Private Function ColIdx(lo As Object, aliases As Variant) As Long
    Dim c As Long, h As String, i As Long
    For c = 1 To lo.ListColumns.Count
        h = LCase$(Trim$(CStr(lo.HeaderRowRange.Cells(1, c).value)))
        For i = LBound(aliases) To UBound(aliases)
            If h = aliases(i) Then ColIdx = c: Exit Function
        Next i
    Next c
    ColIdx = 0
End Function

Private Function RowCount(lo As Object) As Long
    If lo.DataBodyRange Is Nothing Then RowCount = 0 Else RowCount = lo.DataBodyRange.Rows.Count
End Function

Private Function CellStr(lo As Object, r As Long, c As Long) As String
    If c < 1 Then CellStr = "": Exit Function
    Dim v As Variant: v = lo.DataBodyRange.Cells(r, c).value
    If IsError(v) Or IsEmpty(v) Then CellStr = "" Else CellStr = Trim$(CStr(v))
End Function

Private Sub GetCellDate(lo As Object, r As Long, c As Long, ByRef has As Boolean, ByRef d As Date)
    has = False
    If c < 1 Then Exit Sub
    Dim v As Variant: v = lo.DataBodyRange.Cells(r, c).value
    If IsDate(v) Then
        d = DateSerial(Year(v), Month(v), Day(v))   ' local midnight, drop time
        has = True
    End If
End Sub

Private Function SheetStr(ws As Object, r As Long, c As Long) As String
    Dim v As Variant: v = ws.Cells(r, c).value
    If IsError(v) Or IsEmpty(v) Then SheetStr = "" Else SheetStr = CStr(v)
End Function

Private Function GetSheetVal(ws As Object, r As Long, c As Long) As Variant
    GetSheetVal = ws.Cells(r, c).value
End Function

' =====================================================================
'  POWERPOINT DRAWING HELPERS
' =====================================================================
' Add a (left-aligned by default) text box with no fill/line.
Private Sub AddText(sld As Object, left As Single, top As Single, width As Single, height As Single, _
                    text As String, fontSize As Single, bold As Boolean, colour As Long, _
                    align As Long, leftMargin As Single)
    Dim tb As Object
    Set tb = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
    With tb.TextFrame
        .TextRange.text = text
        .TextRange.Font.Size = fontSize
        .TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
        .TextRange.Font.Color.RGB = colour
        .TextRange.ParagraphFormat.Alignment = align
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .MarginLeft = leftMargin: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
End Sub

' =====================================================================
'  DATE / GEOMETRY HELPERS  (mirror roadmap.html)
' =====================================================================
Private Function FirstOfMonth(d As Date) As Date
    FirstOfMonth = DateSerial(Year(d), Month(d), 1)
End Function
Private Function AddMonths(d As Date, n As Long) As Date
    AddMonths = DateSerial(Year(d), Month(d) + n, 1)
End Function
Private Function MonthsBetween(a As Date, b As Date) As Long
    MonthsBetween = (Year(b) - Year(a)) * 12 + (Month(b) - Month(a))
End Function
Private Function DaysInMonth(y As Long, m As Long) As Long
    DaysInMonth = Day(DateSerial(y, m + 1, 0))
End Function
Private Function ClampDate(d As Date, lo As Date, hi As Date) As Date
    If d < lo Then ClampDate = lo: Exit Function
    If d >= hi Then ClampDate = DateSerial(Year(hi), Month(hi), 0): Exit Function  ' last day before boundary
    ClampDate = d
End Function
Private Function XForDate(d As Date, minDate As Date, monthW As Single, plotLeft As Single) As Single
    Dim ms As Date: ms = FirstOfMonth(d)
    Dim frac As Single
    frac = (Day(d) - 1) / DaysInMonth(Year(d), Month(d))
    XForDate = plotLeft + (MonthsBetween(minDate, ms) + frac) * monthW
End Function
Private Function MonthAbbr(m As Long) As String
    MonthAbbr = Format$(DateSerial(2000, m, 1), "mmm")
End Function

' =====================================================================
'  MISC HELPERS
' =====================================================================
Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

' RGB long in the order Office expects (red + green*256 + blue*65536).
Private Function RGBlong(r As Long, g As Long, b As Long) As Long
    RGBlong = RGB(r, g, b)
End Function

Private Function SegmentCount(actId As String) As Long
    If actId = "" Then SegmentCount = 1: Exit Function
    SegmentCount = UBound(Split(actId, "-")) + 1
End Function

Private Function SplitSemis(spec As String, ByRef out() As String) As Long
    If Trim$(spec) = "" Then SplitSemis = 0: Exit Function
    Dim parts() As String: parts = Split(spec, ";")
    Dim i As Long, k As Long: k = 0
    ReDim out(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Trim$(parts(i)) <> "" Then out(k) = Trim$(parts(i)): k = k + 1
    Next i
    If k = 0 Then SplitSemis = 0 Else ReDim Preserve out(0 To k - 1): SplitSemis = k
End Function

Private Function AsBool(v As Variant) As Boolean
    Dim s As String: s = LCase$(Trim$(CStr(v)))
    AsBool = (s = "true" Or s = "yes" Or s = "y" Or s = "1" Or v = True)
End Function

' Grow the parallel sequence arrays to at least size n.
Private Sub GrowSeq(ByRef t() As String, ByRef id() As String, ByRef ix() As Long, ByRef lo() As String, n As Long)
    If n > UBound(t) Then
        ReDim Preserve t(1 To n + 16)
        ReDim Preserve id(1 To n + 16)
        ReDim Preserve ix(1 To n + 16)
        ReDim Preserve lo(1 To n + 16)
    End If
End Sub
