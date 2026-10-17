Attribute VB_Name = "SizePlugin"
'==============================================================================
'  SIZE - CorelDRAW uchun o'lchamlarni avtomatik ko'rsatuvchi plagin (VBA)
'
'  Makroslar:
'    Size          - tanlangan obyektlar / harflarning enini, bo'yini va
'                    ular orasidagi bo'shliqlarni ish sahifasida ko'rsatadi
'    SizeAsk       - avval raqamlar shrift o'lchamini so'raydi, keyin o'lchaydi
'    SizeSettings  - barcha sozlamalar (shrift, rang, ish rejimlari)
'    SizeClear     - joriy sahifadagi barcha Size o'lchamlarini o'chiradi
'
'  Qoidalar:
'    * barcha o'lchamlar santimetrda, nuqtadan keyin bitta raqam: 10.6 cm
'    * oraliqlar obyektlarning bir-biriga qaragan chetlaridan o'lchanadi
'    * o'lchamlar yordamchi chiziqlar va strelkalar bilan chiziladi
'    * yozuvlar obyektlar va boshqa o'lcham yozuvlari bilan ustma-ust tushmaydi
'    * natija alohida "Size - o'lchamlar" qatlamiga, bitta guruh qilib qo'yiladi
'      (Ctrl+Z bitta qadamda bekor qiladi)
'
'  CorelDRAW X7 va undan yangi versiyalar uchun.
'==============================================================================
Option Explicit

Private Const APP_KEY As String = "SizeCorel"
Private Const SET_KEY As String = "Settings"
Private Const LAYER_NAME As String = "Size - o'lchamlar"
Private Const MIN_STEPS As Long = 40
Private Const MAX_STEPS As Long = 600
Private Const BIG As Double = 1E+30
Private Const MIN_SIZE As Double = 0.0005

' To'g'ri to'rtburchak (cm): chap, past, o'ng, yuqori
Private Type TBox
    L As Double
    B As Double
    R As Double
    T As Double
End Type

' Bitta o'lcham. "A" - o'lchanayotgan o'q (gorizontal o'lchamda X, vertikalda Y),
' "P" - unga perpendikulyar o'q.
Private Type TDim
    Orient As Long          ' 0 = gorizontal (eni, gorizontal oraliq), 1 = vertikal
    A1 As Double            ' o'lcham boshi
    A2 As Double            ' o'lcham oxiri
    P1Lo As Double          ' A1 uchidagi obyektning P o'qi bo'yicha chegaralari
    P1Hi As Double
    P2Lo As Double          ' A2 uchidagi obyektning P o'qi bo'yicha chegaralari
    P2Hi As Double
    HasInside As Boolean    ' o'lchamni oraliqning ichiga qo'yish mumkinmi
    InLo As Double
    InHi As Double
    CostLow As Double       ' past / chap tomonga qo'yish "narxi"
    CostHigh As Double      ' yuqori / o'ng tomonga qo'yish "narxi"
    CostInside As Double    ' oraliq ichiga qo'yish "narxi"
    AllowSlide As Boolean   ' yozuvni chiziq uchidan tashqariga surish mumkinmi
    Value As Double         ' qiymat, cm
End Type

' O'lchamni joylashtirish varianti
Private Type TCand
    Side As Long            ' 0 = past/chap, 1 = yuqori/o'ng, 2 = oraliq ichida
    Slide As Long           ' 0 = yozuv o'rtada, 1 = A2 uchidan keyin, 2 = A1 uchidan oldin
    Pos As Double           ' o'lcham chizig'ining P koordinatasi
    LineLo As Double        ' o'lcham chizig'i A o'qi bo'yicha (strelkalar bilan)
    LineHi As Double
    ArrowsOut As Boolean    ' qisqa o'lchamda strelkalar tashqaridan qaraydi
    HasExt As Boolean       ' yordamchi chiziqlar bormi
    E1From As Double
    E1To As Double
    E2From As Double
    E2To As Double
    Lab As TBox             ' yozuv joyi
End Type

' --- Sozlamalar ---
Private gFontPt As Double           ' 0 = avtomatik
Private gFontName As String
Private gBold As Boolean
Private gColR As Long
Private gColG As Long
Private gColB As Long
Private gSplitText As Boolean       ' matn obyektini harflarga ajratib o'lchash
Private gShowOverall As Boolean     ' butun tanlovning umumiy o'lchami
Private gUseOutline As Boolean      ' kontur qalinligini hisobga olish
Private gRotateVertical As Boolean  ' bo'yi yozuvini 90 gradusga burish

' --- Geometriya (cm) ---
Private gPt As Double
Private gEm As Double
Private gLineW As Double
Private gExtW As Double
Private gArrowL As Double
Private gArrowW As Double
Private gGapExt As Double
Private gOver As Double
Private gOff0 As Double
Private gStep As Double
Private gLabelGap As Double
Private gPad As Double
Private gColor As Color

' --- O'lchanadigan obyektlar ---
Private gItem() As TBox
Private gItemN As Long

' --- Joylashtirilgan elementlar (to'qnashuvni tekshirish uchun) ---
Private gEl() As TBox
Private gElKind() As Long           ' 1 = yozuv, 2 = o'lcham chizig'i, 3 = yordamchi chiziq
Private gElOr() As Long             ' yordamchi chiziq: 0 = gorizontal, 1 = vertikal
Private gElC() As Double            ' yordamchi chiziq koordinatasi
Private gElN As Long

' --- O'lchamlar ---
Private gDim() As TDim
Private gDimN As Long
Private gLab() As Shape
Private gLabW() As Double
Private gLabH() As Double

'==============================================================================
'  OCHIQ MAKROSLAR
'==============================================================================

' Tanlangan obyektlarni o'lchash
Public Sub Size()
    RunSize False
End Sub

' Shrift o'lchamini so'rab, keyin o'lchash
Public Sub SizeAsk()
    RunSize True
End Sub

' Joriy sahifadagi barcha Size o'lchamlarini o'chirish
Public Sub SizeClear()
    Dim lyr As Layer
    If ActiveDocument Is Nothing Then Exit Sub
    Set lyr = GetSizeLayer(False)
    If lyr Is Nothing Then
        MsgBox "Bu sahifada Size o'lchamlari yo'q.", vbInformation, "Size"
        Exit Sub
    End If
    ActiveDocument.BeginCommandGroup "Size: tozalash"
    On Error Resume Next
    lyr.Editable = True
    If lyr.Shapes.Count > 0 Then lyr.Shapes.All.Delete
    lyr.Delete
    On Error GoTo 0
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
End Sub

' Sozlamalar
Public Sub SizeSettings()
    Dim r As String, v As Double
    Dim fontPt As Double, fontName As String, bold As Boolean
    Dim cr As Long, cg As Long, cb As Long
    Dim splitTxt As Boolean, overall As Boolean, useOut As Boolean, rotV As Boolean

    LoadSettings

    If Not Ask("Raqamlar shrift o'lchami (pt)." & vbCr & vbCr & _
               "0 = avtomatik (harflar o'lchamiga qarab tanlanadi)", _
               FormatNum(gFontPt), r) Then Exit Sub
    v = ParseNum(r)
    If v < 0 Or v > 3000 Then
        MsgBox "Shrift o'lchami 0 dan 3000 gacha bo'lishi kerak.", vbExclamation, "Size"
        Exit Sub
    End If
    fontPt = v

    If Not Ask("Raqamlar shrifti nomi:", gFontName, r) Then Exit Sub
    fontName = Trim$(r)
    If fontName = "" Then fontName = "Arial"

    If Not Ask("Qalin (bold) raqamlar?" & vbCr & "1 = ha, 0 = yo'q", BoolStr(gBold), r) Then Exit Sub
    bold = (ParseNum(r) <> 0)

    If Not Ask("O'lcham rangi (R,G,B), masalan: 0,90,200", _
               gColR & "," & gColG & "," & gColB, r) Then Exit Sub
    If Not ParseColor(r, cr, cg, cb) Then
        MsgBox "Rang noto'g'ri kiritildi. Namuna: 0,90,200", vbExclamation, "Size"
        Exit Sub
    End If

    If Not Ask("Matn obyektini harflarga ajratib o'lchash?" & vbCr & _
               "1 = ha (har bir harf alohida), 0 = yo'q (matn bitta obyekt)", _
               BoolStr(gSplitText), r) Then Exit Sub
    splitTxt = (ParseNum(r) <> 0)

    If Not Ask("Bir nechta obyekt tanlanganda umumiy (tashqi) o'lchamni ham ko'rsatish?" & vbCr & _
               "1 = ha, 0 = yo'q", BoolStr(gShowOverall), r) Then Exit Sub
    overall = (ParseNum(r) <> 0)

    If Not Ask("Kontur (outline) qalinligini o'lchamga qo'shish?" & vbCr & _
               "1 = ha, 0 = yo'q (faqat shaklning o'zi)", BoolStr(gUseOutline), r) Then Exit Sub
    useOut = (ParseNum(r) <> 0)

    If Not Ask("Bo'yi yozuvini vertikal (90 gradus) yozish?" & vbCr & _
               "1 = ha (ixcham), 0 = yo'q (gorizontal yoziladi)", BoolStr(gRotateVertical), r) Then Exit Sub
    rotV = (ParseNum(r) <> 0)

    gFontPt = fontPt
    gFontName = fontName
    gBold = bold
    gColR = cr: gColG = cg: gColB = cb
    gSplitText = splitTxt
    gShowOverall = overall
    gUseOutline = useOut
    gRotateVertical = rotV
    SaveSettingsAll

    MsgBox "Sozlamalar saqlandi.", vbInformation, "Size"
End Sub

'==============================================================================
'  ASOSIY JARAYON
'==============================================================================

Private Sub RunSize(ByVal askFont As Boolean)
    Dim doc As Document, sel As ShapeRange, lyr As Layer, oldLayer As Layer
    Dim res As ShapeRange, grp As Shape
    Dim oldUnit As cdrUnit, r As String, v As Double
    Dim errMsg As String, started As Boolean

    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Ochiq hujjat yo'q.", vbExclamation, "Size"
        Exit Sub
    End If
    Set sel = ActiveSelectionRange
    If sel.Count = 0 Then
        MsgBox "Avval o'lchanadigan obyekt yoki harflarni tanlang.", vbInformation, "Size"
        Exit Sub
    End If

    LoadSettings
    If askFont Then
        If Not Ask("Raqamlar shrift o'lchami (pt)." & vbCr & vbCr & _
                   "0 = avtomatik (harflar o'lchamiga qarab tanlanadi)", _
                   FormatNum(gFontPt), r) Then Exit Sub
        v = ParseNum(r)
        If v < 0 Or v > 3000 Then
            MsgBox "Shrift o'lchami 0 dan 3000 gacha bo'lishi kerak.", vbExclamation, "Size"
            Exit Sub
        End If
        gFontPt = v
        SaveSetting APP_KEY, SET_KEY, "FontPt", FormatNum(gFontPt)
    End If

    oldUnit = doc.Unit
    Set oldLayer = ActiveLayer

    On Error GoTo Failed
    doc.BeginCommandGroup "Size"
    started = True
    Application.Optimization = True
    doc.Unit = cdrCentimeter

    CollectItems sel
    If gItemN = 0 Then
        errMsg = "Tanlovda o'lchanadigan obyekt topilmadi."
        GoTo CleanUp
    End If

    ComputeGeometry
    Set gColor = CreateRGBColor(gColR, gColG, gColB)

    Set lyr = GetSizeLayer(True)
    lyr.Visible = True
    lyr.Editable = True

    ResetElements
    CollectExisting lyr

    BuildDims
    If gDimN = 0 Then
        errMsg = "Ko'rsatiladigan o'lcham topilmadi."
        GoTo CleanUp
    End If

    CreateLabels lyr
    Set res = CreateShapeRange
    PlaceAll lyr, res
    If res.Count > 0 Then
        Set grp = res.Group
        grp.Name = "Size"
    End If

CleanUp:
    On Error Resume Next
    doc.Unit = oldUnit
    Application.Optimization = False
    If started Then doc.EndCommandGroup
    If Not oldLayer Is Nothing Then oldLayer.Activate
    sel.CreateSelection
    ActiveWindow.Refresh
    Application.Refresh
    On Error GoTo 0
    If errMsg <> "" Then MsgBox errMsg, vbExclamation, "Size"
    Exit Sub

Failed:
    errMsg = "Size: xatolik yuz berdi." & vbCr & Err.Description
    Resume CleanUp
End Sub

'==============================================================================
'  OBYEKTLARNI YIG'ISH
'==============================================================================

Private Sub CollectItems(sel As ShapeRange)
    Dim s As Shape
    gItemN = 0
    ReDim gItem(0 To 31)
    For Each s In sel
        If Not IsSizeShape(s) Then
            If gSplitText And s.Type = cdrTextShape Then
                AddTextLetters s
            Else
                AddShapeBox s
            End If
        End If
    Next s
    SortItems
End Sub

Private Function IsSizeShape(s As Shape) As Boolean
    On Error Resume Next
    IsSizeShape = (s.Layer.Name = LAYER_NAME)
End Function

Private Sub AddShapeBox(s As Shape)
    Dim x As Double, y As Double, w As Double, h As Double, b As TBox
    s.GetBoundingBox x, y, w, h, gUseOutline
    b = MkBox(x, y, x + w, y + h)
    AddItem b
End Sub

Private Sub AddItem(b As TBox)
    If gItemN > UBound(gItem) Then ReDim Preserve gItem(0 To gItemN * 2 + 16)
    gItem(gItemN) = b
    gItemN = gItemN + 1
End Sub

' Matn obyektini harflarga ajratib o'lchash. Matnning nusxasi egri chiziqqa
' aylantiriladi, bo'laklarga ajratiladi, bo'laklar harflarga birlashtiriladi
' (masalan "i" nuqtasi, "O" ichidagi teshik) va nusxa o'chiriladi.
Private Sub AddTextLetters(s As Shape)
    Dim pb() As TBox, n As Long, i As Long, nChars As Long
    Dim txt As String

    On Error Resume Next
    txt = s.Text.Story.Text
    On Error GoTo 0
    nChars = VisibleCharCount(txt)

    If nChars <= 0 Then
        AddShapeBox s
        Exit Sub
    End If
    If Not ExplodeText(s, pb, n) Then
        AddShapeBox s
        Exit Sub
    End If

    ClusterPieces pb, n, nChars
    For i = 0 To n - 1
        AddItem pb(i)
    Next i
End Sub

Private Function ExplodeText(s As Shape, pb() As TBox, n As Long) As Boolean
    Dim dup As Shape, tmp As ShapeRange, pieces As ShapeRange, p As Shape
    Dim x As Double, y As Double, w As Double, h As Double, i As Long

    On Error GoTo Failed
    Set dup = s.Duplicate
    dup.ConvertToCurves

    Set tmp = CreateShapeRange
    If dup.Type = cdrGroupShape Then
        tmp.AddRange dup.UngroupAllEx
    Else
        tmp.Add dup
    End If

    Set pieces = CreateShapeRange
    For Each p In tmp
        If p.Type = cdrCurveShape Then
            If p.Curve.SubPaths.Count > 1 Then
                pieces.AddRange p.BreakApartEx
            Else
                pieces.Add p
            End If
        Else
            pieces.Add p
        End If
    Next p

    n = pieces.Count
    If n > 0 Then
        ReDim pb(0 To n - 1)
        i = 0
        For Each p In pieces
            p.GetBoundingBox x, y, w, h, gUseOutline
            pb(i) = MkBox(x, y, x + w, y + h)
            i = i + 1
        Next p
        ExplodeText = True
    End If

CleanUp:
    ' vaqtinchalik nusxalarni o'chirish (birma-bir, xatolarga e'tibor bermay)
    On Error Resume Next
    DeleteAll pieces
    DeleteAll tmp
    If Not dup Is Nothing Then dup.Delete
    Exit Function

Failed:
    ExplodeText = False
    Resume CleanUp
End Function

Private Sub DeleteAll(sr As ShapeRange)
    Dim i As Long
    On Error Resume Next
    If sr Is Nothing Then Exit Sub
    For i = sr.Count To 1 Step -1
        sr(i).Delete
    Next i
End Sub

' Bo'laklarni harflarga birlashtirish
Private Sub ClusterPieces(pb() As TBox, n As Long, ByVal nChars As Long)
    Dim par() As Long, i As Long, j As Long, ri As Long, rj As Long
    Dim cb() As TBox, nc As Long, idx() As Long
    Dim bi As Long, bj As Long, d As Double, best As Double

    ReDim par(0 To n - 1)
    For i = 0 To n - 1
        par(i) = i
    Next i

    ' bir harfga tegishli bo'laklar: gorizontal bo'yicha ustma-ust va yaqin
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If PiecesBelong(pb(i), pb(j)) Then
                ri = FindRoot(par, i)
                rj = FindRoot(par, j)
                If ri <> rj Then par(rj) = ri
            End If
        Next j
    Next i

    ReDim cb(0 To n - 1)
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        idx(i) = -1
    Next i
    nc = 0
    For i = 0 To n - 1
        ri = FindRoot(par, i)
        If idx(ri) < 0 Then
            idx(ri) = nc
            cb(nc) = pb(i)
            nc = nc + 1
        Else
            cb(idx(ri)) = UnionBox(cb(idx(ri)), pb(i))
        End If
    Next i

    ' bo'laklar soni harflar sonidan ko'p bo'lsa - eng yaqinlarini birlashtirish
    Do While nc > nChars And nc > 1
        best = BIG
        bi = 0: bj = 1
        For i = 0 To nc - 2
            For j = i + 1 To nc - 1
                d = BoxDist(cb(i), cb(j))
                If d < best Then best = d: bi = i: bj = j
            Next j
        Next i
        cb(bi) = UnionBox(cb(bi), cb(bj))
        cb(bj) = cb(nc - 1)
        nc = nc - 1
    Loop

    ReDim pb(0 To nc - 1)
    For i = 0 To nc - 1
        pb(i) = cb(i)
    Next i
    n = nc
End Sub

Private Function PiecesBelong(a As TBox, b As TBox) As Boolean
    Dim ox As Double, mw As Double, gy As Double, mh As Double
    ox = Min2(a.R, b.R) - Max2(a.L, b.L)
    If ox <= 0 Then Exit Function
    mw = Min2(a.R - a.L, b.R - b.L)
    If mw > 0 Then
        If ox < 0.5 * mw Then Exit Function
    End If
    gy = Max2(a.B, b.B) - Min2(a.T, b.T)
    If gy <= 0 Then
        PiecesBelong = True
    Else
        mh = Max2(a.T - a.B, b.T - b.B)
        PiecesBelong = (gy < 0.6 * mh)
    End If
End Function

Private Function FindRoot(par() As Long, ByVal i As Long) As Long
    Do While par(i) <> i
        par(i) = par(par(i))
        i = par(i)
    Loop
    FindRoot = i
End Function

Private Function VisibleCharCount(ByVal t As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(t)
        Select Case AscW(Mid$(t, i, 1))
            Case 9, 10, 11, 12, 13, 32, 160, 8203, 8232, 8233, 12288
                ' bo'shliq belgilari
            Case Else
                n = n + 1
        End Select
    Next i
    VisibleCharCount = n
End Function

Private Sub SortItems()
    Dim i As Long, j As Long, t As TBox
    For i = 1 To gItemN - 1
        t = gItem(i)
        j = i - 1
        Do While j >= 0
            If Not ItemBefore(t, gItem(j)) Then Exit Do
            gItem(j + 1) = gItem(j)
            j = j - 1
        Loop
        gItem(j + 1) = t
    Next i
End Sub

Private Function ItemBefore(a As TBox, b As TBox) As Boolean
    If Abs(a.L - b.L) > 0.000001 Then
        ItemBefore = (a.L < b.L)
    Else
        ItemBefore = (a.B < b.B)
    End If
End Function

'==============================================================================
'  O'LCHAMLAR RO'YXATI
'==============================================================================

Private Sub BuildDims()
    Dim i As Long, j As Long, n As Long
    Dim d As TDim, tb As TBox, fl As Double, fr As Double
    Dim hasNb() As Boolean

    gDimN = 0
    ReDim gDim(0 To 31)
    n = gItemN

    ' 1) har bir obyektning eni (yuqorida)
    For i = 0 To n - 1
        With gItem(i)
            If .R - .L > MIN_SIZE Then
                d = NewDim(0, .L, .R, .B, .T, .B, .T)
                d.CostHigh = 0
                d.CostLow = 2
                AddDim d
            End If
        End With
    Next i

    ' 2) qo'shni obyektlar orasidagi bo'shliqlar (bir-biriga qaragan chetlaridan)
    If n > 1 Then
        ReDim hasNb(0 To n - 1)
        For i = 0 To n - 1
            For j = 0 To n - 1
                If i <> j Then
                    If IsNeighborH(i, j) Then
                        AddGapH i, j
                        hasNb(i) = True: hasNb(j) = True
                    End If
                    If IsNeighborV(i, j) Then
                        AddGapV i, j
                        hasNb(i) = True: hasNb(j) = True
                    End If
                End If
            Next j
        Next i

        ' qo'shnisi topilmagan (diagonal joylashgan) obyektlar - eng yaqini bilan
        For i = 0 To n - 1
            If Not hasNb(i) Then
                j = Nearest(i)
                If j >= 0 Then
                    If gItem(j).L - gItem(i).R > MIN_SIZE Then
                        AddGapH i, j
                    ElseIf gItem(i).L - gItem(j).R > MIN_SIZE Then
                        AddGapH j, i
                    End If
                    If gItem(j).B - gItem(i).T > MIN_SIZE Then
                        AddGapV i, j
                    ElseIf gItem(i).B - gItem(j).T > MIN_SIZE Then
                        AddGapV j, i
                    End If
                    hasNb(i) = True: hasNb(j) = True
                End If
            End If
        Next i
    End If

    ' 3) har bir obyektning bo'yi (bo'sh joyi ko'p tomonda)
    For i = 0 To n - 1
        With gItem(i)
            If .T - .B > MIN_SIZE Then
                d = NewDim(1, .B, .T, .L, .R, .L, .R)
                d.AllowSlide = True
                fl = FreeSide(i, False)
                fr = FreeSide(i, True)
                If fr > fl Then
                    d.CostHigh = 0: d.CostLow = 1
                Else
                    d.CostLow = 0: d.CostHigh = 1
                End If
                AddDim d
            End If
        End With
    Next i

    ' 4) umumiy o'lcham (ixtiyoriy)
    If gShowOverall And n > 1 Then
        tb = gItem(0)
        For i = 1 To n - 1
            tb = UnionBox(tb, gItem(i))
        Next i
        d = NewDim(0, tb.L, tb.R, tb.B, tb.T, tb.B, tb.T)
        d.CostHigh = 0: d.CostLow = 1.5
        AddDim d
        d = NewDim(1, tb.B, tb.T, tb.L, tb.R, tb.L, tb.R)
        d.CostHigh = 0: d.CostLow = 1
        AddDim d
    End If
End Sub

Private Function NewDim(ByVal ori As Long, ByVal a1 As Double, ByVal a2 As Double, _
                        ByVal p1Lo As Double, ByVal p1Hi As Double, _
                        ByVal p2Lo As Double, ByVal p2Hi As Double) As TDim
    Dim d As TDim
    d.Orient = ori
    d.A1 = a1: d.A2 = a2
    d.P1Lo = p1Lo: d.P1Hi = p1Hi
    d.P2Lo = p2Lo: d.P2Hi = p2Hi
    d.HasInside = False
    d.Value = a2 - a1
    NewDim = d
End Function

Private Sub AddDim(d As TDim)
    If gDimN > UBound(gDim) Then ReDim Preserve gDim(0 To gDimN * 2 + 16)
    gDim(gDimN) = d
    gDimN = gDimN + 1
End Sub

' i - chapda, j - o'ngda
Private Sub AddGapH(ByVal i As Long, ByVal j As Long)
    Dim d As TDim, a As TBox, b As TBox
    a = gItem(i): b = gItem(j)
    d = NewDim(0, a.R, b.L, a.B, a.T, b.B, b.T)
    d.InLo = Max2(a.B, b.B)
    d.InHi = Min2(a.T, b.T)
    d.HasInside = (d.InHi > d.InLo)
    d.CostInside = 0: d.CostLow = 1: d.CostHigh = 2
    AddDim d
End Sub

' i - pastda, j - yuqorida
Private Sub AddGapV(ByVal i As Long, ByVal j As Long)
    Dim d As TDim, a As TBox, b As TBox
    a = gItem(i): b = gItem(j)
    d = NewDim(1, a.T, b.B, a.L, a.R, b.L, b.R)
    d.InLo = Max2(a.L, b.L)
    d.InHi = Min2(a.R, b.R)
    d.HasInside = (d.InHi > d.InLo)
    d.CostInside = 0: d.CostLow = 1: d.CostHigh = 2
    AddDim d
End Sub

' j obyekt i ning o'ng tomonidagi eng yaqin qo'shnisimi
Private Function IsNeighborH(ByVal i As Long, ByVal j As Long) As Boolean
    Dim a As TBox, b As TBox, lo As Double, hi As Double, k As Long
    a = gItem(i): b = gItem(j)
    If b.L - a.R <= MIN_SIZE Then Exit Function
    lo = Max2(a.B, b.B): hi = Min2(a.T, b.T)
    If hi <= lo Then Exit Function
    For k = 0 To gItemN - 1
        If k <> i And k <> j Then
            If gItem(k).L < b.L And gItem(k).R > a.R And _
               gItem(k).B < hi And gItem(k).T > lo Then Exit Function
        End If
    Next k
    IsNeighborH = True
End Function

' j obyekt i ning tepasidagi eng yaqin qo'shnisimi
Private Function IsNeighborV(ByVal i As Long, ByVal j As Long) As Boolean
    Dim a As TBox, b As TBox, lo As Double, hi As Double, k As Long
    a = gItem(i): b = gItem(j)
    If b.B - a.T <= MIN_SIZE Then Exit Function
    lo = Max2(a.L, b.L): hi = Min2(a.R, b.R)
    If hi <= lo Then Exit Function
    For k = 0 To gItemN - 1
        If k <> i And k <> j Then
            If gItem(k).B < b.B And gItem(k).T > a.T And _
               gItem(k).L < hi And gItem(k).R > lo Then Exit Function
        End If
    Next k
    IsNeighborV = True
End Function

Private Function Nearest(ByVal i As Long) As Long
    Dim j As Long, d As Double, best As Double
    Nearest = -1
    best = BIG
    For j = 0 To gItemN - 1
        If j <> i Then
            d = BoxDist(gItem(i), gItem(j))
            If d < best Then best = d: Nearest = j
        End If
    Next j
End Function

' obyektning chap yoki o'ng tomonidagi bo'sh joy
Private Function FreeSide(ByVal i As Long, ByVal toRight As Boolean) As Double
    Dim j As Long, a As TBox, b As TBox
    FreeSide = BIG
    a = gItem(i)
    For j = 0 To gItemN - 1
        If j <> i Then
            b = gItem(j)
            If Min2(a.T, b.T) > Max2(a.B, b.B) Then
                If toRight Then
                    If b.L >= a.R Then FreeSide = Min2(FreeSide, b.L - a.R)
                Else
                    If b.R <= a.L Then FreeSide = Min2(FreeSide, a.L - b.R)
                End If
            End If
        End If
    Next j
End Function

'==============================================================================
'  O'LCHAM VA SHRIFT GEOMETRIYASI
'==============================================================================

Private Sub ComputeGeometry()
    If gFontPt > 0 Then
        gPt = gFontPt
    Else
        gPt = AutoFontPt()
    End If
    gEm = gPt * 2.54 / 72#            ' shrift o'lchami cm da

    gLineW = gEm * 0.03               ' o'lcham chizig'i qalinligi
    gExtW = gEm * 0.02                ' yordamchi chiziq qalinligi
    gArrowL = gEm * 0.32              ' strelka uzunligi
    gArrowW = gEm * 0.11              ' strelka yarim kengligi
    gGapExt = gEm * 0.15              ' obyekt va yordamchi chiziq orasidagi tirqish
    gOver = gEm * 0.2                 ' yordamchi chiziqning o'lcham chizig'idan chiqishi
    gOff0 = gEm * 0.55                ' o'lcham chizig'ining obyektdan masofasi
    gStep = gEm * 0.3                 ' joy qidirish qadami
    gLabelGap = gEm * 0.15            ' yozuv va o'lcham chizig'i orasidagi masofa
    gPad = gEm * 0.1                  ' yozuvlar atrofidagi bo'sh joy
End Sub

' Avtomatik shrift: harflar balandligiga qarab. Kichik matnda raqamlar
' harflardan yirikroq, katta harflarda esa ularning ~13% i (aniq ko'rinadi,
' lekin chizmani bosib ketmaydi).
Private Function AutoFontPt() As Double
    Dim hs() As Double, i As Long, j As Long, t As Double, hPt As Double, pt As Double
    ReDim hs(0 To gItemN - 1)
    For i = 0 To gItemN - 1
        hs(i) = gItem(i).T - gItem(i).B
    Next i
    For i = 1 To gItemN - 1
        t = hs(i)
        j = i - 1
        Do While j >= 0
            If hs(j) <= t Then Exit Do
            hs(j + 1) = hs(j)
            j = j - 1
        Loop
        hs(j + 1) = t
    Next i
    hPt = hs(gItemN \ 2) / 2.54 * 72#

    pt = 0.13 * hPt
    If pt < Min2(1.5 * hPt, 28) Then pt = Min2(1.5 * hPt, 28)
    If pt < 8 Then pt = 8
    If pt > 2000 Then pt = 2000
    AutoFontPt = Int(pt + 0.5)
End Function

'==============================================================================
'  JOYLASHTIRISH
'==============================================================================

Private Sub CreateLabels(lyr As Layer)
    Dim i As Long, s As Shape
    Dim x As Double, y As Double, w As Double, h As Double
    ReDim gLab(0 To gDimN - 1)
    ReDim gLabW(0 To gDimN - 1)
    ReDim gLabH(0 To gDimN - 1)
    For i = 0 To gDimN - 1
        Set s = lyr.CreateArtisticText(0, 0, FormatCm(gDim(i).Value), _
                                       Font:=gFontName, Size:=gPt, _
                                       Bold:=IIf(gBold, cdrTrue, cdrFalse))
        s.Fill.ApplyUniformFill gColor
        s.Outline.SetNoOutline
        If gDim(i).Orient = 1 And gRotateVertical Then s.Rotate 90
        s.GetBoundingBox x, y, w, h, False
        Set gLab(i) = s
        gLabW(i) = w
        gLabH(i) = h
    Next i
End Sub

Private Sub PlaceAll(lyr As Layer, res As ShapeRange)
    Dim i As Long
    For i = 0 To gDimN - 1
        PlaceDim i, lyr, res
    Next i
End Sub

' Variantlarni "narxi" bo'yicha saralab, birinchi bo'sh joyni tanlaydi.
'   1-bosqich: hech narsa bilan kesishmaydi
'   2-bosqich: yordamchi chiziqlar obyektlar va boshqa chiziqlar ustidan o'tishi mumkin
'   3-bosqich: o'lcham chizig'i ham obyekt ustidan o'tishi mumkin
'   4-bosqich: chiziqlar yozuvlar ustidan o'tishi mumkin
'   Barcha bosqichlarda raqamlar obyektlar va boshqa raqamlar ustiga tushmaydi.
Private Sub PlaceDim(ByVal di As Long, lyr As Layer, res As ShapeRange)
    Dim cSide() As Long, cK() As Long, cSl() As Long, cCost() As Double, nC As Long
    Dim k As Long, sd As Long, sl As Long, i As Long, j As Long, lvl As Long, kMax As Long
    Dim ts As Long, tk As Long, tl As Long, tc As Double
    Dim c As TCand

    kMax = StepsNeeded(di)
    ReDim cSide(0 To 7 * (kMax + 1))
    ReDim cK(0 To 7 * (kMax + 1))
    ReDim cSl(0 To 7 * (kMax + 1))
    ReDim cCost(0 To 7 * (kMax + 1))

    For k = 0 To kMax
        For sd = 0 To 2
            If sd < 2 Or gDim(di).HasInside Then
                cSide(nC) = sd: cK(nC) = k: cSl(nC) = 0
                Select Case sd
                    Case 0: cCost(nC) = gDim(di).CostLow + k
                    Case 1: cCost(nC) = gDim(di).CostHigh + k
                    Case Else: cCost(nC) = gDim(di).CostInside + k * 0.5
                End Select
                nC = nC + 1
            End If
            ' yozuvni chiziq uchidan tashqariga surish (zich joylashgan harflar uchun)
            If sd < 2 And gDim(di).AllowSlide Then
                For sl = 1 To 2
                    cSide(nC) = sd: cK(nC) = k: cSl(nC) = sl
                    If sd = 0 Then
                        cCost(nC) = gDim(di).CostLow + 4 + k * 0.5
                    Else
                        cCost(nC) = gDim(di).CostHigh + 4 + k * 0.5
                    End If
                    nC = nC + 1
                Next sl
            End If
        Next sd
    Next k

    For i = 1 To nC - 1
        ts = cSide(i): tk = cK(i): tl = cSl(i): tc = cCost(i)
        j = i - 1
        Do While j >= 0
            If cCost(j) <= tc Then Exit Do
            cSide(j + 1) = cSide(j): cK(j + 1) = cK(j): cSl(j + 1) = cSl(j): cCost(j + 1) = cCost(j)
            j = j - 1
        Loop
        cSide(j + 1) = ts: cK(j + 1) = tk: cSl(j + 1) = tl: cCost(j + 1) = tc
    Next i

    For lvl = 1 To 4
        For i = 0 To nC - 1
            If MakeCand(gDim(di), di, cSide(i), cK(i), cSl(i), c) Then
                If CandFree(gDim(di), c, lvl) Then
                    DrawDim di, c, lyr, res
                    Exit Sub
                End If
            End If
        Next i
    Next lvl

    For i = 0 To nC - 1
        If MakeCand(gDim(di), di, cSide(i), cK(i), cSl(i), c) Then
            DrawDim di, c, lyr, res
            Exit Sub
        End If
    Next i
End Sub

' Qidiruv qadamlari soni: hamma narsadan tashqariga chiqa oladigan darajada
Private Function StepsNeeded(ByVal di As Long) As Long
    Dim u As TBox, i As Long, span As Double, n As Double
    u = gItem(0)
    For i = 1 To gItemN - 1
        u = UnionBox(u, gItem(i))
    Next i
    For i = 0 To gElN - 1
        u = UnionBox(u, gEl(i))
    Next i
    span = Max2(u.R - u.L, u.T - u.B) + gLabW(di) + gLabH(di)
    n = span / gStep + 3
    If n < MIN_STEPS Then n = MIN_STEPS
    If n > MAX_STEPS Then n = MAX_STEPS
    StepsNeeded = CLng(Int(n))
End Function

Private Function MakeCand(d As TDim, ByVal di As Long, ByVal sd As Long, _
                          ByVal k As Long, ByVal sl As Long, c As TCand) As Boolean
    Dim lw As Double, lh As Double, m As Double, kk As Long, la As Double

    lw = gLabW(di): lh = gLabH(di)
    c.Side = sd
    c.Slide = sl
    ' surilgan variantda chiziq obyektga yaqin turadi, k esa yozuvni suradi
    kk = k
    If sl <> 0 Then k = 0
    Select Case sd
        Case 1      ' yuqorida / o'ngda
            c.Pos = Max2(d.P1Hi, d.P2Hi) + gOff0 + k * gStep
            c.HasExt = True
            c.E1From = d.P1Hi + gGapExt: c.E1To = c.Pos + gOver
            c.E2From = d.P2Hi + gGapExt: c.E2To = c.Pos + gOver
        Case 0      ' pastda / chapda
            c.Pos = Min2(d.P1Lo, d.P2Lo) - gOff0 - k * gStep
            c.HasExt = True
            c.E1From = d.P1Lo - gGapExt: c.E1To = c.Pos - gOver
            c.E2From = d.P2Lo - gGapExt: c.E2To = c.Pos - gOver
        Case Else   ' oraliqning ichida (yordamchi chiziqsiz)
            If Not d.HasInside Then Exit Function
            m = (d.InLo + d.InHi) / 2
            If k Mod 2 = 0 Then
                c.Pos = m - (k \ 2) * gStep
            Else
                c.Pos = m + ((k + 1) \ 2) * gStep
            End If
            If c.Pos - gArrowW <= d.InLo Or c.Pos + gArrowW >= d.InHi Then Exit Function
            c.HasExt = False
    End Select

    If d.A2 - d.A1 < 2.2 * gArrowL Then
        c.ArrowsOut = True
        c.LineLo = d.A1 - 1.8 * gArrowL
        c.LineHi = d.A2 + 1.8 * gArrowL
    Else
        c.ArrowsOut = False
        c.LineLo = d.A1
        c.LineHi = d.A2
    End If

    ' yozuv o'lcham chizig'ining tashqi tomonida, o'rtada
    m = (d.A1 + d.A2) / 2
    If d.Orient = 0 Then
        c.Lab.L = m - lw / 2: c.Lab.R = m + lw / 2
        If sd = 0 Then
            c.Lab.T = c.Pos - gLabelGap: c.Lab.B = c.Lab.T - lh
        Else
            c.Lab.B = c.Pos + gLabelGap: c.Lab.T = c.Lab.B + lh
        End If
    Else
        c.Lab.B = m - lh / 2: c.Lab.T = m + lh / 2
        If sd = 0 Then
            c.Lab.R = c.Pos - gLabelGap: c.Lab.L = c.Lab.R - lw
        Else
            c.Lab.L = c.Pos + gLabelGap: c.Lab.R = c.Lab.L + lw
        End If
    End If

    ' yozuvni chiziq uchidan tashqariga surish; chiziq yozuvgacha davom etadi
    If d.Orient = 0 Then la = lw Else la = lh
    If sl = 1 Then
        m = d.A2 + gOver + kk * gStep
        If d.Orient = 0 Then
            c.Lab.L = m: c.Lab.R = m + la
        Else
            c.Lab.B = m: c.Lab.T = m + la
        End If
        c.LineHi = Max2(c.LineHi, m + la)
    ElseIf sl = 2 Then
        m = d.A1 - gOver - kk * gStep
        If d.Orient = 0 Then
            c.Lab.L = m - la: c.Lab.R = m
        Else
            c.Lab.B = m - la: c.Lab.T = m
        End If
        c.LineLo = Min2(c.LineLo, m - la)
    End If
    MakeCand = True
End Function

Private Function CandFree(d As TDim, c As TCand, ByVal lvl As Long) As Boolean
    Dim bx As TBox, eo As Long
    bx = c.Lab
    bx.L = bx.L - gPad: bx.R = bx.R + gPad
    bx.B = bx.B - gPad: bx.T = bx.T + gPad
    If Hits(bx, 1, lvl, 0, 0) Then Exit Function

    bx = LineBox(d, c)
    If Hits(bx, 2, lvl, 0, 0) Then Exit Function

    If c.HasExt Then
        eo = 1 - d.Orient
        bx = ExtBox(d, d.A1, c.E1From, c.E1To)
        If Hits(bx, 3, lvl, eo, d.A1) Then Exit Function
        bx = ExtBox(d, d.A2, c.E2From, c.E2To)
        If Hits(bx, 3, lvl, eo, d.A2) Then Exit Function
    End If
    CandFree = True
End Function

' kind: 1 = yozuv, 2 = o'lcham chizig'i, 3 = yordamchi chiziq
Private Function Hits(bx As TBox, ByVal kind As Long, ByVal lvl As Long, _
                      ByVal extOr As Long, ByVal extC As Double) As Boolean
    Dim i As Long, chk As Boolean
    ' obyektlar bilan
    Select Case kind
        Case 1: chk = True
        Case 2: chk = (lvl < 3)
        Case Else: chk = (lvl < 2)
    End Select
    If chk Then
        For i = 0 To gItemN - 1
            If Overlap(bx, gItem(i)) Then Hits = True: Exit Function
        Next i
    End If
    ' oldin chizilgan o'lchamlar bilan
    For i = 0 To gElN - 1
        If Overlap(bx, gEl(i)) Then
            If kind = 1 Then
                If lvl < 4 Or gElKind(i) = 1 Then Hits = True: Exit Function
            ElseIf gElKind(i) = 1 Then
                If lvl < 4 Then Hits = True: Exit Function
            ElseIf lvl < 2 Then
                ' bir to'g'ri chiziqda yotgan yordamchi chiziqlar ustma-ust tushishi mumkin
                If Not (kind = 3 And gElKind(i) = 3 And gElOr(i) = extOr And _
                        Abs(gElC(i) - extC) < gLineW) Then
                    Hits = True: Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function LineBox(d As TDim, c As TCand) As TBox
    Dim lo As Double, hi As Double, e As Double
    e = gLineW * 0.01
    lo = c.LineLo + e: hi = c.LineHi - e
    If d.Orient = 0 Then
        LineBox = MkBox(lo, c.Pos - gArrowW, hi, c.Pos + gArrowW)
    Else
        LineBox = MkBox(c.Pos - gArrowW, lo, c.Pos + gArrowW, hi)
    End If
End Function

Private Function ExtBox(d As TDim, ByVal a As Double, ByVal p1 As Double, ByVal p2 As Double) As TBox
    If d.Orient = 0 Then
        ExtBox = MkBox(a - gExtW, Min2(p1, p2), a + gExtW, Max2(p1, p2))
    Else
        ExtBox = MkBox(Min2(p1, p2), a - gExtW, Max2(p1, p2), a + gExtW)
    End If
End Function

Private Sub DrawDim(ByVal di As Long, c As TCand, lyr As Layer, res As ShapeRange)
    Dim d As TDim, lbl As Shape, eo As Long, bx As TBox
    Dim x As Double, y As Double, w As Double, h As Double

    d = gDim(di)
    eo = 1 - d.Orient

    ' yordamchi chiziqlar
    If c.HasExt Then
        res.Add Seg(lyr, d, d.A1, c.E1From, d.A1, c.E1To, gExtW)
        res.Add Seg(lyr, d, d.A2, c.E2From, d.A2, c.E2To, gExtW)
        bx = ExtBox(d, d.A1, c.E1From, c.E1To)
        AddEl bx, 3, eo, d.A1
        bx = ExtBox(d, d.A2, c.E2From, c.E2To)
        AddEl bx, 3, eo, d.A2
    End If

    ' o'lcham chizig'i va strelkalar
    res.Add Seg(lyr, d, c.LineLo, c.Pos, c.LineHi, c.Pos, gLineW)
    If c.ArrowsOut Then
        res.Add Arrow(lyr, d, d.A1, c.Pos, 1)
        res.Add Arrow(lyr, d, d.A2, c.Pos, -1)
    Else
        res.Add Arrow(lyr, d, d.A1, c.Pos, -1)
        res.Add Arrow(lyr, d, d.A2, c.Pos, 1)
    End If
    bx = LineBox(d, c)
    AddEl bx, 2, 0, 0

    ' yozuv
    Set lbl = gLab(di)
    lbl.GetBoundingBox x, y, w, h, False
    lbl.Move c.Lab.L - x, c.Lab.B - y
    res.Add lbl
    AddEl c.Lab, 1, 0, 0
End Sub

' (A, P) koordinatalaridagi chiziq
Private Function Seg(lyr As Layer, d As TDim, ByVal a1 As Double, ByVal p1 As Double, _
                     ByVal a2 As Double, ByVal p2 As Double, ByVal wid As Double) As Shape
    Dim s As Shape
    If d.Orient = 0 Then
        Set s = lyr.CreateLineSegment(a1, p1, a2, p2)
    Else
        Set s = lyr.CreateLineSegment(p1, a1, p2, a2)
    End If
    On Error Resume Next
    s.Outline.Type = 1              ' cdrOutline - kontur yoqilgan bo'lsin
    On Error GoTo 0
    s.Outline.Width = wid
    s.Outline.Color.RGBAssign gColR, gColG, gColB
    Set Seg = s
End Function

' Uchi (tipA, pos) nuqtada, A o'qi bo'yicha dirSign tomonga qaragan strelka
Private Function Arrow(lyr As Layer, d As TDim, ByVal tipA As Double, ByVal pos As Double, _
                       ByVal dirSign As Long) As Shape
    Dim crv As Curve, sp As SubPath, s As Shape, baseA As Double
    baseA = tipA - dirSign * gArrowL
    Set crv = Application.CreateCurve(ActiveDocument)
    If d.Orient = 0 Then
        Set sp = crv.CreateSubPath(tipA, pos)
        sp.AppendLineSegment baseA, pos + gArrowW
        sp.AppendLineSegment baseA, pos - gArrowW
    Else
        Set sp = crv.CreateSubPath(pos, tipA)
        sp.AppendLineSegment pos + gArrowW, baseA
        sp.AppendLineSegment pos - gArrowW, baseA
    End If
    sp.Closed = True
    Set s = lyr.CreateCurve(crv)
    s.Fill.ApplyUniformFill gColor
    s.Outline.SetNoOutline
    Set Arrow = s
End Function

'==============================================================================
'  TO'SIQLAR (OBSTACLES)
'==============================================================================

Private Sub ResetElements()
    gElN = 0
    ReDim gEl(0 To 63)
    ReDim gElKind(0 To 63)
    ReDim gElOr(0 To 63)
    ReDim gElC(0 To 63)
End Sub

Private Sub AddEl(b As TBox, ByVal kind As Long, ByVal ori As Long, ByVal coord As Double)
    Dim nn As Long
    If gElN > UBound(gEl) Then
        nn = gElN * 2 + 16
        ReDim Preserve gEl(0 To nn)
        ReDim Preserve gElKind(0 To nn)
        ReDim Preserve gElOr(0 To nn)
        ReDim Preserve gElC(0 To nn)
    End If
    gEl(gElN) = b
    gElKind(gElN) = kind
    gElOr(gElN) = ori
    gElC(gElN) = coord
    gElN = gElN + 1
End Sub

' Oldingi Size o'lchamlari ham to'siq hisoblanadi (yangilari ular ustiga tushmaydi)
Private Sub CollectExisting(lyr As Layer)
    Dim sr As ShapeRange, s As Shape
    Dim x As Double, y As Double, w As Double, h As Double, b As TBox
    On Error Resume Next
    Set sr = lyr.Shapes.FindShapes()
    If sr Is Nothing Then Exit Sub
    For Each s In sr
        If s.Type <> cdrGroupShape Then
            s.GetBoundingBox x, y, w, h, False
            b = MkBox(x, y, x + w, y + h)
            If s.Type = cdrTextShape Then
                AddEl b, 1, 0, 0
            Else
                AddEl b, 2, 0, 0
            End If
        End If
    Next s
End Sub

'==============================================================================
'  SOZLAMALAR
'==============================================================================

Private Sub LoadSettings()
    gFontPt = ParseNum(GetSetting(APP_KEY, SET_KEY, "FontPt", "0"))
    If gFontPt < 0 Or gFontPt > 3000 Then gFontPt = 0
    gFontName = GetSetting(APP_KEY, SET_KEY, "FontName", "Arial")
    If Trim$(gFontName) = "" Then gFontName = "Arial"
    gBold = (GetSetting(APP_KEY, SET_KEY, "Bold", "1") <> "0")
    If Not ParseColor(GetSetting(APP_KEY, SET_KEY, "Color", "0,90,200"), gColR, gColG, gColB) Then
        gColR = 0: gColG = 90: gColB = 200
    End If
    gSplitText = (GetSetting(APP_KEY, SET_KEY, "SplitText", "1") <> "0")
    gShowOverall = (GetSetting(APP_KEY, SET_KEY, "ShowOverall", "0") <> "0")
    gUseOutline = (GetSetting(APP_KEY, SET_KEY, "UseOutline", "0") <> "0")
    gRotateVertical = (GetSetting(APP_KEY, SET_KEY, "RotateVertical", "1") <> "0")
End Sub

Private Sub SaveSettingsAll()
    SaveSetting APP_KEY, SET_KEY, "FontPt", FormatNum(gFontPt)
    SaveSetting APP_KEY, SET_KEY, "FontName", gFontName
    SaveSetting APP_KEY, SET_KEY, "Bold", BoolStr(gBold)
    SaveSetting APP_KEY, SET_KEY, "Color", gColR & "," & gColG & "," & gColB
    SaveSetting APP_KEY, SET_KEY, "SplitText", BoolStr(gSplitText)
    SaveSetting APP_KEY, SET_KEY, "ShowOverall", BoolStr(gShowOverall)
    SaveSetting APP_KEY, SET_KEY, "UseOutline", BoolStr(gUseOutline)
    SaveSetting APP_KEY, SET_KEY, "RotateVertical", BoolStr(gRotateVertical)
End Sub

' InputBox: False - "Cancel" bosildi
Private Function Ask(ByVal prompt As String, ByVal def As String, ByRef result As String) As Boolean
    Dim r As String
    r = InputBox(prompt, "Size - sozlamalar", def)
    If StrPtr(r) = 0 Then Exit Function
    If Trim$(r) = "" Then r = def
    result = r
    Ask = True
End Function

'==============================================================================
'  YORDAMCHI FUNKSIYALAR
'==============================================================================

' 10.64 -> "10.6 cm"; nuqta har doim "." (Windows til sozlamasidan qat'i nazar)
Private Function FormatCm(ByVal v As Double) As String
    Dim n As Long
    If v < 0 Then v = 0
    n = CLng(Int(v * 10# + 0.5 + 0.0000001))
    FormatCm = CStr(n \ 10) & "." & CStr(n Mod 10) & " cm"
End Function

Private Function FormatNum(ByVal v As Double) As String
    FormatNum = Trim$(Replace(Str$(v), ",", "."))
End Function

Private Function ParseNum(ByVal s As String) As Double
    ParseNum = Val(Replace(Trim$(s), ",", "."))
End Function

Private Function BoolStr(ByVal b As Boolean) As String
    If b Then BoolStr = "1" Else BoolStr = "0"
End Function

Private Function ParseColor(ByVal s As String, r As Long, g As Long, b As Long) As Boolean
    Dim p() As String
    s = Replace(Replace(s, ";", ","), " ", "")
    p = Split(s, ",")
    If UBound(p) <> 2 Then Exit Function
    r = CLng(Val(p(0))): g = CLng(Val(p(1))): b = CLng(Val(p(2)))
    If r < 0 Or r > 255 Or g < 0 Or g > 255 Or b < 0 Or b > 255 Then Exit Function
    ParseColor = True
End Function

Private Function GetSizeLayer(ByVal create As Boolean) As Layer
    Dim l As Layer
    For Each l In ActivePage.Layers
        If l.Name = LAYER_NAME Then
            Set GetSizeLayer = l
            Exit Function
        End If
    Next l
    If create Then Set GetSizeLayer = ActivePage.CreateLayer(LAYER_NAME)
End Function

Private Function MkBox(ByVal l As Double, ByVal b As Double, ByVal r As Double, ByVal t As Double) As TBox
    MkBox.L = l: MkBox.B = b: MkBox.R = r: MkBox.T = t
End Function

Private Function UnionBox(a As TBox, b As TBox) As TBox
    UnionBox = MkBox(Min2(a.L, b.L), Min2(a.B, b.B), Max2(a.R, b.R), Max2(a.T, b.T))
End Function

Private Function Overlap(a As TBox, b As TBox) As Boolean
    Overlap = (a.L < b.R And a.R > b.L And a.B < b.T And a.T > b.B)
End Function

Private Function BoxDist(a As TBox, b As TBox) As Double
    Dim dx As Double, dy As Double
    dx = Max2(0, Max2(a.L, b.L) - Min2(a.R, b.R))
    dy = Max2(0, Max2(a.B, b.B) - Min2(a.T, b.T))
    BoxDist = Sqr(dx * dx + dy * dy)
End Function

Private Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function
