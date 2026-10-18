Attribute VB_Name = "SizePlugin"
'==============================================================================
'  SIZE - CorelDRAW uchun o'lchamlarni avtomatik ko'rsatuvchi plagin (VBA)
'
'  Makros bitta: Size. U Size oynasini ochadi, hammasi shu oynada:
'    O'lchash   - tanlangan obyektlar / harflarning enini, bo'yini va ular
'                 orasidagi bo'shliqlarni ish sahifasida ko'rsatadi
'    Tozalash   - joriy sahifadagi barcha Size o'lchamlarini o'chiradi
'    sozlamalar - shrift, shrift o'lchami, chiziq qalinligi, strelka o'lchami,
'                 rang va ish rejimlari
'  Oyna ochiq turadi: obyektlarni tanlab, qayta-qayta "O'lchash" bosish mumkin.
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
Private Const FORM_FILE_NAME As String = "SizeForm"       ' SizeForm.frm dan import qilingan oyna
Private Const FORM_NAME As String = "SizeSettingsForm"     ' import qilinmagan bo'lsa - avtomatik yaratiladigan oyna
Private Const FORM_TAG As String = "SizeForm v3"

' --- Windows: ekran DPI (oyna rasmlari aniq chiqishi uchun) ---
#If VBA7 Then
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

' --- oyna dizayn tizimi (pt) ---
Private Const UI_M As Single = 12           ' oyna cheti
Private Const UI_GAP As Single = 10         ' kartalar orasi
Private Const UI_PAD As Single = 14         ' karta ichki chetlari
Private Const UI_CW As Single = 280         ' karta eni
Private Const UI_TITLE As Single = 16       ' karta sarlavhasi balandligi
Private Const UI_TGAP As Single = 8         ' sarlavha va qatorlar orasi
Private Const UI_ROW As Single = 26         ' qator / maydon / tugma balandligi
Private Const UI_RGAP As Single = 6         ' qatorlar orasi
Private Const UI_CTL As Single = 98         ' boshqaruv elementlari boshlanadigan joy (karta ichida)
Private Const UI_R As Single = 7            ' maydon va tugma burchagi
Private Const UI_CR As Single = 12          ' karta burchagi
Private Const UI_TGW As Single = 30         ' almashtirgich eni
Private Const UI_TGH As Single = 18         ' almashtirgich balandligi
Private Const UI_FONT As String = "Segoe UI"
Private Const UI_FONT_B As String = "Segoe UI Semibold"

' To'g'ri to'rtburchak (cm): chap, past, o'ng, yuqori
Private Type TBox
    L As Double
    B As Double
    R As Double
    T As Double
End Type

' Obyektning eng chetki nuqtalari: o'ng / chap chekka qaysi balandlik oralig'ida
' (RLo..RHi, LLo..LHi), yuqori / pastki chekka qaysi gorizontal oraliqda
' (TLo..THi, BLo..BHi) joylashgan
Private Type TExt
    RLo As Double
    RHi As Double
    LLo As Double
    LHi As Double
    TLo As Double
    THi As Double
    BLo As Double
    BHi As Double
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
    S1Lo As Double          ' A1 uchidagi yordamchi chiziq boshlanadigan chetki nuqta
    S1Hi As Double          ' (past/chap va yuqori/o'ng tomon uchun)
    S2Lo As Double          ' A2 uchidagi chetki nuqta
    S2Hi As Double
    I1 As Long              ' A1 va A2 uchlaridagi obyektlar (-1 = umumiy o'lcham)
    I2 As Long
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
Private gLineMM As Double           ' o'lcham chizig'i qalinligi, mm (0 = avtomatik)
Private gArrowMM As Double          ' strelka uzunligi, mm (0 = avtomatik)
Private gFontName As String
Private gBold As Boolean
Private gColR As Long
Private gColG As Long
Private gColB As Long
Private gSplitText As Boolean       ' matn obyektini harflarga ajratib o'lchash
Private gShowOverall As Boolean     ' butun tanlovning umumiy o'lchami
Private gShowItems As Boolean       ' har bir obyekt / harfning eni va bo'yi
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

' --- O'lchanadigan obyektlar (harflar) ---
Private gItem() As TBox
Private gItemE() As TExt
Private gItemN As Long

' --- Harflarga ajratishdagi bo'laklar (konturlar) ---
Private gPc() As TBox
Private gPcE() As TExt
Private gPcN As Long

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

' --- o'rnatilgan shriftlar (bir marta o'qiladi, alifbo tartibida) ---
Private gFonts() As String
Private gFontsN As Long
Private gFontsOk As Boolean

'==============================================================================
'  OCHIQ MAKROSLAR
'==============================================================================

' Yagona makros: Size oynasini ochadi
Public Sub Size()
    If Not ShowSizeWindow() Then SizeFallbackMenu
End Sub

' Oynani yaratib bo'lmasa - oddiy so'rov orqali tanlash
Private Sub SizeFallbackMenu()
    Dim r As String, ok As Boolean, msg As String
    r = InputBox("SizeForm topilmadi: SizeForm.frm ni import qiling." & vbCr & vbCr & _
                 "1 - O'lchash" & vbCr & _
                 "2 - Sozlamalar" & vbCr & _
                 "3 - Tozalash", "Size", "1")
    If StrPtr(r) = 0 Then Exit Sub
    Select Case Trim$(r)
        Case "2"
            SizeSettingsBasic
        Case "3"
            msg = DoClear(ok)
            MsgBox msg, vbInformation, "Size"
        Case Else
            msg = DoSize(ok)
            If Not ok Then MsgBox msg, vbExclamation, "Size"
    End Select
End Sub

' Joriy sahifadagi barcha Size o'lchamlarini o'chirish
Private Function DoClear(ok As Boolean) As String
    Dim lyr As Layer
    ok = False
    If ActiveDocument Is Nothing Then DoClear = "Ochiq hujjat yo'q.": Exit Function
    Set lyr = GetSizeLayer(False)
    If lyr Is Nothing Then DoClear = "Sahifada o'lcham yo'q.": Exit Function
    ActiveDocument.BeginCommandGroup "Size: tozalash"
    On Error Resume Next
    lyr.Editable = True
    If lyr.Shapes.Count > 0 Then lyr.Shapes.All.Delete
    lyr.Delete
    On Error GoTo 0
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    ok = True
    DoClear = "Tozalandi."
End Function

' Oynani yaratib bo'lmasa - sozlamalar ketma-ket so'raladi
Private Sub SizeSettingsBasic()
    Dim r As String, v As Double, lineMM As Double, arrowMM As Double
    Dim fontPt As Double, fontName As String, bold As Boolean
    Dim cr As Long, cg As Long, cb As Long
    Dim splitTxt As Boolean, overall As Boolean, items As Boolean, useOut As Boolean, rotV As Boolean

    LoadSettings

    If Not Ask("Hajm, pt (0 = avto)", _
               FormatNum(gFontPt), r) Then Exit Sub
    v = ParseNum(r)
    If v < 0 Or v > 3000 Then
        MsgBox "Hajm: 0-3000 pt", vbExclamation, "Size"
        Exit Sub
    End If
    fontPt = v

    If Not Ask("Qalinlik, mm (0 = avto)", _
               FormatNum(gLineMM), r) Then Exit Sub
    lineMM = ParseNum(r)
    If lineMM < 0 Or lineMM > 50 Then
        MsgBox "Qalinlik: 0-50 mm", vbExclamation, "Size"
        Exit Sub
    End If

    If Not Ask("Strelka, mm (0 = avto)", _
               FormatNum(gArrowMM), r) Then Exit Sub
    arrowMM = ParseNum(r)
    If arrowMM < 0 Or arrowMM > 500 Then
        MsgBox "Strelka: 0-500 mm", vbExclamation, "Size"
        Exit Sub
    End If

    If Not Ask("Shrift", gFontName, r) Then Exit Sub
    fontName = Trim$(r)
    If fontName = "" Then fontName = "Arial"

    If Not Ask("Qalin (1 / 0)", BoolStr(gBold), r) Then Exit Sub
    bold = (ParseNum(r) <> 0)

    If Not Ask("Rang, R,G,B", _
               gColR & "," & gColG & "," & gColB, r) Then Exit Sub
    If Not ParseColor(r, cr, cg, cb) Then
        MsgBox "Rang: R,G,B, masalan 0,90,200", vbExclamation, "Size"
        Exit Sub
    End If

    If Not Ask("Harflarga ajratish (1 / 0)", BoolStr(gSplitText), r) Then Exit Sub
    splitTxt = (ParseNum(r) <> 0)

    If Not Ask("Umumiy o'lcham (1 / 0)", BoolStr(gShowOverall), r) Then Exit Sub
    overall = (ParseNum(r) <> 0)

    If Not Ask("Eni va bo'yi (1 / 0)", BoolStr(gShowItems), r) Then Exit Sub
    items = (ParseNum(r) <> 0)

    If Not Ask("Kontur qalinligi (1 / 0)", BoolStr(gUseOutline), r) Then Exit Sub
    useOut = (ParseNum(r) <> 0)

    If Not Ask("Vertikal raqam (1 / 0)", BoolStr(gRotateVertical), r) Then Exit Sub
    rotV = (ParseNum(r) <> 0)

    gFontPt = fontPt
    gLineMM = lineMM
    gArrowMM = arrowMM
    gFontName = ResolveFont(fontName)
    gBold = bold
    gColR = cr: gColG = cg: gColB = cb
    gSplitText = splitTxt
    gShowOverall = overall
    gShowItems = items
    gUseOutline = useOut
    gRotateVertical = rotV
    SaveSettingsAll

    MsgBox "Saqlandi.", vbInformation, "Size"
End Sub

'==============================================================================
'  ASOSIY JARAYON
'==============================================================================

' Tanlangan obyektlarni o'lchab chizadi. Natija haqida xabar qaytaradi,
' ok = True - o'lchamlar chizildi.
Private Function DoSize(ok As Boolean) As String
    Dim doc As Document, sel As ShapeRange, lyr As Layer, oldLayer As Layer
    Dim res As ShapeRange, grp As Shape
    Dim oldUnit As cdrUnit
    Dim errMsg As String, started As Boolean

    ok = False
    Set doc = ActiveDocument
    If doc Is Nothing Then
        DoSize = "Ochiq hujjat yo'q."
        Exit Function
    End If
    Set sel = ActiveSelectionRange
    If sel.Count = 0 Then
        DoSize = "Obyekt tanlanmagan."
        Exit Function
    End If

    LoadSettings
    gFontName = ResolveFont(gFontName)
    oldUnit = doc.Unit
    Set oldLayer = ActiveLayer

    On Error GoTo Failed
    doc.BeginCommandGroup "Size"
    started = True
    Application.Optimization = True
    doc.Unit = cdrCentimeter

    CollectItems sel
    If gItemN = 0 Then
        errMsg = "Tanlovda o'lchanadigan obyekt yo'q."
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
        errMsg = "O'lcham topilmadi."
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
    If errMsg = "" Then
        ok = True
        DoSize = "Tayyor: " & gDimN & " ta o'lcham"
    Else
        DoSize = errMsg
    End If
    Exit Function

Failed:
    errMsg = "Xato: " & Err.Description
    Resume CleanUp
End Function

'==============================================================================
'  OBYEKTLARNI YIG'ISH
'==============================================================================

Private Sub CollectItems(sel As ShapeRange)
    Dim s As Shape
    gItemN = 0
    ReDim gItem(0 To 31)
    ReDim gItemE(0 To 31)
    For Each s In sel
        If Not IsSizeShape(s) Then
            If gSplitText Then
                AddSplitShape s
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

Private Function ShapeBox(s As Shape) As TBox
    Dim x As Double, y As Double, w As Double, h As Double
    s.GetBoundingBox x, y, w, h, gUseOutline
    ShapeBox = MkBox(x, y, x + w, y + h)
End Function

Private Sub AddShapeBox(s As Shape)
    Dim b As TBox, e As TExt
    b = ShapeBox(s)
    e = FullExt(b)
    AddItem b, e
End Sub

Private Sub AddItem(b As TBox, e As TExt)
    If gItemN > UBound(gItem) Then
        ReDim Preserve gItem(0 To gItemN * 2 + 16)
        ReDim Preserve gItemE(0 To gItemN * 2 + 16)
    End If
    gItem(gItemN) = b
    gItemE(gItemN) = e
    gItemN = gItemN + 1
End Sub

' Obyektni harflarga ajratib o'lchash. Matn, egri chiziq (masalan, egri chiziqqa
' aylantirilgan va birlashtirilgan yozuv) yoki guruhning har bir konturi
' (subpath) alohida o'lchanadi, so'ng konturlar harflarga yig'iladi:
' "O" ichidagi teshik, "i" nuqtasi, "O'" ustidagi belgilar harfning o'ziga
' qo'shiladi. Asl obyektga tegilmaydi.
Private Sub AddSplitShape(s As Shape)
    Dim i As Long, nChars As Long, allText As Boolean
    gPcN = 0
    ReDim gPc(0 To 63)
    ReDim gPcE(0 To 63)
    allText = True
    CollectPieces s, nChars, allText
    If gPcN = 0 Then
        AddShapeBox s
        Exit Sub
    End If
    ' harflar sonini faqat sof matn uchun bilamiz
    If Not allText Then nChars = 0
    ClusterPieces nChars
    For i = 0 To gPcN - 1
        AddItem gPc(i), gPcE(i)
    Next i
End Sub

Private Sub CollectPieces(s As Shape, nChars As Long, allText As Boolean)
    Dim ch As Shape, n0 As Long, txt As String, b As TBox, e As TExt
    If s.Type = cdrGroupShape Then
        For Each ch In s.Shapes
            CollectPieces ch, nChars, allText
        Next ch
        Exit Sub
    End If

    If s.Type = cdrTextShape Then
        On Error Resume Next
        txt = s.Text.Story.Text
        On Error GoTo 0
        nChars = nChars + VisibleCharCount(txt)
    Else
        allText = False
    End If

    n0 = gPcN
    If Not AddCurvePieces(s) Then
        ' konturlarni o'qib bo'lmadi - butun obyekt bitta bo'lak
        gPcN = n0
        b = ShapeBox(s)
        e = FullExt(b)
        AddPiece b, e
    End If
End Sub

Private Function AddCurvePieces(s As Shape) As Boolean
    Dim crv As Curve
    Set crv = GetShapeCurve(s)
    If Not crv Is Nothing Then
        AddCurvePieces = PiecesFromCurve(crv, s)
    ElseIf s.Type = cdrTextShape Then
        AddCurvePieces = PiecesFromConvertedCopy(s)
    End If
End Function

' Shaklning ko'rinadigan konturi. Matn uchun ham ishlaydi (DisplayCurve),
' matnni egri chiziqqa aylantirish shart emas.
Private Function GetShapeCurve(s As Shape) As Curve
    Dim o As Object, c As Curve
    On Error Resume Next
    If s.Type = cdrCurveShape Then Set c = s.Curve
    If c Is Nothing Then
        Set o = s
        Set c = o.DisplayCurve
    End If
    If Not c Is Nothing Then
        If c.SubPaths.Count = 0 Then Set c = Nothing
    End If
    If Err.Number <> 0 Then Set c = Nothing
    Err.Clear
    On Error GoTo 0
    Set GetShapeCurve = c
End Function

Private Function PiecesFromCurve(crv As Curve, s As Shape) As Boolean
    Dim sp As SubPath, b As TBox, e As TExt, ow As Double, n0 As Long
    n0 = gPcN
    If gUseOutline Then ow = OutlineHalf(s)
    On Error GoTo Failed
    For Each sp In crv.SubPaths
        If SubPathBox(sp, b, e) Then
            If ow > 0 Then
                b.L = b.L - ow: b.R = b.R + ow
                b.B = b.B - ow: b.T = b.T + ow
            End If
            AddPiece b, e
        End If
    Next sp
    PiecesFromCurve = (gPcN > n0)
    Exit Function
Failed:
    gPcN = n0
    PiecesFromCurve = False
End Function

' Matnning konturini o'qib bo'lmasa: vaqtinchalik nusxa egri chiziqqa
' aylantiriladi, o'lchanadi va o'chiriladi.
Private Function PiecesFromConvertedCopy(s As Shape) As Boolean
    Dim dup As Shape, ch As Shape, crv As Curve, ok As Boolean
    On Error GoTo Failed
    Set dup = s.Duplicate
    dup.ConvertToCurves
    If dup.Type = cdrGroupShape Then
        For Each ch In dup.Shapes
            Set crv = GetShapeCurve(ch)
            If Not crv Is Nothing Then
                If PiecesFromCurve(crv, s) Then ok = True
            End If
        Next ch
    Else
        Set crv = GetShapeCurve(dup)
        If Not crv Is Nothing Then ok = PiecesFromCurve(crv, s)
    End If
CleanUp:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    PiecesFromConvertedCopy = ok
    Exit Function
Failed:
    ok = False
    Resume CleanUp
End Function

' Bitta kontur: chegaralari va eng chetki nuqtalari (qaysi balandlikda harf
' eng o'ngga / chapga, qaysi joyda eng tepaga / pastga chiqib turgani)
Private Function SubPathBox(sp As SubPath, b As TBox, e As TExt) As Boolean
    Dim nd As Node, n As Long, i As Long, px() As Double, py() As Double
    Dim x As Double, y As Double, w As Double, h As Double, o As Object, tol As Double
    Dim xb As TBox

    n = sp.Nodes.Count
    If n = 0 Then Exit Function
    ReDim px(0 To n - 1)
    ReDim py(0 To n - 1)
    i = 0
    For Each nd In sp.Nodes
        px(i) = nd.PositionX
        py(i) = nd.PositionY
        i = i + 1
    Next nd

    b = MkBox(px(0), py(0), px(0), py(0))
    For i = 1 To n - 1
        If px(i) < b.L Then b.L = px(i)
        If px(i) > b.R Then b.R = px(i)
        If py(i) < b.B Then b.B = py(i)
        If py(i) > b.T Then b.T = py(i)
    Next i

    ' aniq chegaralar (egri segmentlarning bo'rtib chiqqan joylari bilan)
    On Error Resume Next
    Set o = sp
    w = -1: h = -1
    o.GetBoundingBox x, y, w, h
    If Err.Number = 0 And w >= 0 And h >= 0 Then
        xb = MkBox(x, y, x + w, y + h)
        b = UnionBox(b, xb)
    End If
    Err.Clear
    On Error GoTo 0

    tol = Max2(b.R - b.L, b.T - b.B) * 0.01 + 0.0001
    e.RLo = BIG: e.RHi = -BIG: e.LLo = BIG: e.LHi = -BIG
    e.TLo = BIG: e.THi = -BIG: e.BLo = BIG: e.BHi = -BIG
    For i = 0 To n - 1
        If px(i) >= b.R - tol Then e.RLo = Min2(e.RLo, py(i)): e.RHi = Max2(e.RHi, py(i))
        If px(i) <= b.L + tol Then e.LLo = Min2(e.LLo, py(i)): e.LHi = Max2(e.LHi, py(i))
        If py(i) >= b.T - tol Then e.TLo = Min2(e.TLo, px(i)): e.THi = Max2(e.THi, px(i))
        If py(i) <= b.B + tol Then e.BLo = Min2(e.BLo, px(i)): e.BHi = Max2(e.BHi, px(i))
    Next i
    If e.RLo > e.RHi Then e.RLo = b.B: e.RHi = b.T
    If e.LLo > e.LHi Then e.LLo = b.B: e.LHi = b.T
    If e.TLo > e.THi Then e.TLo = b.L: e.THi = b.R
    If e.BLo > e.BHi Then e.BLo = b.L: e.BHi = b.R
    SubPathBox = True
End Function

Private Function OutlineHalf(s As Shape) As Double
    On Error Resume Next
    If s.Outline.Type <> 0 Then OutlineHalf = s.Outline.Width / 2
    If OutlineHalf < 0 Then OutlineHalf = 0
End Function

Private Sub AddPiece(b As TBox, e As TExt)
    If gPcN > UBound(gPc) Then
        ReDim Preserve gPc(0 To gPcN * 2 + 16)
        ReDim Preserve gPcE(0 To gPcN * 2 + 16)
    End If
    gPc(gPcN) = b
    gPcE(gPcN) = e
    gPcN = gPcN + 1
End Sub

' To'liq chekka (harfning eng chetki nuqtalari noma'lum bo'lsa)
Private Function FullExt(b As TBox) As TExt
    Dim e As TExt
    e.RLo = b.B: e.RHi = b.T
    e.LLo = b.B: e.LHi = b.T
    e.TLo = b.L: e.THi = b.R
    e.BLo = b.L: e.BHi = b.R
    FullExt = e
End Function

' Ikki bo'lak birlashganda eng chetki nuqtalarni saqlab qolish
Private Function MergeExt(a As TBox, ea As TExt, b As TBox, eb As TExt) As TExt
    Dim r As TExt, tol As Double
    tol = Max2(Max2(a.R - a.L, a.T - a.B), Max2(b.R - b.L, b.T - b.B)) * 0.01 + 0.0001
    If a.R > b.R + tol Then
        r.RLo = ea.RLo: r.RHi = ea.RHi
    ElseIf b.R > a.R + tol Then
        r.RLo = eb.RLo: r.RHi = eb.RHi
    Else
        r.RLo = Min2(ea.RLo, eb.RLo): r.RHi = Max2(ea.RHi, eb.RHi)
    End If
    If a.L < b.L - tol Then
        r.LLo = ea.LLo: r.LHi = ea.LHi
    ElseIf b.L < a.L - tol Then
        r.LLo = eb.LLo: r.LHi = eb.LHi
    Else
        r.LLo = Min2(ea.LLo, eb.LLo): r.LHi = Max2(ea.LHi, eb.LHi)
    End If
    If a.T > b.T + tol Then
        r.TLo = ea.TLo: r.THi = ea.THi
    ElseIf b.T > a.T + tol Then
        r.TLo = eb.TLo: r.THi = eb.THi
    Else
        r.TLo = Min2(ea.TLo, eb.TLo): r.THi = Max2(ea.THi, eb.THi)
    End If
    If a.B < b.B - tol Then
        r.BLo = ea.BLo: r.BHi = ea.BHi
    ElseIf b.B < a.B - tol Then
        r.BLo = eb.BLo: r.BHi = eb.BHi
    Else
        r.BLo = Min2(ea.BLo, eb.BLo): r.BHi = Max2(ea.BHi, eb.BHi)
    End If
    MergeExt = r
End Function

' Bo'laklarni harflarga birlashtirish (gPc / gPcE ustida)
Private Sub ClusterPieces(ByVal nChars As Long)
    Dim par() As Long, i As Long, j As Long, ri As Long, rj As Long, n As Long
    Dim cb() As TBox, ce() As TExt, nc As Long, idx() As Long
    Dim bi As Long, bj As Long, d As Double, best As Double

    n = gPcN
    ReDim par(0 To n - 1)
    For i = 0 To n - 1
        par(i) = i
    Next i

    ' bir harfga tegishli bo'laklar: gorizontal bo'yicha ustma-ust va yaqin
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If PiecesBelong(gPc(i), gPc(j)) Then
                ri = FindRoot(par, i)
                rj = FindRoot(par, j)
                If ri <> rj Then par(rj) = ri
            End If
        Next j
    Next i

    ReDim cb(0 To n - 1)
    ReDim ce(0 To n - 1)
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        idx(i) = -1
    Next i
    nc = 0
    For i = 0 To n - 1
        ri = FindRoot(par, i)
        If idx(ri) < 0 Then
            idx(ri) = nc
            cb(nc) = gPc(i)
            ce(nc) = gPcE(i)
            nc = nc + 1
        Else
            j = idx(ri)
            ce(j) = MergeExt(cb(j), ce(j), gPc(i), gPcE(i))
            cb(j) = UnionBox(cb(j), gPc(i))
        End If
    Next i

    ' bo'laklar soni harflar sonidan ko'p bo'lsa - eng yaqinlarini birlashtirish
    Do While nChars > 0 And nc > nChars And nc > 1
        best = BIG
        bi = 0: bj = 1
        For i = 0 To nc - 2
            For j = i + 1 To nc - 1
                d = BoxDist(cb(i), cb(j))
                If d < best Then best = d: bi = i: bj = j
            Next j
        Next i
        ce(bi) = MergeExt(cb(bi), ce(bi), cb(bj), ce(bj))
        cb(bi) = UnionBox(cb(bi), cb(bj))
        cb(bj) = cb(nc - 1)
        ce(bj) = ce(nc - 1)
        nc = nc - 1
    Loop

    For i = 0 To nc - 1
        gPc(i) = cb(i)
        gPcE(i) = ce(i)
    Next i
    gPcN = nc
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
    Dim i As Long, j As Long, t As TBox, te As TExt
    For i = 1 To gItemN - 1
        t = gItem(i)
        te = gItemE(i)
        j = i - 1
        Do While j >= 0
            If Not ItemBefore(t, gItem(j)) Then Exit Do
            gItem(j + 1) = gItem(j)
            gItemE(j + 1) = gItemE(j)
            j = j - 1
        Loop
        gItem(j + 1) = t
        gItemE(j + 1) = te
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
    If gShowItems Or n = 1 Then
        For i = 0 To n - 1
            With gItem(i)
                If .R - .L > MIN_SIZE Then
                    d = NewDim(0, .L, .R, .B, .T, .B, .T)
                    d.S1Lo = gItemE(i).LLo: d.S1Hi = gItemE(i).LHi
                    d.S2Lo = gItemE(i).RLo: d.S2Hi = gItemE(i).RHi
                    d.I1 = i: d.I2 = i
                    d.CostHigh = 0
                    d.CostLow = 2
                    AddDim d
                End If
            End With
        Next i
    End If

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
    If gShowItems Or n = 1 Then
        For i = 0 To n - 1
            With gItem(i)
                If .T - .B > MIN_SIZE Then
                    d = NewDim(1, .B, .T, .L, .R, .L, .R)
                    d.S1Lo = gItemE(i).BLo: d.S1Hi = gItemE(i).BHi
                    d.S2Lo = gItemE(i).TLo: d.S2Hi = gItemE(i).THi
                    d.I1 = i: d.I2 = i
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
    End If

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
    d.S1Lo = p1Lo: d.S1Hi = p1Hi
    d.S2Lo = p2Lo: d.S2Hi = p2Hi
    d.I1 = -1: d.I2 = -1
    d.HasInside = False
    d.Value = a2 - a1
    NewDim = d
End Function

Private Sub AddDim(d As TDim)
    If gDimN > UBound(gDim) Then ReDim Preserve gDim(0 To gDimN * 2 + 16)
    gDim(gDimN) = d
    gDimN = gDimN + 1
End Sub

' i - chapda, j - o'ngda. Oraliq i ning eng o'ngga chiqqan nuqtasidan
' j ning eng chapga chiqqan nuqtasigacha o'lchanadi.
Private Sub AddGapH(ByVal i As Long, ByVal j As Long)
    Dim d As TDim, a As TBox, b As TBox
    a = gItem(i): b = gItem(j)
    d = NewDim(0, a.R, b.L, a.B, a.T, b.B, b.T)
    d.S1Lo = gItemE(i).RLo: d.S1Hi = gItemE(i).RHi
    d.S2Lo = gItemE(j).LLo: d.S2Hi = gItemE(j).LHi
    d.I1 = i: d.I2 = j
    ' strelkalar ikkala harfning chetki nuqtasiga tegadigan balandlik
    d.InLo = Max2(gItemE(i).RLo, gItemE(j).LLo)
    d.InHi = Min2(gItemE(i).RHi, gItemE(j).LHi)
    d.HasInside = (d.InHi >= d.InLo)
    d.CostInside = 0: d.CostLow = 1: d.CostHigh = 2
    AddDim d
End Sub

' i - pastda, j - yuqorida
Private Sub AddGapV(ByVal i As Long, ByVal j As Long)
    Dim d As TDim, a As TBox, b As TBox
    a = gItem(i): b = gItem(j)
    d = NewDim(1, a.T, b.B, a.L, a.R, b.L, b.R)
    d.S1Lo = gItemE(i).TLo: d.S1Hi = gItemE(i).THi
    d.S2Lo = gItemE(j).BLo: d.S2Hi = gItemE(j).BHi
    d.I1 = i: d.I2 = j
    d.InLo = Max2(gItemE(i).TLo, gItemE(j).BLo)
    d.InHi = Min2(gItemE(i).THi, gItemE(j).BHi)
    d.HasInside = (d.InHi >= d.InLo)
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
    gPad = gEm * 0.1                  ' yozuvlar atrofidagi bo'sh joy

    ' qo'lda berilgan chiziq qalinligi va strelka o'lchami (mm -> cm)
    If gLineMM > 0 Then
        gLineW = gLineMM / 10#
        gExtW = gLineW * 0.67
    End If
    If gArrowMM > 0 Then
        gArrowL = gArrowMM / 10#
        gArrowW = gArrowL * 0.34
    End If
    ' yozuv strelka va chiziq ustiga tushmasligi uchun
    gLabelGap = Max2(gEm * 0.15, gArrowW + gLineW + gEm * 0.04)
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
        ApplyLabelFont s
        s.Fill.ApplyUniformFill gColor
        s.Outline.SetNoOutline
        If gDim(i).Orient = 1 And gRotateVertical Then s.Rotate 90
        s.GetBoundingBox x, y, w, h, False
        Set gLab(i) = s
        gLabW(i) = w
        gLabH(i) = h
    Next i
End Sub

' Shrift, o'lcham va qalinlikni butun matnga aniq qo'llash
' (ba'zi versiyalarda CreateArtisticText parametrlari to'liq qo'llanmaydi)
Private Sub ApplyLabelFont(s As Shape)
    Dim tr As Object
    On Error Resume Next
    Set tr = s.Text.Story
    tr.Font = gFontName
    tr.Size = gPt
    tr.Bold = gBold
    tr.Italic = False
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
            c.E1From = d.S1Hi + gGapExt: c.E1To = c.Pos + gOver
            c.E2From = d.S2Hi + gGapExt: c.E2To = c.Pos + gOver
        Case 0      ' pastda / chapda
            c.Pos = Min2(d.P1Lo, d.P2Lo) - gOff0 - k * gStep
            c.HasExt = True
            c.E1From = d.S1Lo - gGapExt: c.E1To = c.Pos - gOver
            c.E2From = d.S2Lo - gGapExt: c.E2To = c.Pos - gOver
        Case Else   ' oraliqning ichida (yordamchi chiziqsiz)
            If Not d.HasInside Then Exit Function
            m = (d.InLo + d.InHi) / 2
            If k Mod 2 = 0 Then
                c.Pos = m - (k \ 2) * gStep
            Else
                c.Pos = m + ((k + 1) \ 2) * gStep
            End If
            ' strelkalar ikkala harfga tegib turishi kerak
            If c.Pos < d.InLo - 0.000001 Or c.Pos > d.InHi + 0.000001 Then Exit Function
            If k > 0 And d.InHi - d.InLo < gStep Then Exit Function
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
    Dim bx As TBox
    bx = c.Lab
    bx.L = bx.L - gPad: bx.R = bx.R + gPad
    bx.B = bx.B - gPad: bx.T = bx.T + gPad
    If Hits(bx, 1, lvl, -1, -1) Then Exit Function

    bx = LineBox(d, c)
    If Hits(bx, 2, lvl, -1, -1) Then Exit Function

    ' yordamchi chiziq o'z harfining chetki nuqtasidan boshlanadi, shuning uchun
    ' shu harf to'siq hisoblanmaydi
    If c.HasExt Then
        bx = ExtBox(d, d.A1, c.E1From, c.E1To)
        If Hits(bx, 3, lvl, d.I1, d.I2) Then Exit Function
        bx = ExtBox(d, d.A2, c.E2From, c.E2To)
        If Hits(bx, 3, lvl, d.I1, d.I2) Then Exit Function
    End If
    CandFree = True
End Function

' kind: 1 = yozuv, 2 = o'lcham chizig'i, 3 = yordamchi chiziq
' skip1, skip2 - tekshirilmaydigan obyektlar (yordamchi chiziqning o'z harflari)
Private Function Hits(bx As TBox, ByVal kind As Long, ByVal lvl As Long, _
                      ByVal skip1 As Long, ByVal skip2 As Long) As Boolean
    Dim i As Long, chk As Boolean
    ' obyektlar bilan
    Select Case kind
        Case 1: chk = True
        Case 2: chk = (lvl < 3)
        Case Else: chk = (lvl < 2)
    End Select
    If chk Then
        For i = 0 To gItemN - 1
            If i <> skip1 And i <> skip2 Then
                If Overlap(bx, gItem(i)) Then Hits = True: Exit Function
            End If
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
                ' yordamchi chiziqlar bir-birini kesib o'tishi mumkin
                If Not (kind = 3 And gElKind(i) = 3) Then Hits = True: Exit Function
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
    gLineMM = ParseNum(GetSetting(APP_KEY, SET_KEY, "LineMM", "0"))
    If gLineMM < 0 Or gLineMM > 50 Then gLineMM = 0
    gArrowMM = ParseNum(GetSetting(APP_KEY, SET_KEY, "ArrowMM", "0"))
    If gArrowMM < 0 Or gArrowMM > 500 Then gArrowMM = 0
    gFontName = GetSetting(APP_KEY, SET_KEY, "FontName", "Arial")
    If Trim$(gFontName) = "" Then gFontName = "Arial"
    gBold = (GetSetting(APP_KEY, SET_KEY, "Bold", "1") <> "0")
    If Not ParseColor(GetSetting(APP_KEY, SET_KEY, "Color", "0,90,200"), gColR, gColG, gColB) Then
        gColR = 0: gColG = 90: gColB = 200
    End If
    gSplitText = (GetSetting(APP_KEY, SET_KEY, "SplitText", "1") <> "0")
    gShowOverall = (GetSetting(APP_KEY, SET_KEY, "ShowOverall", "1") <> "0")
    gShowItems = (GetSetting(APP_KEY, SET_KEY, "ShowItems", "1") <> "0")
    gUseOutline = (GetSetting(APP_KEY, SET_KEY, "UseOutline", "0") <> "0")
    gRotateVertical = (GetSetting(APP_KEY, SET_KEY, "RotateVertical", "1") <> "0")
End Sub

Private Sub SaveSettingsAll()
    SaveSetting APP_KEY, SET_KEY, "FontPt", FormatNum(gFontPt)
    SaveSetting APP_KEY, SET_KEY, "LineMM", FormatNum(gLineMM)
    SaveSetting APP_KEY, SET_KEY, "ArrowMM", FormatNum(gArrowMM)
    SaveSetting APP_KEY, SET_KEY, "FontName", gFontName
    SaveSetting APP_KEY, SET_KEY, "Bold", BoolStr(gBold)
    SaveSetting APP_KEY, SET_KEY, "Color", gColR & "," & gColG & "," & gColB
    SaveSetting APP_KEY, SET_KEY, "SplitText", BoolStr(gSplitText)
    SaveSetting APP_KEY, SET_KEY, "ShowOverall", BoolStr(gShowOverall)
    SaveSetting APP_KEY, SET_KEY, "ShowItems", BoolStr(gShowItems)
    SaveSetting APP_KEY, SET_KEY, "UseOutline", BoolStr(gUseOutline)
    SaveSetting APP_KEY, SET_KEY, "RotateVertical", BoolStr(gRotateVertical)
End Sub

' InputBox: False - "Cancel" bosildi
Private Function Ask(ByVal prompt As String, ByVal def As String, ByRef result As String) As Boolean
    Dim r As String
    r = InputBox(prompt, "Size", def)
    If StrPtr(r) = 0 Then Exit Function
    If Trim$(r) = "" Then r = def
    result = r
    Ask = True
End Function

'==============================================================================
'  SIZE OYNASI
'  Dizayn: och kulrang fon, oq kartalar, yumaloq burchaklar, yengil chegara.
'  MSForms yumaloq burchak chiza olmaydi, shuning uchun karta, maydon, tugma va
'  almashtirgich (toggle) fonlari kichik rasm sifatida chiziladi va %TEMP%\SizeUI
'  papkasida saqlanadi (keyingi ochilishlarda qayta chizilmaydi).
'==============================================================================


Private Function CBg() As Long          ' oyna foni
    CBg = VBA.RGB(242, 242, 245)
End Function
Private Function CCard() As Long      ' karta
    CCard = VBA.RGB(255, 255, 255)
End Function
Private Function CLine() As Long      ' chegara
    CLine = VBA.RGB(230, 230, 235)
End Function
Private Function CField() As Long    ' maydon
    CField = VBA.RGB(242, 242, 245)
End Function
Private Function CText() As Long         ' asosiy matn
    CText = VBA.RGB(28, 28, 30)
End Function
Private Function CMuted() As Long    ' ikkinchi darajali matn
    CMuted = VBA.RGB(142, 142, 147)
End Function
Private Function CDim() As Long        ' o'chirilgan
    CDim = VBA.RGB(199, 199, 204)
End Function
Private Function CPrimary() As Long   ' asosiy tugma
    CPrimary = VBA.RGB(28, 28, 30)
End Function
Private Function COn() As Long            ' yoqilgan almashtirgich
    COn = VBA.RGB(52, 199, 89)
End Function
Private Function COff() As Long        ' o'chirilgan almashtirgich
    COff = VBA.RGB(229, 229, 234)
End Function
Private Function CError() As Long       ' xato
    CError = VBA.RGB(215, 0, 21)
End Function

' O'zbekcha yozuv: ' -> o'/g' belgisi, -- -> chiziqcha
Private Function UZ(ByVal s As String) As String
    s = Replace(s, "'", ChrW(&H2018))
    UZ = Replace(s, "--", ChrW(&H2013))
End Function

' --- oynani ochish ---

' Size oynasini ochadi (oyna ochiq turadi - CorelDRAW'da ishlashda davom etish
' mumkin). Avval SizeForm.frm dan import qilingan oyna ishlatiladi; u bo'lmasa,
' oyna VBA muharriri orqali yaratiladi. False - oynani ochib bo'lmadi.
Private Function ShowSizeWindow() As Boolean
    Dim f As Object
    On Error GoTo Failed
    ' oyna allaqachon yuklangan bo'lsa - shuni ko'rsatamiz
    For Each f In VBA.UserForms
        If f.Name = FORM_FILE_NAME Or f.Name = FORM_NAME Then
            f.Show 0
            FillFontsDeferred f
            ShowSizeWindow = True
            Exit Function
        End If
    Next f
    Set f = Nothing
    Set f = LoadFormByName(FORM_FILE_NAME)
    If f Is Nothing Then
        If EnsureForm() Then Set f = LoadFormByName(FORM_NAME)
    End If
    If f Is Nothing Then Exit Function
    f.Show 0                            ' 0 = vbModeless
    FillFontsDeferred f                 ' shriftlar ro'yxati oyna ko'ringandan keyin
    ShowSizeWindow = True
    Exit Function
Failed:
    ShowSizeWindow = False
End Function

Private Function LoadFormByName(ByVal nm As String) As Object
    On Error Resume Next
    Set LoadFormByName = VBA.UserForms.Add(nm)
End Function

' Oyna (UserForm) GlobalMacros loyihasida bo'lmasa - yaratadi, kodi eski
' bo'lsa - yangilaydi
Private Function EnsureForm() As Boolean
    Dim vbp As Object, comp As Object, cm As Object, cur As Boolean
    On Error GoTo Failed
    Set vbp = MyProject()
    If vbp Is Nothing Then Exit Function
    Set comp = FindComponent(vbp, FORM_NAME)
    If comp Is Nothing Then
        Set comp = vbp.VBComponents.Add(3)      ' 3 = vbext_ct_MSForm
        comp.Name = FORM_NAME
    End If
    Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then cur = (InStr(1, cm.Lines(1, 1), FORM_TAG) > 0)
    If Not cur Then
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        cm.AddFromString FormCode()
    End If
    EnsureForm = True
    Exit Function
Failed:
    EnsureForm = False
End Function

' SizePlugin moduli joylashgan VBA loyihasi
Private Function MyProject() As Object
    Dim app As Object, p As Object
    On Error Resume Next
    Set app = Application
    For Each p In app.VBE.VBProjects
        If Not FindComponent(p, "SizePlugin") Is Nothing Then
            Set MyProject = p
            Exit Function
        End If
    Next p
End Function

Private Function FindComponent(p As Object, ByVal nm As String) As Object
    On Error Resume Next
    Set FindComponent = p.VBComponents(nm)
End Function

' --- oyna hodisalari (SizeForm kodidan chaqiriladi) ---

' Barcha bosishlar shu yerga keladi
Public Sub SizeUIClick(f As Object, ByVal nm As String)
    On Error GoTo Failed
    Select Case nm
        Case "btnRun": SizeFormRun f
        Case "btnClear": SizeFormClear f
        Case "btnClose": SizeFormClose f
        Case "btnColor": SizeFormPickColor f
        Case "btnFontMinus": StepField f, "txtFontPt", "tglFontAuto", -1, 1, 3000, 0
        Case "btnFontPlus": StepField f, "txtFontPt", "tglFontAuto", 1, 1, 3000, 0
        Case "btnLineMinus": StepField f, "txtLine", "tglLineAuto", -0.1, 0.1, 50, 1
        Case "btnLinePlus": StepField f, "txtLine", "tglLineAuto", 0.1, 0.1, 50, 1
        Case "btnArrowMinus": StepField f, "txtArrow", "tglArrowAuto", -0.5, 0.5, 500, 1
        Case "btnArrowPlus": StepField f, "txtArrow", "tglArrowAuto", 0.5, 0.5, 500, 1
        Case Else
            If Left$(nm, 3) = "tgl" Then SetToggle f, nm, Not TogVal(f, nm)
    End Select
    SizeFormSync f
    Exit Sub
Failed:
    SetStatus f, "Xato: " & Err.Description, False
End Sub

Public Sub SizeFormRun(f As Object)
    Dim ok As Boolean, msg As String
    If Not SizeFormSave(f) Then Exit Sub
    SetStatus f, "O'lchanmoqda...", True
    msg = DoSize(ok)
    SetStatus f, msg, ok
End Sub

Public Sub SizeFormClear(f As Object)
    Dim ok As Boolean, msg As String
    msg = DoClear(ok)
    SetStatus f, msg, ok
End Sub

Public Sub SizeFormClose(f As Object)
    SizeFormSave f, True                ' to'g'ri qiymatlar jimgina saqlanadi
    f.Hide
End Sub

Private Sub SetStatus(f As Object, ByVal msg As String, ByVal ok As Boolean)
    On Error Resume Next
    With f.Controls("lblStatus")
        .Caption = UZ(msg)
        If ok Then .ForeColor = CMuted() Else .ForeColor = CError()
    End With
    f.Repaint
End Sub

' "+" / "-": Avto o'chadi, qiymat qadam bilan o'zgaradi
Private Sub StepField(f As Object, ByVal txtName As String, ByVal tglName As String, _
                      ByVal delta As Double, ByVal vMin As Double, ByVal vMax As Double, ByVal decimals As Long)
    Dim v As Double
    If TogVal(f, tglName) Then SetToggle f, tglName, False
    v = ParseNum(f.Controls(txtName).Text) + delta
    If v < vMin Then v = vMin
    If v > vMax Then v = vMax
    f.Controls(txtName).Text = FormatNum(Round(v, decimals))
End Sub

' --- oyna tarkibi ---

Public Sub SizeFormBuild(f As Object)
    Dim c As Object, x As Single, y As Single, y0 As Single, cardH As Single, h3 As Single
    On Error GoTo Failed
    LoadSettings

    f.Caption = "Size"
    f.BackColor = CBg()
    f.Width = UI_M * 2 + UI_CW * 2 + UI_GAP + 6     ' + oyna ramkasi
    cardH = UI_PAD * 2 + UI_TITLE + UI_TGAP + UI_ROW * 3 + UI_RGAP * 2
    f.Height = UI_M * 2 + cardH * 2 + UI_GAP + 28   ' + sarlavha qatori

    ' --- chap ustun: Raqamlar ---
    x = UI_M: y0 = UI_M
    UiCard f, x, y0, UI_CW, cardH, "Raqamlar"
    y = y0 + UI_PAD + UI_TITLE + UI_TGAP
    UiRowLabel f, x, y, "Shrift", "O'lcham raqamlari shrifti"
    UiFieldBg f, x + UI_CTL, y, UI_CW - UI_CTL - UI_PAD, CCard()
    Set c = f.Controls.Add("Forms.ComboBox.1", "cboFont", True)
    c.Left = x + UI_CTL + 6: c.Top = y + 4: c.Width = UI_CW - UI_CTL - UI_PAD - 10: c.Height = UI_ROW - 8
    c.BackColor = CField(): c.ForeColor = CText()
    c.SpecialEffect = 0: c.BorderStyle = 0
    c.ShowDropButtonWhen = 2: c.DropButtonStyle = 1
    UiFont c, False, 9, CText()
    c.AddItem gFontName                 ' to'liq ro'yxat oyna ochilgandan keyin yuklanadi
    c.Text = gFontName
    y = y + UI_ROW + UI_RGAP
    UiRowLabel f, x, y, "Hajm, pt", "Raqamlar hajmi. Avto -- harflar balandligiga qarab"
    UiStepper f, x + UI_CTL, y, "Font", IIf(gFontPt > 0, FormatNum(gFontPt), "36")
    UiAuto f, x, y, "tglFontAuto", (gFontPt <= 0)
    y = y + UI_ROW + UI_RGAP
    UiRowLabel f, x, y, "Qalin", "Qalin raqamlar"
    UiToggle f, "tglBold", x + UI_CW - UI_PAD - UI_TGW, y, gBold

    ' --- chap ustun: Chiziqlar ---
    y0 = UI_M + cardH + UI_GAP
    UiCard f, x, y0, UI_CW, cardH, "Chiziqlar"
    y = y0 + UI_PAD + UI_TITLE + UI_TGAP
    UiRowLabel f, x, y, "Qalinlik, mm", "O'lcham chizig'i qalinligi. Yordamchi chiziqlar 2/3 qalinlikda"
    UiStepper f, x + UI_CTL, y, "Line", IIf(gLineMM > 0, FormatNum(gLineMM), "0.5")
    UiAuto f, x, y, "tglLineAuto", (gLineMM <= 0)
    y = y + UI_ROW + UI_RGAP
    UiRowLabel f, x, y, "Strelka, mm", "Strelka uzunligi"
    UiStepper f, x + UI_CTL, y, "Arrow", IIf(gArrowMM > 0, FormatNum(gArrowMM), "5")
    UiAuto f, x, y, "tglArrowAuto", (gArrowMM <= 0)
    y = y + UI_ROW + UI_RGAP
    UiRowLabel f, x, y, "Rang", "R,G,B, masalan 0,90,200"
    UiFieldBg f, x + UI_CTL, y, 100, CCard()
    Set c = f.Controls.Add("Forms.Label.1", "lblSwatch", True)
    c.Left = x + UI_CTL + 8: c.Top = y + (UI_ROW - 12) / 2: c.Width = 12: c.Height = 12
    c.BackStyle = 0
    Set c = UiText(f, "txtColor", x + UI_CTL + 24, y, 72, gColR & "," & gColG & "," & gColB, 1)
    c.ControlTipText = UZ("R,G,B, masalan 0,90,200")
    UiButton f, "btnColor", "Tanlash", x + UI_CW - UI_PAD - 60, y, 60, False, CCard(), "Rang tanlash oynasi"

    ' --- o'ng ustun: Ko'rsatish ---
    x = UI_M + UI_CW + UI_GAP: y0 = UI_M
    h3 = UI_PAD * 2 + UI_TITLE + UI_TGAP + UI_ROW * 5 + UI_RGAP * 4
    UiCard f, x, y0, UI_CW, h3, "Ko'rsatish"
    y = y0 + UI_PAD + UI_TITLE + UI_TGAP
    UiToggleRow f, x, y, "tglSplit", "Harflarga ajratish", gSplitText, _
                "Matn va egri chiziqdagi har bir harf alohida o'lchanadi"
    y = y + UI_ROW + UI_RGAP
    UiToggleRow f, x, y, "tglItems", "Eni va bo'yi", gShowItems, "Har bir harfning eni va bo'yi"
    y = y + UI_ROW + UI_RGAP
    UiToggleRow f, x, y, "tglOverall", "Umumiy o'lcham", gShowOverall, "Butun tanlovning eni va bo'yi"
    y = y + UI_ROW + UI_RGAP
    UiToggleRow f, x, y, "tglOutline", "Kontur qalinligi", gUseOutline, "Kontur qalinligi o'lchamga qo'shiladi"
    y = y + UI_ROW + UI_RGAP
    UiToggleRow f, x, y, "tglRotate", "Vertikal raqam", gRotateVertical, "Bo'yi raqami 90 gradus burilib yoziladi"

    ' --- o'ng ustun: amallar (pastki chegara chap ustun bilan bir xil) ---
    y0 = UI_M + h3 + UI_GAP
    UiCard f, x, y0, UI_CW, UI_M + cardH * 2 + UI_GAP - y0, ""
    y = y0 + UI_PAD
    UiButton f, "btnRun", "O'lchash", x + UI_PAD, y, 120, True, CCard(), "Tanlangan obyektlarni o'lchash (Enter)"
    UiButton f, "btnClear", "Tozalash", x + UI_PAD + 126, y, 60, False, CCard(), "Sahifadagi barcha Size o'lchamlarini o'chirish"
    UiButton f, "btnClose", "Yopish", x + UI_CW - UI_PAD - 60, y, 60, False, CCard(), ""
    Set c = f.Controls.Add("Forms.Label.1", "lblStatus", True)
    c.Left = x + UI_PAD: c.Top = y + UI_ROW + 8: c.Width = UI_CW - UI_PAD * 2: c.Height = 14
    c.BackStyle = 0
    c.Caption = ""
    UiFont c, False, 8, CMuted()

    ' Enter tugmasi uchun ko'rinmas tugma
    Set c = f.Controls.Add("Forms.CommandButton.1", "btnEnter", True)
    c.Left = -100: c.Top = -100: c.Width = 10: c.Height = 10
    c.Default = True
    c.TabStop = False

    SizeFormSync f
    Exit Sub
Failed:
    MsgBox "Size: oynani tayyorlashda xatolik." & vbCr & Err.Description, vbExclamation, "Size"
End Sub

' Avto holati: maydon va "+"/"-" kulrang; rang namunasi yangilanadi
Public Sub SizeFormSync(f As Object)
    Dim r As Long, g As Long, b As Long
    On Error Resume Next
    SyncStepper f, "Font", "txtFontPt", TogVal(f, "tglFontAuto")
    SyncStepper f, "Line", "txtLine", TogVal(f, "tglLineAuto")
    SyncStepper f, "Arrow", "txtArrow", TogVal(f, "tglArrowAuto")
    If ParseColor(f.Controls("txtColor").Text, r, g, b) Then
        Set f.Controls("lblSwatch").Picture = UiPic("dot", 12, 12, 6, VBA.RGB(r, g, b), CLine(), CField(), False, 0)
    End If
End Sub

Private Sub SyncStepper(f As Object, ByVal key As String, ByVal txtName As String, ByVal isAuto As Boolean)
    Dim col As Long
    If isAuto Then col = CDim() Else col = CText()
    f.Controls(txtName).ForeColor = col
    f.Controls("btn" & key & "Minus").ForeColor = col
    f.Controls("btn" & key & "Plus").ForeColor = col
End Sub

' CorelDRAW'ning rang tanlash oynasi
Public Sub SizeFormPickColor(f As Object)
    Dim c As Object, r As Long, g As Long, b As Long
    If Not ParseColor(f.Controls("txtColor").Text, r, g, b) Then
        r = gColR: g = gColG: b = gColB
    End If
    On Error GoTo NoDialog
    Set c = CreateRGBColor(r, g, b)
    If c.UserAssign() Then
        c.ConvertToRGB
        f.Controls("txtColor").Text = c.RGBRed & "," & c.RGBGreen & "," & c.RGBBlue
    End If
    Exit Sub
NoDialog:
    SetStatus f, "Rang oynasi ochilmadi. R,G,B yozing: 0,90,200", False
End Sub

' Oynadagi qiymatlarni tekshirib saqlaydi (quiet = xato bo'lsa xabar chiqarmaydi)
Public Function SizeFormSave(f As Object, Optional ByVal quiet As Boolean = False) As Boolean
    Dim fontPt As Double, lineMM As Double, arrowMM As Double
    Dim nm As String, r As Long, g As Long, b As Long, er As String
    On Error GoTo Failed

    If TogVal(f, "tglFontAuto") Then
        fontPt = 0
    Else
        fontPt = ParseNum(f.Controls("txtFontPt").Text)
        If fontPt < 1 Or fontPt > 3000 Then er = "Hajm: 1--3000 pt"
    End If
    If TogVal(f, "tglLineAuto") Then
        lineMM = 0
    Else
        lineMM = ParseNum(f.Controls("txtLine").Text)
        If (lineMM < 0.01 Or lineMM > 50) And er = "" Then er = "Qalinlik: 0.01--50 mm"
    End If
    If TogVal(f, "tglArrowAuto") Then
        arrowMM = 0
    Else
        arrowMM = ParseNum(f.Controls("txtArrow").Text)
        If (arrowMM < 0.1 Or arrowMM > 500) And er = "" Then er = "Strelka: 0.1--500 mm"
    End If
    If Not ParseColor(f.Controls("txtColor").Text, r, g, b) And er = "" Then er = "Rang: R,G,B, masalan 0,90,200"
    If er <> "" Then
        If Not quiet Then SetStatus f, er, False
        Exit Function
    End If
    nm = Trim$(f.Controls("cboFont").Text)
    If nm = "" Then nm = "Arial"

    gFontPt = fontPt
    gLineMM = lineMM
    gArrowMM = arrowMM
    gFontName = ResolveFont(nm)
    gBold = TogVal(f, "tglBold")
    gColR = r: gColG = g: gColB = b
    gSplitText = TogVal(f, "tglSplit")
    gShowItems = TogVal(f, "tglItems")
    gShowOverall = TogVal(f, "tglOverall")
    gUseOutline = TogVal(f, "tglOutline")
    gRotateVertical = TogVal(f, "tglRotate")
    SaveSettingsAll
    SizeFormSave = True
    Exit Function
Failed:
    If Not quiet Then SetStatus f, "Saqlanmadi: " & Err.Description, False
End Function

' --- dizayn elementlari ---

Private Sub UiFont(c As Object, ByVal semibold As Boolean, ByVal sz As Single, ByVal col As Long)
    If semibold Then c.Font.Name = UI_FONT_B Else c.Font.Name = UI_FONT
    c.Font.Size = sz
    c.Font.Bold = False
    c.ForeColor = col
End Sub

' Oq karta (yumaloq, yengil chegara va soya) va sarlavha
Private Sub UiCard(f As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                   ByVal h As Single, ByVal title As String)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1")
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
    c.BackStyle = 0
    c.PicturePosition = 12                          ' markazda
    Set c.Picture = UiPic("card", w, h, UI_CR, CCard(), CLine(), CBg(), True, 0)
    If title <> "" Then
        Set c = f.Controls.Add("Forms.Label.1")
        c.Left = x + UI_PAD: c.Top = y + UI_PAD - 1: c.Width = w - UI_PAD * 2: c.Height = UI_TITLE
        c.BackStyle = 0
        c.Caption = UZ(title)
        UiFont c, True, 10, CText()
    End If
End Sub

Private Sub UiRowLabel(f As Object, ByVal cardX As Single, ByVal y As Single, ByVal txt As String, ByVal tip As String)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1")
    c.Left = cardX + UI_PAD: c.Top = y + (UI_ROW - 14) / 2: c.Width = UI_CTL - UI_PAD - 4: c.Height = 14
    c.BackStyle = 0
    c.Caption = UZ(txt)
    If tip <> "" Then c.ControlTipText = UZ(tip)
    UiFont c, False, 9, CText()
End Sub

' Kulrang yumaloq maydon foni
Private Sub UiFieldBg(f As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal bg As Long)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1")
    c.Left = x: c.Top = y: c.Width = w: c.Height = UI_ROW
    c.BackStyle = 0
    c.PicturePosition = 12
    Set c.Picture = UiPic("field", w, UI_ROW, UI_R, CField(), CField(), bg, False, 0)
End Sub

' Maydon ichidagi shaffof matn kiritish qutisi
Private Function UiText(f As Object, ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                        ByVal w As Single, ByVal v As String, ByVal align As Long) As Object
    Dim c As Object
    Set c = f.Controls.Add("Forms.TextBox.1", nm, True)
    c.Left = x: c.Top = y + (UI_ROW - 18) / 2: c.Width = w: c.Height = 18
    c.SpecialEffect = 0: c.BorderStyle = 0
    c.BackStyle = 0
    c.TextAlign = align                             ' 1 = chap, 2 = markaz
    UiFont c, False, 9, CText()
    c.Text = v
    Set UiText = c
End Function

' Sonli maydon:  [ -   qiymat   + ]
Private Sub UiStepper(f As Object, ByVal x As Single, ByVal y As Single, ByVal key As String, ByVal v As String)
    Const W As Single = 100, B As Single = 24
    Dim txtName As String
    UiFieldBg f, x, y, W, CCard()
    UiStepButton f, "btn" & key & "Minus", ChrW(&H2212), x, y, B
    UiStepButton f, "btn" & key & "Plus", "+", x + W - B, y, B
    If key = "Font" Then txtName = "txtFontPt" Else txtName = "txt" & key
    UiText f, txtName, x + B, y, W - B * 2, v, 2
End Sub

Private Sub UiStepButton(f As Object, ByVal nm As String, ByVal cap As String, ByVal x As Single, _
                         ByVal y As Single, ByVal w As Single)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1", nm, True)
    c.Left = x: c.Top = y + 3: c.Width = w: c.Height = UI_ROW - 6
    c.BackStyle = 0
    c.TextAlign = 2
    c.Caption = cap
    UiFont c, False, 12, CText()
End Sub

' "Avto" va almashtirgich qator oxirida
Private Sub UiAuto(f As Object, ByVal cardX As Single, ByVal y As Single, ByVal nm As String, ByVal isOn As Boolean)
    Dim c As Object, tx As Single
    tx = cardX + UI_CW - UI_PAD - UI_TGW
    Set c = f.Controls.Add("Forms.Label.1")
    c.Left = tx - 34: c.Top = y + (UI_ROW - 14) / 2: c.Width = 30: c.Height = 14
    c.BackStyle = 0
    c.TextAlign = 3                                 ' o'ngga
    c.Caption = "Avto"
    c.ControlTipText = UZ("Harflar o'lchamiga qarab avtomatik")
    UiFont c, False, 8, CMuted()
    UiToggle f, nm, tx, y, isOn
End Sub

Private Sub UiToggleRow(f As Object, ByVal cardX As Single, ByVal y As Single, ByVal nm As String, _
                        ByVal txt As String, ByVal isOn As Boolean, ByVal tip As String)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1")
    c.Left = cardX + UI_PAD: c.Top = y + (UI_ROW - 14) / 2: c.Width = UI_CW - UI_PAD * 2 - UI_TGW - 8: c.Height = 14
    c.BackStyle = 0
    c.Caption = UZ(txt)
    c.ControlTipText = UZ(tip)
    UiFont c, False, 9, CText()
    UiToggle f, nm, cardX + UI_CW - UI_PAD - UI_TGW, y, isOn
    f.Controls(nm).ControlTipText = UZ(tip)
End Sub

' Almashtirgich (toggle): holati Tag da ("1" / "0")
Private Sub UiToggle(f As Object, ByVal nm As String, ByVal x As Single, ByVal y As Single, ByVal isOn As Boolean)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1", nm, True)
    c.Left = x: c.Top = y + (UI_ROW - UI_TGH) / 2: c.Width = UI_TGW: c.Height = UI_TGH
    c.BackStyle = 0
    c.PicturePosition = 12
    SetToggle f, nm, isOn
End Sub

Private Sub SetToggle(f As Object, ByVal nm As String, ByVal isOn As Boolean)
    With f.Controls(nm)
        If isOn Then
            .Tag = "1"
            Set .Picture = UiPic("tgl", UI_TGW, UI_TGH, UI_TGH / 2, COn(), COn(), CCard(), False, 1)
        Else
            .Tag = "0"
            Set .Picture = UiPic("tgl", UI_TGW, UI_TGH, UI_TGH / 2, COff(), COff(), CCard(), False, -1)
        End If
    End With
End Sub

Private Function TogVal(f As Object, ByVal nm As String) As Boolean
    TogVal = (f.Controls(nm).Tag = "1")
End Function

' Tugma: yumaloq fon rasmi va markazdagi yozuv
Private Sub UiButton(f As Object, ByVal nm As String, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                     ByVal w As Single, ByVal primary As Boolean, ByVal bg As Long, ByVal tip As String)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1", nm, True)
    c.Left = x: c.Top = y: c.Width = w: c.Height = UI_ROW
    c.BackStyle = 0
    c.PicturePosition = 12
    c.TextAlign = 2
    c.Caption = UZ(txt)
    If tip <> "" Then c.ControlTipText = UZ(tip)
    If primary Then
        Set c.Picture = UiPic("btn", w, UI_ROW, UI_R, CPrimary(), CPrimary(), bg, False, 0)
        UiFont c, True, 9, VBA.RGB(255, 255, 255)
    Else
        Set c.Picture = UiPic("btn", w, UI_ROW, UI_R, CField(), CField(), bg, False, 0)
        UiFont c, False, 9, CText()
    End If
End Sub

' --- yumaloq shakllar rasmi (24-bit BMP, silliq chetlar) ---

' knob: 0 = yo'q, 1 = o'ngda (yoqilgan), -1 = chapda (o'chirilgan)
Private Function UiPic(ByVal kind As String, ByVal wPt As Single, ByVal hPt As Single, ByVal rPt As Single, _
                       ByVal fill As Long, ByVal border As Long, ByVal bg As Long, _
                       ByVal shadow As Boolean, ByVal knob As Long) As Object
    Dim dpi As Long, w As Long, h As Long, r As Double, path As String
    dpi = ScreenDpi()
    w = CLng(wPt * dpi / 72): h = CLng(hPt * dpi / 72)
    If w < 2 Then w = 2
    If h < 2 Then h = 2
    r = rPt * dpi / 72
    path = UiDir() & "\" & kind & "_" & w & "x" & h & "_" & CLng(r * 10) & "_" & Hex$(fill) & "_" & _
           Hex$(border) & "_" & Hex$(bg) & "_" & IIf(shadow, "s", "n") & knob & "_v1.bmp"
    If Len(Dir$(path)) = 0 Then WriteShapeBmp path, w, h, r, fill, border, bg, shadow, knob, dpi
    Set UiPic = LoadPicture(path)
End Function

Private Function UiDir() As String
    Dim p As String
    p = Environ$("TEMP")
    If p = "" Then p = Environ$("TMP")
    p = p & "\SizeUI"
    On Error Resume Next
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    UiDir = p
End Function

Private Function ScreenDpi() As Long
    Static cached As Long
#If VBA7 Then
    Dim hdc As LongPtr
#Else
    Dim hdc As Long
#End If
    If cached > 0 Then ScreenDpi = cached: Exit Function
    cached = 96
    On Error Resume Next
    hdc = GetDC(0)
    If hdc <> 0 Then
        cached = GetDeviceCaps(hdc, 88)             ' LOGPIXELSX
        ReleaseDC 0, hdc
    End If
    If cached < 72 Or cached > 480 Then cached = 96
    ScreenDpi = cached
End Function

' Rasmni chizish: fon -> soya -> chegara -> to'ldirish -> (dumaloq tutqich)
Private Sub WriteShapeBmp(ByVal path As String, ByVal w As Long, ByVal h As Long, ByVal r As Double, _
                          ByVal fill As Long, ByVal border As Long, ByVal bg As Long, _
                          ByVal shadow As Boolean, ByVal knob As Long, ByVal dpi As Long)
    Dim buf() As Byte, rowLen As Long, x As Long, y As Long, i As Long, fn As Integer
    Dim px As Double, py As Double, x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim cr As Double, cg As Double, cb As Double, a As Double, d As Double, bw As Double
    Dim kx As Double, ky As Double, kr As Double, ppm As Long
    Dim bgR As Double, bgG As Double, bgB As Double, fR As Byte, fG As Byte, fB As Byte
    Dim bandY As Boolean, bandX0 As Double, bandX1 As Double, inX0 As Double, inX1 As Double
    Dim inY0 As Double, inY1 As Double, bandYOk As Boolean

    rowLen = ((w * 3 + 3) \ 4) * 4
    ReDim buf(0 To 54 + rowLen * h - 1)
    ppm = CLng(dpi / 0.0254)
    PutLong buf, 0, 0: buf(0) = 66: buf(1) = 77                ' "BM"
    PutLong buf, 2, 54 + rowLen * h
    PutLong buf, 10, 54
    PutLong buf, 14, 40
    PutLong buf, 18, w
    PutLong buf, 22, h
    buf(26) = 1: buf(28) = 24
    PutLong buf, 34, rowLen * h
    PutLong buf, 38, ppm
    PutLong buf, 42, ppm

    ' shakl chegaralari (soya uchun joy qoldiriladi)
    If shadow Then
        x0 = 1.5: y0 = 1: x1 = w - 1.5: y1 = h - 2.5
    Else
        x0 = 0: y0 = 0: x1 = w: y1 = h
    End If
    If r > (y1 - y0) / 2 Then r = (y1 - y0) / 2
    bw = dpi / 96#                                  ' 1 px chegara (96 DPI da)
    If border = fill Then bw = 0
    If knob <> 0 Then
        kr = (y1 - y0) / 2 - 2 * dpi / 96#
        ky = (y0 + y1) / 2
        If knob > 0 Then kx = x1 - (y1 - y0) / 2 Else kx = x0 + (y1 - y0) / 2
    End If

    ' tezlashtirish: shaklning ichki qismi bir xil rangda, faqat chetlar hisoblanadi
    bgR = ColR(bg): bgG = ColG(bg): bgB = ColB(bg)
    fR = CByte(ColR(fill)): fG = CByte(ColG(fill)): fB = CByte(ColB(fill))
    inX0 = x0 + bw + 1.5: inX1 = x1 - bw - 1.5          ' to'g'ri chetlardan ichkari
    inY0 = y0 + bw + 1.5: inY1 = y1 - bw - 1.5
    bandX0 = x0 + r + 1: bandX1 = x1 - r - 1            ' burchaklardan tashqari
    For y = 0 To h - 1
        py = y + 0.5
        bandY = (knob = 0 And py > y0 + r + 1 And py < y1 - r - 1)
        bandYOk = (knob = 0 And py > inY0 And py < inY1)
        For x = 0 To w - 1
            px = x + 0.5
            i = 54 + (h - 1 - y) * rowLen + x * 3
            If (bandY And px > inX0 And px < inX1) Or (bandYOk And px > bandX0 And px < bandX1) Then
                buf(i) = fB: buf(i + 1) = fG: buf(i + 2) = fR
                GoTo NextPixel
            End If
            cr = bgR: cg = bgG: cb = bgB
            If shadow Then
                d = SdRoundRect(px, py - 1, x0, y0, x1, y1, r)
                If d < 3 Then
                    a = 0.06 * (1 - Max2(d, 0) / 3)
                    cr = cr * (1 - a): cg = cg * (1 - a): cb = cb * (1 - a)
                End If
            End If
            d = SdRoundRect(px, py, x0, y0, x1, y1, r)
            a = Clamp01(0.5 - d)
            If a > 0 Then
                If bw > 0 Then
                    Blend cr, cg, cb, border, a
                    a = Clamp01(0.5 - (d + bw))
                End If
                If a > 0 Then Blend cr, cg, cb, fill, a
            End If
            If knob <> 0 Then
                d = Sqr((px - kx) ^ 2 + (py - ky) ^ 2) - kr
                a = Clamp01(0.5 - d)
                If a > 0 Then Blend cr, cg, cb, VBA.RGB(255, 255, 255), a
            End If
            buf(i) = CByte(Int(cb + 0.5)): buf(i + 1) = CByte(Int(cg + 0.5)): buf(i + 2) = CByte(Int(cr + 0.5))
NextPixel:
        Next x
    Next y

    On Error Resume Next
    fn = FreeFile
    Open path For Binary Access Write As #fn
    Put #fn, 1, buf
    Close #fn
End Sub

' Yumaloq to'rtburchakgacha masofa (ichida manfiy)
Private Function SdRoundRect(ByVal px As Double, ByVal py As Double, ByVal x0 As Double, ByVal y0 As Double, _
                             ByVal x1 As Double, ByVal y1 As Double, ByVal r As Double) As Double
    Dim qx As Double, qy As Double, ox As Double, oy As Double
    qx = Abs(px - (x0 + x1) / 2) - ((x1 - x0) / 2 - r)
    qy = Abs(py - (y0 + y1) / 2) - ((y1 - y0) / 2 - r)
    ox = Max2(qx, 0): oy = Max2(qy, 0)
    SdRoundRect = Sqr(ox * ox + oy * oy) + Min2(Max2(qx, qy), 0) - r
End Function

Private Sub Blend(cr As Double, cg As Double, cb As Double, ByVal col As Long, ByVal a As Double)
    cr = cr + (ColR(col) - cr) * a
    cg = cg + (ColG(col) - cg) * a
    cb = cb + (ColB(col) - cb) * a
End Sub

Private Function Clamp01(ByVal v As Double) As Double
    If v < 0 Then v = 0
    If v > 1 Then v = 1
    Clamp01 = v
End Function

Private Function ColR(ByVal c As Long) As Double
    ColR = c And &HFF&
End Function
Private Function ColG(ByVal c As Long) As Double
    ColG = (c \ &H100&) And &HFF&
End Function
Private Function ColB(ByVal c As Long) As Double
    ColB = (c \ &H10000) And &HFF&
End Function

Private Sub PutLong(b() As Byte, ByVal off As Long, ByVal v As Long)
    b(off) = v And &HFF&
    b(off + 1) = (v \ &H100&) And &HFF&
    b(off + 2) = (v \ &H10000) And &HFF&
    b(off + 3) = (v \ &H1000000) And &HFF&
End Sub

' Shriftlar ro'yxatini CorelDRAW'dan bir marta o'qiydi va saralaydi
Private Sub LoadFontCache()
    Dim app As Object, fl As Object, n As Long, i As Long, k As Long
    If gFontsOk Then Exit Sub
    gFontsOk = True
    gFontsN = 0
    On Error Resume Next
    Set app = Application
    Set fl = app.FontList
    n = 0
    n = fl.Count
    If n <= 0 Then Exit Sub
    ReDim gFonts(0 To n - 1)
    For i = 1 To n
        gFonts(gFontsN) = CStr(fl(i))
        If Len(gFonts(gFontsN)) > 0 Then gFontsN = gFontsN + 1
    Next i
    If gFontsN = 0 Then Exit Sub
    QuickSortStr gFonts, 0, gFontsN - 1
    ' takrorlarni olib tashlash
    k = 0
    For i = 1 To gFontsN - 1
        If StrComp(gFonts(i), gFonts(k), vbTextCompare) <> 0 Then
            k = k + 1
            gFonts(k) = gFonts(i)
        End If
    Next i
    gFontsN = k + 1
    ReDim Preserve gFonts(0 To gFontsN - 1)
End Sub

Private Sub QuickSortStr(a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pv As String, tmp As String
    Do While lo < hi
        i = lo: j = hi
        pv = a((lo + hi) \ 2)
        Do While i <= j
            Do While StrComp(a(i), pv, vbTextCompare) < 0
                i = i + 1
            Loop
            Do While StrComp(a(j), pv, vbTextCompare) > 0
                j = j - 1
            Loop
            If i <= j Then
                tmp = a(i): a(i) = a(j): a(j) = tmp
                i = i + 1: j = j - 1
            End If
        Loop
        ' kichik qism rekursiya bilan, kattasi sikl bilan
        If j - lo < hi - i Then
            If lo < j Then QuickSortStr a, lo, j
            lo = i
        Else
            If i < hi Then QuickSortStr a, i, hi
            hi = j
        End If
    Loop
End Sub

' Ro'yxat bo'sh bo'lsa (oyna endi ochilgan) - to'ldiradi
Private Sub FillFontsDeferred(f As Object)
    Dim cbo As Object
    On Error Resume Next
    Set cbo = f.Controls("cboFont")
    If cbo Is Nothing Then Exit Sub
    If cbo.ListCount > 1 Then Exit Sub
    DoEvents                            ' oyna avval chizilib olsin
    FillFonts cbo
End Sub

' O'rnatilgan shriftlar ro'yxati (alifbo tartibida, bir martada)
Private Sub FillFonts(cbo As Object)
    Dim cur As String
    On Error Resume Next
    cur = cbo.Text
    LoadFontCache
    If gFontsN > 0 Then
        cbo.List = gFonts
    Else
        cbo.Clear
        cbo.AddItem "Arial"
        cbo.AddItem "Arial Narrow"
        cbo.AddItem "Calibri"
        cbo.AddItem "Segoe UI"
        cbo.AddItem "Tahoma"
        cbo.AddItem "Times New Roman"
        cbo.AddItem "Verdana"
    End If
    cbo.Text = cur
End Sub

' Shrift kompyuterda bormi; nomini to'g'ri yozilishi bilan qaytaradi.
' Topilmasa - Arial. (saralangan ro'yxatda ikkilik qidiruv)
Private Function ResolveFont(ByVal nm As String) As String
    Dim lo As Long, hi As Long, m As Long, c As Long
    ResolveFont = nm
    LoadFontCache
    If gFontsN <= 0 Then Exit Function          ' ro'yxatni o'qib bo'lmadi - boricha qoldiramiz
    lo = 0: hi = gFontsN - 1
    Do While lo <= hi
        m = (lo + hi) \ 2
        c = StrComp(gFonts(m), nm, vbTextCompare)
        If c = 0 Then
            ResolveFont = gFonts(m)
            Exit Function
        ElseIf c < 0 Then
            lo = m + 1
        Else
            hi = m - 1
        End If
    Loop
    ResolveFont = "Arial"
End Function

' Oyna (UserForm) moduli kodi - SizeForm.frm bilan bir xil. SizeForm import
' qilinmagan bo'lsa, oyna VBA muharriri orqali shu kod bilan yaratiladi.
Private Function FormCode() As String
    Dim s As String
    s = "' " & FORM_TAG & vbCrLf
    s = s & "' Size oynasi. Oyna bo'sh: barcha elementlar oyna ochilganda" & vbCrLf
    s = s & "' SizePlugin.SizeFormBuild orqali yaratiladi, bosishlar SizePlugin.SizeUIClick ga uzatiladi." & vbCrLf
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents e_btnRun As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnClear As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnClose As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnColor As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnFontMinus As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnFontPlus As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnLineMinus As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnLinePlus As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnArrowMinus As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnArrowPlus As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglBold As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglFontAuto As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglLineAuto As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglArrowAuto As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglSplit As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglItems As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglOverall As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglOutline As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_tglRotate As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents e_btnEnter As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents e_txtColor As MSForms.TextBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    SizePlugin.SizeFormBuild Me" & vbCrLf
    s = s & "    Set e_btnRun = Me.Controls(""btnRun"")" & vbCrLf
    s = s & "    Set e_btnClear = Me.Controls(""btnClear"")" & vbCrLf
    s = s & "    Set e_btnClose = Me.Controls(""btnClose"")" & vbCrLf
    s = s & "    Set e_btnColor = Me.Controls(""btnColor"")" & vbCrLf
    s = s & "    Set e_btnFontMinus = Me.Controls(""btnFontMinus"")" & vbCrLf
    s = s & "    Set e_btnFontPlus = Me.Controls(""btnFontPlus"")" & vbCrLf
    s = s & "    Set e_btnLineMinus = Me.Controls(""btnLineMinus"")" & vbCrLf
    s = s & "    Set e_btnLinePlus = Me.Controls(""btnLinePlus"")" & vbCrLf
    s = s & "    Set e_btnArrowMinus = Me.Controls(""btnArrowMinus"")" & vbCrLf
    s = s & "    Set e_btnArrowPlus = Me.Controls(""btnArrowPlus"")" & vbCrLf
    s = s & "    Set e_tglBold = Me.Controls(""tglBold"")" & vbCrLf
    s = s & "    Set e_tglFontAuto = Me.Controls(""tglFontAuto"")" & vbCrLf
    s = s & "    Set e_tglLineAuto = Me.Controls(""tglLineAuto"")" & vbCrLf
    s = s & "    Set e_tglArrowAuto = Me.Controls(""tglArrowAuto"")" & vbCrLf
    s = s & "    Set e_tglSplit = Me.Controls(""tglSplit"")" & vbCrLf
    s = s & "    Set e_tglItems = Me.Controls(""tglItems"")" & vbCrLf
    s = s & "    Set e_tglOverall = Me.Controls(""tglOverall"")" & vbCrLf
    s = s & "    Set e_tglOutline = Me.Controls(""tglOutline"")" & vbCrLf
    s = s & "    Set e_tglRotate = Me.Controls(""tglRotate"")" & vbCrLf
    s = s & "    Set e_btnEnter = Me.Controls(""btnEnter"")" & vbCrLf
    s = s & "    Set e_txtColor = Me.Controls(""txtColor"")" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnRun_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnRun""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnClear_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnClear""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnClose_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnClose""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnColor_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnColor""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnFontMinus_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnFontMinus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnFontPlus_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnFontPlus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnLineMinus_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnLineMinus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnLinePlus_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnLinePlus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnArrowMinus_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnArrowMinus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnArrowPlus_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnArrowPlus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglBold_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglBold""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglFontAuto_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglFontAuto""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglLineAuto_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglLineAuto""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglArrowAuto_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglArrowAuto""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglSplit_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglSplit""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglItems_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglItems""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglOverall_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglOverall""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglOutline_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglOutline""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_tglRotate_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""tglRotate""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnFontMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnFontMinus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnFontPlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnFontPlus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnLineMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnLineMinus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnLinePlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnLinePlus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnArrowMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnArrowMinus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnArrowPlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnArrowPlus""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_btnEnter_Click()" & vbCrLf
    s = s & "    SizePlugin.SizeUIClick Me, ""btnRun""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub e_txtColor_Change()" & vbCrLf
    s = s & "    SizePlugin.SizeFormSync Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        SizePlugin.SizeFormClose Me" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCode = s
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
    Dim t As String
    t = Trim$(Replace(Str$(v), ",", "."))
    If Left$(t, 1) = "." Then t = "0" & t           ' Str$(0.5) = " .5"
    FormatNum = t
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
