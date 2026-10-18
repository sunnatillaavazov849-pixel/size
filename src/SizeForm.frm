VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SizeForm 
   Caption         =   "Size"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "SizeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SizeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Size oynasi. Oyna bo'sh: barcha elementlar oyna ochilganda
' SizePlugin.SizeFormBuild orqali yaratiladi, bosishlar SizePlugin.SizeUIClick ga uzatiladi.
Option Explicit

Private WithEvents e_btnRun As MSForms.Label
Private WithEvents e_btnClear As MSForms.Label
Private WithEvents e_btnClose As MSForms.Label
Private WithEvents e_btnColor As MSForms.Label
Private WithEvents e_btnFontMinus As MSForms.Label
Private WithEvents e_btnFontPlus As MSForms.Label
Private WithEvents e_btnLineMinus As MSForms.Label
Private WithEvents e_btnLinePlus As MSForms.Label
Private WithEvents e_btnArrowMinus As MSForms.Label
Private WithEvents e_btnArrowPlus As MSForms.Label
Private WithEvents e_tglBold As MSForms.Label
Private WithEvents e_tglFontAuto As MSForms.Label
Private WithEvents e_tglLineAuto As MSForms.Label
Private WithEvents e_tglArrowAuto As MSForms.Label
Private WithEvents e_tglSplit As MSForms.Label
Private WithEvents e_tglItems As MSForms.Label
Private WithEvents e_tglOverall As MSForms.Label
Private WithEvents e_tglOutline As MSForms.Label
Private WithEvents e_tglRotate As MSForms.Label
Private WithEvents e_btnEnter As MSForms.CommandButton
Private WithEvents e_txtColor As MSForms.TextBox

Private Sub UserForm_Initialize()
    SizePlugin.SizeFormBuild Me
    Set e_btnRun = Me.Controls("btnRun")
    Set e_btnClear = Me.Controls("btnClear")
    Set e_btnClose = Me.Controls("btnClose")
    Set e_btnColor = Me.Controls("btnColor")
    Set e_btnFontMinus = Me.Controls("btnFontMinus")
    Set e_btnFontPlus = Me.Controls("btnFontPlus")
    Set e_btnLineMinus = Me.Controls("btnLineMinus")
    Set e_btnLinePlus = Me.Controls("btnLinePlus")
    Set e_btnArrowMinus = Me.Controls("btnArrowMinus")
    Set e_btnArrowPlus = Me.Controls("btnArrowPlus")
    Set e_tglBold = Me.Controls("tglBold")
    Set e_tglFontAuto = Me.Controls("tglFontAuto")
    Set e_tglLineAuto = Me.Controls("tglLineAuto")
    Set e_tglArrowAuto = Me.Controls("tglArrowAuto")
    Set e_tglSplit = Me.Controls("tglSplit")
    Set e_tglItems = Me.Controls("tglItems")
    Set e_tglOverall = Me.Controls("tglOverall")
    Set e_tglOutline = Me.Controls("tglOutline")
    Set e_tglRotate = Me.Controls("tglRotate")
    Set e_btnEnter = Me.Controls("btnEnter")
    Set e_txtColor = Me.Controls("txtColor")
End Sub

Private Sub e_btnRun_Click()
    SizePlugin.SizeUIClick Me, "btnRun"
End Sub

Private Sub e_btnClear_Click()
    SizePlugin.SizeUIClick Me, "btnClear"
End Sub

Private Sub e_btnClose_Click()
    SizePlugin.SizeUIClick Me, "btnClose"
End Sub

Private Sub e_btnColor_Click()
    SizePlugin.SizeUIClick Me, "btnColor"
End Sub

Private Sub e_btnFontMinus_Click()
    SizePlugin.SizeUIClick Me, "btnFontMinus"
End Sub

Private Sub e_btnFontPlus_Click()
    SizePlugin.SizeUIClick Me, "btnFontPlus"
End Sub

Private Sub e_btnLineMinus_Click()
    SizePlugin.SizeUIClick Me, "btnLineMinus"
End Sub

Private Sub e_btnLinePlus_Click()
    SizePlugin.SizeUIClick Me, "btnLinePlus"
End Sub

Private Sub e_btnArrowMinus_Click()
    SizePlugin.SizeUIClick Me, "btnArrowMinus"
End Sub

Private Sub e_btnArrowPlus_Click()
    SizePlugin.SizeUIClick Me, "btnArrowPlus"
End Sub

Private Sub e_tglBold_Click()
    SizePlugin.SizeUIClick Me, "tglBold"
End Sub

Private Sub e_tglFontAuto_Click()
    SizePlugin.SizeUIClick Me, "tglFontAuto"
End Sub

Private Sub e_tglLineAuto_Click()
    SizePlugin.SizeUIClick Me, "tglLineAuto"
End Sub

Private Sub e_tglArrowAuto_Click()
    SizePlugin.SizeUIClick Me, "tglArrowAuto"
End Sub

Private Sub e_tglSplit_Click()
    SizePlugin.SizeUIClick Me, "tglSplit"
End Sub

Private Sub e_tglItems_Click()
    SizePlugin.SizeUIClick Me, "tglItems"
End Sub

Private Sub e_tglOverall_Click()
    SizePlugin.SizeUIClick Me, "tglOverall"
End Sub

Private Sub e_tglOutline_Click()
    SizePlugin.SizeUIClick Me, "tglOutline"
End Sub

Private Sub e_tglRotate_Click()
    SizePlugin.SizeUIClick Me, "tglRotate"
End Sub

Private Sub e_btnFontMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SizePlugin.SizeUIClick Me, "btnFontMinus"
End Sub

Private Sub e_btnFontPlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SizePlugin.SizeUIClick Me, "btnFontPlus"
End Sub

Private Sub e_btnLineMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SizePlugin.SizeUIClick Me, "btnLineMinus"
End Sub

Private Sub e_btnLinePlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SizePlugin.SizeUIClick Me, "btnLinePlus"
End Sub

Private Sub e_btnArrowMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SizePlugin.SizeUIClick Me, "btnArrowMinus"
End Sub

Private Sub e_btnArrowPlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SizePlugin.SizeUIClick Me, "btnArrowPlus"
End Sub

Private Sub e_btnEnter_Click()
    SizePlugin.SizeUIClick Me, "btnRun"
End Sub

Private Sub e_txtColor_Change()
    SizePlugin.SizeFormSync Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        SizePlugin.SizeFormClose Me
    End If
End Sub
