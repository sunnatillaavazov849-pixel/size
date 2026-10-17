VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SizeForm 
   Caption         =   "Size"
   ClientHeight    =   9200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6800
   OleObjectBlob   =   "SizeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SizeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Size plaginining oynasi. Oyna bo'sh: barcha tugma va maydonlar oyna
' ochilganda SizePlugin.SizeFormBuild orqali yaratiladi.
' SizePlugin.bas bilan birga import qiling (SizeForm.frx shu papkada bo'lsin).
Option Explicit

Private WithEvents mRun As MSForms.CommandButton
Private WithEvents mClear As MSForms.CommandButton
Private WithEvents mClose As MSForms.CommandButton
Private WithEvents mColor As MSForms.CommandButton
Private WithEvents mFontAuto As MSForms.CheckBox
Private WithEvents mLineAuto As MSForms.CheckBox
Private WithEvents mArrowAuto As MSForms.CheckBox
Private WithEvents mColorTxt As MSForms.TextBox

Private Sub UserForm_Initialize()
    SizePlugin.SizeFormBuild Me
    Set mRun = Me.Controls("btnRun")
    Set mClear = Me.Controls("btnClear")
    Set mClose = Me.Controls("btnClose")
    Set mColor = Me.Controls("btnColor")
    Set mFontAuto = Me.Controls("chkFontAuto")
    Set mLineAuto = Me.Controls("chkLineAuto")
    Set mArrowAuto = Me.Controls("chkArrowAuto")
    Set mColorTxt = Me.Controls("txtColor")
End Sub

Private Sub mRun_Click()
    SizePlugin.SizeFormRun Me
End Sub

Private Sub mClear_Click()
    SizePlugin.SizeFormClear Me
End Sub

Private Sub mClose_Click()
    SizePlugin.SizeFormClose Me
End Sub

Private Sub mColor_Click()
    SizePlugin.SizeFormPickColor Me
End Sub

Private Sub mFontAuto_Click()
    SizePlugin.SizeFormSync Me
End Sub

Private Sub mLineAuto_Click()
    SizePlugin.SizeFormSync Me
End Sub

Private Sub mArrowAuto_Click()
    SizePlugin.SizeFormSync Me
End Sub

Private Sub mColorTxt_Change()
    SizePlugin.SizeFormSync Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        SizePlugin.SizeFormClose Me
    End If
End Sub
