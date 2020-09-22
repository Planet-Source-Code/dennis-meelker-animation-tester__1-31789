VERSION 5.00
Begin VB.Form frmPicture 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Loaded Picture"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Const SWW_HPARENT = -8

Private Sub Form_Load()
Call SetWindowWord(hwnd, SWW_HPARENT, frmMain.hwnd)
StickToMainForm
End Sub

Public Sub StickToMainForm()
frmPicture.Left = frmMain.Left + frmMain.Width
frmPicture.Top = frmMain.Top
End Sub

Private Sub Form_Resize()
StickToMainForm
End Sub
