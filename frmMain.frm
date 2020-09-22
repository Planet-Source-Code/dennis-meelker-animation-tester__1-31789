VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Animation Tester"
   ClientHeight    =   3420
   ClientLeft      =   270
   ClientTop       =   2520
   ClientWidth     =   6270
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   Begin VB.CommandButton Command3 
      Caption         =   "&Stop Animation"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer FPSTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Load"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox cntAni 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   11
      Top             =   0
      Width           =   4455
      Begin VB.PictureBox picAni 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblFPS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 Fps"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3945
         TabIndex        =   14
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   2415
      Left            =   4560
      TabIndex        =   2
      Top             =   0
      Width           =   1695
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Text            =   "200"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "40"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "40"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "3"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "2"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   2085
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Rows:"
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   10
         Top             =   1635
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Columns:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Height:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   780
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Width:"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   7
         Top             =   435
         Width           =   465
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Animation"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Tag             =   "0"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5400
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Dim Loaded As Boolean
Dim TimeToEnd As Boolean
Dim FPS As Integer
Dim EscapePressed As Boolean

Private Sub Command1_Click()
Command1.Visible = False
Command3.Visible = True

Dim TickDifference As Long
Dim LastTick As Long

Dim Amount As Integer
Dim CurrentPos As Integer
Dim Tmp As Integer
Dim TmpX As Integer

Dim Width As Integer
Dim Height As Integer

Dim XAm As Integer
Dim YAm As Integer

TickDifference = Val(txtSpeed.Text)

Width = Val(Text1.Text)
Height = Val(Text2.Text)

XAm = Val(Text3.Text)
YAm = Val(Text4.Text)

picAni.Width = Width
picAni.Height = Height

picAni.Left = (cntAni.Width / 2) - (picAni.Width / 2)
picAni.Top = (cntAni.Height / 2) - (picAni.Height / 2)


Amount = XAm * YAm

CurrentPos = 1
TimeToEnd = False

FPS = 0
FPSTimer = True

Do
    If EscapePressed = True Then TimeToEnd = True
    
    If GetTickCount - LastTick > TickDifference Then
    LastTick = GetTickCount
        If CurrentPos >= Amount + 1 Then CurrentPos = 1
        picAni.Cls
        If CurrentPos <= XAm Then
            BitBlt picAni.hDC, 0, 0, Width, Height, picTiles.hDC, (CurrentPos - 1) * Width, 0, SRCCOPY
        Else
            Tmp = Int((CurrentPos - 1) / XAm)
            TmpX = (CurrentPos - (Tmp * XAm)) - 1
            BitBlt picAni.hDC, 0, 0, Width, Height, picTiles.hDC, TmpX * Width, Tmp * Height, SRCCOPY
        End If
        picAni.Refresh
        CurrentPos = CurrentPos + 1
        
        FPS = FPS + 1
    End If
    DoEvents
Loop Until TimeToEnd

FPSTimer = False
Command1.Visible = True
Command3.Visible = False
End Sub

Private Sub Command2_Click()
On Error GoTo canceled

CDlg.Filter = "All Files(*.*)|*.*|Graphic Files|*.bmp;*.jpg;*.jpeg;*.gif"
CDlg.DialogTitle = "Load Animation..."
CDlg.ShowOpen

frmPicture.Picture = LoadPicture(CDlg.FileName)
picTiles.Picture = LoadPicture(CDlg.FileName)
picAni.Cls

Loaded = True
Command1.Enabled = True

Exit Sub
canceled:
Exit Sub
End Sub

Private Sub Command3_Click()
TimeToEnd = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then EscapePressed = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then EscapePressed = False

End Sub

Private Sub Form_Load()
EscapePressed = False
Loaded = False
frmPicture.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub FPSTimer_Timer()
lblFPS = FPS & " Fps"
FPS = 0
End Sub
