VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   555
   ClientLeft      =   5415
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   -120
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Config"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   -30
      Width           =   255
   End
   Begin VB.Label lblSec 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim start As Boolean
Dim windowred As Boolean
Dim tMin As Integer
Dim tSec As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '总在最前
Const SWP_NOMOVE = &H2 '不更动目前视窗位置
Const SWP_NOSIZE = &H1 '不更动目前视窗大小
Const HWND_TOPMOST = -1 '设定为最上层
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Sub Command1_Click()
    If start = False Then
        start = True
        Command1.Caption = "Pause"
    Else
        start = False
        Command1.Caption = "Start"
    End If
End Sub

Private Sub Command2_Click()
    lblMin.Caption = Right(Str(tMin), Len(Str(tMin)) - 1)
    If tSec < 10 Then lblSec.Caption = "0" & Right(Str(tSec), 1) Else lblSec.Caption = Right(Str(tSec), 2)
    Form1.BackColor = &HFFFFFF
    Label2.BackColor = &HFFFFFF
    lblMin.BackColor = &HFFFFFF
    lblSec.BackColor = &HFFFFFF
    Command1.Caption = "Start"
    windowred = False
    start = False
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Command4_Click()
    Dim tStr
    tStr = InputBox("请输入四位数时间，格式为mmss，即1000代表10:00。", "Timer 1.1 by Sykie Chen")
    Rem tStr = InputBox("请输入四位数时间，格式为mmss，即1000代表10:00。" & vbCrLf & "请不要输入奇怪的东西(ˉ﹃ˉ)", "Timer 1.1 by Sykie Chen")
    If tStr <> "" Then
        tMin = Int(Left(tStr, 2))
        tSec = Int(Right(tStr, 2))
        Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
    SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    start = False
    windowred = False
    tMin = 10
    tSec = 0
    Form1.Left = (Screen.Width - Form1.Width) / 2
End Sub


Private Sub Timer1_Timer()
    If start = True Then
        If Int(lblSec.Caption) = 0 Then
            If Int(lblMin.Caption) = 0 Then
                If windowred = False Then
                    Form1.BackColor = &HFF&
                    lblMin.BackColor = &HFF&
                    lblSec.BackColor = &HFF&
                    Label2.BackColor = &HFF&
                    windowred = True
                Else
                    Form1.BackColor = &HFFFFFF
                    lblMin.BackColor = &HFFFFFF
                    lblSec.BackColor = &HFFFFFF
                    Label2.BackColor = &HFFFFFF
                    windowred = False
                End If
            Else
                lblMin.Caption = Right(Str(Int(lblMin.Caption) - 1), Len(Str(Int(lblMin.Caption)) - 1))
                lblSec.Caption = "59"
            End If
        Else
            If Int(lblSec.Caption) - 1 < 10 Then
                lblSec.Caption = "0" & Right(Str(Int(lblSec.Caption) - 1), 1)
            Else
                lblSec.Caption = Right(Str(Int(lblSec.Caption) - 1), 2)
            End If
        End If
    End If
End Sub
