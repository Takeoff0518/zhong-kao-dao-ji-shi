VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleMode       =   0  'User
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   65535
      Left            =   0
      Top             =   960
   End
   Begin VB.Label DayLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "天"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "中考倒计时"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'设置窗口位置和状态API

Dim ZK As Date
Dim nian As Integer
Dim yue As Integer
Dim ri As Integer
Dim Today As Date

Private Const HWND_TOPMOST& = -1 ' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1 ' 保持窗口大小
Private Const SWP_NOMOVE& = &H2 ' 保持窗口位置


Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE '窗口置于最前

Main.Top = 0
Main.Left = 0
'Main.Left = Screen.Width - Main.Width

nian = Year(Date)
yue = Month(Date)
ri = Day(Date)
ZK = DateSerial(nian, 6, 12)
Today = DateSerial(nian, yue, ri)
If ZK - Today < 0 Then '如果不同年
    ZK = DateSerial(nian + 1, 6, 12)
    nian = nian + 1
End If
DayLabel.Caption = ZK - Today

With nidProgramData
.cbSize = Len(nidProgramData)
.hwnd = Me.hwnd
.uID = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = "倒计时：" & (ZK - Today) & "天（单击关闭）" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nidProgramData
End Sub

Private Sub Timer1_Timer()
nian = Year(Date)
yue = Month(Date)
ri = Day(Date)
ZK = DateSerial(nian, 6, 12)
Today = DateSerial(nian, yue, ri)
If ZK - Today < 0 Then '如果不同年
    ZK = DateSerial(nian + 1, 6, 12)
    nian = nian + 1
End If
DayLabel.Caption = ZK - Today
Me.Top = 0
Main.Left = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Form_MouseMove_err:
        Dim Result, MSG As Long, I As Integer
        If Me.ScaleMode = vbPixels Then
            MSG = X
        Else
            MSG = X / Screen.TwipsPerPixelX
        End If
    Select Case MSG
        Case WM_LBUTTONUP
            SetForegroundWindow Me.hwnd '这个函数用来当你不或得焦点时弹出菜单能自动消失
            End
        Case WM_LBUTTONDOWN '双击托盘
            SetForegroundWindow Me.hwnd
            End
    End Select
    Exit Sub
Form_MouseMove_err:
        End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim blnExitMe As Boolean
        If blnExitMe = False Then
        Cancel = True '取消退出
    Me.Hide
End If
End Sub
Private Sub MnuQuit_Click() '单击退出按钮时
    Shell_NotifyIcon NIM_DELETE, nidProgramData
    End

'***********************************************

Form_MouseMove_err:
End Sub
