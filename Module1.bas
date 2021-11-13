Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Type NOTIFYICONDATA
cbSize As Long 'NOTIFYICONDATA类型的大小
hwnd As Long '你的应用程序窗体的名柄
uID As Long '应用程序图标资源的ID号
uFlags As Long '使那些参数有效它是以下枚举类型中的NIF_MESSAGE、NIF_ICON、NIF_TIP三组的组合
uCallbackMessage As Long '鼠标移动时把此消息发给该图标的窗体
hIcon As Long '图标名柄
szTip As String * 64 '当鼠标在图标上时显示的Tip文本
End Type
Public Enum enm_NIM_Shell
NIM_ADD = &H0 '增加图标
NIM_MODIFY = &H1 '修改图标
NIM_DELETE = &H2 '删除图标
NIF_MESSAGE = &H1 '使类型"NOTIFYICONDATA"中的uCallbackMessage有效
NIF_ICON = &H2 '使类型"NOTIFYICONDATA"中的hIcon有效
NIF_TIP = &H4 '使类型"NOTIFYICONDATA"中的szTip有效
End Enum
Public Const WM_MOUSEMOVE = &H200 '在图标上移动鼠标
Public Const WM_LBUTTONDOWN = &H201 '鼠标左键按下
Public Const WM_LBUTTONUP = &H202 '鼠标左键释放
Public Const WM_LBUTTONDBLCLK = &H203 '双击鼠标左键
Public Const WM_RBUTTONDOWN = &H204 '鼠标右键按下
Public Const WM_RBUTTONUP = &H205 '鼠标右键释放
Public Const WM_RBUTTONDBLCLK = &H206 '双击鼠标右键
Public Const WM_SETHOTKEY = &H32 '响应您定义的热键
Public nidProgramData As NOTIFYICONDATA
