Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Type NOTIFYICONDATA
cbSize As Long 'NOTIFYICONDATA���͵Ĵ�С
hwnd As Long '���Ӧ�ó����������
uID As Long 'Ӧ�ó���ͼ����Դ��ID��
uFlags As Long 'ʹ��Щ������Ч��������ö�������е�NIF_MESSAGE��NIF_ICON��NIF_TIP��������
uCallbackMessage As Long '����ƶ�ʱ�Ѵ���Ϣ������ͼ��Ĵ���
hIcon As Long 'ͼ������
szTip As String * 64 '�������ͼ����ʱ��ʾ��Tip�ı�
End Type
Public Enum enm_NIM_Shell
NIM_ADD = &H0 '����ͼ��
NIM_MODIFY = &H1 '�޸�ͼ��
NIM_DELETE = &H2 'ɾ��ͼ��
NIF_MESSAGE = &H1 'ʹ����"NOTIFYICONDATA"�е�uCallbackMessage��Ч
NIF_ICON = &H2 'ʹ����"NOTIFYICONDATA"�е�hIcon��Ч
NIF_TIP = &H4 'ʹ����"NOTIFYICONDATA"�е�szTip��Ч
End Enum
Public Const WM_MOUSEMOVE = &H200 '��ͼ�����ƶ����
Public Const WM_LBUTTONDOWN = &H201 '����������
Public Const WM_LBUTTONUP = &H202 '�������ͷ�
Public Const WM_LBUTTONDBLCLK = &H203 '˫��������
Public Const WM_RBUTTONDOWN = &H204 '����Ҽ�����
Public Const WM_RBUTTONUP = &H205 '����Ҽ��ͷ�
Public Const WM_RBUTTONDBLCLK = &H206 '˫������Ҽ�
Public Const WM_SETHOTKEY = &H32 '��Ӧ��������ȼ�
Public nidProgramData As NOTIFYICONDATA
