VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer topmostTimer 
      Interval        =   100
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer volumeTimer 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer exitTimer 
      Interval        =   500
      Left            =   480
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SND_ASYNC = &H1       '将控制权立即转让给程序
Private Const SND_NODEFAULT = &H2   '不使用缺省声音
Private Const SND_MEMORY = &H4      '指向一个内存文件
Private Const SND_LOOP = &H8        '循环播放
Private Declare Function sndPlaySoundFromMemory Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Const WM_APPCOMMAND As Long = 793
Private Const APPCOMMAND_VOLUME_UP As Long = 10
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1    '将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1     '保持窗口大小
Private Const SWP_NOMOVE& = &H2     '保持窗口位置
Dim exitCount As Integer

Private Sub Form_Activate()
    Me.Move 0, 0, Screen.Width, Screen.Height
    Dim myMusic() As Byte
    myMusic = LoadResData(101, "CUSTOM")
    sndPlaySoundFromMemory myMusic(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP
End Sub

Private Sub topmostTimer_Timer()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub volumeTimer_Timer()
    SendMessage Me.hwnd, WM_APPCOMMAND, &H30292, 10 * &H10000
End Sub

Private Sub Form_Click()
    exitCount = exitCount + 1
    If exitCount = 10 Then End
End Sub

Private Sub exitTimer_Timer()
    If exitCount > 0 Then exitCount = exitCount - 1
End Sub
