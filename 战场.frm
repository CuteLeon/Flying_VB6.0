VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��˿��! ��ɻ��!      �q�ɨr�����ᣩ�q�ɨr"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8655
   Icon            =   "ս��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8655
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2460
      Top             =   180
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   660
   End
   Begin ��һ�.ProgressBar HP 
      Height          =   255
      Left            =   120
      Top             =   5700
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   450
      Max             =   10
      Value           =   10
      Theme           =   7
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   255
      Text            =   "HP"
      TextEffectColor =   65535
      TextEffect      =   4
      PBSCustomeColor2=   16777215
      PBSCustomeColor1=   16777215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   300
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4320
      Top             =   4980
   End
   Begin ��һ�.Enemy Enemy 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   -735
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   3863
      _ExtentY        =   2566
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   35
      Left            =   3960
      Top             =   3540
   End
   Begin ��һ�.MyFly MyFly 
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      ToolTipText     =   "�������ս��Ŷ"
      Top             =   4500
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1296
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���Ƕ��ǌ�˿��"
      BeginProperty Font 
         Name            =   "������κ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Image Image2 
      Height          =   315
      Index           =   0
      Left            =   60
      Picture         =   "ս��.frx":57E2
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label ZhanJi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   5280
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ս����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   555
      Index           =   0
      Left            =   4080
      Picture         =   "ս��.frx":5A1A
      Top             =   4140
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image BackPicture 
      Height          =   4125
      Left            =   780
      Picture         =   "ս��.frx":5F2F
      Stretch         =   -1  'True
      Top             =   780
      Visible         =   0   'False
      Width           =   7275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

Private Const EnmeyCount = 7
Private Const PaodanCount = 6
Dim TX As Long, HurtN As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 87 Then
    MyFly.Top = IIf(MyFly.Top > 400, MyFly.Top - 400, 0)      '�ϼ�
  ElseIf KeyCode = 40 Or KeyCode = 83 Then
    If MyFly.Top + MyFly.Height < HP.Top - 400 Then MyFly.Top = IIf(MyFly.Top < Me.ScaleHeight - MyFly.Height - 400, MyFly.Top + 400, Me.ScaleHeight - MyFly.Height) '�¼�
  ElseIf KeyCode = 37 Or KeyCode = 65 Then
    MyFly.Left = IIf(MyFly.Left > 400, MyFly.Left - 400, 0)       '���
  ElseIf KeyCode = 39 Or KeyCode = 68 Then
    MyFly.Left = IIf(MyFly.Left < Me.ScaleWidth - MyFly.Width - 400, MyFly.Left + 400, Me.ScaleWidth - MyFly.Width)  '�Ҽ�
  End If
End Sub

Private Sub Form_Load()
  Dim N As Long
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  Me.PaintPicture BackPicture.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
  AddEnemy
  For N = 1 To PaodanCount
    Load Image1(N)
    Load Timer2(N)
    Timer2(N).Tag = N
  Next
  
  MyFly.ZOrder 0
  Me.SetFocus
End Sub

Private Sub Timer1_Timer()      '�����ڵ�
  Dim N As Long
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  For N = 1 To PaodanCount
    If Image1(N).Visible = False Then
      Image1(N).Move MyFly.Left + (MyFly.Width - Image1(N).Width) / 2, MyFly.Top - Image1(N).Height + 15
      Image1(N).Visible = True
      Timer2(N).Enabled = True
      Timer2(N).Tag = N
      Exit Sub
    End If
  Next
End Sub

Private Sub AddEnemy() '���ӵл�
  Dim N As Long
  For N = 1 To EnmeyCount
    Load Enemy(N)
    Randomize
    Enemy(N).PicN = Int(6 * Rnd())
    Enemy(N).Father_hWnd = Me.hwnd
    Load Image2(N)
    Load Timer3(N)
    Timer3(N).Tag = N
    Enemy(N).Visible = True
    Enemy(N).Start
  Next
End Sub

Private Sub Timer2_Timer(Index As Integer)   '�ڵ��˶�
  Dim MyIndex As Long, N As Long
  MyIndex = Timer2(Index).Tag
  Image1(MyIndex).Top = Image1(MyIndex).Top - 300
  
  If Image1(MyIndex).Top <= -Image1(MyIndex).Height Then
    Image1(MyIndex).Visible = False
    Timer2(MyIndex).Enabled = False
  End If
  
  For N = 0 To EnmeyCount
    If (Image1(MyIndex).Left >= Enemy(N).FlyX - Image1(MyIndex).Width And Image1(MyIndex).Left <= Enemy(N).FlyX + Enemy(N).Width) And (Image1(MyIndex).Top >= Enemy(N).FlyY - Image1(MyIndex).Height And Image1(MyIndex).Top <= Enemy(N).FlyY + Enemy(N).Height) Then
      Image1(MyIndex).Visible = False
      Timer2(MyIndex).Enabled = False
      ZhanJi.Caption = CLng(ZhanJi) + 1
      Enemy(N).Over
      Randomize
      Enemy(N).PicN = Int(6 * Rnd())
    End If
  Next
End Sub

Private Sub Enemy_Hurt(Index As Integer)  '�л�����
  If Image2(Index).Visible = False And Enemy(Index).FlyY < Me.ScaleHeight * 0.5 Then
    Image2(Index).Move Enemy(Index).FlyX + (Enemy(Index).Width - Image2(Index).Width) / 2, Enemy(Index).FlyY + Enemy(Index).Height
    Image2(Index).Visible = True
    Timer3(Index).Enabled = True
  End If
End Sub

Private Sub Timer3_Timer(Index As Integer)
On Error Resume Next
  Dim MyIndex As Integer, N As Long
  MyIndex = CInt(Timer3(Index).Tag)
  Image2(MyIndex).Top = Image2(MyIndex).Top + 200
  
  If Image2(MyIndex).Top >= Me.ScaleHeight + Image2(MyIndex).Height Then
    Image2(MyIndex).Visible = False
    Timer3(MyIndex).Enabled = False
  End If

  If (Image2(MyIndex).Left >= MyFly.Left - Image2(MyIndex).Width And Image2(MyIndex).Left <= MyFly.Left + MyFly.Width) And (Image2(MyIndex).Top >= MyFly.Top - Image2(MyIndex).Height And Image2(MyIndex).Top <= MyFly.Top + MyFly.Height) Then
    Image2(MyIndex).Visible = False
    Timer3(MyIndex).Enabled = False
    HP.Value = HP.Value - 1
    If HP.Value = 0 Then
      MyFly.Move -MyFly.Width, -MyFly.Height
      Timer1.Enabled = False
      For N = 1 To PaodanCount
        Image1(N).Visible = False
        Timer2(N).Enabled = False
      Next
      
      MessageBox Me.hwnd, "��ã��װ���24K����˿��" & vbCrLf & "����δ�ɻ���У���һ��������" & ZhanJi & "�����ˣ�" & vbCrLf & "      o(�Rv�Q)o~~", "���Ƕ�����ɻ�!", 48
  
      HP.Value = HP.Max
      ZhanJi = "0"
      MyFly.Move 3960, 4800
      Timer1.Enabled = True
      
      For N = 1 To EnmeyCount
        Enemy(N).Over
        Image2(N).Visible = False
        Timer3(N).Enabled = False
      Next
    Else
      Info.Visible = True
      Info = "���ܵ��˹�����"
      HurtN = 0
      Timer4.Enabled = True
      Timer5.Enabled = True
    End If
  End If
End Sub

Private Sub Timer4_Timer()
  If HurtN < 6 Then Info.Visible = Not Info.Visible
  HurtN = HurtN + 1
  
  If HurtN >= 15 Then
    HurtN = 0
    Info.Visible = False
    Timer4.Enabled = False
    Timer5.Enabled = False
  End If
End Sub

Private Sub Timer5_Timer()
  Static mFlash As Boolean
  FlashWindow hwnd, Not mFlash
End Sub
