VERSION 5.00
Begin VB.UserControl Enemy 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ClipBehavior    =   0  'нч
   FillStyle       =   0  'Solid
   ForwardFocus    =   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   146
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   420
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   720
      Picture         =   "Enemy.ctx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   0
      Picture         =   "Enemy.ctx":5C64
      Top             =   720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   1440
      Picture         =   "Enemy.ctx":B2AE
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   720
      Picture         =   "Enemy.ctx":10DDE
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "Enemy.ctx":16992
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   1440
      Picture         =   "Enemy.ctx":1BE27
      Top             =   720
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "Enemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long   'left + width
  Bottom As Long  'top + height
End Type
Dim Father_RECT As RECT
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Dim TX, TY As Long, XD As Boolean, YD As Boolean
Public FlyX As Long, FlyY As Long
Public Event Hurt()
Public PicN As Long
Public Father_hWnd As Long

Private Sub Timer1_Timer()
  Static NowX As Long, NowY As Long

NowX = FlyX: NowY = FlyY
  
  UserControl.PaintPicture Image1(PicN).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  DoEvents
  
  If NowX >= TX + 100 Then
    NowX = NowX - 100
  ElseIf NowX <= TX - 100 Then
    NowX = NowX + 100
  ElseIf NowX > TX - 100 And NowX < TX + 100 Then
    XD = True
  End If
  
  If NowY >= TY + 100 Then
    NowY = NowY - 100
  ElseIf NowY <= TY - 100 Then
    NowY = NowY + 100
  ElseIf NowY > TY - 100 And NowY < TY + 100 Then
    YD = True
  End If
  
  If XD = True And YD = True Then
    TX = ((Father_RECT.Right - Father_RECT.Left) * 15 - UserControl.Width) * Rnd
    TY = Rnd(1) * (Father_RECT.Bottom - Father_RECT.Top) * 8
    XD = False: YD = False: Exit Sub
  End If
  
  MoveWindow UserControl.hwnd, NowX / 15, NowY / 15, UserControl.Width / 15, UserControl.Height / 15, 1
  FlyX = NowX: FlyY = NowY
  
  If Rnd > 0.9 Then RaiseEvent Hurt
End Sub

Private Sub UserControl_Initialize()
  Dim Hround As Long
  Hround = CreateRoundRectRgn(0, 0, ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels), ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels), 7, 7)
  SetWindowRgn UserControl.hwnd, Hround, True
  DeleteObject Hround
  
  YD = True: XD = True
  UserControl.PaintPicture Image1(0).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 2190
  UserControl.Height = 1455
  Dim Hround As Long
  Hround = CreateRoundRectRgn(0, 0, ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels), ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels), 7, 7)
  SetWindowRgn UserControl.hwnd, Hround, True
  DeleteObject Hround
End Sub

Private Sub UserControl_Show()
  UserControl.PaintPicture Image1(0).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  Dim Hround As Long
  Hround = CreateRoundRectRgn(0, 0, ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels), ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels), 7, 7)
  SetWindowRgn UserControl.hwnd, Hround, True
  DeleteObject Hround
End Sub

Public Sub Start()
  Call GetWindowRect(Father_hWnd, Father_RECT)
  FlyX = ((Father_RECT.Right - Father_RECT.Left) * 15 - UserControl.Width) * Rnd
  FlyY = -UserControl.Height
  MoveWindow UserControl.hwnd, FlyX / 15, FlyY / 15, UserControl.Width / 15, UserControl.Height / 15, 1
  Timer1.Enabled = True
End Sub

Public Sub Over()
  FlyX = ((Father_RECT.Right - Father_RECT.Left) * 15 - UserControl.Width) * Rnd: FlyY = -UserControl.Height
  Timer1_Timer
End Sub
