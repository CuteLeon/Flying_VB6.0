VERSION 5.00
Begin VB.UserControl MyFly 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   ClipBehavior    =   0  'нч
   FillStyle       =   0  'Solid
   ForwardFocus    =   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   50
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "MyFly.ctx":0000
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "MyFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long   'left + width
  Bottom As Long  'top + height
End Type
Dim Father_RECT As RECT
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Sub UserControl_Initialize()
  Dim Hround As Long
  Hround = CreateRoundRectRgn(0, 0, ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels), ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels), 7, 7)
  SetWindowRgn UserControl.hwnd, Hround, True
  DeleteObject Hround
  UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 750
  UserControl.Height = 735
  Dim Hround As Long
  Hround = CreateRoundRectRgn(0, 0, ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels), ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels), 7, 7)
  SetWindowRgn UserControl.hwnd, Hround, True
  DeleteObject Hround
End Sub

Private Sub UserControl_Show()
  UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  Dim Hround As Long
  Hround = CreateRoundRectRgn(0, 0, ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels), ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels), 7, 7)
  SetWindowRgn UserControl.hwnd, Hround, True
End Sub
