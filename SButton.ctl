VERSION 5.00
Begin VB.UserControl SButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   ClipBehavior    =   0  '无
   FillStyle       =   0  'Solid
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   60
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1980
      Top             =   240
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "SButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINT_API
  x As Long
  y As Long
End Type

Public Event Click()

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.x, MousePos.y) Then Exit Sub
  UserControl.PaintPicture Image2.Picture, -1, -1, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Image3.Picture <> 0 Then UserControl.PaintPicture Image3.Picture, -1, -1, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.x, MousePos.y) Then Exit Sub
  
  If Timer1.Enabled = False Then
  Timer1.Enabled = True
    UserControl.PaintPicture Image2.Picture, -1, -1, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
  End If
End Sub

Private Sub Timer1_Timer()
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.x, MousePos.y) Then
    UserControl.PaintPicture Image1.Picture, -1, -1, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
    Timer1.Enabled = False
  End If
End Sub

Private Sub UserControl_InitProperties()
  If Image1.Picture <> 0 Then UserControl.PaintPicture Image1.Picture, -1, -1, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
  Dim LonRect As Long
  LonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 60, 60)
  SetWindowRgn UserControl.hwnd, LonRect, True
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 900
  UserControl.Height = 900
  Dim LonRect As Long
  LonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 60, 60)
  SetWindowRgn UserControl.hwnd, LonRect, True
End Sub

Private Sub UserControl_Show()
  If Image1.Picture <> 0 Then UserControl.PaintPicture Image1.Picture, -1, -1, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set PictureMove = PropBag.ReadProperty("PictureMove", Nothing)
  Set PictureNormal = PropBag.ReadProperty("PictureNormal", Nothing)
  Set PictureDown = PropBag.ReadProperty("PictureDown", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("PictureMove", PictureMove, Nothing)
  Call PropBag.WriteProperty("PictureNormal", PictureNormal, Nothing)
  Call PropBag.WriteProperty("PictureDown", PictureDown, Nothing)
End Sub

'MappingInfo=Image1,Image1,-1,Picture
Public Property Get PictureNormal() As Picture
Attribute PictureNormal.VB_Description = "返回/设置控件中显示的图形。"
  Set PictureNormal = Image1.Picture
End Property

Public Property Set PictureNormal(ByVal New_PictureNormal As Picture)
  Set Image1.Picture = New_PictureNormal
  If Image1.Picture <> 0 Then UserControl.PaintPicture Image1.Picture, -1, -1, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
  PropertyChanged "PictureNormal"
End Property

'MappingInfo=Image2,Image2,-1,Picture
Public Property Get PictureMove() As Picture
  Set PictureMove = Image2.Picture
End Property

Public Property Set PictureMove(ByVal New_PictureMove As Picture)
  Set Image2.Picture = New_PictureMove
  PropertyChanged "PictureMove"
End Property

'MappingInfo=Image3,Image3,-1,Picture
Public Property Get PictureDown() As Picture
  Set PictureDown = Image3.Picture
End Property

Public Property Set PictureDown(ByVal New_PictureDown As Picture)
  Set Image3.Picture = New_PictureDown
  PropertyChanged "PictureDown"
End Property
