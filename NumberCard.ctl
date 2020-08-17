VERSION 5.00
Begin VB.UserControl NumberCard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H007C8F97&
   CanGetFocus     =   0   'False
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   975
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   65
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "NumberCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Public Number As Long

Private Sub UserControl_InitProperties()
  If Image1.Picture <> 0 Then UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
  Dim LonRect As Long
  LonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 15, 15)
  SetWindowRgn UserControl.hwnd, LonRect, True
End Sub

Private Sub UserControl_Show()
  If Image1.Picture <> 0 Then UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture() As Picture
  Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
  Set Image1.Picture = New_Picture
  If Image1.Picture <> 0 Then UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
  PropertyChanged "Picture"
End Property

Private Sub UserControl_Resize()
  UserControl.Width = 975
  UserControl.Height = 1245
  Dim LonRect As Long
  LonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 15, 15)
  SetWindowRgn UserControl.hwnd, LonRect, True
End Sub
