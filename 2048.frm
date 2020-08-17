VERSION 5.00
Begin VB.Form GameBack 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "2048"
   ClientHeight    =   7920
   ClientLeft      =   11760
   ClientTop       =   4260
   ClientWidth     =   5400
   Icon            =   "2048.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "2048.frx":5E62
   ScaleHeight     =   7920
   ScaleWidth      =   5400
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "GameBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Sub Form_Activate()
  ForeForm.SetFocus
  ForeForm.Move Me.Left, Me.Top
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ReleaseCapture
  SendMessageA Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
  ForeForm.Move Me.Left, Me.Top
End Sub

Private Sub Form_Load()
On Error Resume Next
  Dim LonRect  As Long
  LonRect = CreateRoundRectRgn(0, 0, ScaleX(Me.ScaleWidth, vbTwips, vbPixels), ScaleY(Me.ScaleHeight, vbTwips, vbPixels), 60, 60)
  SetWindowRgn Me.hwnd, LonRect, True
  Call SetWindowLong(Me.hwnd, -20, GetWindowLong(Me.hwnd, -20) Or &H80000)
  Call SetLayeredWindowAttributes(Me.hwnd, Me.BackColor, 200, 2)
  ForeForm.Show , Me
  Call SetWindowLong(ForeForm.hwnd, -20, GetWindowLong(ForeForm.hwnd, -20) Or &H80000)
  Call SetLayeredWindowAttributes(ForeForm.hwnd, ForeForm.BackColor, 0, 1)
  ForeForm.Move Me.Left, Me.Top
  ForeForm.SetFocus
End Sub
