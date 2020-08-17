VERSION 5.00
Begin VB.Form ForeForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin Game2048.SButton Up 
      Height          =   900
      Left            =   675
      Top             =   6765
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      PictureMove     =   "ForeForm.frx":0000
      PictureNormal   =   "ForeForm.frx":579B
      PictureDown     =   "ForeForm.frx":AD5F
   End
   Begin Game2048.SButton ReStart 
      Height          =   900
      Left            =   3135
      Top             =   105
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      PictureMove     =   "ForeForm.frx":10170
      PictureNormal   =   "ForeForm.frx":1609B
   End
   Begin Game2048.SButton Exit 
      Height          =   900
      Left            =   4275
      Top             =   105
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      PictureMove     =   "ForeForm.frx":1BFA8
      PictureNormal   =   "ForeForm.frx":21FE5
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   1
      Left            =   615
      Top             =   1425
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   2
      Left            =   1635
      Top             =   1425
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   3
      Left            =   2655
      Top             =   1425
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   4
      Left            =   3675
      Top             =   1425
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   5
      Left            =   615
      Top             =   2745
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   6
      Left            =   1635
      Top             =   2745
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   7
      Left            =   2655
      Top             =   2745
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   8
      Left            =   3675
      Top             =   2745
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   9
      Left            =   615
      Top             =   4065
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   10
      Left            =   1635
      Top             =   4065
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   11
      Left            =   2655
      Top             =   4065
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   12
      Left            =   3675
      Top             =   4065
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   13
      Left            =   615
      Top             =   5385
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   14
      Left            =   1635
      Top             =   5385
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   15
      Left            =   2655
      Top             =   5385
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.NumberCard Card 
      Height          =   1245
      Index           =   16
      Left            =   3675
      Top             =   5385
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2196
   End
   Begin Game2048.SButton Down 
      Height          =   900
      Left            =   1695
      Top             =   6765
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      PictureMove     =   "ForeForm.frx":281E7
      PictureNormal   =   "ForeForm.frx":2DA11
      PictureDown     =   "ForeForm.frx":33049
   End
   Begin Game2048.SButton Left 
      Height          =   900
      Left            =   2715
      Top             =   6765
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      PictureMove     =   "ForeForm.frx":384AD
      PictureNormal   =   "ForeForm.frx":3DD77
      PictureDown     =   "ForeForm.frx":4337F
   End
   Begin Game2048.SButton Right 
      Height          =   900
      Left            =   3735
      Top             =   6765
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      PictureMove     =   "ForeForm.frx":487C6
      PictureNormal   =   "ForeForm.frx":4E0BF
      PictureDown     =   "ForeForm.frx":5384F
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   2
      Left            =   -960
      Picture         =   "ForeForm.frx":58E1C
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   4
      Left            =   -960
      Picture         =   "ForeForm.frx":5DE4C
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   8
      Left            =   -960
      Picture         =   "ForeForm.frx":62D34
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   16
      Left            =   -960
      Picture         =   "ForeForm.frx":67D03
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   32
      Left            =   -960
      Picture         =   "ForeForm.frx":6CC63
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   64
      Left            =   -960
      Picture         =   "ForeForm.frx":72058
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   128
      Left            =   -960
      Picture         =   "ForeForm.frx":773E7
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   256
      Left            =   -960
      Picture         =   "ForeForm.frx":7C6A5
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   2048
      Left            =   -960
      Picture         =   "ForeForm.frx":818FC
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   1024
      Left            =   -960
      Picture         =   "ForeForm.frx":86D16
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   512
      Left            =   -960
      Picture         =   "ForeForm.frx":8BE02
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image N_Pic 
      Height          =   1245
      Index           =   0
      Left            =   -960
      Picture         =   "ForeForm.frx":90F82
      Top             =   -1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   915
      TabIndex        =   0
      Top             =   -15
      Width           =   2160
   End
End
Attribute VB_Name = "ForeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

Private Sub Exit_Click()
  If MsgBox("真的要退出游戏么?", vbQuestion + vbYesNo, "退出游戏?") = vbYes Then End
End Sub

Private Sub Form_Load()
On Error Resume Next
  Score = "0"
  Dim N As Long
  For N = Card.LBound To Card.UBound
    Set Card(N).Picture = N_Pic(0)
    Card(N).Number = 0
  Next
  
  Dim RndX As Double, AddN As Long
  For N = 1 To 2
    Randomize
    RndX = Rnd()
    Randomize
    AddN = Int(12 * Rnd() + 1)
    If RndX < 0.5 Then
      Set Card(AddN).Picture = N_Pic(2)
      Card(AddN).Number = 2
    Else
      Set Card(AddN).Picture = N_Pic(4)
      Card(AddN).Number = 4
    End If
  Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 87 Then
    Up_Click     '上
  ElseIf KeyCode = 40 Or KeyCode = 83 Then
    Down_Click   '下
  ElseIf KeyCode = 37 Or KeyCode = 65 Then
    Left_Click   '左
  ElseIf KeyCode = 39 Or KeyCode = 68 Then
    Right_Click  '右
  End If
End Sub

Private Sub ReStart_Click()
  If MsgBox("真的要重新开始游戏么?", vbQuestion + vbYesNo, "重新开始?") = vbYes Then
    Form_Load
  End If
End Sub

Private Sub Merge(ByVal StartN As Long, ByVal EndN As Long, ByVal StepN As Long, ByVal DifferenceN As Long)
  Dim N As Long, Merged As Boolean
  
  For N = StartN To EndN Step StepN
      If Card(N).Number = Card(N + DifferenceN).Number Then '1=2
          If Card(N).Number = 0 Then  '1=2=0
              If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                  '3=4=0   0000
                  If Card(N + 2 * DifferenceN).Number = 2048 Then '      3=4=2048   00XX
                      Card(N).Number = 2048
                      Card(N + DifferenceN).Number = 2048
                      Card(N + 2 * DifferenceN).Number = 0
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  ElseIf 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048   0011
                      Card(N).Number = 2 * Card(N + 2 * DifferenceN).Number
                      Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                      Card(N + 2 * DifferenceN).Number = 0
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  End If
              Else '    3<>4
                  If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0    000*
                      Card(N).Number = Card(N + 3 * DifferenceN).Number
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0    00*0
                      Card(N).Number = Card(N + 2 * DifferenceN).Number
                      Card(N + 2 * DifferenceN).Number = 0
                      Merged = True
                  Else '      0<3<2048 0<4<2048 3<>4     0012
                      Card(N).Number = Card(N + 2 * DifferenceN).Number
                      Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                      Card(N + 2 * DifferenceN).Number = 0
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  End If
              End If
          ElseIf Card(N).Number = 2048 Then '  1=2=2048
              If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                  '3=4=0     XX00
                  '3=4=2048   XXXX
                  If 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048   XX11
                      Card(N + 2 * DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                      Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  End If
              Else '    3<>4
                  If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0  XX0*
                      Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  End If
                  '3>0 4=0   XX*0
                  '0<3<2048 0<4<2048 3<>4    XX12
              End If
          Else '  0<1<2048 0<2<2048 1=2
              If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                  If Card(N + 2 * DifferenceN).Number = 0 Then '      3=4=0   1100
                      Card(N).Number = 2 * Card(N).Number
                      Score = Str(CLng(Score) + Card(N).Number)
                      Card(N + DifferenceN).Number = 0
                      Merged = True
                  ElseIf Card(N + 2 * DifferenceN).Number = 2048 Then '      3=4=2048    11XX
                      Card(N).Number = 2 * Card(N).Number
                      Score = Str(CLng(Score) + Card(N).Number)
                      Card(N + DifferenceN).Number = 2048
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  ElseIf 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048  1122/1111
                      Card(N).Number = 2 * Card(N).Number
                      Score = Str(CLng(Score) + Card(N).Number)
                      Card(N + DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                      Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                      Card(N + 2 * DifferenceN).Number = 0
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  End If
              Else '    3<>4
                  If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0   110*
                      Card(N).Number = 2 * Card(N).Number
                      Score = Str(CLng(Score) + Card(N).Number)
                      Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0   11*0
                      Card(N).Number = 2 * Card(N).Number
                      Score = Str(CLng(Score) + Card(N).Number)
                      Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                      Card(N + 2 * DifferenceN).Number = 0
                      Merged = True
                  Else '      0<3<2048 0<4<2048 3<>4   1123/1112/1121
                      Card(N).Number = 2 * Card(N).Number
                      Score = Str(CLng(Score) + Card(N).Number)
                      Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                      Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                      Card(N + 3 * DifferenceN).Number = 0
                      Merged = True
                  End If
              End If
          End If
      Else '1<>2
          If Card(N).Number = 0 And Card(N + DifferenceN).Number > 0 Then  '      1=0 2>0
              If Card(N + DifferenceN).Number = 2048 Then  '        2=2048
                  If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                      If Card(N + 2 * DifferenceN).Number = 0 Then '      3=4=0  0X00
                          Card(N).Number = 2048
                          Card(N + DifferenceN).Number = 0
                          Merged = True
                      ElseIf Card(N + 2 * DifferenceN).Number = 2048 Then '      3=4=2048 OXXX
                          Card(N).Number = 2048
                          Card(N + 3 * DifferenceN).Number = 0
                          Merged = True
                      ElseIf 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048   0X11
                          Card(N).Number = 2048
                          Card(N + DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                          Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                          Card(N + 2 * DifferenceN).Number = 0
                          Card(N + 3 * DifferenceN).Number = 0
                          Merged = True
                      End If
                  Else '    3<>4
                      If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0
                          If Card(N + 3 * DifferenceN).Number = 2048 Then '        4=2048      0X0X
                              Card(N).Number = 2048
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '        0<4<2048             0X01
                              Card(N).Number = 2048
                              Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0
                          If Card(N + 2 * DifferenceN).Number = 2048 Then '        3=2048    0XX0
                              Card(N).Number = 2048
                              Card(N + 2 * DifferenceN).Number = 0
                              Merged = True
                          Else '        0<3<2048     0X10
                              Card(N).Number = 2048
                              Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                              Card(N + 2 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      Else '      0<3<2048 0<4<2048 3<>4     0X12
                          Card(N).Number = 2048
                          Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                          Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                          Card(N + 3 * DifferenceN).Number = 0
                          Merged = True
                      End If
                  End If
              Else '        0<2<2048
                  If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                      If Card(N + 2 * DifferenceN).Number = 0 Then '      3=4=0    0100
                          Card(N).Number = Card(N + DifferenceN).Number
                          Card(N + DifferenceN).Number = 0
                          Merged = True
                      ElseIf Card(N + 2 * DifferenceN).Number = 2048 Then '      3=4=2048  01XX
                          Card(N).Number = Card(N + DifferenceN).Number
                          Card(N + DifferenceN).Number = 2048
                          Card(N + 3 * DifferenceN).Number = 0
                          Merged = True
                      ElseIf 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048
                          If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3  0111
                              Card(N).Number = 2 * Card(N + DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                              Card(N + 2 * DifferenceN).Number = 0
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '          2<>3  0122
                              Card(N).Number = Card(N + DifferenceN).Number
                              Card(N + DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                              Card(N + 2 * DifferenceN).Number = 0
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      End If
                  Else '    3<>4
                      If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0
                          If Card(N + 3 * DifferenceN).Number = 2048 Then '        4=2048   010X
                              Card(N).Number = Card(N + DifferenceN).Number
                              Card(N + DifferenceN).Number = 2048
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '        0<4<2048  0101
                          
                              If Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '  2=4  0101
                                  Card(N).Number = 2 * Card(N + DifferenceN).Number
                                  Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                                  Card(N + DifferenceN).Number = 0
                                  Card(N + 3 * DifferenceN).Number = 0
                                  Merged = True
                              Else     '2<>4 0102
                                  Card(N).Number = Card(N + DifferenceN).Number
                                  Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                                  Card(N + 3 * DifferenceN).Number = 0
                                  Merged = True
                              End If
                          End If
                      ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0
                          If Card(N + 2 * DifferenceN).Number = 2048 Then '        3=2048   01X0
                              Card(N).Number = Card(N + DifferenceN).Number
                              Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                              Card(N + 2 * DifferenceN).Number = 0
                              Merged = True
                          Else '        0<3<2048
                              If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3   0110
                                  Card(N).Number = 2 * Card(N + DifferenceN).Number
                                  Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                                  Card(N + DifferenceN).Number = 0
                                  Card(N + 2 * DifferenceN).Number = 0
                                  Merged = True
                              Else '          2<>3   0120
                                  Card(N).Number = Card(N + DifferenceN).Number
                                  Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                                  Card(N + 2 * DifferenceN).Number = 0
                                  Merged = True
                              End If
                          End If
                      Else '      0<3<2048 0<4<2048 3<>4
                          If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3 0112
                              Card(N).Number = 2 * Card(N + DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                              Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                              Card(N + 2 * DifferenceN).Number = 0
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '          2<>3    0121/0123
                              Card(N).Number = Card(N + DifferenceN).Number
                              Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                              Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      End If
                  End If
              End If
          ElseIf Card(N).Number > 0 And Card(N + DifferenceN).Number = 0 Then  '      1>0 2=0
              If Card(N).Number = 2048 Then   '        1=2048
                If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                    '3=4=0  X000
                    If Card(N + 2 * DifferenceN).Number = 2048 Then '3=4=2048   X0XX
                        Card(N + DifferenceN).Number = 2048
                        Card(N + 3 * DifferenceN).Number = 0
                        Merged = True
                    ElseIf 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048  X011
                        Card(N + DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                        Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                        Card(N + 2 * DifferenceN).Number = 0
                        Card(N + 3 * DifferenceN).Number = 0
                        Merged = True
                    End If
                Else '    3<>4
                    If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0   X00*
                        Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                        Card(N + 3 * DifferenceN).Number = 0
                        Merged = True
                    ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0  XO*0
                        Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                        Card(N + 2 * DifferenceN).Number = 0
                        Merged = True
                    Else '      0<3<2048 0<4<2048 3<>4    X012
                        Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                        Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                        Card(N + 3 * DifferenceN).Number = 0
                        Merged = True
                    End If
                End If

              Else '        0<1<2048
                If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                    '3=4=0    1000
                    If Card(N + 2 * DifferenceN).Number = 2048 Then '      3=4=2048 10XX
                        Card(N + DifferenceN).Number = 2048
                        Card(N + 3 * DifferenceN).Number = 0
                        Merged = True
                    ElseIf 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048
                        If Card(N).Number = Card(N + 2 * DifferenceN).Number Then '          1=3   1011
                            Card(N).Number = 2 * Card(N).Number
                            Score = Str(CLng(Score) + Card(N).Number)
                            Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                            Card(N + 2 * DifferenceN).Number = 0
                            Card(N + 3 * DifferenceN).Number = 0
                            Merged = True
                        Else '          1<>3   1022
                            Card(N + DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                            Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                            Card(N + 2 * DifferenceN).Number = 0
                            Card(N + 3 * DifferenceN).Number = 0
                            Merged = True
                        End If
                    End If
                Else '    3<>4
                    If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0
                        If Card(N + 3 * DifferenceN).Number = 2048 Then '        4=2048  100X
                            Card(N + DifferenceN).Number = 2048
                            Card(N + 3 * DifferenceN).Number = 0
                            Merged = True
                        Else '        0<4<2048
                            If Card(N).Number = Card(N + 3 * DifferenceN).Number Then '          1=4  1001
                                Card(N).Number = 2 * Card(N).Number
                                Score = Str(CLng(Score) + Card(N).Number)
                                Card(N + 3 * DifferenceN).Number = 0
                                Merged = True
                            Else '          1<>4   1002
                                Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                                Card(N + 3 * DifferenceN).Number = 0
                                Merged = True
                            End If
                        End If
                    ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0
                        If Card(N + 2 * DifferenceN).Number = 2048 Then '        3=2048   10X0
                            Card(N + DifferenceN).Number = 2048
                            Card(N + 2 * DifferenceN).Number = 0
                            Merged = True
                        Else '        0<3<2048
                            If Card(N).Number = Card(N + 2 * DifferenceN).Number Then '          1=3   1010
                                Card(N).Number = 2 * Card(N).Number
                                Score = Str(CLng(Score) + Card(N).Number)
                                Card(N + 2 * DifferenceN).Number = 0
                                Merged = True
                            Else '          1<>3   1020
                                Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                                Card(N + 2 * DifferenceN).Number = 0
                                Merged = True
                            End If
                        End If
                    Else '      0<3<2048 0<4<2048 3<>4
                        If Card(N).Number = Card(N + 2 * DifferenceN).Number Then '          1=3   1012
                            Card(N).Number = 2 * Card(N).Number
                            Score = Str(CLng(Score) + Card(N).Number)
                            Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                            Card(N + 2 * DifferenceN).Number = 0
                            Card(N + 3 * DifferenceN).Number = 0
                            Merged = True
                        Else '          1<>3   1021
                            Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number
                            Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                            Card(N + 3 * DifferenceN).Number = 0
                            Merged = True
                        End If
                    End If
                End If
              End If
          
          Else '      0<1<=2048 0<2<=2048 1<>2
              If 0 < Card(N).Number < 2048 And Card(N + DifferenceN).Number = 2048 Then '        0<1<2048 2=2048
                  If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                      '3=4=0    1X00
                      '3=4=2048      1XXX
                      If 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048  1X11
                          Card(N + 2 * DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                          Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                          Card(N + 3 * DifferenceN).Number = 0
                          Merged = True
                      End If
                  Else '    3<>4
                      If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0  1X0*
                          Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                          Card(N + 3 * DifferenceN).Number = 0
                          Merged = True
                      End If
                      '3>0 4=0   1X*0
                      '0<3<2048 0<4<2048 3<>4    1X23
                  End If
              ElseIf Card(N).Number = 2048 And 0 < Card(N + DifferenceN).Number < 2048 Then '        1=2048 0<2<2048
                  If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                      '3=4=0   X100
                      '3=4=2048      X1XX
                      If 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048
                          If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3    X111
                              Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '          2<>3  X122
                              Card(N + 2 * DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      End If
                  Else '    3<>4
                      If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0
                          If Card(N + 3 * DifferenceN).Number = 2048 Then '        4=2048    X10X
                              Card(N + 2 * DifferenceN).Number = 2048
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '        0<4<2048
                              If Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '          2=4   X101
                                  Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                                  Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                                  Card(N + 3 * DifferenceN).Number = 0
                                  Merged = True
                              Else '          2<>4    X102
                                  Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                                  Card(N + 3 * DifferenceN).Number = 0
                                  Merged = True
                              End If
                          End If
                      ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0
                           '3=2048  X1X0
                          If 0 < Card(N + 2 * DifferenceN).Number < 2048 Then '        0<3<2048
                              '2<>3     X120
                              If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3   X110
                                  Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                                  Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                                  Card(N + 2 * DifferenceN).Number = 0
                                  Merged = True
                              End If
                          End If
                      Else '      0<3<2048 0<4<2048 3<>4
                          '2<>3    X123
                          If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3    X112
                              Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                              Card(N + 2 * DifferenceN).Number = 2 * Card(N + 3 * DifferenceN).Number
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      End If
                  End If
              Else '        0<1<2048 0<2<2048 1<>2
                  If Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '    3=4
                      '3=4=0     1200
                      '3=4=2048     12XX
                      If 0 < Card(N + 2 * DifferenceN).Number And Card(N + 2 * DifferenceN).Number < 2048 Then '      0<3=4<2048
                          If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3    1222
                              Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '          2<>3  1233
                              Card(N + 2 * DifferenceN).Number = 2 * Card(N + 2 * DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + 2 * DifferenceN).Number)
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      End If
                  Else '    3<>4
                      If Card(N + 2 * DifferenceN).Number = 0 And Card(N + 3 * DifferenceN).Number > 0 Then '      3=0 4>0
                          If Card(N + 3 * DifferenceN).Number = 2048 Then '        4=2048    120X
                              Card(N + 2 * DifferenceN).Number = 2048
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          Else '        0<4<2048
                              If Card(N + DifferenceN).Number = Card(N + 3 * DifferenceN).Number Then '          2=4   1202
                                  Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                                  Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                                  Card(N + 3 * DifferenceN).Number = 0
                                  Merged = True
                              Else '          2<>4    1203
                                  Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                                  Card(N + 3 * DifferenceN).Number = 0
                                  Merged = True
                              End If
                          End If
                      ElseIf Card(N + 2 * DifferenceN).Number > 0 And Card(N + 3 * DifferenceN).Number = 0 Then '      3>0 4=0
                          '3=2048    12X0
                          If 0 < Card(N + 2 * DifferenceN).Number < 2048 Then '        0<3<2048
                              '2<>3    1230
                              If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3  1220
                                  Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                                  Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                                  Card(N + 2 * DifferenceN).Number = 0
                                  Merged = True
                              End If
                          End If
                      Else '      0<3<2048 0<4<2048 3<>4
                          '2<>3    1234
                          If Card(N + DifferenceN).Number = Card(N + 2 * DifferenceN).Number Then '          2=3     1223
                              Card(N + DifferenceN).Number = 2 * Card(N + DifferenceN).Number
                              Score = Str(CLng(Score) + Card(N + DifferenceN).Number)
                              Card(N + 2 * DifferenceN).Number = Card(N + 3 * DifferenceN).Number
                              Card(N + 3 * DifferenceN).Number = 0
                              Merged = True
                          End If
                      End If
                  End If
              End If
          End If
      End If
  Next
  
  For N = 1 To 16
    Set Card(N).Picture = N_Pic(Card(N).Number).Picture
  Next
  
  If Merged = True Then AddNewCard
End Sub

Private Sub AddNewCard()
On Error Resume Next
  Dim N As Long
  Dim ZoreCard() As Long
  ReDim ZoreCard(0)
  For N = 1 To 16
    If Card(N).Number = 0 Then
      ZoreCard(UBound(ZoreCard)) = N
      ReDim Preserve ZoreCard(UBound(ZoreCard) + 1)
    End If
  Next
  ReDim Preserve ZoreCard(UBound(ZoreCard) - 1)
  
  Dim RndX As Double, AddN As Long
  Randomize
  RndX = Rnd()
  Randomize
  AddN = Int(UBound(ZoreCard) * Rnd())
  If RndX < 0.5 Then
    Set Card(ZoreCard(AddN)).Picture = N_Pic(2)
    Card(ZoreCard(AddN)).Number = 2
  Else
    Set Card(ZoreCard(AddN)).Picture = N_Pic(4)
    Card(ZoreCard(AddN)).Number = 4
  End If
  
  If UBound(ZoreCard) = 0 And GameOver = True Then  '卡片堆满  失败
    MsgBox "很可惜，你失败了！" & String(20, " ") & vbCrLf & "你的得分： " & Score.Caption, 64, "You Lost !"
    Form_Load
    Exit Sub
  End If
  
  For N = 1 To 16    '出现2048的卡片  胜利
    If Card(N).Number = 2048 Then
      MsgBox "太棒了，你赢了！" & String(20, " ") & vbCrLf & "你的得分： " & Score.Caption, 64, "You Win !"
      Form_Load
      Exit For
    End If
  Next
End Sub

Private Function GameOver() As Boolean
  Dim x, y As Long
  For x = 1 To 12
    If Card(x).Number = Card(x + 4).Number Then Exit Function
  Next
  For y = 1 To 13 Step 4
    For x = y To y + 2
      If Card(x).Number = Card(x + 1).Number Then Exit Function
    Next
  Next
  GameOver = True
End Function

'1  2  3  4
'5  6  7  8
'9  10 11 12
'13 14 15 16

Private Sub Up_Click()
  Merge 1, 4, 1, 4
End Sub

Private Sub Down_Click()
  Merge 13, 16, 1, -4
End Sub

Private Sub Left_Click()
  Merge 1, 13, 4, 1
End Sub

Private Sub Right_Click()
  Merge 4, 16, 4, -1
End Sub

