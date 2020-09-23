VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDraw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Draw!"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmDraw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Brush Color"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Pic"
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   0
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   3195
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   840
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   20
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label1 
      Caption         =   "- Brush Size -"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastx As Integer 'x co-ordinate for line
Dim lasty As Integer 'y co-ordinate for line
Dim wsdata As String 'winsock data that will be send
Dim draw As String 'string for the draw feature
Dim color 'color of brush
Dim backcolor1 'color of the background
Dim size As Integer 'size of the brush

Private Sub Command1_Click()
CommonDialog1.ShowColor
color = CommonDialog1.color
End Sub

Private Sub Command2_Click()
frmClient.client.SendData "101DC" 'send data to the otherside to clear the picturebox for you and the other person
p1.Cls 'clears your picturebox
End Sub

Private Sub Command3_Click()
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Bitmap Image (*.bmp)|*.bmp" 'just filters all the file types to bitmaps (our chosen file format for saving)
CommonDialog1.DialogTitle = "Save Drawing" 'just the title of the save dialog
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then 'if there is text , then save it
SavePicture p1.Image, CommonDialog1.FileName 'save picture to desired path and name
End If
End Sub

Private Sub Form_Load()
color = vbBlack 'the color of the brush on start-up will be black
size = 2 'the size of the brush will be 2
p1.DrawWidth = 2
Slider1.Value = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClient.client.SendData "101DCL"
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then 'if the user is holding down the left mouse button...
    p1.Line (lastx, lasty)-(X, Y), color 'this draws the line using the two variables we chose at top
    draw = draw & lastx & "$" & lasty & "$" & X & "$" & Y & "$" & color & "$" & size & "," 'this is the string that I have named 'draw'. In this string, winsock sends all the data needed for the construction of a line on the other computer, including color
ElseIf Button = 2 Then
    Dim prevcolor
    prevcolor = color
    color = vbWhite
    p1.Line (lastx, lasty)-(X, Y), color 'this draws the line using the two variables we chose at top
    draw = draw & lastx & "$" & lasty & "$" & X & "$" & Y & "$" & color & "$" & size & "," 'this is the string that I have named 'draw'. In this string, winsock sends all the data needed for the construction of a line on the other computer, including color
    color = prevcolor
End If

lastx = X 'assigns the variable a value
lasty = Y 'assigns the variable a value

End Sub

Private Sub Slider1_Change()
size = Slider1.Value
p1.DrawWidth = size
End Sub

Private Sub Timer1_Timer()
If Len(draw) <> 0 Then 'if the length of the string 'draw' is no 0 then send the data

If Len(draw) > 4500 Then 'this is the limit that I have made. It can be altered, but note that much higher will cause winsock to crash
draw = "" 'clear the string so that it doesnt accumulate
Exit Sub 'if the data is too large, don't send it, and exit this sub
End If

frmClient.client.SendData "DRAW" & draw 'sends a draw protocal followed by the string containing all the info
draw = "" 'clear the string so that it doesnt accumulate
End If

End Sub
