Attribute VB_Name = "Module1"
Global nochat As Integer
Global username As String
Global lastsaid As String
Global smallmsgs As Integer
Global istyping As Boolean
Global oldcaption As String
Global mycolor As String
Global theircolor As String
Global bold As Boolean
Global bold2 As Boolean
Global mysize As Integer
Global theirsize As Integer
Global override As Boolean

Global box1s As String
Global box2s As String
Global box3s As String
Global box4s As String
Global box5s As String
Global box6s As String
Global box7s As String
Global box8s As String
Global box9s As String
Global allboxs As Integer

Global rsp As String
Global rsp2 As String

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

