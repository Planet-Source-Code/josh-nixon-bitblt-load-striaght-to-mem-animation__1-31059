VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   1395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1365
      Left            =   -11880
      Picture         =   "Penguin.frx":0000
      ScaleHeight     =   1305
      ScaleWidth      =   19425
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   19485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'By Joshua Nixon Yar Interactive
'This code may be used in any of your programs.
'JNixon21@excite.com/Nit3shift
'*************************************
Dim Counter 'setup a counter for the x value
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 Const SRCAND = &H8800C6
 Const SRCPAINT = &HEE0086
Private Sub Command1_Click()
Unload Form1
End
End Sub
Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Counter = 0 'set the x value to 0
'set the value for interval its max to
'be 1000 miliseconds
If Right(App.Path, 1) = "\" Then
LoadSprite App.Path + "penguinsani.jpg", 0
Else
LoadSprite App.Path + "\penguinsani.jpg", 0
End If
End Sub

Private Sub Timer1_Timer()
PictureAni Picture1(0), Picture2, 1295, 87, 72
End Sub
Private Sub Form_Unload(Cancel As Integer)
DeleteDC Source
End Sub
Private Function PictureAni(Picture As PictureBox, pictureS As PictureBox, Width As Integer, Height As Integer, Shift As Integer)
DoEvents
Dim i As Integer 'Just a Variable
BitBlt Picture.hdc, 0, 0, Width, Height, pictureS.hdc, Counter, 0, SRCPAINT
'Copy the picture with SRCAND onto picture1
BitBlt Picture.hdc, 0, 0, Width, Height, pictureS.hdc, Counter, 0, SRCAND
Counter = Counter + Shift  'add 110 to x moves the picture
'horizontally
Picture.Refresh
Let i = Width  ' This line makes sure that it will cycle
'through the whole picture correctly
If Counter >= i Then Counter = 0 'checks to see if all the frames have been displayed. if they have then restart from the first frame
End Function
'Load all picture into memory by this simple function
Private Function LoadSprite(SpriteString As String, CompHDC As Long) As Long
Dim Source As Long
    Source = CreateCompatibleDC(CompHDC)
    SelectObject Source, LoadPicture(SpriteString)
End Function
