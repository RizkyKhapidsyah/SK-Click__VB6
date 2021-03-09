VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Click v2.0"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   Icon            =   "game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "game.frx":0442
   ScaleHeight     =   3795
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2400
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   600
      Top             =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   2055
      Left            =   3720
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   2055
      Left            =   1920
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Left            =   120
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Left -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblTimeLeft 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Score-->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   1815
      Left            =   3840
      MouseIcon       =   "game.frx":074C
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   2040
      MouseIcon       =   "game.frx":0A56
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   240
      MouseIcon       =   "game.frx":0D60
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim points As Integer 'this calculates the score
Dim start As Boolean 'this starts the game
Dim time_count As Integer 'this will count how much time has passed

Private Sub cmdAbout_Click()
about = MsgBox("This is a simple game by Satish Chandra." + " All you have to do is click the colored boxes that come and when the black box appears, the number of times you click will get you points(10). Caution: There are negative points(-5) for clicking the red ones." + "     Enjoy!!! :-) " + " You can contact me at satish4u@yahoo.com", vbOKOnly, "About Click v2.0")
End Sub

Private Sub cmdStart_Click()

start = True
Command3.Enabled = False
'cmdHS.Enabled = False
cmdStart.Enabled = False
cmdAbout.Enabled = False
Image1.Enabled = True
Image2.Enabled = True
Image3.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
lblScore = CStr(0)
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
time_count = 0
Image1.Enabled = False
Image2.Enabled = False
Image3.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
End Sub

Private Sub Image1_Click()

If Image1.Tag = "black" Then
    points = points + 10
ElseIf Image1.Tag = "red" Then
    points = points - 5
End If
lblScore.Caption = CStr(points)

End Sub

Private Sub Image2_Click()
If Image2.Tag = "black" Then
    points = points + 10
ElseIf Image2.Tag = "red" Then
    points = points - 5
End If
lblScore.Caption = CStr(points)
End Sub

Private Sub Image3_Click()
If Image3.Tag = "black" Then
    points = points + 10
ElseIf Image3.Tag = "red" Then
    points = points - 5
End If
lblScore.Caption = CStr(points)
End Sub


Private Sub Timer1_Timer()
Dim x, y, z As Integer 'Here x y and z take care of the images.


If (start = True) And (time_count < 30) Then
    x = Int(10 * Rnd(3))
    y = Int(10 * Rnd(3))
    z = Int(10 * Rnd(3))

    ' set the picture for the first image
    Select Case x
    Case 1
        Image1.Picture = LoadPicture("white.gif")
        Image1.Tag = "white"
    Case 2
        Image1.Picture = LoadPicture("black.gif")
        Image1.Tag = "black"
    Case 3
        Image1.Picture = LoadPicture("red.gif")
        Image1.Tag = "red"
    End Select

    'set the picture for the second picture
    Select Case y
    Case 1
        Image2.Picture = LoadPicture("white.gif")
        Image2.Tag = "white"
    Case 2
        Image2.Picture = LoadPicture("black.gif")
        Image2.Tag = "black"
    Case 3
        Image2.Picture = LoadPicture("red.gif")
        Image2.Tag = "red"
    End Select

    'set the picture for the third image
    Select Case z
    Case 1
        Image3.Picture = LoadPicture("white.gif")
        Image3.Tag = "white"
    Case 2
        Image3.Picture = LoadPicture("black.gif")
        Image3.Tag = "black"
    Case 3
        Image3.Picture = LoadPicture("red.gif")
        Image3.Tag = "red"
    End Select
End If
End Sub

Private Sub Timer2_Timer()


If start = True Then
time_count = time_count + 1
lblTimeLeft.Caption = CStr(30 - time_count)

If time_count = 30 Then
    start = False
    cmdStart.Enabled = True
    Command3.Enabled = True
    cmdAbout.Enabled = True
    MsgBox ("Time Up! You scored " + lblScore.Caption)
    time_count = 0
    Image1.Enabled = False
    Image2.Enabled = False
    Image3.Enabled = False
    points = 0
End If
End If
End Sub


