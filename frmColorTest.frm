VERSION 5.00
Begin VB.Form frmColorTest 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.HScrollBar ScrollRed 
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.HScrollBar ScrollGreen 
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.HScrollBar ScrollBlue 
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   3600
      ScaleHeight     =   2715
      ScaleWidth      =   2955
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblblue 
      Caption         =   "0"
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
      Left            =   960
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblGreen 
      Caption         =   "0"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblRed 
      Caption         =   "0"
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
      Left            =   960
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmColorTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This snippet of code will change the background
'color of a picture box to reflect the value of the
'Red, Green and Blue scroll bars and indicated what the color it is:
'A shade of blue, purple, orange etc...
'This code is fairly accurate but right now isn't
'able to distinguish brown/tan colors
'This code is still in its infancy so
'any comments/suggestions can be directed to
'me @ jenutech@sasktel.net
'
'This code may be used and distrubuted free of change
'providing that these comments are left intact.

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    'now find out what type of color it is. IE yellow, orange, blue, grey etc...
    Dim MyRVal As String
    Dim MyGVal As String
    Dim MyBVal As String
    
    MyRVal = Brightness(Me.ScrollRed.Value) & "R"
    MyGVal = Brightness(Me.ScrollGreen.Value) & "G"
    MyBVal = Brightness(Me.ScrollBlue.Value) & "B"
    
    Dim MyWholeColor As String
    MyWholeColor = MyRVal & MyGVal & MyBVal
    Dim colorname As String
    Select Case MyWholeColor
        'example: a Dark Red and a Dk Green and a Dark Blue values are similar in color to a dk grey or blk.
        Case Is = "DRDGDB"
            colorname = "Dk Grey or black"
        Case Is = "DRDGMB"
            colorname = "Blue"
        Case Is = "DRDGLB"
            colorname = "Blue"
        Case Is = "DRMGDB"
            colorname = "Green"
        Case Is = "DRMGMB"
            colorname = "BlueGreen"
        Case Is = "DRMGLB"
            colorname = "Blue"
        Case Is = "DRLGDB"
            colorname = "Green"
        Case Is = "DRLGMB"
            colorname = "BlueGreen"
        Case Is = "DRLGLB"
            colorname = "Blue"
        Case Is = "MRDGDB"
            colorname = "Red"
        Case Is = "MRDGMB"
            colorname = "Purple"
        Case Is = "MRDGLB"
            colorname = "Purple"
        Case Is = "MRMGDB"
            colorname = "Yellow"
        Case Is = "MRMGMB"
            colorname = "Grey"
        Case Is = "MRMGLB"
            colorname = "BluePurple"
        Case Is = "MRLGDB"
            colorname = "Green"
        Case Is = "MRLGMB"
            colorname = "Green"
        Case Is = "MRLGLB"
            colorname = "BlueGreen"
        Case Is = "LRDGDB"
            colorname = "Red"
        Case Is = "LRDGMB"
            colorname = "RedPurple"
        Case Is = "LRDGLB"
            colorname = "Purple"
        Case Is = "LRMGDB"
            colorname = "Orange"
        Case Is = "LRMGMB"
            colorname = "Pink/Red"
        Case Is = "LRMGLB"
            colorname = "Purple"
        Case Is = "LRLGDB"
            colorname = "Yellow"
        Case Is = "LRLGMB"
            colorname = "Yellow"
        Case Is = "LRLGLB"
            colorname = "White/Grey"
    End Select
    MsgBox ("The color in the picture is " & colorname)
End Sub


Private Sub ScrollBlue_Change()
    'change the caption and the back color of the picture
    Me.lblblue.Caption = Me.ScrollBlue.Value
    Call ChangeBackColor(Me.ScrollRed, Me.ScrollGreen, Me.ScrollBlue)
End Sub

Private Sub ScrollGreen_Change()
    'change the caption and the back color of the picture
    Me.lblGreen.Caption = Me.ScrollGreen.Value
    Call ChangeBackColor(Me.ScrollRed, Me.ScrollGreen, Me.ScrollBlue)
End Sub

Private Sub ScrollRed_Change()
    'change the caption and the back color of the picture
    Me.lblRed.Caption = Me.ScrollRed.Value
    Call ChangeBackColor(Me.ScrollRed, Me.ScrollGreen, Me.ScrollBlue)
End Sub
Private Sub ChangeBackColor(RVal As Integer, GVal As Integer, BVal As Integer)
    'change the back color of the picture to reflect the croll bar values
    RVal = Me.ScrollRed.Value
    GVal = Me.ScrollGreen.Value
    BVal = Me.ScrollBlue.Value
    Me.Picture1.BackColor = RGB(RVal, GVal, BVal)
End Sub

Public Function Brightness(mycolor As Integer) As String
    'see what the brightness of the color is: dark, medium or light
    '"D" for Dark, "L" for light and "M" for med
    'If R value and G value and B Value are all less than 85 then it is a dk color
    'so will probably be either black or a dk grey (Example (50,50,50) would be a dk grey)
    'Same goes for R,G and B value greater than 170.  The closer the 3 values get to 255, the closer the
    'color will be to white.
    
    If mycolor < 86 Then
        Brightness = "D"
    ElseIf mycolor < 171 Then
        Brightness = "M"
    Else
        Brightness = "L"
    End If
End Function
