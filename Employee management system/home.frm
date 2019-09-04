VERSION 5.00
Begin VB.Form home 
   BackColor       =   &H0080FF80&
   Caption         =   "Welcome to Engineering Library- Employee management System"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   Picture         =   "home.frx":0000
   ScaleHeight     =   5895
   ScaleMode       =   0  'User
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   10320
      Picture         =   "home.frx":C0E1
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   11040
      Picture         =   "home.frx":D638
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   120
      Picture         =   "home.frx":EC0A
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Twitter - www.twitter.com/engglibrary"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   4320
      Width           =   7095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Instagram - www.instagram.com/engglibrary"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   3840
      Width           =   7095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Facebook - www.facebook.com/engglibrary"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   3360
      Width           =   7095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Website - www.engglibrary.com"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2880
      Width           =   7095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Follow Us on"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2400
      Width           =   7095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to ENGINEERING LIBRARY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   7095
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   120
      X2              =   2040
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   120
      X2              =   2040
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   120
      X2              =   2040
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "    SEARCH"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "    REMOVE        RECORD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "     UPDATE        RECORD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "       ADD           EMPLOYEE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   120
      X2              =   2040
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   120
      X2              =   2040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2040
      X2              =   2040
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"home.frx":FFA7
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label2_Click()
home.Visible = False
add.Visible = True
End Sub

Private Sub Label3_Click()
home.Visible = False
login.Visible = True
End Sub

Private Sub Label4_Click()
home.Visible = False
delete.Visible = True
End Sub

Private Sub Label5_Click()
home.Visible = False
list.Visible = True
End Sub

Private Sub Picture2_Click()
End

End Sub

Private Sub Picture3_Click()
home.Visible = True
End Sub
