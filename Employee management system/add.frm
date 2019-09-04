VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form add 
   BackColor       =   &H0080FF80&
   Caption         =   "Add new Employee"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   Picture         =   "add.frx":0000
   ScaleHeight     =   6045
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   10200
      Picture         =   "add.frx":C0E1
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      Picture         =   "add.frx":D638
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   10920
      Picture         =   "add.frx":E9D5
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   0
      Width           =   495
   End
   Begin MSAdodcLib.Adodc a1 
      Height          =   375
      Left            =   4920
      Top             =   5520
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\shubh\Desktop\VB Project\emp.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\shubh\Desktop\VB Project\emp.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "emp"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6600
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      DataField       =   "doj"
      DataSource      =   "a1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   11
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      DataField       =   "salary"
      DataSource      =   "a1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "post"
      DataSource      =   "a1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "name"
      DataSource      =   "a1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "password"
      DataSource      =   "a1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "userid"
      DataSource      =   "a1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"add.frx":FFA7
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
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11415
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1920
      X2              =   1920
      Y1              =   1680
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label11 
      BackColor       =   &H000080FF&
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
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label10 
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
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label9 
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
      Left            =   0
      TabIndex        =   16
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label8 
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
      Left            =   0
      TabIndex        =   15
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   1680
      Y2              =   5040
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3615
      Left            =   5280
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   " DATE OF JOINING"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "      SALARY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "        POST"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "        NAME"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "    PASSWORD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "      USER ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a1.Recordset.AddNew
End Sub


Private Sub Label10_Click()
home.Visible = False
login.Visible = True
End Sub

Private Sub Label8_Click()
home.Visible = False
list.Visible = True
End Sub

Private Sub Label9_Click()
home.Visible = False
delete.Visible = True
End Sub

Private Sub Picture2_Click()
End
End Sub

Private Sub Picture3_Click()
home.Visible = True
End Sub
