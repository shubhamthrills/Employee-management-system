VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   BackColor       =   &H0080FF80&
   Caption         =   "Update Employee Details"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   Picture         =   "login.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   10080
      Picture         =   "login.frx":C0E1
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   24
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      Picture         =   "login.frx":D638
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   10920
      Picture         =   "login.frx":E9D5
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Text8 
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
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text7 
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
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   3480
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
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3960
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
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   4440
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
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
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
      Left            =   6000
      TabIndex        =   6
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5400
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc a1 
      Height          =   330
      Left            =   8400
      Top             =   6120
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
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
   Begin VB.TextBox Text1 
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
      Left            =   6000
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text2 
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
      Left            =   6000
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "LOGIN"
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
      Left            =   5400
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"login.frx":FFA7
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11415
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1920
      X2              =   1920
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label13 
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
      Left            =   0
      TabIndex        =   23
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H000080FF&
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
      TabIndex        =   22
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label11 
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
      TabIndex        =   21
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label10 
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
      TabIndex        =   20
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   0
      X2              =   1920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Label Label9 
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
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   3600
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
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   4080
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
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
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
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   5040
      Width           =   1455
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
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   3735
      Left            =   3960
      Top             =   2880
      Width           =   4335
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
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
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
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1815
      Left            =   3960
      Top             =   1080
      Width           =   4335
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer

Private Sub Command1_Click()
'If a1.Recordset.Fields(0) = Text1.Text And a1.Recordset.Fields(1) = Text2.Text Then
'MsgBox "Login successful"
'f = 0
a1.Recordset.MoveFirst
Do While Not a1.Recordset.EOF = True
If a1.Recordset.Fields(0) = Text1.Text And a1.Recordset.Fields(1) = Text2.Text Then
MsgBox "Login successful"
Text8.Text = a1.Recordset.Fields(0)
Text7.Text = a1.Recordset.Fields(1)
Text3.Text = a1.Recordset.Fields(2)
Text4.Text = a1.Recordset.Fields(3)
Text6.Text = a1.Recordset.Fields(5)
Text5.Text = a1.Recordset.Fields(4)
f = 1
Exit Do
Else
a1.Recordset.MoveNext
End If
Loop
If f = 0 Then
MsgBox "USERNAME or PASSWORD INCRRECT"
End If
'Else
'MsgBox "USERNAME / PASSWORD INCRRECT"
'End If

End Sub

Private Sub Command2_Click()
a1.Recordset.MoveFirst
Do While Not a1.Recordset.EOF = True
If a1.Recordset.Fields(0) = Text1.Text And a1.Recordset.Fields(1) = Text2.Text Then
MsgBox "PRESS OK TO UPDATE"
a1.Recordset.Fields(0) = Text8.Text
a1.Recordset.Fields(1) = Text7.Text
a1.Recordset.Fields(2) = Text3.Text
a1.Recordset.Fields(3) = Text4.Text
a1.Recordset.Fields(5) = Text6.Text
a1.Recordset.Fields(4) = Text5.Text
f = 1
Exit Do
Else
a1.Recordset.MoveNext
End If
Loop
If f = 0 Then
MsgBox "USERNAME or PASSWORD INCORRECT"
Else
MsgBox "DATA UPDATED"
End If

'End If
End Sub


Private Sub Label10_Click()
home.Visible = False
list.Visible = True
End Sub

Private Sub Label11_Click()
home.Visible = False
delete.Visible = True
End Sub

Private Sub Label13_Click()
home.Visible = False
add.Visible = True
End Sub

Private Sub Picture2_Click()
End
End Sub

Private Sub Picture3_Click()
home.Visible = True
End Sub
