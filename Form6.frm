VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form6"
   ScaleHeight     =   9645
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   975
      Left            =   5400
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   975
      Left            =   1800
      TabIndex        =   5
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   5040
      TabIndex        =   4
      Top             =   3960
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   5040
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   2
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reviewer Login"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from ReviewTeam where revid='" & Text1.Text & "'", cn.ConnectionString, adOpenDynamic, adLockOptimistic
If rs.EOF Or rs.BOF Then
MsgBox "INVALID USER ID!!", vbOKOnly + vbCritical, "ID ERROR"
ElseIf (Text2.Text = rs!passlink) Then
MsgBox "LOGIN SUCCESSFULL!!", vbOKOnly + vbInformation, "LOGIN"
Unload Me
Form3.Show
Else
MsgBox "INVALID PASSWORD!!", vbOKOnly + vbCritical, "PASSWORD ERROR"
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=it13;Initial Catalog=coms;Data Source=MSECSERVER1;Password=Msec*100"
End Sub

Private Sub Text2_Change()
Text2.PasswordChar = "*"
End Sub
