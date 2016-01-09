VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form3"
   ScaleHeight     =   8520
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form3.frx":0000
      Left            =   4440
      List            =   "Form3.frx":0002
      TabIndex        =   10
      Top             =   3840
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form3.frx":0004
      Left            =   4560
      List            =   "Form3.frx":0011
      TabIndex        =   9
      Top             =   7680
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit"
      Height          =   1215
      Left            =   1560
      TabIndex        =   7
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   1095
      Left            =   6240
      TabIndex        =   6
      Top             =   9120
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   4440
      TabIndex        =   5
      Top             =   5400
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Reviewed Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Paper Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Participant ID "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Reviewer ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from RevMan where Parid=' " & Combo2.Text & "'", cn.ConnectionString, adOpenDynamic, adLockOptimistic
If rs.EOF Or rs.BOF Then
rs.AddNew
rs!revid = Text1.Text
rs!Parid = Combo2.Text
rs!PaperName = Text3.Text
rs!RevStat = Combo1.Text
rs.Update
rs.Close
MsgBox "Details Submitted"
Else
MsgBox "Wrong Information"
End If
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=it13;Initial Catalog=coms;Data Source=MSECSERVER1;Password=Msec*100"
Set rs1 = New ADODB.Recordset
rs1.Open "select * from Participants ", cn.ConnectionString, adOpenDynamic, adLockOptimistic
While Not rs1.EOF
Combo2.AddItem rs1!Parid
rs1.MoveNext
Wend
rs1.Close
End Sub

