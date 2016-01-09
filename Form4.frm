VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form4"
   ScaleHeight     =   9615
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   4680
      TabIndex        =   8
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   4680
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   1095
      Left            =   5400
      TabIndex        =   5
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   1095
      Left            =   1440
      TabIndex        =   4
      Top             =   7320
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   4560
      TabIndex        =   1
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Participant ID"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Paper Name "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reviewed Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from RevMan where Parid='" & Text1.Text & "'", cn.ConnectionString, adOpenDynamic, adLockOptimistic
If rs.EOF Or rs.BOF Then
MsgBox "Your paper will be reviewed very shortly!!", vbOKOnly + vbExclamation, "TECHSAT TEAM"
Else
Text1.Text = rs!Parid
Text2.Text = rs!PaperName
Text3.Text = rs!RevStat
rs.Close
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=it13;Initial Catalog=coms;Data Source=MSECSERVER1;Password=Msec*100"
End Sub
