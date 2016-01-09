VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form2"
   ScaleHeight     =   9435
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
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
      Left            =   6840
      TabIndex        =   14
      Top             =   8400
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   4320
      List            =   "Form2.frx":000D
      TabIndex        =   13
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate Participant ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   8400
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   4320
      TabIndex        =   9
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   735
      Left            =   4320
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Date of Paper to be presented"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Area Of Paper Presentation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Paper Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   135
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   15
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Participant Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Participant ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please Enter the Participant Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from Participants where Parid='" & Text1.Text & "'", cn.ConnectionString, adOpenDynamic, adLockOptimistic
If (rs.EOF Or rs.BOF) And ((Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Combo1.Text <> "")) Then
rs.AddNew
rs!Parid = Text1.Text
rs!ParticipantName = Text2.Text
rs!PaperName = Text3.Text
rs!AreaofPPT = Text4.Text
rs!DateOfPPT = Combo1.Text
rs.Update
rs.Close
MsgBox "Thanks For Registering !", vbOKOnly + vbInformation, "TECHSAT 2011"
Else
MsgBox "Please fill in all the fields"
End If
End Sub

Private Sub Command2_Click()
Randomize Timer
RandomNumber = Int(Rnd * 5000)
Text1.Text = Int(Rnd * 5000)
Command2.Enabled = False
Text1.Enabled = False
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=it13;Initial Catalog=coms;Data Source=MSECSERVER1;Password=Msec*100"
End Sub
