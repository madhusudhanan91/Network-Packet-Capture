VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000008&
      Caption         =   "Exit"
      Height          =   1215
      Left            =   4920
      TabIndex        =   10
      Top             =   9120
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000008&
      Caption         =   "TECHSAT Details"
      Height          =   1215
      Left            =   4920
      TabIndex        =   8
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000008&
      Caption         =   "Details"
      Height          =   1095
      Left            =   4920
      TabIndex        =   6
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000008&
      Caption         =   "Login "
      Height          =   975
      Left            =   4920
      TabIndex        =   3
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      Caption         =   "Participants Registration"
      Height          =   855
      Left            =   4920
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "To leave The Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   9120
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "To know more about TECHSAT 2011"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Reviewed Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Registration for the conference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1125
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Reviewer Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   2685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Welcome To TECHSAT 2011      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Form6.Show
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub

Private Sub Command5_Click()
Form1.Visible = False
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=it13;Initial Catalog=coms;Data Source=MSECSERVER1;Password=Msec*100"
End Sub


