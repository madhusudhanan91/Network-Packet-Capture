VERSION 5.00
Object = "Word.Document.8"; "WINWORD.EXE"
Begin VB.Form Form5 
   BackColor       =   &H80000016&
   Caption         =   "Form5"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form5"
   MousePointer    =   4  'Icon
   ScaleHeight     =   8865
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin WordCtl.Document Document1 
      Height          =   4695
      Left            =   300
      OleObjectBlob   =   "Form5.frx":0000
      TabIndex        =   2
      Top             =   3000
      Width           =   8640
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000000&
      Caption         =   "ABOUT US"
      Height          =   375
      Left            =   240
      MousePointer    =   6  'Size NE SW
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "TECHSAT 2011"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
