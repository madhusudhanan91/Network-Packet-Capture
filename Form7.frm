VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from RevDet where Parid='" & Text1.Text & "'", cn.ConnectionString, adOpenDynamic, adLockOptimistic
If rs.EOF Or rs.BOF Then
rs.AddNew
rs!Parid = Text1.Text
rs!PaperName = Text2.Text
rs!RevStat = Text3.Text
rs.Update
rs.Close
Else
MsgBox "Wrong Information"
End If

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=it13;Initial Catalog=coms;Data Source=MSECSERVER1;Password=Msec*100"
End Sub

