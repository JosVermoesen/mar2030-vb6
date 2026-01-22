VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As ADODB.Connection

Private Sub Form_Load()

' Open a connection using a DSN and ODBC tags.
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "DSN=test001;UID=sa;"
  cnn.Open
  Stop
  
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.CursorLocation = adUseServer
  rs.Source = "Klanten"
  rs.ActiveConnection = cnn
  rs.CursorType = adOpenKeyset
  rs.LockType = adLockOptimistic
  rs.Open , , , , adCmdTableDirect
  
  Stop
  

End Sub
