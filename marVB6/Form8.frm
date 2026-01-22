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
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   795
      Left            =   1740
      Picture         =   "Form8.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1380
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()

   Dim cnn1 As ADODB.Connection
   Dim rstEmployees As ADODB.Recordset
   Dim strCnn As String
   Dim varDate As Variant

   Dim daoCnn2 As DAO.Connection
   Dim daoRstEmpoyee As DAO.Recordset
   Dim daows As Workspace
      
     
   'open connection DAO
   daows.OpenDatabase ""
   
   
   ' Open connection ADO
      strCnn = "Provider=sqloledb;" & _
      "Data Source=srv;Initial Catalog=pubs;User Id=sa;Password=; "
   Set cnn1 = New ADODB.Connection
   cnn1.Open strCnn
   
   ' Open employee table.
   Set rstEmployees = New ADODB.Recordset
   rstEmployees.CursorType = adOpenKeyset
   rstEmployees.LockType = adLockOptimistic
   rstEmployees.Open "employee", cnn1, , , adCmdTable

   ' Assign the first employee record's hire date
   ' to a variable, then change the hire date.
   varDate = rstEmployees!hire_date
   Debug.Print "Original data"
   Debug.Print "  Name - Hire Date"
   Debug.Print "  " & rstEmployees!fName & " " & _
      rstEmployees!lName & " - " & rstEmployees!hire_date
   rstEmployees!hire_date = #1/1/1900#
   rstEmployees.Update
   Debug.Print "Changed data"
   Debug.Print "  Name - Hire Date"
   Debug.Print "  " & rstEmployees!fName & " " & _
      rstEmployees!lName & " - " & rstEmployees!hire_date

   ' Requery Recordset and reset the hire date.
   rstEmployees.Requery
   rstEmployees!hire_date = varDate
   rstEmployees.Update
   Debug.Print "Data after reset"
   Debug.Print "  Name - Hire Date"
   Debug.Print "  " & rstEmployees!fName & " " & _
      rstEmployees!lName & " - " & rstEmployees!hire_date

   rstEmployees.Close
   cnn1.Close

End Sub

