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
      Height          =   495
      Left            =   3180
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpen_Click()

' Create ADO objects.
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command

    ' Open a connection using a client cursor.
    conn.CursorLocation = adUseServer
    conn.ConnectionString = "Provider=SQLOLEDB;DataSource=(local);Database=ZakenkantoorRV;uid=sa;pwd=;"
    conn.Open
    
    ' Set options on the Command object.
    cmd.CommandText = "SELECT * FROM Produkten"
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = conn

    ' Open a recordset and then disconnect.
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open cmd
    'Set rs.ActiveConnection = Nothing
       
    ' Close the connection and release 2 of the ADO objects.
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    
    ' Assign the recordset to the ADO data control on the form.
    'Set dsoRecordset.Recordset = rs
Stop
SupportsX

End Sub

Public Sub SupportsX()

   Dim aintCursorType(4) As Integer
   Dim rstTitles As ADODB.Recordset
   Dim strCnn As String
   Dim intIndex As Integer

   ' Open connections.
      strCnn = "Provider=sqloledb;" & _
      "Data Source=(local);Initial Catalog=pubs;User Id=sa;Password=; "

   ' Fill array with CursorType constants.
   aintCursorType(0) = adOpenForwardOnly
   aintCursorType(1) = adOpenKeyset
   aintCursorType(2) = adOpenDynamic
   aintCursorType(3) = adOpenStatic
      
   ' Open recordset using each CursorType and
   ' optimistic locking. Then call the DisplaySupport
   ' procedure to display the supported options.
   For intIndex = 0 To 3
      Set rstTitles = New ADODB.Recordset
      rstTitles.CursorType = aintCursorType(intIndex)
      rstTitles.LockType = adLockOptimistic
      rstTitles.Open "Titles", strCnn, , , adCmdTable
      
      Select Case aintCursorType(intIndex)
         Case adOpenForwardOnly
            Debug.Print "ForwardOnly cursor supports:"
         Case adOpenKeyset
            Debug.Print "Keyset cursor supports:"
         Case adOpenDynamic
            Debug.Print "Dynamic cursor supports:"
         Case adOpenStatic
            Debug.Print "Static cursor supports:"
            
      End Select

      DisplaySupport rstTitles
      rstTitles.Close
   Next intIndex

End Sub

Public Sub DisplaySupport(rstTemp As ADODB.Recordset)

   Dim alngConstants(13) As Long
   Dim booSupports As Boolean
   Dim intIndex As Integer

   ' Fill array with cursor option constants.
   alngConstants(0) = adAddNew
   alngConstants(1) = adApproxPosition
   alngConstants(2) = adBookmark
   alngConstants(3) = adDelete
   alngConstants(4) = adFind
   alngConstants(5) = adHoldRecords
   alngConstants(6) = adMovePrevious
   alngConstants(7) = adNotify
   alngConstants(8) = adResync
   alngConstants(9) = adUpdate
   alngConstants(10) = adUpdateBatch
   alngConstants(11) = adSeek
   alngConstants(12) = adIndex
   
   For intIndex = 0 To 12
      booSupports = _
         rstTemp.Supports(alngConstants(intIndex))
      If booSupports Then
         Select Case alngConstants(intIndex)
            Case adAddNew
               Debug.Print "   AddNew"
            Case adApproxPosition
               Debug.Print "   AbsolutePosition and AbsolutePage"
            Case adBookmark
               Debug.Print "   Bookmark"
            Case adDelete
               Debug.Print "   Delete"
            Case adFind
               Debug.Print "   Find"
            Case adHoldRecords
               Debug.Print "   Holding Records"
            Case adMovePrevious
               Debug.Print "   MovePrevious and Move"
            Case adNotify
               Debug.Print "   Notifications"
            Case adResync
               Debug.Print "   Resyncing data"
            Case adUpdate
               Debug.Print "   Update"
            Case adUpdateBatch
               Debug.Print "   batch updating"
         End Select
      End If
   Next intIndex

End Sub



