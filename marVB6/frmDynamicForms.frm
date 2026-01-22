VERSION 5.00
Begin VB.Form frmDynaVensters 
   Caption         =   "Dynamisch Vensters Openen"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmDynaVensters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'Forms object lets you refer to items either by index or by name,
'as seen below:
'Set frm = Forms(1)
'or
'Set frm = Forms("frmEditor")
'And of course, you can substitute a variable for any of the index
'values, like so:
'Set frm = Forms(strName)
'With this collection available, it becomes a relatively easy
'matter to load a form dynamically. The following code shows a
'simplified way to do so:

With List1
   .AddItem "frmGroeneKaart"
   .AddItem "Form2"
   .AddItem "Form3"
End With

End Sub

Private Sub List1_Click()

    Dim frm As Form
    Dim selForm As String
    
    With List1
        selForm = .List(.ListIndex)
    End With
    Set frm = Forms.Add(selForm)
    frm.Show
    Set frm = Nothing

End Sub
