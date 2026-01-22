VERSION 5.00
Begin VB.Form ProductGroepen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ProductGroepen"
   ClientHeight    =   3225
   ClientLeft      =   5070
   ClientTop       =   3675
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cbItemsWijzigen 
      Caption         =   "Items &Wijzigen"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cbBewaren 
      Caption         =   "&Bewaren"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cbGroepItemToevoegen 
      Caption         =   "&Keuze Bijvoegen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox tbGroepItem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton cbGroepToevoegen 
      Caption         =   "&Groep Bijvoegen"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cbGroepDefinitie 
      BackColor       =   &H00C0FFC0&
      Height          =   960
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "cbGroepDefinitie"
      Top             =   120
      Width           =   3375
   End
   Begin VB.ListBox lbGroepItems 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
   End
End
Attribute VB_Name = "ProductGroepen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGroepenHier As ADODB.Recordset

Private Sub cbBewaren_Click()

    Dim StringHier As String
    Dim COUNT_TOHier As Integer
    
    On Local Error Resume Next
    For COUNT_TOHier = 0 To Me.lbGroepItems.ListCount - 1
        StringHier = StringHier & Me.lbGroepItems.List(COUNT_TOHier) & ";"
    Next
    rsGroepenHier.Fields("GroepItems") = StringHier
    rsGroepenHier.Update
    Knoppen (True)
    
End Sub

Private Sub cbGroepDefinitie_Click()
    
    Dim splitstring() As String
    Dim COUNT_TOHier As Integer
    
    On Error Resume Next
    Me.lbGroepItems.Clear
    Err = 0
    rsGroepenHier.Seek cbGroepDefinitie.List(cbGroepDefinitie.ListIndex), adSeekFirstEQ
    If Err Then
    Else
        'lijstje samenstellen
        splitstring = Split(rsGroepenHier.Fields("GroepItems"), ";")
        For COUNT_TOHier = 0 To UBound(splitstring) - 1
            lbGroepItems.AddItem splitstring(COUNT_TOHier)
        Next
        
    End If
    
End Sub

Private Sub cbGroepItemToevoegen_Click()

    If Me.tbGroepItem.Text = "" Then Exit Sub
    If lbGroepItems.ListIndex < 0 Then
        Me.lbGroepItems.AddItem Me.tbGroepItem.Text
    Else
        Me.lbGroepItems.AddItem Me.tbGroepItem.Text, lbGroepItems.ListIndex
    End If
    With Me.tbGroepItem
        .Text = ""
        .SetFocus
    End With
    
End Sub

Private Sub cbGroepToevoegen_Click()

    If MsgBox("Nieuwe groep " & Me.cbGroepDefinitie.Text & " bijvoegen.  Bent U zeker", vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then
        On Error Resume Next
        rsGroepenHier.AddNew
        rsGroepenHier.Fields("GroepsNaam") = Me.cbGroepDefinitie.Text
        rsGroepenHier.Fields("GroepItems") = ""
        Err = 0
        rsGroepenHier.Update
        If Err Then
            MsgBox Error, vbInformation
        Else
            refreshGroepen
        End If
    End If
    
End Sub

Private Sub cbItemsWijzigen_Click()

     Knoppen (False)
     
End Sub

Private Sub cbSluiten_Click()

    Unload Me
        
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Err = 0
    adntDB.Execute ("CREATE TABLE p_Groepen ( ID int IDENTITY (1,1), GroepsNaam varchar(60) UNIQUE, GroepItems MEMO, CONSTRAINT p_PK PRIMARY KEY (ID))")
    If Err Then
        'niks het is reeds geïnitialiseerd voor gebruik
    Else
        adntDB.Execute ("INSERT INTO p_Groepen (GroepsNaam, GroepItems) VALUES ('Categorie', 'zonder voorwerp');")
        adntDB.Execute ("INSERT INTO p_Groepen (GroepsNaam, GroepItems) VALUES ('Merk', 'zonder voorwerp');")
        adntDB.Execute ("CREATE UNIQUE INDEX GroepsNaam ON p_Groepen (GroepsNaam)")
    End If
    Set rsGroepenHier = New ADODB.Recordset
    rsGroepenHier.CursorLocation = adUseServer
    rsGroepenHier.Open "p_Groepen", adntDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect '  adLockReadOnly, adCmdTableDirect
    rsGroepenHier.Index = "GroepsNaam"
    refreshGroepen
                  
End Sub

Function refreshGroepen()

    rsGroepenHier.MoveFirst
    Me.cbGroepDefinitie.Clear
    Do While Not rsGroepenHier.EOF
        Me.cbGroepDefinitie.AddItem rsGroepenHier.Fields("GroepsNaam")
        rsGroepenHier.MoveNext
    Loop
    If Me.cbGroepDefinitie.ListCount > 0 Then
        Me.cbGroepDefinitie.ListIndex = 0
    End If

End Function

Function Knoppen(Vlag As Boolean)

    Me.cbGroepToevoegen.Enabled = Vlag
    Me.cbGroepDefinitie.Enabled = Vlag
    Me.cbItemsWijzigen.Enabled = Vlag
    
    Me.lbGroepItems.Enabled = Not Vlag
    Me.tbGroepItem.Enabled = Not Vlag
    Me.cbGroepItemToevoegen.Enabled = Not Vlag
    Me.cbBewaren.Enabled = Not Vlag
        
End Function


Private Sub lbGroepItems_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    If KeyCode = 46 Then
        lbGroepItems.RemoveItem lbGroepItems.ListIndex
    End If
    
End Sub

