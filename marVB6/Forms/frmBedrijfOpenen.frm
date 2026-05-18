VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form BedrijfOpenen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bedrijf Openen"
   ClientHeight    =   2940
   ClientLeft      =   2130
   ClientTop       =   2100
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   11000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   7500
   Begin VB.CommandButton ButtonOpenFolder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      Picture         =   "frmBedrijfOpenen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "MARNT\DATA Locatie openen"
      Top             =   2160
      Width           =   435
   End
   Begin VB.OptionButton opLocatie 
      Alignment       =   1  'Right Justify
      Caption         =   "ServerData"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton opLocatie 
      Caption         =   "Data lokaal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox lblLokatie 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   6135
   End
   Begin VB.CommandButton cmdInit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      Picture         =   "frmBedrijfOpenen.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Locatie MARNT\DATA wijzigen"
      Top             =   2160
      Width           =   435
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   2990
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Benaming"
         Object.Width           =   11192
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Map"
         Object.Width           =   1005
      EndProperty
   End
   Begin VB.CommandButton CmbCompactDatabase 
      Caption         =   "Database &Vernieuwen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1740
   End
   Begin VB.CommandButton OK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   720
      Left            =   1140
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   1440
   End
   Begin VB.CommandButton Annuleren 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Annuleren"
      Height          =   720
      Left            =   2700
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   540
      Width           =   1440
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "BedrijfOpenen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Dim flgVerwijderen As Boolean
Dim strDataLocatie As String

Sub KeuzeLijstVullen()
Dim FlTemp As Integer

Dim A As String
Dim NaamDetail As String

Dim MyPath As String
Dim MyName As String

' Declare a variable to add ListItem objects.
Dim itmX As ListItem

ListView1.ListItems.Clear

MyPath = Me.lblLokatie & "\" 'LOCATION_ ' Set the path.

Err = 0
On Error Resume Next
MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
If Err Then Exit Sub
Do While MyName <> ""   ' Start the loop.
    ' Ignore the current directory and the encompassing directory.
    If MyName <> "." And MyName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
            On Error Resume Next
            Err = 0
            FlTemp = FreeFile
            Open MyPath + MyName + "\marnt.txt" For Input As FlTemp
            If Err Then
            Else
                Line Input #FlTemp, A
                Close FlTemp
                NaamDetail = A
                ' Add a ListItem object.
                Set itmX = ListView1.ListItems.Add(, , NaamDetail)
                itmX.SubItems(1) = MyName
            End If
        End If   ' it represents a directory.
   End If
   MyName = Dir   ' Get next entry.
Loop

End Sub

Private Sub Annuleren_Click()

Unload Me

End Sub


Private Sub ButtonOpenFolder_Click()

    Dim r As Long
    r = ShellExecute(0, "open", Me.lblLokatie.text, 0, 0, 1)
    
End Sub

Private Sub cbVerwijderen_Click()

    flgVerwijderen = Not flgVerwijderen
    ButtonOpenFolder.Enabled = Not ButtonOpenFolder.Enabled
    cmdInit.Enabled = Not cmdInit.Enabled
    
End Sub

Private Sub CmbCompactDatabase_Click()

If ListView1.ListItems.Count = 0 Then Exit Sub

    
On Error GoTo ErGingIetsFout:
jetConnect = ADOJET_PROVIDER & _
    "Data Source=" & LOCATION_ & ListView1.SelectedItem.ListSubItems(1) _
    & "\marnt.mdv;" & _
    "Persist Security Info=False"

Set adntDB = New ADODB.Connection
adntDB.Open jetConnect
Msg = "Huidige database in JetVersie 4.x vernieuwen" & vbCr
Msg = Msg & "Microsoft ADO Versie " & adntDB.Version & vbCr & vbCr
Msg = Msg & "LOCATION_ : " & LOCATION_ & ListView1.SelectedItem.ListSubItems(1) & "\marnt.MDV" & vbCr & vbCr
adntDB.Close
Set adntDB = Nothing

Msg = Msg & "Regelmatig te gebruiken indien U zelf tabellen, velden, indexen, SQL-opvraagdefinities aanmaakt en/of verwijdert + nadat journalen en dokumenten van een of meer boekjaren werden opgekuist." & vbCr & vbCr
Msg = Msg & "Onderhoud database.  Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2, "Database vernieuwen (compact maken)")
If KtrlBox = vbYes Then
    Screen.MousePointer = vbHourglass
    If Dir(LOCATION_ & ListView1.SelectedItem.ListSubItems(1) & "\marnt.$$$") <> "" Then
        Kill LOCATION_ & ListView1.SelectedItem.ListSubItems(1) & "\marnt.$$$"
    End If
    SnelHelpPrint "Bezig...", BL_LOGGING
    Mim.SnelHelp.Refresh
    Me.Refresh
    
    Dim JRO As JRO.JetEngine
    Set JRO = New JRO.JetEngine

    'You must provide the old database name and the new database name.
    'You must also provide a valid connection string, as well as
    'indicate the type of Jet engine to compact the database as.
    'We used 5 to create a Jet Version 4.x database suitable for
    'Access 2000 users. If you work with Access 97 use the
    'value 4 to create a Jet Version 3.x database.
    
    JRO.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" _
        & LOCATION_ & ListView1.SelectedItem.ListSubItems(1) & "\marnt.mdv", _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" _
        & LOCATION_ + ListView1.SelectedItem.ListSubItems(1) + "\marnt.$$$" _
        & ";Jet OLEDB:Engine Type=5"
    Kill LOCATION_ + ListView1.SelectedItem.ListSubItems(1) + "\marnt.mdv"
    Name LOCATION_ + ListView1.SelectedItem.ListSubItems(1) + "\marnt.$$$" As LOCATION_ + ListView1.SelectedItem.ListSubItems(1) + "\marnt.mdv"
    SnelHelpPrint "Klaar !", BL_LOGGING
    Screen.MousePointer = vbNormal
End If
ListView1.SetFocus
Exit Sub

ErGingIetsFout:
Screen.MousePointer = vbNormal
If Err = 3301 Then
    MsgBox "U kan geen database vernieuwen naar een lagere versie !"
Else
    MsgBox Error
End If
Exit Sub

End Sub


Private Sub cmdInit_Click()
    
    lblLokatie.Enabled = Not lblLokatie.Enabled
    If lblLokatie.Enabled = True Then
        lblLokatie.SetFocus
    Else
        If Me.opLocatie(0).Value = True Then
            If lblLokatie <> LaadTekst(App.Title, "Bedrijfsinhoudsopgave" & "2025") Then
                KtrlBox = MsgBox(lblLokatie + vbCrLf + vbCrLf + "Wordt dit de nieuwe 'lokale' opstartinhoudsopgave ?", vbQuestion + vbDefaultButton2 + vbYesNo)
                If KtrlBox = vbYes Then
                    BeWaarTekst App.Title, "Bedrijfsinhoudsopgave" & "2025", lblLokatie
                    MsgBox "Hierna wordt er afgesloten.  Start het programma opnieuw op"
                    Unload Mim
                Else
                    lblLokatie = LaadTekst(App.Title, "Bedrijfsinhoudsopgave" & "2025")
                End If
            End If
        Else
            If lblLokatie <> LaadTekst(App.Title, "ServerBedrijfsinhoudsopgave") Then
                KtrlBox = MsgBox(lblLokatie + vbCrLf + vbCrLf + "Wordt dit de nieuwe 'server' opstartinhoudsopgave ?", vbQuestion + vbDefaultButton2 + vbYesNo)
                If KtrlBox = vbYes Then
                    BeWaarTekst App.Title, "ServerBedrijfsinhoudsopgave", lblLokatie
                    MsgBox "Hierna wordt er afgesloten.  Start het programma opnieuw op"
                    Unload Mim
                Else
                    lblLokatie = LaadTekst(App.Title, "ServerBedrijfsinhoudsopgave")
                End If
            End If
        End If
    End If
    
End Sub


Private Sub Form_Load()

flgVerwijderen = False

Top = 0
Left = 0
' Add ColumnHeaders. The width of the columns is
' the width of the control divided by the number of
' ColumnHeader objects.
'ListView1.ColumnHeaders.Add , , "Benaming", 6345
'ListView1.ColumnHeaders.Add , , "Map", 570
  
' Set View property to Report.
ListView1.View = lvwReport

strDataLocatie = LaadTekst("BedrijfOpenen", "DataDefault")
If Trim$(strDataLocatie) = "" Then
    strDataLocatie = "lokaal"
End If
If strDataLocatie = "server" Then
    Me.opLocatie(1).Value = True
    Mim.Basis(4).Enabled = False
Else
    Me.opLocatie(0).Value = True
    Mim.Basis(4).Enabled = True
End If

VSF_PRO = False
Err = 0
On Error Resume Next
'dlbFolder.path = LOCATION_
If Err Then
    MsgBox "LOCATION_ bedrijven onvindbaar.  Kontroleer manueel a.u.b.", vbCritical
Else
    KeuzeLijstVullen
End If

End Sub


'Private Sub ListView1_Click()

    'Ok_Click

'End Sub


Private Sub ListView1_DblClick()

    Ok_Click

End Sub


Private Sub ListView1_GotFocus()

    SnelHelpPrint "[Enter] of [Dubbelklikken om te openen], [Delete] om te verwijderen", BL_LOGGING

End Sub


Private Sub Ok_Click()

If ListView1.ListItems.Count = 0 Then Exit Sub
LOCATION_COMPANYDATA = LOCATION_ + ListView1.SelectedItem.ListSubItems(1) + "\"
Mim.Caption = appTitleAndVersion & " - [" & Trim$(ListView1.SelectedItem) & "]"
AutoLoadBedrijf
If Mim.Caption = appTitleAndVersion Then
    BedrijfOpenen.SetFocus
Else
    Annuleren_Click
    If Val(ProducentNummer) <> 0 Then Venster.Show: DoEvents: Venster.SetFocus
End If

End Sub

Private Sub opLocatie_Click(Index As Integer)

    Err = 0
    On Error Resume Next
    'dlbFolder.path = LOCATION_
    If Me.opLocatie(1).Value = True Then
        lblLokatie.text = LaadTekst(App.Title, "ServerBedrijfsinhoudsopgave")  'Server anders
        BeWaarTekst "BedrijfOpenen", "DataDefault", "server"
        Me.CmbCompactDatabase.Enabled = False
        Mim.Basis(4).Enabled = False
    Else
        lblLokatie.text = LaadTekst(App.Title, "Bedrijfsinhoudsopgave" & "2025") 'Lokaal is standaard
        BeWaarTekst "BedrijfOpenen", "DataDefault", "lokaal"
        Me.CmbCompactDatabase.Enabled = True
        Mim.Basis(4).Enabled = True
    End If
    If Err Or Trim$(lblLokatie.text) = "" Then
        SnelHelpPrint "Verbinding controleren en/of instellen a.u.b.", False
    End If
    LOCATION_ = lblLokatie.text + "\"
    KeuzeLijstVullen
    
End Sub
