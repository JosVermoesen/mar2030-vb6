VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmOntwerpSQLVenster 
   Caption         =   "Allerlei Testen"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4485
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1755
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   3096
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdRTFOpen 
      Caption         =   "RTF Bestand Openen"
      Height          =   330
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1755
   End
   Begin VB.CommandButton cmdSpelling 
      Caption         =   "Spellingscontrole"
      Height          =   330
      Left            =   2880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1515
   End
   Begin VB.CommandButton cmdDataOmgeving 
      Caption         =   "DataOmgeving Test"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin VB.OptionButton optBewerkbaar 
      Caption         =   "&Bewerkbaar"
      Height          =   255
      Left            =   1740
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton optRooster 
      Caption         =   "&Niet bewerkbaar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdMaakVenster 
      Caption         =   "&SQL Resultaat"
      Height          =   330
      Left            =   3120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   495
      Left            =   1620
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
End
Attribute VB_Name = "frmOntwerpSQLVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub chkEditeerbaar_Click()

End Sub

Private Sub cmdDataOmgeving_Click()

frmDataOmgeving.Show

End Sub

Private Sub cmdMaakVenster_Click()
Dim FormHier As Object
Set FormHier = New mijnSQLEditor

If rtbTekst.Text = "" Then Exit Sub

If optBewerkbaar.Value = True Then
    FormHier!msfSQL.AllowUpdate = True
    FormHier!msfSQL.AllowAddNew = True
    FormHier!msfSQL.AllowDelete = True
Else
    FormHier!msfSQL.AllowUpdate = False
    FormHier!msfSQL.AllowAddNew = False
    FormHier!msfSQL.AllowDelete = False
End If
FormHier.txtSQL = rtbTekst.Text
FormHier.Show

FormHier.SetFocus
SendKeys "{ENTER}"

End Sub

Private Sub cmdRTFOpen_Click()

    Mim.Teken.Filter = "Rich Text Format bestanden|*.rtf"
    Mim.Teken.ShowOpen
    frmOntwerpSQLVenster.WebBrowser1.Navigate Mim.Teken.filename
    
End Sub

Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub

Private Sub cmdSpelling_Click()
   
    Err = 0
    On Error Resume Next
    frmOntwerpSQLVenster.WebBrowser1.ExecWB OLECMDID_SPELL, OLECMDEXECOPT_DODEFAULT
    
    If Err Then MsgBox Error, vbInformation
    
End Sub


