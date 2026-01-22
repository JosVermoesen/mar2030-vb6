VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmSimpleScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "marIntegraal NT Script Demo"
   ClientHeight    =   3495
   ClientLeft      =   870
   ClientTop       =   1455
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSchoon 
      Caption         =   "&Schoon"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDemo 
      Caption         =   "DemoKode"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin MSScriptControlCtl.ScriptControl scDemo 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ListBox lstProcedures 
      Height          =   2985
      Left            =   5880
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtScript 
      Height          =   3015
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdToevoegen 
      Caption         =   "&Toevoegen aan Procedures"
      Height          =   795
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblInstructies 
      Caption         =   "Instructies"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblProcedures 
      Caption         =   "Procedures"
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSimpleScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private MijnObject As New MijnClass

Private Sub cmdDemo_Click()

   frmSimpleScript.txtScript.Text = _
   "Sub ZetVoet()" & vbCrLf & _
    "Dim Voet" & vbCrLf & _
    "Voet = InputBox(""Taksvoet:"",,.21)" & _
    vbCrLf & _
    "objTaksScript.Voet = Voet" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Sub TotalePrijs()" & vbCrLf & _
    "  Dim Prijs, Totaal" & vbCrLf & _
    "  Prijs = InputBox(""Prijs:"",,100)" & _
    vbCrLf & _
    "  Totaal =objTaksScript.MetTaks(Prijs)" & vbCrLf & _
    "  MsgBox Totaal" & vbCrLf & _
    "End Sub"
    
    cmdToevoegen_Click
    
End Sub

Private Sub cmdSchoon_Click()
    
    ' MijnObject toevoegen aan de Scriptcontrol.
    scDemo.Reset
    'scDemo.AddObject "objTaksScript", MijnObject
        
    txtScript = ""
    lstProcedures.Clear
    
End Sub

Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub


Private Sub cmdToevoegen_Click()
    On Error Resume Next
    Err = 0
    
    ' Add the code in the TextBox to the control.
    scDemo.AddCode txtScript.Text
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Exit Sub
    End If

    ' Clear the ListBox, and add the name of each
    ' procedure in the Procedures collection.
    lstProcedures.Clear
    Dim p As Procedure
    For Each p In scDemo.Procedures
        lstProcedures.AddItem p.Name
    Next
    lstProcedures.SetFocus
    
End Sub

Private Sub Form_Load()

    MsgBox Now & ": nog in ontwikkeling"
    cmdSchoon_Click
    
End Sub


Private Sub lstProcedures_Click()
    On Error Resume Next
    Err = 0
    
    ' Run the procedure in the ListBox.
    scDemo.Run lstProcedures.Text
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Exit Sub
    End If
    
End Sub

