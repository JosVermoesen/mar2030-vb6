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
Option Explicit


Private Sub Form_Load()

    OpenSQL
    
End Sub

Sub OpenSQL()

Dim daoRuimte       As DAO.Workspace
Dim daoDBase        As DAO.Database
Dim daorsKlanten    As DAO.Recordset

Dim dbNaam          As String


Set daoRuimte = CreateWorkspace("", "admin", "", dbUseJet)
Err = 0
' voor SQL SERVER
Set daoDBase = OpenDatabase("", dbDriverPrompt)
If Err Then
    Stop
    Exit Sub
ElseIf InStr(daoDBase.Connect, "DATABASE=") Then
    dbNaam = Mid(daoDBase.Connect, InStr(daoDBase.Connect, "DATABASE=") + 9)
    dbNaam = Left(dbNaam, InStr(dbNaam, ";") - 1)
Else
    MsgBox "Er is geen database opgegeven", vbCritical
    Exit Sub
End If

Set daorsKlanten = daoDBase.OpenRecordset("SELECT * FROM Klanten;", dbOpenDynamic)
Stop

End Sub

