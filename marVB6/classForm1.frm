VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim Stock As cmGemiddelde
Set Stock = New cmGemiddelde

Stock.aantalElementen = 2
Stock.AddItem 2366
Stock.AddItem 7889
If Stock.Volledig Then
    MsgBox Stock.GeMiddelde
End If
Set Stock = Nothing

End Sub
