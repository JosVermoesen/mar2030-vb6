Attribute VB_Name = "Module3"
DefInt A-Z
Option Explicit

Dim DataPos                As Integer
Dim DataLen                As Integer
Dim OptieTxt               As String
Dim TempoTxt               As String
Dim CrText                 As String
Dim CrText2                As String
Dim T                       As Integer


Function Toegankelijk(frmHier As Form) As Boolean

'If objUsers.CheckRights(Left(frmHier.Name, 20), 1) Then
    Toegankelijk = True
'    objAudit.WriteLog frmHier.Name, "Toegang met Succes"
'Else
'    Toegankelijk = False
'    MsgBox "Deze taak is (nog) niet toegekend aan huidige gebruiker. Neem contact op met systeembeheerder.", vbExclamation, frmHier.Name
'    objAudit.WriteLog "LogIn", frmHier.Name + ":Toegang voorbehouden"
'End If

End Function

Sub GridSchoon(GelijkWelkRooster As MSFlexGrid)
Dim T As Integer

GelijkWelkRooster.Rows = 2
GelijkWelkRooster.Row = 1
For T = 0 To GelijkWelkRooster.Cols - 1
    GelijkWelkRooster.Col = T
    GelijkWelkRooster.text = ""
Next

End Sub

