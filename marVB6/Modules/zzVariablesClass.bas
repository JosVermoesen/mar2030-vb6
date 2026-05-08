Attribute VB_Name = "thisTextToolsClass"
DefInt A-Z
Option Explicit

Function vSet(fTekst As String, fLengte As Integer) As String
Dim b As String

b = Left(fTekst, fLengte)
vSet = b + Space$(fLengte - Len(b))

End Function


Sub Cijfermaskers()

If bhEuro Then
    MASK_2002 = MASK_EUR
Else
    MASK_2002 = MASK_BEF
End If

MASK_SY(0) = "#########"
MASK_SY(1) = "###0"
MASK_SY(2) = "######0.00"
MASK_SY(3) = "##0.00000000"
MASK_SY(4) = "#######0.00"
MASK_SY(5) = "##0"
MASK_SY(6) = "#0"
MASK_SY(7) = "#####0.0"
MASK_SY(8) = "#######0"

End Sub

Public Function IsValidEmail(ByVal sEmail As String) As Boolean
    Dim oReg As Object
    Set oReg = CreateObject("VBScript.RegExp")

    With oReg
        .Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With

    IsValidEmail = oReg.Test(Trim$(sEmail))
End Function


Function Dec(fGetal As Double, fMasker As String) As String
    
    Dim MaskerLengte As Integer
    Dim TempoString As String

    MaskerLengte = Len(fMasker)
    TempoString = Format(fGetal, fMasker)
    If (MaskerLengte - Len(TempoString)) > 0 Then
        TempoString = Space$(MaskerLengte - Len(TempoString)) + TempoString
    End If
    If InStr(TempoString, ",") Then
        Mid(TempoString, InStr(TempoString, ","), 1) = "."
    End If
    Dec = TempoString

End Function


Function ScrMaakTekstBestand(TekstZelf As String, Bestandsnaam As String) As Boolean

    On Error Resume Next
    Kill Bestandsnaam
    
    Screen.MousePointer = vbHourglass
    Call MarWriteUtf8File(Bestandsnaam, TekstZelf)
    Screen.MousePointer = vbNormal
    ScrMaakTekstBestand = True
    
End Function

Function ScrLeesTekstBestand(TekstZelf As String, Bestandsnaam As String) As Boolean
    
    On Error Resume Next
    TekstZelf = MarReadUtf8File(Bestandsnaam)
    ScrLeesTekstBestand = True
        
End Function

Function ScrLeesBestandAlleTekst(TekstZelf As String, Bestandsnaam As String) As Boolean
    
    On Error Resume Next
    
    TekstZelf = MarReadUtf8File(Bestandsnaam)
    ScrLeesBestandAlleTekst = True
        
End Function


