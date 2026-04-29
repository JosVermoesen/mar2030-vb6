Attribute VB_Name = "thisTextToolsClass"
DefInt A-Z
Option Explicit

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


