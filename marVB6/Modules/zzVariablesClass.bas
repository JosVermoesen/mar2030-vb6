Attribute VB_Name = "thisVariablesClass"
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

