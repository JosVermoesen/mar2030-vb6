Attribute VB_Name = "modScripting"
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

Function scrFolderBestaat(Folder As String) As Boolean

    scrFolderBestaat = fs.FolderExists(Folder)
        
End Function


Function LineCalculating(StartWith As Variant) As Variant

Dim StartWithHere As Variant
Dim Resulting As Variant
StartWithHere = StartWith

JumpLijnRekenen:
StartWith = InputBox("Rekenen met '" & Trim(StartWithHere) & "'" & vbCrLf & "(voer in CLR de startgegevens te verwijderen)", "1-Lijn Rekenen")
If InStr(UCase(StartWith), "CLR") Then Stop: StartWithHere = "": GoTo JumpLijnRekenen
On Error Resume Next
Resulting = Mim.ScriptControl1.Eval(StartWithHere & StartWith)
If Err Then
    Select Case Err
        Case 11
            MsgBox "Deling door nul", vbExclamation
        Case 1002
            MsgBox "Schrijffout of ongeldige notering", vbExclamation

        Case Else
            MsgBox Error
    End Select
ElseIf Resulting <> "" Then
    LineCalculating = Resulting
End If

End Function

