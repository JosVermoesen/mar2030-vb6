Attribute VB_Name = "modShared"
Option Explicit

Public Enum FormAction
   faAdd = 0
   faEdit = 1
   faDelete = 2
End Enum
Public faAction As FormAction

Private mUserID As String

Public Property Get UserID() As String
    UserID = mUserID
End Property

Public Property Let UserID(ByVal vNewValue As String)
    mUserID = vNewValue
End Property

Public Sub ReportError(Number As Long, Source As String, Description As String)
    
    Dim Msg As String
    Dim dKtrl As Double
    Dim path As String * 260
    Dim deFolder As String
    Dim deSubFolder As String
    Dim Ret As Integer

    '
    If Number = -2147467259 Then
        Msg = "Eerst " & frmUsers.adcUsers.ConnectionString
        Msg = Msg & " (ODBC-verbinding) definiëren a.u.b." & vbCrLf & vbCrLf
        Msg = Msg & "Neem contact op met uw systeembeheerder of mail binnen uw contract naar info@rv.be (indien programma- of servicecontract, bel 053/21.59.25" & vbCrLf & vbCrLf
        Msg = Msg & "Wenst U hierna proberen een DOS venster te openen."
        
        If MsgBox(Msg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            On Local Error Resume Next
            path = ""
            Ret = SHGetFolderPath(0, rvProgramFiles, 0, 0, path)
            deFolder = Left(path, InStr(path, Chr(0)) - 1)
            If InStr(deFolder, "(x86)") Then
                deSubFolder = "\SysWOW64\"
            Else
                deSubFolder = "\System32\"
            End If
        
            path = ""
            Ret = SHGetFolderPath(0, rvWindows, 0, 0, path)
            deFolder = Left(path, InStr(path, Chr(0)) - 1)
        
            Msg = "marIntegraal gaat nu het DOS venster openen" & vbCrLf & vbCrLf
            Msg = Msg & "in: " & deFolder & deSubFolder & "cmd.exe (of command.com)" & vbCrLf
            Msg = Msg & "Mocht dit niet lukken, start manueel zelf dos venster, zie uw windows handleiding" & vbCrLf & vbCrLf
            Msg = Msg & "Zie handleiding voor het instellen van ODBC voor marIntegraal"
            MsgBox Msg, , "DOS VENSTER"
            
            ChDir (deFolder & deSubFolder)
            dKtrl = Shell(deFolder & deSubFolder & "cmd.exe", vbNormalFocus)
            If dKtrl Then
                dKtrl = Shell(deFolder & deSubFolder & "command.com", vbNormalFocus)
            End If
            'dKtrl = Shell(deFolder & deSubFolder & "odbcad32.exe", vbNormalFocus)
            Exit Sub
        End If
    ElseIf Err.Number = 91 Then
    Else
        'Err.Raise Number, Source, Description
        MsgBox Format(Err.Number) & vbCrLf & Err.Source & vbCrLf & Err.Description
    End If
End Sub

Private Function waituntilspring()


End Function
