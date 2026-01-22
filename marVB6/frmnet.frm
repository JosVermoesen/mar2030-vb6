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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function NetGetDCName Lib "netapi32.dll" (strServerName As Any, _
    strDomainName As Any, pBuffer As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (buffer As Any) As _
    Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)

' Return the name of the Primary Domain Controller (PDC)
'
' server names are preceded by a double slash

Function GetPDCName() As String
    Dim lpBuffer As Long
    Dim buffer As String
    
    If NetGetDCName(vbNullString, vbNullString, lpBuffer) = 0 Then
        ' a zero return value means success
        buffer = Space$(128)
        ' copy into local buffer
        CopyMemory ByVal StrPtr(buffer), ByVal lpBuffer, LenB(buffer)
        ' trim excess chars
        GetPDCName = Left$(buffer, InStr(buffer & vbNullChar, vbNullChar) - 1)
        ' release memory
        NetApiBufferFree lpBuffer
    End If

End Function


Private Sub Command1_Click()

    GetPDCName
    
End Sub


