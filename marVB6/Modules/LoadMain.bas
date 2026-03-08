Attribute VB_Name = "aLoadMain"
Option Explicit

'Public objUsers As Beveiliging10.Users
Public strConnect As String

'Public objAudit As Beveiliging10.Audit
Public strLogFile As String

Public Sub Main()

MAR_VERSION = App.major & "." & App.minor & "." & App.Revision
BeWaarTekst "marIntegraal", "Version", MAR_VERSION
appTitleAndVersion = App.Title & " v." & MAR_VERSION '& "b"
PeppolFlag = False

Dim T As Variant

DecimalKTRL = False

'Screen.MousePointer = vbHourglass
'MsgBox "controle van XP", vbApplicationModal
'Set xpW = New Word.dokument
'If xpW.Application.Version = "10.0" Then
'Else
'    MsgBox "XP functionaliteit enkel met Microsoft Office XP en Microsoft Windows XP (PRO).", vbInformation
'End If
'Set xpW = Nothing

T = App.PrevInstance
If T <> 0 Then
    Msg = App.ProductName + " draait reeds op dit systeem !  Wenst U een bijkomende instantie van marIntegraal te openen voor deze computer én gebruiker (enkel aanbevolen voor netwerktesten)" & vbCr & vbCr
    Msg = Msg & "Via CTRL+ALT+DEL kan U eventueel het venster TAAKBEHEER opstarten voor extra opties (enkel bij NT/2000/XP/Vista-versies van Windows)" & vbCr & vbCr
    If MsgBox(Msg, vbDefaultButton2 + vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
End If

BL_LOGGING = False

For COUNT_TO = 0 To 9
    Set rsMAR(COUNT_TO) = New ADODB.Recordset
Next

ChDir App.path
Set fs = New FileSystemObject

Dim Antwoord As String
Dim dPip As Double
Dim sPip As String
Dim FlFree As Integer

ProbeerNogEens:
If App.EXEName = "marntBrokersExpress" Then
    FlFree = FreeFile
    Open App.path + "\us103.lic" For Output As FlFree
    Print #FlFree, Now
    Close FlFree
ElseIf Dir(App.path + "\us103.lic") = "" Then
    Msg = "Eerste opstart 10.3.xx licentie op dit systeem. "
    Msg = Msg & "Raadpleeg binnen het contract 2020 uw MAIL of FAX voor de juiste "
    Msg = Msg & "opstartcode.  Code kwijt?  Mail ons: info@rv.be of fax +32.53781922 voor nieuwe code" & vbCrLf & vbCrLf
    Msg = Msg & "STAP 1: Naam van deze computer :"
    Antwoord = InputBox(Msg, "Naam deze Computer", "Computer1")
    If Antwoord = "" Then GoTo ProbeerNogEens
    
    Msg = "STAP 2: Opstartcode." & vbCrLf
    Msg = Msg & "Code kwijt?  Mail ons: info@rv.be of fax +32.53781922 voor nieuwe code" & vbCrLf & vbCrLf & "Om uit te proberen (max. 50 journaalLIJNEN) en voor beperkte functionaliteit geef het volgende woord in: DEMO"
    Antwoord = InputBox(Msg, "Opstartkode? ")
    If Antwoord = "" Then
        MsgBox "Het programma wordt beëindigd.", vbInformation
        End
    End If
    If UCase(Antwoord) = "DEMO" Then
        usrLicentieInfo = "DemoModus"
    Else
        dPip = Val(Left(Antwoord, 10))
        sPip = Format(dPip - Int(dPip / 97) * 97, "00")
        If sPip <> Right(Antwoord, 2) Then
            MsgBox "Ongeldige toegangskode."
            End
        ElseIf Mid(Antwoord, 9, 2) <> "10" Then
            MsgBox "Ongeldige toegangskode."
            End
        Else
            FlFree = FreeFile
            Open App.path + "\us103.lic" For Output As FlFree
            Print #FlFree, Now
            Close FlFree
        End If
    End If
End If

'Set objUsers = CreateObject("Beveiliging10.Users")
'With objUsers
'    .AppTitle = App.Title
'    .ConnectString = "DSN=MARDSN10"
'End With

strLogFile = App.path & "\" & App.EXEName & ".log"
'Set objAudit = CreateObject("Beveiliging10.Audit")
'With objAudit
'    .FileHeader = App.EXEName
'    .filename = strLogFile
'End With

'If objUsers.Login = True Then
'    objAudit.WriteLog "Login", "Met succes"
    Mim.Show
'Else
'    objAudit.WriteLog "Login", "Zonder succes"
'    Exit Sub
'    End
'End If

End Sub



