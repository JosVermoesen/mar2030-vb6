Attribute VB_Name = "thisLoadMainClass"
Option Explicit

Public strConnect As String
Public strLogFile As String

Public Sub Main()

    MAR_VERSION = App.major & "." & App.minor & "." & App.Revision & "b"
    BeWaarTekst "marIntegraal", "Version", MAR_VERSION
    appTitleAndVersion = App.Title & " v." & MAR_VERSION '& "b"
    PeppolFlag = False

    Dim T As Variant

    DecimalKTRL = False

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
    Mim.Show

End Sub



