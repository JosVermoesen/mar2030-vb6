Attribute VB_Name = "modMakelaars"

DefInt A-Z

Dim TeleBibRC As String


Sub MijStatistiek()
Dim FlNetTemp As Integer

'On Error GoTo SchrijfWegError
Dim A As String * 140
Dim aa As String * 420

Msg$ = "Beschikbare toepassingen maatschappijen opvragen ?" + vbCrLf + vbCrLf + "(enkel nuttig bij eerste gebruik Assurnet !)"
Ktrl = MsgBox(Msg$, 292, "Verzoek 912")
If Ktrl = 6 Then
Else
    Exit Sub
End If

A = ""
aa = ""
Mid(A, 1, 3) = "AS1"                   'constante
Mid(A, 16, 1) = "0"                    'constante
Mid(A, 17, 3) = "912"                  'verzoek
Mid(A, 20, 1) = "1"                    'gestruktureerd
Mid(A, 21, 1) = "1"                    'producent
Mid(A, 22, 8) = vSet(ProducentNummer, 8)    'ASNET prod.nr.
Mid(A, 30, 12) = Format(Now, "yymmddhhmmss") 'datum
Mid(A, 42, 1) = "1"                    'gewoon niet vertrouwelijk
Mid(A, 43, 1) = "4"                    'bestemmeling=Assurnet"
Mid(A, 44, 4) = "XXXX" 'BVVO kode
Mid(A, 48, 2) = "01"                   'volgnummer
Mid(A, 50, 1) = "X"                    'einde tekst

Mid(aa, 1, 8) = vSet(ProducentNummer, 8)    'ASNET prod.nr.
Mid(aa, 9, 1) = "1"                        'alle mijen

FlNetTemp = FreeFile
Open LOCATION_ASWEB + "AS1.SND" For Append As FlNetTemp
Print #FlNetTemp, A + aa
Close FlNetTemp

End Sub

Public Function SnippetXEH(UserArea As String, PolisNummer As String, blIndent As Boolean) As String

    Dim XEHarray() As String
    Dim COUNT_TO As Integer
    Dim txtHier As String

    Dim zoekXET As String
    Dim zoekXEH As String
    
    If InStr(UserArea, "XET+03") Then
        zoekXEH = "XEH+03" 'termijnen en contanten via borderels
        zoekXET = "XET+03"
    ElseIf InStr(UserArea, "XRT+2") Then
        zoekXET = "XRT+1" 'commissies via rekeninguittreksels
        zoekXEH = "XRH+1"
    Else
        MsgBox "geen XET te vinden!", vbCritical
        SnippetXEH = ""
    End If
        
    'termijnhernieuwing opzoeken voor dit contract in de hernieuwingen van een bepaalde maand
    XEHarray = Split(UserArea, zoekXET)
        SnippetXEH = ""
        For COUNT_TO = 0 To UBound(XEHarray)
            If InStr(XEHarray(COUNT_TO), "RFF+001:" & PolisNummer) Then
                If blIndent = False Then
                    SnippetXEH = Mid(XEHarray(COUNT_TO), InStr(XEHarray(COUNT_TO), "'" & zoekXEH)) & zoekXET & "'"
                Else
                    SnippetXEH = tb2Indent(XEHarray(COUNT_TO) & zoekXET & "'")
                End If
                Exit Function
            End If
        Next
    
End Function

Sub EdiFactTERMIJN(UserArea As String)

Dim Infolijn As String, Bestand As String, VBestand As String, iolijn As String, PolisNummer As String * 12
Dim FlBestand As Integer, Operatie As Integer
Dim HuidigeBM As String * 2
Dim Dummy As String * 30
Dim fKtrl As Integer
Dim InfoTekst As String
Dim DCBedrag As Currency
Dim MaandVerwerkingTermijn As Integer
Dim BoekjaarKontroleTermijn As String
Dim BorderelTxt As String
Dim Datum2000 As String

Dim Bedrag0 As Currency
Dim Bedrag1 As Currency
Dim Bedrag2 As Currency
Dim Bedrag3 As Currency
Dim Bedrag4 As Currency
Dim Bedrag5 As Currency
Dim Bedrag6 As Currency
Dim Bedrag7 As Currency
Dim Bedrag8 As Currency
Dim Bedrag9 As Currency

Dim Bedrag10 As Currency
Dim Bedrag11 As Currency
Dim Bedrag12 As Currency

Dim TotaalNettoPremie As Currency
Dim TotaalLasten As Currency
Dim BedragPremie As String

Dim MaskerHier As String
Dim Xarray() As String
Dim ATTarray() As String
Dim XEHarray() As String
Dim XGHarray() As String
Dim MOAarray() As String
Dim DTMarray() As String
Dim PTYarray() As String
Dim ftxArray() As String
Dim FTXTXTarray() As String
Dim POLArray() As String
Dim XRHniveau As Integer

Dim Telhier As Integer
Dim LijnTeller As Integer
Dim Maatschappij As String
Dim MaandVerwerking As String * 2
Dim BoekjaarKontrole As String * 4
Dim TempoNaamKlant As String
Dim HuidigeIndex As String
Dim DatumKwijting As String
Dim DagKwijting As String * 2
Dim MaandKwijting As String * 2

Dim PolisTeWijzigen As Boolean
Dim DagKtrl1 As String * 2
Dim DagKtrl2 As String * 2
Dim mijWijzigen As Boolean

Dim contanteFlag As Boolean
Dim inningmijFlag As Boolean
Dim typekwijtingFlag As Integer

mijWijzigen = False
Bedrag0 = 0
Bedrag1 = 0
Bedrag2 = 0
Bedrag3 = 0
Bedrag4 = 0
Bedrag5 = 0
Bedrag7 = 0
Bedrag8 = 0
Bedrag9 = 0

Bedrag10 = 0
Bedrag11 = 0
Bedrag12 = 0

HuidigeIndex = ""
HuidigeBM = ""
Xarray = Split(UserArea, "'")
XGHarray = Split(Xarray(0), "+")
Maatschappij = Format(Val(XGHarray(2)), "0000")

MaskerHier = "######0.00"
BorderelTxt = ""
GridText = ""
MaandVerwerkingTermijn = 0
XRHniveau = 0
LijnTeller = 0
Dim ftxTeller As Integer
Dim txtMSG As String
    txtMSG = ""
Dim txtQTY As String
    txtQTY = ""
    
Dim isTotaalMOA As Boolean

Dim check06 As String

contanteFlag = False
For Telhier = 0 To UBound(Xarray) - 1
    BL_LOGGING = True
    'Bedrag1 = 0
    'Bedrag2 = 0
    'Bedrag3 = 0
    'Bedrag4 = 0
    'Bedrag5 = 0
    'Bedrag7 = 0
    'Bedrag8 = 0
    'Bedrag9 = 0

    Select Case Left(Xarray(Telhier), 3)
        Case "XEH" 'Volgend dokumentlijntje (codering nog controleren!)
            'vb. : XEH+01 = CONTRACT
            'vb. : XEH+03 = KWIJTING
            
            isTotaalMOA = True 'nieuwe
            Select Case Left(Xarray(Telhier), 6)
                Case "XEH+01"
                    'nieuw contract(wijziging) mogelijk met kwijting
                    contanteFlag = True
                    inningmijFlag = False
                    check06 = "01"
                    
                    
                Case "XEH+03"
                    LijnTeller = LijnTeller + 1
                    inningmijFlag = False
                    check06 = "03"
                    
                Case "XEH+06"
                    'Stop
                    'Case "XEH+06" is ?
                    check06 = Left(Xarray(Telhier), 6)
                    
                    
                Case Else
                    '?
                    Stop
            
            End Select
            
            
            
        Case "XRH"
            'MsgBox "hier controleren of het geen bedragen voor indivi. wb zijn"
            isTotaalMOA = False
        
        Case "DTM"
            DTMarray = Split(Xarray(Telhier), ":")

            If Left$(Xarray(Telhier), 8) = "DTM+005:" Then
                If DTMarray(UBound(DTMarray)) = "010" Then
                    
                    If Len(DTMarray(1)) = 8 Then
                        MaandVerwerking = Mid(DTMarray(1), 5, 2)
                        BoekjaarKontrole = Mid(DTMarray(1), 1, 4)
                    Else
                        MsgBox "stoppen en controleren a.u.b"
                    End If
                
                Else
                    If Len(DTMarray(1)) = 8 Then
                        MaandVerwerking = Mid(DTMarray(1), 3, 2)
                        BoekjaarKontrole = Mid(DTMarray(1), 5, 4)
                    ElseIf DTMarray(UBound(DTMarray)) = "001" Then
                        MaandVerwerking = Mid$(Xarray(Telhier), 11, 2) 'boekMAAND
                        BoekjaarKontrole = Mid$(Xarray(Telhier), 13, 4) 'boekJAAR
                    ElseIf DTMarray(UBound(DTMarray)) = "005" Then
                        MaandVerwerking = Mid$(Xarray(Telhier), 9, 2)    'boekMAAND
                        BoekjaarKontrole = Mid$(Xarray(Telhier), 11, 4)  'boekJAAR
                    Else
                        MsgBox "stoppen en controleren a.u.b"
                    End If
                End If
            Else
                If Left$(Xarray(Telhier), 8) = "DTM+004:" Then
                    DagKwijting = Mid$(DTMarray(1), 1, 2)
                    MaandKwijting = Mid$(DTMarray(1), 3, 2)
                    DatumKwijting = DagKwijting & "/" & MaandKwijting & "/" & BoekjaarKontrole 'Mid$(DTMarray(1), 5, 4)
                ElseIf Left$(Xarray(Telhier), 8) = "DTM+041:" Then
                    If DTMarray(UBound(DTMarray)) = "010" Then
                        DagKwijting = Mid$(DTMarray(1), 7, 2)
                        MaandKwijting = Mid$(DTMarray(1), 5, 2)
                        DatumKwijting = DagKwijting & "/" & MaandKwijting & "/" & BoekjaarKontrole 'Mid$(DTMarray(1), 5, 4)
                    Else
                
                        DagKwijting = Mid$(DTMarray(1), 1, 2)
                        MaandKwijting = Mid$(DTMarray(1), 3, 2)
                        DatumKwijting = DagKwijting & "/" & MaandKwijting & "/" & BoekjaarKontrole 'Mid$(DTMarray(1), 5, 4)
                    End If
                End If
            End If
        'Stop
                       
            
        Case "RFF"
            'niks vroeger polisnummer ophalen
            If Left$(Xarray(Telhier), 8) = "RFF+001:" Then
                POLArray = Split(Xarray(Telhier), ":")
                PolisNummer = POLArray(1)
            End If
            
            
        Case "IND"  'huidige index
            If Mid$(Xarray(Telhier), 1, 10) = "IND+002+1:" Then
                HuidigeIndex = Format(Val(Mid$(Xarray(Telhier), 11, 5)) / 100, "000.00")
            Else
                HuidigeIndex = Format(Val(Mid$(Xarray(Telhier), 10, 5)) / 100, "000.00")
                
            End If
            'MsgBox "index= " & HuidigeIndex, vbInformation
                                
        Case "MOA"
            If isTotaalMOA = True Then
                MOAarray = Split(Xarray(Telhier), ":")
                'WAT MET AUTO ?
                'Bedrag1 = Val(kt(Mid$(iolijn, 42, 6))) 'premie B.A. auto
                'Bedrag2 = Val(kt(Mid$(iolijn, 48, 5))) 'taksen B.A. auto
                'Bedrag3 = Val(kt(Mid$(iolijn, 53, 4))) 'kosten B.A. auto
            
                'Bedrag4 = Val(kt(Mid$(iolijn, 57, 8))) 'andere premie
                'Bedrag5 = Val(kt(Mid$(iolijn, 65, 7))) 'taksen+kosten andere
                'Bedrag9 = Val(kt(Mid$(iolijn, 72, 7))) 'commissie
           
                'Bedrag7 = Bedrag1 + Bedrag4
                'Bedrag8 = Bedrag2 + Bedrag3 + Bedrag5
               
                'WAT MET ANDERE?
                'Bedrag7 = Val(kt(Mid$(iolijn, 57, 8))) 'premie
                'Bedrag8 = Val(kt(Mid$(iolijn, 65, 7))) 'taksen&kosten
                'Bedrag9 = Val(kt(Mid$(iolijn, 72, 7))) 'commissie
            
                Select Case MOAarray(0)
                    Case "MOA+039"  'Splitsingskosten 2018
                        If MOAarray(2) = "EUR" Then
                            If Left(MOAarray(3), 1) = 2 Then
                                Bedrag10 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag10 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag10 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag10 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                
                    Case "MOA+210"  'schatting aanwervingskosten 2018
                        If MOAarray(2) = "EUR" Then
                            If Left(MOAarray(3), 1) = 2 Then
                                Bedrag11 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag11 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag11 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag11 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                
                    Case "MOA+211"  'Schatting administratiekosten 2018
                        If MOAarray(2) = "EUR" Then
                            If Left(MOAarray(3), 1) = 2 Then
                                Bedrag12 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag12 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag12 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag12 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    
            
                    Case "MOA+012"  'Totale premie
                        If MOAarray(2) = "EUR" Then
                            If Left(MOAarray(3), 1) = 2 Then
                                Bedrag0 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag0 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag0 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag0 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    
                    Case "MOA+013"  'netto gefractioneerde premie
                        
                        If MOAarray(2) = "EUR" Then
                            
                            If Left(MOAarray(3), 1) = 2 Then
                                Bedrag1 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag1 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag1 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag1 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                        TotaalNettoPremie = Bedrag1
                    
                    Case "MOA+015"  'commissie
                        If XRHniveau > 0 Then
                            SnelHelpPrint "XRHniveau is : " & XRHniveau, BL_LOGGING
                        Else
                            If MOAarray(2) = "EUR" Then
                                If Left(MOAarray(3), 1) = 2 Then
                                    Bedrag9 = Val(MOAarray(1)) / 100
                                ElseIf MOAarray(3) = 4 Then
                                    Bedrag9 = Val(MOAarray(1)) / 10000
                                ElseIf MOAarray(3) = 0 Then
                                    MsgBox "stop"
                                    Bedrag9 = Val(MOAarray(1))
                                Else
                                    MsgBox "Stop"
                                End If
                            ElseIf MOAarray(2) = "BEF" Then
                                Bedrag9 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        End If
                                    
                       
                    Case "MOA+016"  'Lastenbedrag BA auto
                        If MOAarray(2) = "EUR" Then
                            If MOAarray(3) = 2 Then
                                Bedrag2 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag2 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag2 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag2 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                        TotaalLasten = Bedrag2 + Bedrag5
                                        
                    Case "MOA+017"  'Splitsingskosten? + enkel BA auto ?
                        'MsgBox "nog manueel te controleren, voor AG reeds in orde", vbExclamation
                        If MOAarray(2) = "EUR" Then
                            If MOAarray(3) = 2 Then
                                Bedrag2 = Bedrag2 + Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag2 = Bedrag2 + Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag2 = Bedrag2 + Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag2 = Bedrag2 + Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                        TotaalLasten = Bedrag2 + Bedrag5
                                        
                    'ANDERE DAN AUTO !
                    Case "MOA+097"  'premie andere waarborgen dan BA auto
                        If MOAarray(2) = "EUR" Then
                            If MOAarray(3) = 2 Then
                                Bedrag4 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag4 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag4 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag4 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    
                    Case "MOA+098"  'taksen en kosten andere dan BA auto
                        If MOAarray(2) = "EUR" Then
                            If MOAarray(3) = 2 Then
                                Bedrag5 = Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag5 = Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag5 = Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag5 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                        TotaalLasten = Bedrag2 + Bedrag5
                        
                    Case "MOA+100"  'Kwitantiekosten
                        'MsgBox "nog manueel te controleren, voor MAURETUS reeds in orde", vbExclamation
                        If MOAarray(2) = "EUR" Then
                            If MOAarray(3) = 2 Then
                                Bedrag2 = Bedrag2 + Val(MOAarray(1)) / 100
                            ElseIf MOAarray(3) = 4 Then
                                Bedrag2 = Bedrag2 + Val(MOAarray(1)) / 10000
                            ElseIf MOAarray(3) = 0 Then
                                MsgBox "stop"
                                Bedrag2 = Bedrag2 + Val(MOAarray(1))
                            Else
                                MsgBox "Stop"
                            End If
                        ElseIf MOAarray(2) = "BEF" Then
                            Bedrag2 = Bedrag2 + Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                        TotaalLasten = Bedrag2 + Bedrag5
                        
                    Case "MOA+105"  'Andere aan netto premie toe te voegen kosten
                            
                    
                    Case Else
                        SnelHelpPrint "MOA-stop voor " & MOAarray(0), BL_LOGGING
                End Select
            End If
    
            
        Case "NME"
            TempoNaamKlant = Mid(Xarray(Telhier), 9)
        
        Case "ATT"
            ATTarray = Split(Xarray(Telhier), "+")
            Select Case ATTarray(1)
                Case "5300" 'bonus malus
                    HuidigeBM = ATTarray(2)
                    
                Case "B001" 'type kwijting
                    
                    
                    typekwijtingFlag = Val(Right(Xarray(Telhier), 1))
                    
                            
                Case "B003", "A600"
                    'ATT+B003+3  =  inning maatschappij vanaf 2008
                    'ATT+A600+3  =  inning maatschappij tot 2008
                    If ATTarray(2) = "3" Then  'Inning Maatschappij
                        inningmijFlag = True
                                                
                        Bedrag0 = 0
                        Bedrag1 = 0
                        Bedrag2 = 0
                        Bedrag3 = 0
                        Bedrag4 = 0
                        Bedrag5 = 0
                        Bedrag7 = 0
                        Bedrag8 = 0
                        Bedrag9 = 0
                        Bedrag10 = 0
                        Bedrag11 = 0
                        Bedrag12 = 0
                        isTotaalMOA = True
                        
                        HuidigeIndex = ""
                        HuidigeBM = ""
                            
                        If contanteFlag = False Then
                            MsgBox "Kwitantie voor polis " & Trim$(PolisNummer) & " staat inning Maatschappij"
                            
                            
                            Do
                                If Xarray(Telhier) = "XET+03" Then
                                    Exit Do
                                ElseIf Xarray(Telhier) = "XET+06" Then
                                    Exit Do
                                Else
                                    'Debug.Print Xarray(Telhier)
                                    Ktrl = DoEvents()
                                    Telhier = Telhier + 1
                                    If Telhier > 250 Then
                                    'Stop
                                        Exit Do
                                    End If
                                End If
                            Loop
                        End If
                    End If
            End Select
            
            
        Case "XET" 'Afsluiten/lijn bewaren
        
            If Left(Xarray(Telhier), 6) = "XET+03" Then

                Bedrag7 = Bedrag1 + Bedrag4
                Bedrag8 = Bedrag2 + Bedrag3 + Bedrag5
                                
                'GoSub Wegschrijven
                On Error Resume Next
                If bhEuro = True And MOAarray(2) = "EUR" Then
                'boekhouding en bericht in EURO, niets te doen
                ElseIf bhEuro = False And MOAarray(2) = "BEF" Then
                'boekhouding en bericht in BEF, niets te doen
                ElseIf MOAarray(2) = "BEF" Then
                'boekhouding in EURO en bericht in BEF !
                    Bedrag7 = Round(Bedrag7 / EURO, 2)
                    Bedrag8 = Round(Bedrag8 / EURO, 2)
                    Bedrag9 = Round(Bedrag9 / EURO, 2)
                    Bedrag0 = Round(Bedrag0 / EURO, 2)
    
                ElseIf MOAarray(2) = "EUR" Then
                'boekhouding in EURO en bericht in BEF !
                    Bedrag7 = Round(Bedrag7 * EURO)
                    Bedrag8 = Round(Bedrag8 * EURO)
                    Bedrag9 = Round(Bedrag9 * EURO)
                    Bedrag0 = Round(Bedrag0 * EURO)
                Else
                    MsgBox "ONMOGELIJKE SITUATIE"
                End If

                If Bedrag0 = 0 Then
                    BedragPremie = Dec$((Bedrag7 + Bedrag8), MaskerHier)
                Else
                    BedragPremie = Dec$((Bedrag0), MaskerHier)
                    'MsgBox "Totaal premie werd meegegeven: " & BedragPremie
                End If
                bGet TABLE_CONTRACTS, 0, PolisNummer
                If Ktrl Then
                    MsgBox "polis niet aanwezig. EDIFACT nieuw nog in te brengen !!!", vbCritical
                    TLB_RECORD(TABLE_CONTRACTS) = ""
                    MsgBox "Stop.  Polis bestaat nog niet :" + PolisNummer
                    vBib TABLE_CONTRACTS, MaandKwijting, "v164"
                    vBib TABLE_CONTRACTS, "NONAME", "A110"
                    vBib TABLE_CONTRACTS, Maatschappij, "A010"
                    vBib TABLE_CONTRACTS, PolisNummer, "A000"
                    vBib TABLE_CONTRACTS, TempoNaamKlant, "vs99"
                    bInsert TABLE_CONTRACTS, 0
                    If Ktrl Then
                        MsgBox "onbekende soldaat ! STOP!!!"
                    End If
                    Infolijn = Infolijn + " " + TempoNaamKlant + " aanvullen a.u.b."
                    Dummy = "Kontroleer !!! " + TempoNaamKlant
                    'As2TelebibIn.UserInfo(3).Caption = Str$(Val(As2TelebibIn.UserInfo(3).Caption) + 1)
                    ddag$ = "01"
                Else
                    RecordToVeld TABLE_CONTRACTS
                    PolisTeWijzigen = False
                    DagKtrl1 = vBibTekst(TABLE_CONTRACTS, "#v165 #")
                    DagKtrl2 = Mid$(vBibTekst(TABLE_CONTRACTS, "#AW_2 #"), 7, 2)
                    If rsMAR(TABLE_CONTRACTS)("vs97") > "2" Then
                        MsgBox "Polis " & PolisNummer & vbCrLf & "Actiecode staat op: " & rsMAR(TABLE_CONTRACTS)("vs97") & vbCr & vbCr & "Wordt automatisch op 1 (Post) geplaatst", vbInformation
                        vBib TABLE_CONTRACTS, "1", "vs97"
                        PolisTeWijzigen = True
                    End If
                    If MaandKwijting = vBibTekst(TABLE_CONTRACTS, "#v164 #") Then
                    Else
                        vBib TABLE_CONTRACTS, MaandKwijting, "v164"
                        PolisTeWijzigen = True
                    End If
                    If DagKtrl1 = DagKtrl2 Then
                        'alles ok
                    Else
                        KtrlBox = MsgBox(DagKtrl2 & " correctie vervalDAG", vbYesNo + vbQuestion + vbDefaultButton1)
                        If KtrlBox = vbYes Then
                            vBib TABLE_CONTRACTS, DagKtrl2, "v165"
                            PolisTeWijzigen = True
                        End If
                    End If
                    DagKwijting = DagKtrl2
                    If Maatschappij <> vBibTekst(TABLE_CONTRACTS, "#A010 #") Then
                        'As2TelebibIn.UserInfo(1).Caption = "Maatschappij conflict!  Polis " + PolisNummer + "bestaat reeds bij maatschappij !" + vBibTekst(TABLE_CONTRACTS, "#A010 #")
                        vBib TABLE_CONTRACTS, Maatschappij, "A010"
                        PolisTeWijzigen = True
                    End If
                    If PolisTeWijzigen = True Then
                        bUpdate TABLE_CONTRACTS, 0
                    End If
                    'As2TelebibIn.UserInfo(2).Caption = Str$(Val(As2TelebibIn.UserInfo(2).Caption) + 1)
                    bGet TABLE_CUSTOMERS, 0, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12)
                    If Ktrl Then
                        Dummy = "Verbeter !!! " + TempoNaamKlant
                    Else
                        RecordToVeld TABLE_CUSTOMERS
                        Dummy = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
                    End If
                End If
        
                'GoSub TeleBibKTRL
                bGet TABLE_VARIOUS, 1, "25" + Maatschappij + PolisNummer
                If Ktrl Then
                    TLB_RECORD(TABLE_VARIOUS) = ""
                Else
                    RecordToVeld TABLE_VARIOUS
                End If
        
                Select Case Mid$(iolijn, 42, 1)
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        'AUTO ?
                        'If Maatschappij = "0145" And Mid$(iolijn, 42, 6) = "000000" Then
                        'Else
                        '    vBib TABLE_VARIOUS, "P11", "AW00"
                        '    vBib TABLE_VARIOUS, "4", "AW06"
                        '    vBib TABLE_VARIOUS, Str$(Bedrag1 + Bedrag2 + Bedrag3), "AW04"
                        '    vBib TABLE_VARIOUS, Mid$(iolijn, 40, 2), "5315"
                        'End If
                End Select
        
                If Trim$(HuidigeIndex) <> "" Then
                    vBib TABLE_VARIOUS, HuidigeIndex, "AW.R"                            'index
                End If
                If Trim$(HuidigeBM) <> "" Then
                    'Debug.Assert False
                    vBib TABLE_VARIOUS, HuidigeBM, "5315"
                    vBib TABLE_VARIOUS, HuidigeBM, "5300"
                End If
                
                vBib TABLE_VARIOUS, Str$(Bedrag10), "v400"                                 '2018 Splitsingskosten
                vBib TABLE_VARIOUS, Str$(Bedrag11), "v401"                                 '2018 Schatting aanwervingskosten
                vBib TABLE_VARIOUS, Str$(Bedrag12), "v402"                                 '2018 Schatting administratiekosten
                
                'klopt dat wel??
                'MsgBox "kontrole voor totaalmoa enkel"
                vBib TABLE_VARIOUS, Str$(Bedrag7 + Bedrag8), "B010"                      'premie taksen en kosten inbegrepen
                vBib TABLE_VARIOUS, Str$(Bedrag8), "B011"                                 'taksen en kosten
                vBib TABLE_VARIOUS, Str$(Bedrag7), "B013"                                 'premie zonder taksen en kosten
                vBib TABLE_VARIOUS, Str$(Bedrag9), "B014"                                 'commissieloon
                
                'best toevoegen
                vBib TABLE_VARIOUS, Str$(Bedrag0), "v390"
                vBib TABLE_VARIOUS, Str$(Bedrag1), "v391"
                vBib TABLE_VARIOUS, Str$(Bedrag2), "v392"
                vBib TABLE_VARIOUS, Str$(Bedrag4), "v393"
                vBib TABLE_VARIOUS, Str$(Bedrag5), "v394"
                
                
                
                vBib TABLE_VARIOUS, vSet("K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), 13), "v004"
                vBib TABLE_VARIOUS, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12), "A110"
                vBib TABLE_VARIOUS, Maatschappij, "A010"
                vBib TABLE_VARIOUS, PolisNummer, "A000"
                vBib TABLE_VARIOUS, vSet("25" + Maatschappij + PolisNummer, 20), "v005"
                         
                If Ktrl Then
                    bInsert TABLE_VARIOUS, 1
                Else
                    bUpdate TABLE_VARIOUS, 1
                End If
                If Ktrl Then
                    MsgBox "Stopkode " + Str$(Ktrl)
                End If
                 
                On Error Resume Next
                If Val(DagKwijting) = 0 Then
                    DagKwijting = "01"

                End If
                Mid$(DatumKwijting, 1, 2) = DagKwijting
                'contanteFlag As Boolean
                'inningmijFlag As Boolean

                If inningmijFlag = False Then
                    GridText = GridText & PolisNummer & vbTab
                    GridText = GridText & DatumKwijting & vbTab
                    GridText = GridText & BedragPremie & vbTab
                    GridText = GridText & Dec((Bedrag9), MaskerHier) & vbTab
                    GridText = GridText & Dummy & vbTab & SnippetXEH(UserArea, Trim(PolisNummer), False) & vbTab & Str(TotaalNettoPremie) & vbTab & Str(TotaalLasten) & vbCr
                End If
                
                Bedrag0 = 0
                Bedrag1 = 0
                Bedrag2 = 0
                Bedrag3 = 0
                Bedrag4 = 0
                Bedrag5 = 0
                Bedrag7 = 0
                Bedrag8 = 0
                Bedrag9 = 0
                HuidigeIndex = ""
                HuidigeBM = ""
            Else
                'Stop
            End If
                       
            
        Case "XRH"
            XRHniveau = XRHniveau + 1
            
        Case "XRT"
            XRHniveau = XRHniveau - 1
            
        Case "PTY"
            'Maatschappij = Format(Val(XGHarray(2)), "0000")
            PTYarray = Split(Xarray(Telhier), "+")
            If PTYarray(1) = "006" Then
                If LijnTeller = 1 Then
                    If Format(Val(PTYarray(2)), "0000") <> Maatschappij Then
                        If MsgBox("Maatschappijcode wijzigen" & vbCr & vbCr & Format(Val(PTYarray(2)), "0000") & " <> " & Maatschappij, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                            Maatschappij = Format(Val(PTYarray(2)), "0000")
                            mijWijzigen = True
                        Else
                            mijWijzigen = False
                        End If
                    End If
                ElseIf mijWijzigen = True Then
                    If Format(Val(PTYarray(2)), "0000") <> Maatschappij Then
                        Maatschappij = Format(Val(PTYarray(2)), "0000")
                    End If
                End If
            End If
            
        Case "IPD"
            'voorlopig over te slaan...
            
        Case "FTX"
            'ftxArray = Split(Xarray(Telhier), "+")
            'If ftxArray(1) = "018" Then 'infobericht?
            '    FTXTXTarray = Split(ftxArray(2), ":")
            '    txtMSG = ""
            '    For ftxTeller = 0 To UBound(FTXTXTarray) - 1
            '        txtMSG = txtMSG & FTXTXTarray(ftxTeller) & vbCrLf
            '    Next
            '    'MsgBox txtMSG, , "FTX 018"
            'ElseIf ftxArray(1) = "019" Then
            '    txtMSG = ftxArray(2)
            '    MsgBox txtMSG, , "FTX 019"
            'Else
            '    MsgBox "FTX type anders dan 019: " & ftxArray(1)
            'End If
            
        Case "QTY"
            'txtQTY = Xarray(Telhier)
            'MsgBox txtQTY
    End Select
    'Stop
    
Next

Dim PeriodeVoor As String
Dim KtrlFlag As Integer

If GridText = "" Then
Else
    BJPERDAT!Boekjaar.ListIndex = 0
    PeriodeVoor = BoekjaarKontrole + MaandVerwerking
    If PeriodeVoor < Left$(BOOKYEAR_FROMTO, 6) Then
        If BJPERDAT!Boekjaar.ListCount > 1 Then
            BJPERDAT!Boekjaar.ListIndex = 1
        End If
    End If
    KtrlFlag = False
    For Tel = 0 To BJPERDAT!PeriodeBoekjaar.ListCount - 1
        BJPERDAT!PeriodeBoekjaar.ListIndex = Tel
        If PeriodeVoor >= Left$(PERIOD_FROMTO, 6) And PeriodeVoor <= Mid$(PERIOD_FROMTO, 9, 6) Then
            'bClose 99
            Unload KwijtingBoeken

            Select Case typekwijtingFlag
                Case 1
                    GridTextIs = "001" & vbTab & Maatschappij & vbTab & "Termijn"
                    
                Case 2
                    GridTextIs = "002" & vbTab & Maatschappij & vbTab & "Contante"
                    
                Case 3
                    GridTextIs = "003" & vbTab & Maatschappij & vbTab & "Teruggave"
                    
                Case Else
                    MsgBox "stop voor controle type kwijting  EUROMEX?"
                    GridTextIs = "001" & vbTab & Maatschappij & vbTab & "Termijn"
            End Select
            BL_LOGGING = False
            
            KwijtingBoeken.Show 1
            GridText = ""
            KtrlFlag = True
            Exit For
        End If
    Next
    If KtrlFlag = False Then
        MsgBox "Grote gebruikersfout.  Nieuw boekjaar nog niet beschikbaar !" + vbCrLf + "Manuele inbreng is noodzakelijk..."
    End If
End If

End Sub



Sub EdiFactCONTANT(UserArea As String)

Dim Xarray() As String
Dim IPDarray() As String
Dim RFFarray() As String
Dim DTMarray() As String
Dim PTYarray() As String
Dim ATTarray() As String
Dim MOAarray() As String
Dim NMEarray() As String

Dim XEHnext As String
    XEHnext = ""
                              
'Deel PolisInfo
Dim PolisNummer As String
Dim BijvoegselNummer As String
Dim Verzekeraar As String
Dim PolisType As String
Dim HoofdVervaldag As String
Dim InningswijzeTermijn As String
Dim SplitsingsCode As String
Dim StatusPolis As String
Dim TempoNaamKlant As String
Dim Infolijn As String

'Deel BoekingsInfo
Dim BoekDatum As String
Dim InningswijzeContant As String
Dim TypeKwijting As String
Dim TotaalNettoPremie As Currency
Dim TotaalTeBetalen As Currency
Dim TotaalCommissie As Currency
Dim TotaalLasten As Currency
Dim TotaalKosten As Currency
Dim VerzekerdVanDatum As String
Dim VerzekerdTotDatum As String
Dim Dummy As String


Dim bezigaanKwijting As Boolean
    bezigaanKwijting = False

Dim BoekjaarKontrole As String
Dim MaandVerwerking As String
Dim PeriodeVoor As String
Dim MaskerHier As String
    MaskerHier = "######0.00"

Xarray = Split(UserArea, "'")

For Telhier = 0 To UBound(Xarray) - 1
    Select Case Left(Xarray(Telhier), 3)
      
        Case "ATT"
            ATTarray = Split(Xarray(Telhier), "+")
            If ATTarray(1) = "A600" Then 'Inningswijze termijn 4= producent
                InningswijzeTermijn = ATTarray(2)
            ElseIf ATTarray(1) = "A325" Then 'Splitsingscode premie 1= jaarlijks
                SplitsingsCode = ATTarray(2)
            ElseIf ATTarray(1) = "A003" Then 'Status polis 1= lopend
                StatusPolis = ATTarray(2)
            ElseIf ATTarray(1) = "A602" Then 'Inningswijze contantkwijting 4= Producent
                InningswijzeContant = ATTarray(2)
            ElseIf ATTarray(1) = "B001" Then 'Type kwijting 2= contant
                TypeKwijting = ATTarray(2)
            End If
                    
        Case "DTM"
            DTMarray = Split(Xarray(Telhier), ":")
            If Left(Xarray(Telhier), 7) = "DTM+004" Then 'HoofdVervaldag
                HoofdVervaldag = DTMarray(1)
            ElseIf Left(Xarray(Telhier), 7) = "DTM+005" Then 'Boekdatum
                BoekDatum = DTMarray(1)
            ElseIf Left(Xarray(Telhier), 7) = "DTM+041" Then 'Begindatum Periode
                VerzekerdVanDatum = DTMarray(1)
            ElseIf Left(Xarray(Telhier), 7) = "DTM+022" Then 'Einddatum periode
                VerzekerdTotDatum = DTMarray(1)
            End If

        
        Case "IPD"
            IPDarray = Split(Xarray(Telhier), "+")
            If IPDarray(1) = "A502" Then 'Polistype
                PolisType = IPDarray(2)
            End If
        
        Case "NME"
            NMEarray = Split(Xarray(Telhier), "+")
            If NMEarray(1) = "001" Then 'officiele naam klant
                TempoNaamKlant = Replace(NMEarray(2), ":", " ")
            End If
        
        Case "MOA"
            MOAarray = Split(Xarray(Telhier), ":")
            Select Case MOAarray(0)
                Case "MOA+012" 'Totaal te betalen
                    TotaalTeBetalen = Val(MOAarray(1)) / 100
                Case "MOA+013" 'Netto gefractioneerde premie
                    TotaalNettoPremie = Val(MOAarray(1)) / 100
                Case "MOA+015" 'Totaal commissie
                    TotaalCommissie = Val(MOAarray(1)) / 100
                Case "MOA+016" 'Totaal lasten
                    TotaalLasten = Val(MOAarray(1)) / 100
                Case "MOA+017"  'Totaal kosten
                    TotaalKosten = Val(MOAarray(1)) / 100
            End Select
            
            
        Case "PTY"
            PTYarray = Split(Xarray(Telhier), "+")
            If PTYarray(1) = "006" Then 'Verzekeraar
                Verzekeraar = PTYarray(2)
            End If
            
        Case "RFF"
            RFFarray = Split(Xarray(Telhier), ":")
            If RFFarray(0) = "RFF+001" Then 'PolisNummer
                If UBound(RFFarray) > 1 Then
                    If RFFarray(2) = "001" Then 'Definitief polisnummer
                        'maar wie doet daar iets mee ??
                        PolisNummer = RFFarray(1)
                    End If
                Else
                    'hier het nodige indien geen definitief polisnummer."
                    PolisNummer = RFFarray(1)
                End If
            ElseIf RFFarray(0) = "RFF+002" Then 'BijvoegselNummer
                BijvoegselNummer = RFFarray(1)
            End If
            
        Case "XEH"
            'vb. : XEH+01 = CONTRACT
            'vb. : XEH+03 = KWIJTING
            If Mid(Xarray(Telhier), 1, 6) = "XEH+03" Then
                bezigaanKwijting = True
             End If
            
        Case Else
            
            'Stop
        
    End Select
    'Stop
Next

bGet TABLE_CONTRACTS, 0, PolisNummer
If Ktrl Then
    MsgBox "polis " & Trim(PolisNummer) & " niet aanwezig. EDIFACT nieuw nog in te brengen !!!", vbCritical
    TLB_RECORD(TABLE_CONTRACTS) = ""
    'MsgBox "Stop.  Polis bestaat nog niet :" + PolisNummer
    vBib TABLE_CONTRACTS, Mid(VerzekerdVanDatum, 3, 2), "v164"
    'klant zoeken !

    vBib TABLE_CONTRACTS, "NONAME", "A110"
    If Verzekeraar = "2652" Then
        MsgBox "2652: Nateus wordt nog weggeschreven onder nummer 0196"
        Verzekeraar = "0196"
    End If
    vBib TABLE_CONTRACTS, Verzekeraar, "A010"
    vBib TABLE_CONTRACTS, PolisNummer, "A000"
    vBib TABLE_CONTRACTS, TempoNaamKlant, "vs99"
    bInsert TABLE_CONTRACTS, 0
    If Ktrl Then
        MsgBox "onbekende soldaat ! STOP!!!"
    End If
    Infolijn = Infolijn + " " + TempoNaamKlant + " aanvullen a.u.b."
    Dummy = "Kontroleer !!! " + TempoNaamKlant
    'As2TelebibIn.UserInfo(3).Caption = Str$(Val(As2TelebibIn.UserInfo(3).Caption) + 1)
Else
    RecordToVeld TABLE_CONTRACTS
    bGet TABLE_CUSTOMERS, 0, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12)
    If Ktrl Then
        Dummy = "Verbeter !!! " + TempoNaamKlant
    Else
        RecordToVeld TABLE_CUSTOMERS
        Dummy = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
    End If
        
    If TypeKwijting = "2" Then
        bGet TABLE_VARIOUS, 1, "25" + Verzekeraar + PolisNummer
        If Ktrl Then
            TLB_RECORD(TABLE_VARIOUS) = ""
        Else
            RecordToVeld TABLE_VARIOUS
        End If
        vBib TABLE_VARIOUS, Str(TotaalTeBetalen), "B010"                      'premie taksen en kosten inbegrepen
        vBib TABLE_VARIOUS, Str(TotaalLasten + TotaalKosten), "B011"                               'taksen en kosten
        'vBib TABLE_VARIOUS, Str$(Bedrag7), "B013"                                 'premie zonder taksen en kosten
        vBib TABLE_VARIOUS, Str(TotaalCommissie), "B014"                                 'commissieloon
    
        vBib TABLE_VARIOUS, vSet("K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), 13), "v004"
        vBib TABLE_VARIOUS, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12), "A110"
        vBib TABLE_VARIOUS, Verzekeraar, "A010"
        vBib TABLE_VARIOUS, PolisNummer, "A000"
        vBib TABLE_VARIOUS, vSet("25" + Verzekeraar + PolisNummer, 20), "v005"
    
        If Ktrl Then
            bInsert TABLE_VARIOUS, 1
        Else
            bUpdate TABLE_VARIOUS, 1
        End If
        If Ktrl Then
            MsgBox "Stopkode " + Str$(Ktrl)
        End If
    End If
End If
On Error Resume Next
Dim newBoekdatum As String

newBoekdatum = Mid(BoekDatum, 1, 2) & "/" & Mid(BoekDatum, 3, 2) & "/" & Mid(BoekDatum, 5, 4)
'Mid(DatumKwijting, 1, 2) = DagKwijting
'Dim TotaalNettoPremie As Currency
'Dim TotaalLasten As Currency
'Dim TotaalKosten As Currency

GridText = PolisNummer & vbTab
GridText = GridText & newBoekdatum & vbTab
If TypeKwijting = "2" Then
    GridTextIs = "002" & vbTab & Verzekeraar & vbTab & "Contant" & vbTab & newBoekdatum
ElseIf TypeKwijting = "3" Then
    GridTextIs = "003" & vbTab & Verzekeraar & vbTab & "Teruggave" & vbTab & newBoekdatum
Else
    MsgBox "controle type contante wat anders? " & TypeKwijting
End If
'    GridText = GridText & Dec((-TotaalTeBetalen), MaskerHier) & vbTab
'    GridText = GridText & Dec((-TotaalCommissie), MaskerHier) & vbTab
'Else
    GridText = GridText & Dec((TotaalTeBetalen), MaskerHier) & vbTab
    GridText = GridText & Dec((TotaalCommissie), MaskerHier) & vbTab
'End If
GridText = GridText & Dummy & vbTab & SnippetXEH(UserArea, Trim(PolisNummer), False) & vbTab & Str(TotaalNettoPremie) & vbTab & Str(TotaalLasten) & vbCr

BoekjaarKontrole = Mid(BoekDatum, 5, 4)
MaandVerwerking = Mid(BoekDatum, 3, 2)

If GridText = "" Then
Else
    BJPERDAT!Boekjaar.ListIndex = 0
    PeriodeVoor = BoekjaarKontrole + MaandVerwerking
    If PeriodeVoor < Left(BOOKYEAR_FROMTO, 6) Then
        If BJPERDAT!Boekjaar.ListCount > 1 Then
            BJPERDAT!Boekjaar.ListIndex = 1
        End If
    End If
    KtrlFlag = False
    For Tel = 0 To BJPERDAT!PeriodeBoekjaar.ListCount - 1
        BJPERDAT!PeriodeBoekjaar.ListIndex = Tel
        If PeriodeVoor >= Left(PERIOD_FROMTO, 6) And PeriodeVoor <= Mid(PERIOD_FROMTO, 9, 6) Then
            'bClose 99
            Unload KwijtingBoeken
            'GridTextIs = "002" & vbTab & Verzekeraar & vbTab & "Contant"
            KwijtingBoeken.Show 1
            GridText = ""
            KtrlFlag = True
            Exit For
        End If
    Next
    If KtrlFlag = False Then
        MsgBox "Grote gebruikersfout.  Nieuw boekjaar nog niet beschikbaar !" + vbCrLf + "Manuele inbreng is noodzakelijk..."
    End If
End If
    Exit Sub

End Sub


Sub EdiFactREKENINGUITTREKSEL(UserArea As String)

Dim Infolijn As String, Bestand As String, VBestand As String, iolijn As String, PolisNummer As String * 12
Dim FlBestand As Integer, Operatie As Integer
Dim HuidigeBM As String * 2
Dim Dummy As String * 30
Dim fKtrl As Integer
Dim InfoTekst As String
Dim DCBedrag As Currency
Dim MaandVerwerkingTermijn As Integer
Dim BoekjaarKontroleTermijn As String
Dim BorderelTxt As String
Dim Datum2000 As String

Dim Bedrag9 As Currency
Dim BedragPremie As String

Dim MaskerHier As String
Dim Xarray() As String
Dim XGHarray() As String
Dim MOAarray() As String
Dim PTYarray() As String

Dim boekLijn As String

Dim Telhier As Integer
Dim Teljaar As Integer
Dim LijnTeller As Integer
Dim Maatschappij As String
Dim MaandVerwerking As String * 2
Dim BoekjaarKontrole As String * 4
Dim TempoNaamKlant As String
Dim HuidigeIndex As String
Dim DatumKwijting As String
Dim DagKwijting As String * 2
Dim MaandKwijting As String * 2

Dim PolisTeWijzigen As Boolean
Dim DagKtrl1 As String * 2
Dim DagKtrl2 As String * 2

GoSub ruInitialiseren
Xarray = Split(UserArea, "'")
XGHarray = Split(Xarray(0), "+")
Maatschappij = Format(Val(XGHarray(2)), "0000")

BL_LOGGING = True
boekLijn = ""
If bhEuro Then
    MaskerHier = "######0.00"
Else
    MaskerHier = "#########0"
End If

BorderelTxt = ""
GridText = ""
MaandVerwerkingTermijn = 0

LijnTeller = 0
For Telhier = 0 To UBound(Xarray) - 1
    Select Case Left(Xarray(Telhier), 3)
        Case "LIN" 'Volgende verrichting
            'eerst kijken of men niet reeds bezig is met een boekingslijn !
            If boekLijn <> "" Then
                GoSub ruWegschrijven
            End If
            LijnTeller = LijnTeller + 1
            boekLijn = Mid(Xarray(Telhier), 5, 3) & vbTab
            Select Case Mid(Xarray(Telhier), 5, 3)
                Case "001"  'Termijnhernieuwingen
                Case "009"  'Betaling
                
                'Voorlopig enkel de commissies boeken !
                Case "006"  'Commissie inning maatschappij
                    boekLijn = boekLijn & DatumKwijting & vbTab
                    
                Case Else
                    SnelHelpPrint "Nog geen controle voor " & Mid(Xarray(Telhier), 5, 3), BL_LOGGING
            End Select
        
        Case "DTM"
            'Alleen datum uittreksel interesseert ons
            If Left(Xarray(Telhier), 8) = "DTM+069:" Then
                'datum documentlijn
                DatumKwijting = Mid(Xarray(Telhier), 9)
                MaandKwijting = Mid(DatumKwijting, 3, 2)
                MaandVerwerking = MaandKwijting
                BoekjaarKontrole = Mid(DatumKwijting, 5, 4)
                DatumKwijting = Mid(DatumKwijting, 1, 2) & "/" & Mid(DatumKwijting, 3, 2) & "/" & Mid(DatumKwijting, 5, 4)
            End If
        
        Case "RFF"
            'Voorlopig alleen nodig voor de commissies
            If Left(Xarray(Telhier), 8) = "RFF+001:" Then
                PolisNummer = Mid(Xarray(Telhier), 9)
            End If
                               
        Case "MOA"
            MOAarray = Split(Xarray(Telhier), ":")
            Select Case Mid(MOAarray(0), 5, 3)
                Case "083"
                    'beginsaldo Debet ?
            
                Case "084"
                    'beginsaldo Credit ?
                    
                Case "087"
                    'Totaal bedrag premies (+) ?
                    'Totaal bedrag betaling (-) ?
                    
                Case "088"
                    'Commissiebedrag (+) = Credit ?
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag9 = Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 4 Then
                            Bedrag9 = Val(MOAarray(1)) / 10000
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag9 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag9 = Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                    
                Case "MOA+012"  'Totale premie
                
                Case Else
                    SnelHelpPrint "MOA-stop voor " & MOAarray(0), BL_LOGGING
            End Select
            
        Case "NME"
            TempoNaamKlant = Mid(Xarray(Telhier), 9)
            
        Case "PTY"
            PTYarray = Split(Xarray(Telhier), "+")
            If PTYarray(1) = "006" Then
                SnelHelpPrint "maatschappij info" & Xarray(Telhier), BL_LOGGING
                Maatschappij = Format(Val(PTYarray(2)), "0000")
            End If
            
            
        Case "ATT", "XRH", "XRT", "XEH", "XET", "XGH", "XGT", "GIS"
            'voorlopig over te slaan...
        
        Case Else
            SnelHelpPrint "nog niets voorzien voor " & Left(Xarray(Telhier), 3), BL_LOGGING
    End Select
Next

'vervolgen met afsluiting rekeninguittreksel

Dim PeriodeVoor As String
Dim KtrlFlag As Integer

If Len(boekLijn) > 5 Then
    GoSub ruWegschrijven
End If

If GridText = "" Then
    SnelHelpPrint "Er zijn geen commissielonen te boeken", BL_LOGGING
Else
    'BJPERDAT!Boekjaar.ListIndex = 0
    PeriodeVoor = BoekjaarKontrole + MaandVerwerking
    'If PeriodeVoor < Left(BOOKYEAR_FROMTO, 6) Then
    '    If BJPERDAT!Boekjaar.ListCount > 1 Then
    '        BJPERDAT!Boekjaar.ListIndex = 1
    '    End If
    'End If
    KtrlFlag = False
    For Teljaar = 0 To 1
        BJPERDAT!Boekjaar.ListIndex = Teljaar
        For Tel = 0 To BJPERDAT!PeriodeBoekjaar.ListCount - 1
            BJPERDAT!PeriodeBoekjaar.ListIndex = Tel
            If PeriodeVoor >= Left(PERIOD_FROMTO, 6) And PeriodeVoor <= Mid(PERIOD_FROMTO, 9, 6) Then
                'bClose 99
                Unload KwijtingBoeken
                GridTextIs = "006" & vbTab & Maatschappij & vbTab & "Commissies"
                BL_LOGGING = False
                KwijtingBoeken.Show 1
                GridText = ""
                KtrlFlag = True
                Exit For
            End If
        Next
        If KtrlFlag = True Then
            Exit For
        End If
    Next
    If KtrlFlag = False Then
        MsgBox "Gebruikersfout.  Vermoedelijk is een Nieuw boekjaar nog niet beschikbaar !" + vbCrLf + "Manuele inbreng is noodzakelijk... of kies negeren en maak eerst het nieuwe boekjaar aan", vbExclamation
    End If
End If
    Exit Sub

ruInitialiseren:
Bedrag9 = 0
Return

ruWegschrijven:
If Left(boekLijn, 3) <> "006" Then
    Return
End If

If bhEuro = True And MOAarray(2) = "EUR" Then
    'boekhouding en bericht in EURO, niets te doen
ElseIf bhEuro = False And MOAarray(2) = "BEF" Then
    'boekhouding en bericht in BEF, niets te doen
ElseIf MOAarray(2) = "BEF" Then
    'boekhouding in EURO en bericht in BEF !
    Bedrag0 = Round(Bedrag9 / EURO, 2)
ElseIf MOAarray(2) = "EUR" Then
    'boekhouding in EURO en bericht in BEF !
    Bedrag9 = Round(Bedrag9 * EURO)
Else
    MsgBox "ONMOGELIJKE SITUATIE"
End If

bGet TABLE_CONTRACTS, 0, PolisNummer
If Ktrl Then
    MsgBox "polis niet aanwezig. EDIFACT nieuw nog in te brengen !!!", vbCritical
    TLB_RECORD(TABLE_CONTRACTS) = ""
    MsgBox "Stop.  Polis bestaat nog niet :" + PolisNummer
    vBib TABLE_CONTRACTS, Mid(iolijn, 23, 2), "v164"
    vBib TABLE_CONTRACTS, "NONAME", "A110"
    vBib TABLE_CONTRACTS, Maatschappij, "A010"
    vBib TABLE_CONTRACTS, PolisNummer, "A000"
    vBib TABLE_CONTRACTS, Mid(iolijn, 30, 10), "vs99"
    bInsert TABLE_CONTRACTS, 0
    If Ktrl Then
        MsgBox "onbekende soldaat ! STOP!!!"
    End If
    Infolijn = Infolijn + " " + TempoNaamKlant + " aanvullen a.u.b."
    Dummy = "Kontroleer !!! " + TempoNaamKlant
    As2TelebibIn.UserInfo(3).Caption = Str$(Val(As2TelebibIn.UserInfo(3).Caption) + 1)
    ddag$ = "01"
Else
    RecordToVeld TABLE_CONTRACTS
    PolisTeWijzigen = False
    DagKtrl1 = vBibTekst(TABLE_CONTRACTS, "#v165 #")
    DagKtrl2 = Mid(vBibTekst(TABLE_CONTRACTS, "#AW_2 #"), 7, 2)
    If MaandKwijting = vBibTekst(TABLE_CONTRACTS, "#v164 #") Then
    Else
        vBib TABLE_CONTRACTS, MaandKwijting, "v164"
        PolisTeWijzigen = True
    End If
    If DagKtrl1 = DagKtrl2 Then
        'alles ok
    Else
        KtrlBox = MsgBox(DagKtrl2 & " correctie vervalDAG", vbYesNo + vbQuestion + vbDefaultButton1)
        If KtrlBox = vbYes Then
            vBib TABLE_CONTRACTS, DagKtrl2, "v165"
            PolisTeWijzigen = True
        End If
    End If
    DagKwijting = DagKtrl2
    If Maatschappij <> vBibTekst(TABLE_CONTRACTS, "#A010 #") Then
        As2TelebibIn.UserInfo(1).Caption = "Maatschappij conflict!  Polis " + PolisNummer + "bestaat reeds bij maatschappij !" + vBibTekst(TABLE_CONTRACTS, "#A010 #")
        vBib TABLE_CONTRACTS, Maatschappij, "A010"
        PolisTeWijzigen = True
    End If
    If PolisTeWijzigen = True Then
        bUpdate TABLE_CONTRACTS, 0
    End If
    As2TelebibIn.UserInfo(2).Caption = Str$(Val(As2TelebibIn.UserInfo(2).Caption) + 1)
    bGet TABLE_CUSTOMERS, 0, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12)
    If Ktrl Then
        Dummy = "Verbeter !!! " + TempoNaamKlant
    Else
        RecordToVeld TABLE_CUSTOMERS
        Dummy = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
    End If
End If
GoSub ruTeleBibKTRL

GridText = GridText & PolisNummer & vbTab
GridText = GridText & DatumKwijting & vbTab
GridText = GridText & Dec$(0, MaskerHier) & vbTab
GridText = GridText & Dec$((Bedrag9), MaskerHier) & vbTab
GridText = GridText & Dummy & vbTab
'GridText = GridText & SnippetXEH(UserArea, Trim(PolisNummer), False) & vbCr
GridText = GridText & "" & vbCr
GoSub ruInitialiseren
Return

ruTeleBibKTRL:
bGet TABLE_VARIOUS, 1, "25" + Maatschappij + PolisNummer
If Ktrl Then
    TLB_RECORD(TABLE_VARIOUS) = ""
Else
    RecordToVeld TABLE_VARIOUS
End If

Select Case Mid(iolijn, 42, 1)
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
        'AUTO ?
        'If Maatschappij = "0145" And Mid(iolijn, 42, 6) = "000000" Then
        'Else
        '    vBib TABLE_VARIOUS, "P11", "AW00"
        '    vBib TABLE_VARIOUS, "4", "AW06"
        '    vBib TABLE_VARIOUS, Str$(Bedrag1 + Bedrag2 + Bedrag3), "AW04"
        '    vBib TABLE_VARIOUS, Mid(iolijn, 40, 2), "5315"
        'End If
End Select

vBib TABLE_VARIOUS, Str$(Bedrag9), "B014"                                 'commissieloon
vBib TABLE_VARIOUS, vSet("K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), 13), "v004"
vBib TABLE_VARIOUS, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12), "A110"
vBib TABLE_VARIOUS, Maatschappij, "A010"
vBib TABLE_VARIOUS, PolisNummer, "A000"
vBib TABLE_VARIOUS, vSet("25" + Maatschappij + PolisNummer, 20), "v005"

If Ktrl Then
    bInsert TABLE_VARIOUS, 1
Else
    bUpdate TABLE_VARIOUS, 1
End If
If Ktrl Then
    MsgBox "Stopkode " + Str$(Ktrl)
End If
Return

End Sub

Sub Borderel(Maatschappij As String, UserArea As String)
Dim Infolijn As String, Bestand As String, VBestand As String, iolijn As String, PolisNummer As String * 12
Dim FlBestand As Integer, Operatie As Integer
Dim VorigeBM As String * 2
Dim Dummy As String * 30
Dim fKtrl As Integer
Dim InfoTekst As String
Dim DCBedrag As Currency

Dim MaandVerwerking As Integer
Dim BoekjaarKontrole As String

Dim MaandVerwerkingTermijn As Integer
Dim BoekjaarKontroleTermijn As String

Dim BorderelTxt As String
Dim Datum2000 As String

Dim Bedrag1 As Currency
Dim Bedrag2 As Currency
Dim Bedrag3 As Currency
Dim Bedrag4 As Currency
Dim Bedrag5 As Currency
Dim Bedrag6 As Currency
Dim Bedrag7 As Currency
Dim Bedrag8 As Currency
Dim Bedrag9 As Currency

Dim MaskerHier As String

If bhEuro Then
    MaskerHier = "######0.00"
Else
    MaskerHier = "#########0"
End If
'On Local Error GoTo PrintError

iolijn = UserArea
Debug.Assert False

BorderelTxt = ""
GridText = ""
MaandVerwerkingTermijn = 0
Do While iolijn <> Space$(Len(iolijn))
    MaandVerwerking = Val(Mid(iolijn, 19, 2)) 'boekMAAND
    BoekjaarKontrole = Mid(iolijn, 21, 2)     'boekJAAR
    
    PolisNummer = LTrim$(Mid(iolijn, 7, 12))
    Operatie = Val(Mid(iolijn, 4, 1))          '0 tot 9
    Select Case Val(Mid(iolijn, 3, 1))
        Case 1  'Kwijting inning producent
            Select Case Operatie
                Case 1  'Termijnhernieuwing
                    MaandVerwerkingTermijn = Val(Mid(iolijn, 19, 2)) 'boekMAAND
                    BoekjaarKontroleTermijn = Mid(iolijn, 21, 2)     'boekJAAR
                    'InfoLijn = InfoLijn + " " + Mid(IOLijn, 25, 5) 'Index
                    Select Case Mid(iolijn, 42, 1)
                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                            If bhEuro Then
                                Bedrag1 = Round(Val(kt(Mid(iolijn, 42, 6))) / EURO, 2) 'premie B.A. auto
                                Bedrag2 = Round(Val(kt(Mid(iolijn, 48, 5))) / EURO, 2) 'taksen B.A. auto
                                Bedrag3 = Round(Val(kt(Mid(iolijn, 53, 4))) / EURO, 2) 'kosten B.A. auto
                                Bedrag4 = Round(Val(kt(Mid(iolijn, 57, 8))) / EURO, 2) 'andere premie
                                Bedrag5 = Round(Val(kt(Mid(iolijn, 65, 7))) / EURO, 2) 'taksen+kosten andere
                                Bedrag7 = Bedrag1 + Bedrag4
                                Bedrag8 = Bedrag2 + Bedrag3 + Bedrag5
                                Bedrag9 = Round(Val(kt(Mid(iolijn, 72, 7))) / EURO, 2) 'commissie
                            Else
                                Bedrag1 = Val(kt(Mid(iolijn, 42, 6))) 'premie B.A. auto
                                Bedrag2 = Val(kt(Mid(iolijn, 48, 5))) 'taksen B.A. auto
                                Bedrag3 = Val(kt(Mid(iolijn, 53, 4))) 'kosten B.A. auto
                                Bedrag4 = Val(kt(Mid(iolijn, 57, 8))) 'andere premie
                                Bedrag5 = Val(kt(Mid(iolijn, 65, 7))) 'taksen+kosten andere
                                Bedrag7 = Bedrag1 + Bedrag4
                                Bedrag8 = Bedrag2 + Bedrag3 + Bedrag5
                                Bedrag9 = Val(kt(Mid(iolijn, 72, 7))) 'commissie
                            End If
                            GoSub Wegschrijven
                        Case Else
                            If bhEuro Then
                                Bedrag7 = Round(Val(kt(Mid(iolijn, 57, 8))) / EURO, 2)
                                Bedrag8 = Round(Val(kt(Mid(iolijn, 65, 7))) / EURO, 2)
                                Bedrag9 = Round(Val(kt(Mid(iolijn, 72, 7))) / EURO, 2)
                            Else
                                Bedrag7 = Val(kt(Mid(iolijn, 57, 8)))
                                Bedrag8 = Val(kt(Mid(iolijn, 65, 7)))
                                Bedrag9 = Val(kt(Mid(iolijn, 72, 7)))
                            End If
                            GoSub Wegschrijven
                    End Select
                Case Else
                    MsgBox "Kwijting inning producent, andere dan termijnhernieuwing: " + Format(Operatie)
            End Select
            
        Case 2 'Rekeninguittreksel
            Select Case Operatie
                Case 0  'Vorig saldo
                          Msg = "Vorig saldo maatschappij " + Maatschappij + vbCrLf
                    Msg = Msg + "Boekmaand : " + Format(MaandVerwerking, "00") + " Boekjaar : " + BoekjaarKontrole + vbCrLf + vbCrLf
                    Msg = Msg + Format(Val(kt(Mid(iolijn, 57, 8)))) + vbCrLf
                    'MsgBox Msg, , "Rekeninguittreksel"

                Case 1  'Termijntotalen (te kontroleren met bordereltotaal)
                    Msg = "Termijnborderel in rekeninguittreksel voor maatschappij " + Maatschappij + vbCrLf
                    Msg = Msg + "Boekmaand : " + Format(MaandVerwerking, "00") + " Boekjaar : " + BoekjaarKontrole + vbCrLf + vbCrLf
                    Msg = Msg + kt(Mid(iolijn, 57, 8)) + vbCrLf
                    Msg = Msg + kt(Mid(iolijn, 72, 7))
                    MsgBox Msg, , "Rekeninguittreksel"
                    
                Case 2 To 5  '2= Kontante kwijting
                             '3= Terugbetaling van premies
                             '4= Vernietiging van verrichtingen
                             '5= Terugzending van kwitanties
                        If Operatie = 2 Then
                            InfoTekst = "Kontante kwijting"
                        ElseIf Operatie = 3 Then
                            InfoTekst = "Terugbetaling van premie"
                        ElseIf Operatie = 4 Then
                            InfoTekst = "Vernietiging van verrichting"
                        ElseIf Operatie = 5 Then
                            InfoTekst = "Terugzending van kwitantie"
                        Else
                            InfoTekst = "ONBEKENDE OPERATIE"
                        End If
                        
                        Datum2000 = Format("28/" + Format(MaandVerwerking, "00") + "/" + BoekjaarKontrole, "dd/mm/yyyy")
                                                                
                          Msg = "Beweging: " + InfoTekst + " " + Maatschappij + vbCrLf
                    Msg = Msg + "Boekmaand : " + Format(MaandVerwerking, "00") + " Boekjaar : " + BoekjaarKontrole + vbCrLf + vbCrLf
                    Msg = Msg + kt(Mid(iolijn, 57, 8)) + vbCrLf
                    Msg = Msg + kt(Mid(iolijn, 72, 7)) + vbCrLf + vbCrLf
                    Msg = Msg + "Polisnummer : " + PolisNummer + vbCrLf
                    Msg = Msg + Mid(iolijn, 30, 27)
                    MsgBox Msg, , "Rekeninguittreksel"
                                        
                    GridTextPolis = GridTextPolis + Maatschappij + vbTab
                    GridTextPolis = GridTextPolis + PolisNummer + vbTab
                    GridTextPolis = GridTextPolis + Datum2000 + vbTab
                    GridTextPolis = GridTextPolis + Format(CDbl(kt(Mid(iolijn, 57, 8)))) + vbTab
                    GridTextPolis = GridTextPolis + Format(CDbl(kt(Mid(iolijn, 72, 7)))) + vbTab
                    GridTextPolis = GridTextPolis + Mid(iolijn, 30, 27) + vbTab
                    GridTextPolis = GridTextPolis + Format(Operatie) + vbCrLf
                                                            
                Case 6  'Commissielonen
                          Msg = "Commissielonen, maatschappij: " + Maatschappij + vbCrLf
                    Msg = Msg + "Boekmaand : " + Format(MaandVerwerking, "00") + " Boekjaar : " + BoekjaarKontrole + vbCrLf + vbCrLf
                    Msg = Msg + kt(Mid(iolijn, 57, 8)) + vbCrLf
                    Msg = Msg + kt(Mid(iolijn, 72, 7)) + vbCrLf + vbCrLf
                    Msg = Msg + "Polisnummer : " + PolisNummer + vbCrLf
                    Msg = Msg + Mid(iolijn, 30, 27)
                    MsgBox Msg, , "Rekeninguittreksel"
                    
                    InfoTekst = "Commissielonen"
                    Datum2000 = Format("28/" + Format(MaandVerwerking, "00") + "/" + BoekjaarKontrole, "dd/mm/yyyy")
                                            
                    GridTextPolis = GridTextPolis + Maatschappij + vbTab
                    GridTextPolis = GridTextPolis + PolisNummer + vbTab
                    GridTextPolis = GridTextPolis + Datum2000 + vbTab
                    GridTextPolis = GridTextPolis + Format(CDbl(kt(Mid(iolijn, 57, 8)))) + vbTab
                    GridTextPolis = GridTextPolis + Format(CDbl(kt(Mid(iolijn, 72, 7)))) + vbTab
                    GridTextPolis = GridTextPolis + Mid(iolijn, 30, 27) + vbTab
                    GridTextPolis = GridTextPolis + Format(Operatie) + vbCrLf
                    
                Case 7  'Schadegevallen
                    MsgBox "Stop", vbOKOnly, "Schaderegelingen"

                Case 8  'Allerlei
                Case 9  'Financiele verrichtingen
                    'Msg = "Betalingen voor maatschappij " + Maatschappij + vbCrLf
                    'Msg = Msg + "Boekmaand : " + Format(MaandVerwerking, "00") + " Boekjaar : " + BoekjaarKontrole + vbCrLf + vbCrLf
                    'Msg = Msg + Format(Val(kt(Mid(iolijn, 57, 8)))) + vbCrLf
                    'MsgBox Msg, , "Rekeninguittreksel"
            End Select
            
        Case 3  'Kwijting Inning Maatschappij
            If Len(RTrim$(PolisNummer)) = 0 Then
                'MsgBox "Termijnkwijting INNING maatschappij zonder iets, vermoedelijk onze vrienden DELPHI ?"
            Else
                MsgBox "KWIJTING INNING MAATSCHAPPIJ"
            End If
            
        Case Else
            MsgBox "onbekende kode"
    End Select
    iolijn = Right(iolijn, Len(iolijn) - 80)
Loop

Dim PeriodeVoor As String
Dim KtrlFlag As Integer

If GridText = "" Then
Else
    BJPERDAT!Boekjaar.ListIndex = 0
    If BoekjaarKontroleTermijn < "86" Then
        PeriodeVoor = "20" + BoekjaarKontroleTermijn + Format(MaandVerwerkingTermijn, "00")
    Else
        MsgBox "stop"
        PeriodeVoor = "19" + BoekjaarKontroleTermijn + Format(MaandVerwerkingTermijn, "00")
    End If
    If PeriodeVoor < Left(BOOKYEAR_FROMTO, 6) Then
        If BJPERDAT!Boekjaar.ListCount > 1 Then
            BJPERDAT!Boekjaar.ListIndex = 1
        End If
    End If
    KtrlFlag = False
    For Tel = 0 To BJPERDAT!PeriodeBoekjaar.ListCount - 1
        BJPERDAT!PeriodeBoekjaar.ListIndex = Tel
        If PeriodeVoor >= Left(PERIOD_FROMTO, 6) And PeriodeVoor <= Mid(PERIOD_FROMTO, 9, 6) Then
            'bClose 99
            Unload KwijtingBoeken
            KwijtingBoeken.Show 1
            GridText = ""
            KtrlFlag = True
            Exit For
        End If
    Next
    If KtrlFlag = False Then
        MsgBox "Grote gebruikersfout.  Nieuw boekjaar nog niet beschikbaar !" + vbCrLf + "Manuele inbreng is noodzakelijk..."
    End If
End If
    Exit Sub

Wegschrijven:
bGet TABLE_CONTRACTS, 0, PolisNummer
If Ktrl Then
    TLB_RECORD(TABLE_CONTRACTS) = ""
    MsgBox "Stop.  Polis bestaat nog niet :" + PolisNummer
    vBib TABLE_CONTRACTS, Mid(iolijn, 23, 2), "v164"
    vBib TABLE_CONTRACTS, "NONAME", "A110"
    vBib TABLE_CONTRACTS, Maatschappij, "A010"
    vBib TABLE_CONTRACTS, PolisNummer, "A000"
    vBib TABLE_CONTRACTS, Mid(iolijn, 30, 10), "vs99"
    bInsert TABLE_CONTRACTS, 0
    If Ktrl Then
        MsgBox "onbekende soldaat ! STOP!!!"
    End If
    Infolijn = Infolijn + " " + Mid(iolijn, 30, 10) + " aanvullen a.u.b."
    Dummy = "Kontroleer !!! " + Mid(iolijn, 30, 10)
    As2TelebibIn.UserInfo(3).Caption = Str$(Val(As2TelebibIn.UserInfo(3).Caption) + 1)
    ddag$ = "01"
Else
    RecordToVeld TABLE_CONTRACTS
    ddag$ = vBibTekst(TABLE_CONTRACTS, "#v165 #")
    If Maatschappij <> vBibTekst(TABLE_CONTRACTS, "#A010 #") Then
        As2TelebibIn.UserInfo(1).Caption = "Maatschappij konflikt!  Polis " + PolisNummer + "bestaat reeds bij maatschappij !" + vBibTekst(TABLE_CONTRACTS, "#A010 #")
        vBib TABLE_CONTRACTS, Maatschappij, "A010"
        bUpdate TABLE_CONTRACTS, 0
    End If
    As2TelebibIn.UserInfo(2).Caption = Str$(Val(As2TelebibIn.UserInfo(2).Caption) + 1)
    bGet TABLE_CUSTOMERS, 0, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12)
    If Ktrl Then
        Dummy = "Verbeter !!! " + Mid(iolijn, 30, 10)
    Else
        RecordToVeld TABLE_CUSTOMERS
        Dummy = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
    End If
End If
GoSub TeleBibKTRL

GridText = GridText + PolisNummer + vbTab
If BoekjaarKontroleTermijn < "86" Then
    GridText = GridText + ddag$ + "/" + Mid(iolijn, 23, 2) + "/20" + BoekjaarKontroleTermijn + vbTab
Else
    GridText = GridText + ddag$ + "/" + Mid(iolijn, 23, 2) + "/19" + BoekjaarKontroleTermijn + vbTab
End If
GridText = GridText + Dec$((Bedrag7 + Bedrag8), MaskerHier) + vbTab
GridText = GridText + Dec$((Bedrag9), MaskerHier) + vbTab
GridText = GridText + Dummy + Chr$(13)
Return

TeleBibKTRL:
bGet TABLE_VARIOUS, 1, "25" + Maatschappij + PolisNummer
If Ktrl Then
    TLB_RECORD(TABLE_VARIOUS) = ""
Else
    RecordToVeld TABLE_VARIOUS
End If

Select Case Mid(iolijn, 42, 1)
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
        If Maatschappij = "0145" And Mid(iolijn, 42, 6) = "000000" Then
        Else
            vBib TABLE_VARIOUS, "P11", "AW00"
            vBib TABLE_VARIOUS, "4", "AW06"
            vBib TABLE_VARIOUS, Str$(Bedrag1 + Bedrag2 + Bedrag3), "AW04"
            VorigeBM = vBibTekst(TABLE_VARIOUS, "#5315 #")
            vBib TABLE_VARIOUS, VorigeBM, "5310"
            vBib TABLE_VARIOUS, Mid(iolijn, 40, 2), "5315"
        End If
End Select

vBib TABLE_VARIOUS, Mid(iolijn, 25, 5), "AW.R"                            'index
vBib TABLE_VARIOUS, Str$(Bedrag7 + Bedrag8), "B010"                      'premie taksen en kosten inbegrepen
vBib TABLE_VARIOUS, Str$(Bedrag8), "B011"                                 'taksen en kosten
vBib TABLE_VARIOUS, Str$(Bedrag7), "B013"                                 'premie zonder taksen en kosten
vBib TABLE_VARIOUS, Str$(Bedrag9), "B014"                                 'commissieloon
vBib TABLE_VARIOUS, vSet("K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), 13), "v004"
vBib TABLE_VARIOUS, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12), "A110"
vBib TABLE_VARIOUS, Maatschappij, "A010"
vBib TABLE_VARIOUS, PolisNummer, "A000"
vBib TABLE_VARIOUS, vSet("25" + Maatschappij + PolisNummer, 20), "v005"

If Ktrl Then
    bInsert TABLE_VARIOUS, 1
Else
    bUpdate TABLE_VARIOUS, 1
End If
If Ktrl Then
    MsgBox "Stopkode " + Str$(Ktrl)
End If
Return

PrintError:
    FlBestand = FreeFile
    Open LOCATION_COMPANYDATA + Bestand + ".txt" For Append As FlBestand
    Resume
    Exit Sub

End Sub



Function kt$(fBedrag As String)

Select Case Asc(Right(fBedrag, 1))
    Case 48 To 57, 32
        kt = fBedrag
    Case 232, 233
        kt = Left(fBedrag, Len(fBedrag) - 1) + "0"
    Case 65 To 73
        kt = Left(fBedrag, Len(fBedrag) - 1) + Format(Asc(Right(fBedrag, 1)) - 64, "0")
    Case 74 To 82
        Mid(fBedrag, 1, 1) = "-"
        kt = Left(fBedrag, Len(fBedrag) - 1) + Format(Asc(Right(fBedrag, 1)) - 73, "0")
    Case Else
        MsgBox "Foutieve waarde in conversietafel voor '" + fBedrag + "'" + vbCrLf + vbCrLf + "Kontakteer onmiddellijk de maatschappij !!"
        kt = String(Len(fBedrag), "0")
End Select

End Function

Function KTRLVerzoek(Verzoekstring As String, Maatschappij As String) As Integer

KTRLVerzoek = False
bGet TABLE_VARIOUS, 1, vSet("23CO" + Maatschappij, 20)
If Ktrl Then
    'MsgBox "Parameters voor maatschappij " + Maatschappij + " bestaan niet.  Vernieuw eventueel de kontrolelijst !", , "AS1 Verzoekkontrole"
Else
    RecordToVeld TABLE_VARIOUS
    If Mid(vBibTekst(TABLE_VARIOUS, "#v220 #"), Val(Verzoekstring), 1) = "1" Then
        KTRLVerzoek = True
    End If
End If

End Function

Sub LogBoekUpdate()
Dim T As Integer
Dim FlAsNet As Integer
Dim FlAsDum As Integer
Dim RecordOK As Integer
Dim RecordNOT As Integer
Dim A As String
Dim Msg As String
Dim PolisNummer As String * 12
Dim As1VerzoekNummer As String * 3
Dim AS1VerzoekTekst As String

On Local Error GoTo OpenDummy

If Dir$(LOCATION_ASWEB + "AS1.SND") = "" Then
    Exit Sub
Else
    FlAsNet = FreeFile
    Open LOCATION_ASWEB + "AS1.SND" For Input As FlAsNet
End If

bBegin
Do While Not EOF(FlAsNet)
    A = ""
    Line Input #FlAsNet, A
    As1VerzoekNummer = Mid(A, 17, 3)
    If As1VerzoekNummer = "912" Or As1VerzoekNummer = "915" Then
        GoTo NextLoop
    End If
    
    If Mid(A, 16, 1) = "1" Then
    TLB_RECORD(TABLE_VARIOUS) = ""
    Select Case As1VerzoekNummer
            Case "020"  'Vrij Bericht
                PolisNummer = Mid(A, 92, 12)
                AS1VerzoekTekst = Mid(A, 17, 3) + ":" + Mid(A, 51, 40)
                               
            Case "022"  'Wijziging Inlichtigen Klient
                PolisNummer = Mid(A, 141, 12)
                AS1VerzoekTekst = Mid(A, 17, 3) + ":" + "Wijziging inlichtingen Klient"
                
            Case "024"  'Wijziging Inlichtingen Polis Auto
                PolisNummer = Mid(A, 141, 12)
                AS1VerzoekTekst = Mid(A, 17, 3) + ":" + "Wijzig. inlichtingen polis auto"
                
            Case "027", "028"
                PolisNummer = Mid(A, 141, 12)
                AS1VerzoekTekst = Mid(A, 17, 3) + ":" + Mid(A, 245, 2) + "/" + Mid(A, 140 + 71, 9) + "/" + Mid(A, 220, 8)

            Case "913"
                MsgBox "913 *error*"
                GoTo NextLoop
            Case Else
                MsgBox "Niet beschikbare funktiekode : " + Mid(A, 17, 3) + vbCrLf + vbCrLf + "Raadpleeg R&V 053/21.59.25"
                RecordNOT = RecordNOT + 1
                Print #FlAsDum, A
                GoTo NextLoop
        End Select
        RecordOK = RecordOK + 1
        TLB_RECORD(TABLE_VARIOUS) = ""
        vBib TABLE_VARIOUS, Mid(A, 4, 12), "v219" 'IDA
        vBib TABLE_VARIOUS, PolisNummer, "A000"    'polis
        vBib TABLE_VARIOUS, A, "v228"              'record
        vBib TABLE_VARIOUS, "26" + vBibTekst(TABLE_VARIOUS, "#v219 #"), "v005"
        bInsert TABLE_VARIOUS, 1
        If Ktrl Then
            MsgBox "Fout bij invoegen 26-record"
        End If
        If Mid(A, 50, 1) = "X" Then
            vBib TABLE_VARIOUS, AS1VerzoekTekst, "v105"
            vBib TABLE_VARIOUS, PolisNummer, "A000"
            vBib TABLE_VARIOUS, Mid(A, 4, 12), "v219"    'IDA
            vBib TABLE_VARIOUS, Left(Format(Now, "yyyy"), 2) + Mid(A, 30, 12), "v128" 'Datum, uur, min, sec
            vBib TABLE_VARIOUS, "22" + vBibTekst(TABLE_VARIOUS, "#A000 #"), "v005"
            bInsert TABLE_VARIOUS, 1
            If Ktrl Then
                MsgBox "Fout bij invoegen 22-record"
                GoTo NextLoop
            End If
        End If
    Else
        RecordNOT = RecordNOT + 1
        Print #FlAsDum, A
    End If
NextLoop:
Loop
Close FlAsNet
Close FlAs1Log
If FlAsDum Then
    Close FlAsDum
End If
bEnd

      Msg = Format(RecordOK) + " verzoeken werden in het logboek ingeschreven." + vbCrLf + vbCrLf
Msg = Msg + "Nog " + Format(RecordNOT) + " verzoeken te versturen naar ASSURNET" + vbCrLf
Msg = Msg + "via de keuze [B]estandsoverdracht in de AS/1 menu."
MsgBox Msg, 0, "AS/1 Logboek bestandsoverdracht"
Kill LOCATION_ASWEB + "AS1.snd"
If Dir$(LOCATION_ASWEB + "AS1.$$$") <> "" Then
    Name LOCATION_ASWEB + "AS1.$$$" As LOCATION_ASWEB + "AS1.SND"
End If
Exit Sub

OpenDummy:
    FlAsDum = FreeFile
    Open LOCATION_ASWEB + "AS1.$$$" For Output As FlAsDum
    Resume
    
End Sub

Sub KTRLMijStatistiek(VerzoekTxt As String)
Dim FlNetTemp As Integer
Dim A As String * 140
Dim aa As String

Select Case VerzoekTxt
    Case "912"
        MsgTxt$ = "Beschikbare toepassingen "
        aa = Space$(420)
    Case "915"
        MsgTxt$ = "Bepalingen "
        aa = Space$(150)
End Select
Msg$ = MsgTxt$ + "maatschappijen opvragen ?" + vbCrLf + vbCrLf + "(enkel nuttig bij eerste gebruik Assurnet !)"
Ktrl = MsgBox(Msg$, 292, "Verzoek " + VerzoekTxt)
If Ktrl = 6 Then
Else
    Exit Sub
End If

A = ""
Mid(A, 1, 3) = "AS1"                   'constante
Mid(A, 16, 1) = "0"                    'constante
Mid(A, 17, 3) = VerzoekTxt             'verzoek
Mid(A, 20, 1) = "1"                    'gestruktureerd
Mid(A, 21, 1) = "1"                    'producent
Mid(A, 22, 8) = vSet(ProducentNummer, 8)    'ASNET prod.nr.
Mid(A, 30, 12) = Format(Now, "yymmddhhmmss") 'datum
Mid(A, 42, 1) = "1"                    'gewoon niet vertrouwelijk
Mid(A, 43, 1) = "4"                    'bestemmeling=Assurnet"
Mid(A, 44, 4) = "XXXX" 'BVVO kode
Mid(A, 48, 2) = "01"                   'volgnummer
Mid(A, 50, 1) = "X"                    'einde tekst

Mid(aa, 1, 8) = vSet(ProducentNummer, 8)    'ASNET prod.nr.

Select Case VerzoekTxt
    Case "912"
        Mid(aa, 9, 1) = "1"                        'alle mijen
    Case "915"
        Mid(aa, 9, 1) = "9"                        'gewenste taalkode Nederlands en frans
        Mid(aa, 10, 1) = "1"                       'alle BVVO kodes
        Mid(aa, 61, 1) = "1"                       'alle takken/ondertakken
        Mid(aa, 92, 1) = "1"                       'verklarende tekst gewenst
        Mid(aa, 93, 1) = "1"                       'verbindingen tussen gegevens
End Select

FlNetTemp = FreeFile
Open LOCATION_ASWEB + "AS1.SND" For Append As FlNetTemp
Print #FlNetTemp, A + aa
Close FlNetTemp

End Sub

Function Puur$(Nummer As String)
Dim zoek    As String * 3
Dim zoek1   As String * 1
Dim T       As Integer
Dim Plaats  As Integer

zoek = "./-"
For T = 1 To 3
    zoek1$ = Mid(zoek, T, 1)
    Do While InStr(Nummer, zoek1)
        Plaats = InStr(Nummer, zoek1)
        Nummer = Left(Nummer, Plaats - 1) + Right(Nummer, Len(Nummer) - Plaats)
    Loop
Next
Puur = Nummer

End Function

Function Tk$(fBedrag As String, lengte As Integer)

Tk$ = Format(Val(fBedrag), String$(lengte, "0"))

End Function

Function UpdatePolisDatabase(Nummer As Integer, Bericht As String) As Integer
Dim Tekst As String
Dim PolisNummer As String * 12
Dim Maatschappij As String * 4
Dim TempoKtrl As Integer

UpdatePolisDatabase = False
If TLBPag3("AS1" + Format(Nummer, "000")) Then
    PolisNummer = LTrim$(RTrim$(Mid(Bericht, TELEBIB_POS(0), TELEBIB_LENGTH(0))))
    Maatschappij = RTrim$(Mid(Bericht, TELEBIB_POS(1), TELEBIB_LENGTH(1)))
    bGet TABLE_VARIOUS, 1, "25" + Maatschappij + PolisNummer
    If Ktrl Then
        TempoKtrl = Ktrl
        TLB_RECORD(TABLE_VARIOUS) = ""
    Else
        TempoKtrl = 0
        RecordToVeld TABLE_VARIOUS
    End If
    T = 0
    Do While TELEBIB_CODE(T) <> Space$(10)
        Tekst = LTrim$(RTrim$(Mid(Bericht, TELEBIB_POS(T), TELEBIB_LENGTH(T))))
        If Tekst <> "" Then
            If TempoKtrl Then
                vBib TABLE_VARIOUS, Tekst, Mid(TELEBIB_CODE(T), 5, 5)
            Else
                If LTrim$(RTrim$(vBibTekst(TABLE_VARIOUS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#"))) <> Tekst Then
                    Msg = LTrim$(RTrim$(vBibTekst(TABLE_VARIOUS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#"))) + " vervangen door" + vbCrLf
                    Msg = Msg + Tekst + vbCrLf + vbCrLf
                    Msg = Msg + "Bent U zeker ?"
                    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton1)
                    If KtrlBox = vbYes Then
                        vBib TABLE_VARIOUS, Tekst, Mid(TELEBIB_CODE(T), 5, 5)
                    End If
                End If
            End If
        End If
        T = T + 1
    Loop
End If

End Function


