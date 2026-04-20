Attribute VB_Name = "thisEnvironmentClass"
Option Explicit
DefInt A-Z

Sub InitEerst()

    FULL_LINE = String$(128, 173)

    TABLEDEF_ONT(TABLE_VARIOUS) = "0000000.ONT"         '00
    TABLEDEF_ONT(TABLE_CUSTOMERS) = "0010000.ONT"            '01
    TABLEDEF_ONT(TABLE_SUPPLIERS) = "0020000.ONT"      '02
    TABLEDEF_ONT(TABLE_LEDGERACCOUNTS) = "0030000.ONT"         '03
    TABLEDEF_ONT(TABLE_PRODUCTS) = "0040000.ONT"          '04
    TABLEDEF_ONT(TABLE_JOURNAL) = "0600000.ONT"         '05
    TABLEDEF_ONT(TABLE_INVOICES) = "0200000.ONT"         '06
    TABLEDEF_ONT(TABLE_CONTRACTS) = "0700000.ONT"            '07
    TABLEDEF_ONT(TABLE_DUMMY) = "90DUMMY.ONT"            '08
    TABLEDEF_ONT(TABLE_COUNTERS) = "00.ONT"                '09

    bstNaam(TABLE_VARIOUS) = "Allerlei"        '00
    bstNaam(TABLE_CUSTOMERS) = "Klanten"            '01
    bstNaam(TABLE_SUPPLIERS) = "Leveranciers" '02
    bstNaam(TABLE_LEDGERACCOUNTS) = "Rekeningen"      '03
    bstNaam(TABLE_PRODUCTS) = "Produkten"        '04
    bstNaam(TABLE_JOURNAL) = "Journalen"       '05
    bstNaam(TABLE_INVOICES) = "dokumenten"      '06
    bstNaam(TABLE_CONTRACTS) = "Polissen"           '07
    bstNaam(TABLE_DUMMY) = "TmpBestand"         '08
    bstNaam(TABLE_COUNTERS) = "Tell"              '09

    DAYS_IN_MONTH(1) = 31
    DAYS_IN_MONTH(2) = 29
    DAYS_IN_MONTH(3) = 31
    DAYS_IN_MONTH(4) = 30
    DAYS_IN_MONTH(5) = 31
    DAYS_IN_MONTH(6) = 30
    DAYS_IN_MONTH(7) = 31
    DAYS_IN_MONTH(8) = 31
    DAYS_IN_MONTH(9) = 30
    DAYS_IN_MONTH(10) = 31
    DAYS_IN_MONTH(11) = 30
    DAYS_IN_MONTH(12) = 31

    MONTH_AS_TEXT(1) = "Januari  "
    MONTH_AS_TEXT(2) = "Februari "
    MONTH_AS_TEXT(3) = "Maart    "
    MONTH_AS_TEXT(4) = "April    "
    MONTH_AS_TEXT(5) = "Mei      "
    MONTH_AS_TEXT(6) = "Juni     "
    MONTH_AS_TEXT(7) = "Juli     "
    MONTH_AS_TEXT(8) = "Augustus "
    MONTH_AS_TEXT(9) = "September"
    MONTH_AS_TEXT(10) = "October  "
    MONTH_AS_TEXT(11) = "November "
    MONTH_AS_TEXT(12) = "December "

End Sub

Function InstellingenBewaard(frmVenster As Form) As Boolean

    Err = 0: InstellingenBewaard = True
    On Error Resume Next
    SaveSetting App.ProductName, frmVenster.Name, "Top", frmVenster.Top
    SaveSetting App.ProductName, frmVenster.Name, "Links", frmVenster.Left
    SaveSetting App.ProductName, frmVenster.Name, "Breedte", frmVenster.Width
    SaveSetting App.ProductName, frmVenster.Name, "Hoogte", frmVenster.Height
    If Err Then InstellingenBewaard = False

End Function


Sub LaadInstellingen(frmVenster As Form)

    On Local Error GoTo EenFoutBijINLaden

    frmVenster.Top = GetSetting(App.ProductName, frmVenster.Name, "Top")
    frmVenster.Left = GetSetting(App.ProductName, frmVenster.Name, "Links")
    frmVenster.Width = GetSetting(App.ProductName, frmVenster.Name, "Breedte")
    frmVenster.Height = GetSetting(App.ProductName, frmVenster.Name, "Hoogte")
    Exit Sub

EenFoutBijINLaden:
    Exit Sub

End Sub


Sub BeWaarTekst(Onderdeel As String, SubDeel As String, Element As String)

    SaveSetting App.ProductName, Onderdeel, SubDeel, Element

End Sub

Function LaadTekst(Onderdeel As String, SubDeel As String) As Variant

    On Local Error Resume Next
    If InStr(Onderdeel, ";") Then
        LaadTekst = GetSetting(Left(Onderdeel, InStr(Onderdeel, ";") - 1), Mid(Onderdeel, InStr(Onderdeel, ";") + 1), SubDeel)
    Else
        LaadTekst = GetSetting(App.ProductName, Onderdeel, SubDeel)
    End If

End Function

