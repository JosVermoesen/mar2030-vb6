Attribute VB_Name = "thisGlobals"
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long

'Mijn dokumenten, ApplicatieData
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PROGRAM_FILES = &H26
  
Public Const SND_ASYNC = &H1     ' Play asynchronously
Public Const SND_NODEFAULT = &H2 ' Don't use default sound
Public Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Public SoundBuffer As String

Public Const TABLE_STR = "Tabel"
Public Const ATTACHED_STR = "Verbonden"
Public Const QUERY_STR = "Opzoeking"
Public Const FIELD_STR = "Kolom"
Public Const FIELDS_STR = "Kolommen"
Public Const INDEX_STR = "Index"
Public Const INDEXES_STR = "Indexen"
Public Const PROPERTY_STR = "Eigenschap"
Public Const PROPERTIES_STR = "Eigenschappen"

Public Const ADOJET_PROVIDER = "Provider=Microsoft.Jet.OLEDB.4.0;"

Public gnodDBNode    As node        'current database node in treeview
Public gnodDBNode2   As node        'backup of current database node i

'marNT constanten
Public Const NUMBER_TABLES = 9
Public Const TABLE_VARIOUS = 0
Public Const TABLE_CUSTOMERS = 1
Public Const TABLE_SUPPLIERS = 2
Public Const TABLE_LEDGERACCOUNTS = 3
Public Const TABLE_PRODUCTS = 4
Public Const TABLE_CONTRACTS = 5
Public Const TABLE_INVOICES = 6
Public Const TABLE_JOURNAL = 7
Public Const TABLE_DUMMY = 8
Public Const TABLE_COUNTERS = 9
'Public Const TABLE_AS1LOG = 10

Public Const PERIODAS_TEXT = 0
Public Const BOOKYEARAS_TEXT = 1
Public Const PERIODAS_KEY = 2
Public Const BOOKYEAR_KEY = 3
Public Const SISO = "001*002*002*003*004*005*006*007*008*009*010*011*030*032*038*046*053*054*055*060*061*063*064*091*600*"
Public Const MAX_TELEBIB = 150
'Public Const MAX_INDEX = 5
'Public Const MAX_PLUS = 6
Public Const READING = True
Public Const READING_LOCK = False

Public Const MASK_EURX = "######0.0000"
Public Const MASK_EURBH = "########0.00"

Public Const MASK_BEF = "##########"
Public Const MASK_EUR = "######0.00"

Public Const EURO = 40.3399
'Public Const BELGIAN_FRANC = 1

Public FileNameQR As String
Public PeppolFlag As Boolean

Public MASK_SY(0 To 8) As String
Public MASK_2002       As String * 10
Public VSF_PRO           As Boolean

Public SYS_VAR(0 To 6) As String
Public FILE_NR(0 To NUMBER_TABLES) As Integer
Public TLB_RECORD(0 To NUMBER_TABLES) As String
Public KEY_BUF(0 To NUMBER_TABLES) As String
Public TABLEDEF_ONT(0 To NUMBER_TABLES) As String
Public KEY_INDEX(0 To NUMBER_TABLES) As Integer
Public INSERT_FLAG(0 To NUMBER_TABLES) As Integer
Public FL_NUMBEROFINDEXEN(0 To 10) As Integer
Public JETTABLEUSE_INDEX(0 To NUMBER_TABLES, 0 To 10) As String
Public FLINDEX_LEN(0 To NUMBER_TABLES, 0 To 10) As Integer
Public FLINDEX_CAPTION(0 To NUMBER_TABLES, 0 To 10) As String
'Public LIST_IDX(0 To 5, 0 To 10) As String
Public FVT(0 To NUMBER_TABLES, 0 To 10) As String
Public DAYS_IN_MONTH(12) As Integer
Public MONTH_AS_TEXT(1 To 12) As String * 9
'Public SETUP_RECNUM(25) As Integer

'Public ESC_CODES_PRINTER(2) As Integer
'Public PAPERLENGTH(2) As Integer
'Public PRINTER_INI(2) As String

Public REPORT_FIELD(23) As String
Public REPORT_TAB(23) As Integer

Public TELEBIB_CODE(-1 To MAX_TELEBIB) As String * 10
Public TELEBIB_TEXT(0 To MAX_TELEBIB) As String * 35
Public TELEBIB_TYPE(0 To MAX_TELEBIB) As String * 1
Public TELEBIB_LENGTH(0 To MAX_TELEBIB) As Integer
Public TELEBIB_POS(0 To MAX_TELEBIB) As Integer
Public TELEBIB_LAST As Integer

Public FL99 As Integer
Public FL99_RECORD As String
Public PRINTER_CURRENT_Y As Integer
Public PAGE_COUNTER As Integer
Public FULL_LINE As String * 128

Public MAR_VERSION As String
Public LOG_PRINT As String
Public BL_LOGGING As Boolean

Public DKTRL_CUMUL As Currency
Public DKTRL_BEF As Currency
Public DKTRL_EUR As Currency

Public B_MODUS As Integer
Public COUNT_TO As Integer

Public PERIOD_FROMTO As String * 16
Public BOOKYEAR_FROMTO As String * 16
Public ACTIVE_BOOKYEAR As Integer
Public MIM_GLOBAL_DATE As String * 10
Public VAT_BOBTHEBUILDERS As Boolean
Public DIRECTSELL_STRING As String

Public LOCATION_DESKTOP As String
Public LOCATION_COMPANYDATA As String
Public LOCATION_NETDATA As String
Public PROGRAM_LOCATION As String
Public LOCATION_ As String
Public LOCATION_ASWEB As String
Public LOCATION_MYDOCUMENTS As String
Public SYSTEM_MYPERSONALDOCUMENTS As String

Public ProducentNummer As String * 8
Public Eigenaar As String * 8
Public Fl As Integer
Public SharedFl As Integer
Public SharedScanFl As Integer
Public Ktrl As Integer
Public KtrlLong As Long
Public aIndex As Integer
Public AktieveFiche As Integer

Public blMilieu As Boolean
Public MilieuGridText As String
Public GridText As String
Public GridTextIs As String
Public GridTextPolis As Variant
Public GridText9 As String
Public GridRows As Integer
Public XLogKey As String

Public XLogKassa As String

Public dKtrCumul As Double
Public SetupVelden As Integer
Public BedrijfKeuze As String
Public dMuntL As Double
Public Msg As String
Public KtrlBox As Integer
Public SQLBevel As String
Public DoEventsStatus As Integer
Public VsoftLog As Integer
Public ProgrammaVersie As String
Public LockHold As Integer

'Public KBTable As DAO.Recordset
Public ntDB As DAO.Database
Public ntRS(0 To 9) As DAO.Recordset
Public NTRuimte As DAO.Workspace

Public adKBDB As ADODB.Connection
Public adKBTable As ADODB.Recordset

Public adntDB As ADODB.Connection
Public adntDBSQLS As ADODB.Connection

Public adTBIB As ADODB.Connection
Public rsWaarden As ADODB.Recordset
Public rsJournaal As ADODB.Recordset
Public rsMAR(0 To 9)  As ADODB.Recordset
Public jetConnect As String
Public SQLConnect As String
    
Public XDoEvents As Integer
Public bstNaam(0 To 9) As String
Public AddNewStatus(0 To 9) As Integer
Public vBC(0 To 9, 0 To 200) As String * 4
Public BAModus As Integer

Public TestEuroModus As Boolean
Public bhEuro As Boolean
Public XisEuroWisBEF As Boolean

Public TimerTijd As Date
Public RetVal As Variant
Public Figuur1 As Variant
Public Figuur2 As Variant

Public LijstPrinterNr As Integer
Public dokumentPrinterNr As Integer
Public KassaPrinterNr As Integer

Public FormReference As BasicTable
Public BasisB(1 To 3) As New BasicTable
Public JumpVenster As Variant

Public fs As FileSystemObject

Public KasTicketTotaal As Currency
Public KasTotaal As Currency
Public KasBetalingBEF As Currency
Public KasBetalingEUR As Currency
Public KasTerugEUR As Currency

Public KasTotaalBEF As Currency
Public KasTotaalEUR As Currency

Public DecimalKTRL As Boolean

'marIntegraal.NET
'Public xpW As Word.dokument
Public usrLicentieInfo As String
Public JournaalLocked As Boolean
Public usrMailAdres As String
Public usrPW As String

Public pdfVsoftVanaf As Double
Public pdfVsoftTot As Double
Public pdfadresXpos As Double
Public pdfadresYpos As Double
Public pdfadresXpos2 As Double
Public pdfadresYpos2 As Double

Public strTELEBIBIO As String

Public uitwisselingOMS As String
Public uitwisselingDATA As String
Public documentLinesOMS As String
Public documentLinesDATA As String

Public uitwisselingOMSArray() As String
Public uitwisselingDATAArray() As String
Public documentLinesOMSArray() As String
Public documentLinesDATAArray() As String

' LegalMonetaryTotal
Public legalMonetaryTotalOMS As String
Public legalMonetaryTotalDATA As String
Public legalMonetaryTotalOMSArray() As String
Public legalMonetaryTotalDATAArray() As String

' AllowanceCharge
' TODO: Later extending in VB6 and .NET
Public allowanceChargeOMS As String
Public allowanceChargeDATA As String
Public allowanceChargeOMSArray() As String
Public allowanceChargeDATAArray() As String

Public xdaOMS As String
Public xdaDATA As String
Public xdaLinesOMS As String
Public xdaLinesDATA As String

Public xdaOMSArray() As String
Public xdaDATAArray() As String
Public xdaLinesOMSArray() As String
Public xdaLinesDATAArray() As String

Public purchasePeppolTODOShowed As Boolean
Public appTitleAndVersion As String

Public investWarning As Boolean
Public bankNumberArray() As String

