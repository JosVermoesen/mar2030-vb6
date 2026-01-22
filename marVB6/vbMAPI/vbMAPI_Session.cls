VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *****************************************************************************
' *  Virtual-COM library: vbMAPI  http://www.everythingaccess.com/vbmapi.htm  *
' *  v1.16.3.0 DEMO              built Nov  8 2023                            *
' *---------------------------------------------------------------------------*
' *  This is a demonstration version of the library for TESTING purposes only.*
' *****************************************************************************

Option Explicit

Enum ExceptionCodesEnum
    E_OUT_OF_MEMORY = &H80042210
    E_FILE_ERROR = &H8004220C
    E_CLOSING_FILE = &H8004220F
    E_FILE_NOT_FOUND2 = &H8004220A
    E_INVALIDARG = &H80070057
    E_UNEXPECTED = &H8000FFFF
    E_ACCESSDENIED = &H80070005
    E_FAIL = &H80004005
    DISP_E_BADPARAMCOUNT = &H8002000E
    DISP_E_BADVARTYPE = &H80020008
    E_FILE_NOT_FOUND = &H80070002
    E_ALREADY_LOGGED_ON = &H80042201
    E_SESSION_NOT_LOGGED_ON = &H80042202
    E_INVALID_STORE = &H80042203
    E_INVALID_FOLDER = &H80042204
    E_INVALID_MAILITEM = &H80042205
    E_INVALID_SESSION = &H80042206
    E_INVALID_ACCOUNT = &H80042207
    E_UNABLE_TO_OPEN_ADDRESSBOOK = &H80042208
    E_FAILED_TO_CREATE_ATTACHMENT = &H80042209
    E_OPENING_STREAM = &H8004220B
    E_CREATING_FILE = &H8004220D
    E_WRITING_FILE = &H8004220E
    E_READING_STREAM = &H80042211
    E_OPENING_ACCOUNT = &H80042212
    E_UNABLE_TO_ACCESS_OUTLOOK_ACCOUNTS = &H80042213
    E_OBTAINING_ACCOUNT_INFO = &H80042214
    E_ENUMERATING_PROPERTIES = &H80042215
    E_UNSUPPORTED = &H80042216
    E_WRITING_PROPERTY = &H80042217
    E_FAILED_TO_WRAP_OBJECT = &H80042218
    E_UNRESOLVABLE_RECIPIENTS = &H80042219
    E_OBJECT_MOVED_OR_DELETED = &H8004221A
    E_DLL_ERROR = &H8004221B
    E_DELETED_RECIPIENT = &H80042220
    E_OLK_ACCOUNT_MANAGER = &H80042221
    E_INVALID_INDEX = &H80042222
    E_INVALID_ADDRESSENTRY = &H80042223
    E_INVALID_ENTRYID = &H80042224
    E_PROP_NOT_SET = &H80042225
    E_INVALID_SIGNATURENAME = &H80042226
    E_WRITING_STREAM = &H80042227
    E_INVALID_PROFILESERVICE = &H80042228
    E_INVALID_PROFILESERVICEPROVIDER = &H80042229
    E_INVALID_PROFILESECTION = &H8004222A
    E_DELETED_ATTACHMENT = &H8004222B
    E_INVALID_ADDRESSLIST = &H8004222C
    E_INVALID_ADDRESSBOOK = &H8004222D
    E_NO_DEFAULT_PROFILE = &H8004222E
    E_MAPI_DLL_BITNESS = &H8004222F
    E_SYNC_MANAGER_INUSE = &H80042230
    E_OLK_SYNC_MANAGER = &H80042231
    E_RESTRICTION_PARSER = &H80042231
    E_OLK_MIME_CONVERTER = &H80042232
    MAPI_E_CALL_FAILED = &H80004005
    MAPI_E_NOT_ENOUGH_MEMORY = &H8007000E
    MAPI_E_INVALID_PARAMETER = &H80070057
    MAPI_E_INTERFACE_NOT_SUPPORTED = &H80004002
    MAPI_E_NO_ACCESS = &H80070005
    MAPI_E_NO_SUPPORT = &H80040102
    MAPI_E_BAD_CHARWIDTH = &H80040103
    MAPI_E_STRING_TOO_LONG = &H80040105
    MAPI_E_UNKNOWN_FLAGS = &H80040106
    MAPI_E_INVALID_ENTRYID = &H80040107
    MAPI_E_INVALID_OBJECT = &H80040108
    MAPI_E_OBJECT_CHANGED = &H80040109
    MAPI_E_OBJECT_DELETED = &H8004010A
    MAPI_E_BUSY = &H8004010B
    MAPI_E_NOT_ENOUGH_DISK = &H8004010D
    MAPI_E_NOT_ENOUGH_RESOURCES = &H8004010E
    MAPI_E_NOT_FOUND = &H8004010F
    MAPI_E_VERSION = &H80040110
    MAPI_E_LOGON_FAILED = &H80040111
    MAPI_E_SESSION_LIMIT = &H80040112
    MAPI_E_USER_CANCEL = &H80040113
    MAPI_E_UNABLE_TO_ABORT = &H80040114
    MAPI_E_NETWORK_ERROR = &H80040115
    MAPI_E_DISK_ERROR = &H80040116
    MAPI_E_TOO_COMPLEX = &H80040117
    MAPI_E_BAD_COLUMN = &H80040118
    MAPI_E_COMPUTED = &H8004011A
    MAPI_E_CORRUPT_DATA = &H8004011B
    MAPI_E_UNCONFIGURED = &H8004011C
    MAPI_E_FAILONEPROVIDER = &H8004011D
    MAPI_E_UNKNOWN_CPID = &H8004011E
    MAPI_E_UNKNOWN_LCID = &H8004011F
    MAPI_E_PASSWORD_CHANGE_REQUIRED = &H80040120
    MAPI_E_PASSWORD_EXPIRED = &H80040121
    MAPI_E_INVALID_WORKSTATION_ACCOUNT = &H80040122
    MAPI_E_INVALID_ACCESS_TIME = &H80040123
    MAPI_E_ACCOUNT_DISABLED = &H80040124
    MAPI_E_END_OF_SESSION = &H80040200
    MAPI_E_UNKNOWN_ENTRYID = &H80040201
    MAPI_E_MISSING_REQUIRED_COLUMN = &H80040202
    MAPI_W_NO_SERVICE = &H40203
    MAPI_E_BAD_VALUE = &H80040301
    MAPI_E_INVALID_TYPE = &H80040302
    MAPI_E_TYPE_NO_SUPPORT = &H80040303
    MAPI_E_UNEXPECTED_TYPE = &H80040304
    MAPI_E_TOO_BIG = &H80040305
    MAPI_E_DECLINE_COPY = &H80040306
    MAPI_E_UNEXPECTED_ID = &H80040307
    MAPI_W_ERRORS_RETURNED = &H40380
    MAPI_E_UNABLE_TO_COMPLETE = &H80040400
    MAPI_E_TIMEOUT = &H80040401
    MAPI_E_TABLE_EMPTY = &H80040402
    MAPI_E_TABLE_TOO_BIG = &H80040403
    MAPI_E_INVALID_BOOKMARK = &H80040405
    MAPI_W_POSITION_CHANGED = &H40481
    MAPI_W_APPROX_COUNT = &H40482
    MAPI_E_WAIT = &H80040500
    MAPI_E_CANCEL = &H80040501
    MAPI_E_NOT_ME = &H80040502
    MAPI_W_CANCEL_MESSAGE = &H40580
    MAPI_E_CORRUPT_STORE = &H80040600
    MAPI_E_NOT_IN_QUEUE = &H80040601
    MAPI_E_NO_SUPPRESS = &H80040602
    MAPI_E_COLLISION = &H80040604
    MAPI_E_NOT_INITIALIZED = &H80040605
    MAPI_E_NON_STANDARD = &H80040606
    MAPI_E_NO_RECIPIENTS = &H80040607
    MAPI_E_SUBMITTED = &H80040608
    MAPI_E_HAS_FOLDERS = &H80040609
    MAPI_E_HAS_MESSAGES = &H8004060A
    MAPI_E_FOLDER_CYCLE = &H8004060B
    MAPI_W_PARTIAL_COMPLETION = &H40680
    MAPI_E_AMBIGUOUS_RECIP = &H80040700
    STG_E_INVALIDFUNCTION = &H80030001
    STG_E_FILENOTFOUND = &H80030002
    STG_E_PATHNOTFOUND = &H80030003
    STG_E_TOOMANYOPENFILES = &H80030004
    STG_E_ACCESSDENIED = &H80030005
    STG_E_INVALIDHANDLE = &H80030006
    STG_E_INSUFFICIENTMEMORY = &H80030008
    STG_E_INVALIDPOINTER = &H80030009
    STG_E_NOMOREFILES = &H80030012
    STG_E_DISKISWRITEPROTECTED = &H80030013
    STG_E_SEEKERROR = &H80030019
    STG_E_WRITEFAULT = &H8003001D
    STG_E_READFAULT = &H8003001E
    STG_E_SHAREVIOLATION = &H80030020
    STG_E_LOCKVIOLATION = &H80030021
    STG_E_FILEALREADYEXISTS = &H80030050
    STG_E_INVALIDPARAMETER = &H80030057
    STG_E_MEDIUMFULL = &H80030070
    STG_E_ABNORMALAPIEXIT = &H800300FA
    STG_E_INVALIDHEADER = &H800300FB
    STG_E_INVALIDNAME = &H800300FC
    STG_E_UNKNOWN = &H800300FD
    STG_E_UNIMPLEMENTEDFUNCTION = &H800300FE
    STG_E_INVALIDFLAG = &H800300FF
    STG_E_INUSE = &H80030100
    STG_E_NOTCURRENT = &H80030101
    STG_E_REVERTED = &H80030102
    STG_E_CANTSAVE = &H80030103
    STG_E_OLDFORMAT = &H80030104
    STG_E_OLDDLL = &H80030105
    STG_E_SHAREREQUIRED = &H80030106
    STG_E_NOTFILEBASEDSTORAGE = &H80030107
    STG_E_EXTANTMARSHALLINGS = &H80030108
    STG_E_DOCFILECORRUPT = &H80030109
    STG_E_BADBASEADDRESS = &H80030110
    STG_E_DOCFILETOOLARGE = &H80030111
    STG_E_NOTSIMPLEFORMAT = &H80030112
    STG_E_INCOMPLETE = &H80030201
    STG_E_TERMINATED = &H80030202
    MAPI_E_EXTENDED_ERROR = &H80040119
    E_CLASSNOTREG = &H80040154
    CO_E_CLASSSTRING = &H800401F3
    E_ACCT_NOT_FOUND = &H800C8101
    E_ACCT_UI_BUSY = &H800C8102
    E_ACCT_WRONG_SORT_ORDER = &H800C8105
    E_OLK_PARAM_NOT_SUPPORTED = &H800C8003
    E_OLK_PROP_READ_ONLY = &H800C800D
End Enum

Enum EnumCreateStoreFlags
    CreateStoreFlags_USE_ADMIN_PRIVILEGE = &H1
    CreateStoreFlags_PUBLIC = &H2
    CreateStoreFlags_HOME_LOGON = &H4
    CreateStoreFlags_TAKE_OWNERSHIP = &H8
    CreateStoreFlags_OVERRIDE_HOME_MDB = &H10
    CreateStoreFlags_TRANSPORT = &H20
    CreateStoreFlags_REMOTE_TRANSPORT = &H40
    CreateStoreFlags_INTERNET_ANONYMOUS = &H80
    CreateStoreFlags_ALTERNATE_SERVER = &H100
    CreateStoreFlags_IGNORE_HOME_MDB = &H200
    CreateStoreFlags_NO_MAIL = &H400
    CreateStoreFlags_OVERRIDE_LAST_MODIFIER = &H800
    CreateStoreFlags_CALLBACK_LOGON = &H1000
    CreateStoreFlags_LOCAL = &H2000
    CreateStoreFlags_FAIL_IF_NO_MAILBOX = &H4000
    CreateStoreFlags_CACHE_EXCHANGE = &H8000
    CreateStoreFlags_CLI_WITH_NAMEDPROP_FIX = &H10000
    CreateStoreFlags_ENABLE_LAZY_LOGGING = &H20000
    CreateStoreFlags_CLI_WITH_REPLID_GUID_MAPPING_FIX = &H40000
    CreateStoreFlags_NO_LOCALIZATION = &H80000
    CreateStoreFlags_RESTORE_DATABASE = &H100000
    CreateStoreFlags_XFOREST_MOVE = &H0
End Enum

Enum EnumOpenStoreFlags
    OpenStoreFlags_MAPI_BEST_ACCESS = &H10
    OpenStoreFlags_MAPI_DEFERRED_ERRORS = &H8
    OpenStoreFlags_MDB_NO_DIALOG = &H1
    OpenStoreFlags_MDB_NO_MAIL = &H80
    OpenStoreFlags_MDB_ONLINE = &H100
    OpenStoreFlags_MDB_TEMPORARY = &H20
    OpenStoreFlags_MDB_WRITE = &H0
End Enum

Enum EnumAddressbookOpenFlags
    AddressBook_MAPI_BEST_ACCESS = &H10
    AddressBook_MAPI_CACHE_ONLY = &H4000
    AddressBook_MAPI_DEFERRED_ERRORS = &H8
    AddressBook_MAPI_MODIFY = &H1
    AddressBook_MAPI_NO_CACHE = &H0
End Enum

Enum EnumPSTFileFormat
    PSTFileFormat_Default = &H0
    PSTFileFormat_Ansi = &H1
    PSTFileFormat_Unicode = &H2
End Enum

Enum EnumROHAuthenticateMode
    ROHAuthenticate_Basic = &H1
    ROHAuthenticate_NTLM = &H2
End Enum

Event OnModelessFormEvent(ByVal Command As String, ByVal Item As vbMAPI_MailItem)
Sub LogOn(Optional ByVal ProfileName As String, Optional ByVal Password As String, Optional ByVal ShowDialog As Boolean, Optional ByVal NewSession As Boolean = True, Optional ByVal ParentWindowHandle As Long, Optional ByVal AllowOthers As Boolean): End Sub
Attribute LogOn.VB_UserMemId = -101
Sub LogOnPSTStore(ByVal FilePath As String, Optional ByVal Format As EnumPSTFileFormat, Optional ByVal DisplayName As String, Optional ByVal Password As String, Optional ByVal Encryption As Long): End Sub
Attribute LogOnPSTStore.VB_UserMemId = -102
Sub LogOnExchangeMailbox(ByVal User As String, ByVal Server As String, Optional ByVal EncryptConnection As Boolean, Optional ByVal IgnoreNoPF As Boolean): End Sub
Attribute LogOnExchangeMailbox.VB_UserMemId = -103
Property Get Stores() As vbMAPI_Stores: End Property
Attribute Stores.VB_UserMemId = -104
Property Get Accounts() As Object: End Property
Attribute Accounts.VB_UserMemId = -105
Property Get Options() As vbMAPI_Options: End Property
Attribute Options.VB_UserMemId = -106
Property Get ProfileName() As String: End Property
Attribute ProfileName.VB_UserMemId = -107
Function GetDefaultFolder(ByVal DefaultFolder As EnumDefaultFolderType) As vbMAPI_Folder: End Function
Attribute GetDefaultFolder.VB_UserMemId = -108
Property Get CurrentUser() As Object: End Property
Attribute CurrentUser.VB_UserMemId = -109
Sub OutlookSendReceiveAll(): End Sub
Attribute OutlookSendReceiveAll.VB_UserMemId = -110
Function GetItemFromID(ByVal EntryID As String, Optional ByVal EntryIDStore As String) As vbMAPI_MailItem: End Function
Attribute GetItemFromID.VB_UserMemId = -111
Function GetFolderFromID(ByVal EntryID As String, Optional ByVal EntryIDStore As String) As vbMAPI_Folder: End Function
Attribute GetFolderFromID.VB_UserMemId = -112
Function GetStoreFromID(ByVal EntryID As String, Optional ByVal OpenFlags As EnumOpenStoreFlags = OpenStoreFlags_MAPI_BEST_ACCESS) As vbMAPI_Store: End Function
Attribute GetStoreFromID.VB_UserMemId = -113
Function CompareEntryIDs(ByVal EntryID1 As String, ByVal EntryID2 As String) As Boolean: End Function
Attribute CompareEntryIDs.VB_UserMemId = -114
Function GetMessageFromMsgFile(ByVal FilePath As String, Optional ByVal blnCreateNew As Boolean) As vbMAPI_MailItem: End Function
Attribute GetMessageFromMsgFile.VB_UserMemId = -115
Property Get Profiles() As Object: End Property
Attribute Profiles.VB_UserMemId = -116
Property Get OutlookSignatures() As Object: End Property
Attribute OutlookSignatures.VB_UserMemId = -117
Function GetSharedMailbox(ByVal NameOrAddressOrEntryIDOrObject As Variant, Optional ByVal EntryIDFlags As EnumCreateStoreFlags = &HC, Optional ByVal OpenStoreFlags As EnumOpenStoreFlags = &HB1) As vbMAPI_Store: End Function
Attribute GetSharedMailbox.VB_UserMemId = -118
Function GetSharedDefaultFolder(ByVal NameOrAddressOrEntryIDOrObject As Variant, ByVal DefaultFolder As EnumDefaultFolderType) As vbMAPI_Folder: End Function
Attribute GetSharedDefaultFolder.VB_UserMemId = -119
Property Get MAPIOBJECT() As Variant: End Property
Attribute MAPIOBJECT.VB_UserMemId = -120
Property Let MAPIOBJECT(ByVal Value As Variant): End Property
Attribute MAPIOBJECT.VB_UserMemId = -120
Sub DeliverNow(): End Sub
Attribute DeliverNow.VB_UserMemId = -121
Property Get AddressBook() As Object: End Property
Attribute AddressBook.VB_UserMemId = -122
