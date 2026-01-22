VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Options"
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

Enum EnumAutoResolveMode
    AutoResolve_AddressBook_With_Dialog = &H0
    AutoResolve_AddressBook_No_Dialog = &H1
    AutoResolve_OneOff = &H2
    AutoResolve_Manual = &H3
End Enum

Enum EnumWindowHandleConstants
    HWND_USEFOREGROUNDWINDOW = &HFFFFFFFF
    HWND_USEMAPIDEFAULT = &H0
End Enum

Property Get AutoResolveMode() As EnumAutoResolveMode: End Property
Attribute AutoResolveMode.VB_UserMemId = -101
Property Let AutoResolveMode(ByVal Value As EnumAutoResolveMode): End Property
Attribute AutoResolveMode.VB_UserMemId = -101
Property Get FallbackWindowHandle() As EnumWindowHandleConstants: End Property
Attribute FallbackWindowHandle.VB_UserMemId = -102
Property Let FallbackWindowHandle(ByVal Value As EnumWindowHandleConstants): End Property
Attribute FallbackWindowHandle.VB_UserMemId = -102
Property Get CollectionBase() As Long: End Property
Attribute CollectionBase.VB_UserMemId = -103
Property Let CollectionBase(ByVal Value As Long): End Property
Attribute CollectionBase.VB_UserMemId = -103
Property Get ForceInternalConversionOfEmbeddedHTML() As Boolean: End Property
Attribute ForceInternalConversionOfEmbeddedHTML.VB_UserMemId = -104
Property Let ForceInternalConversionOfEmbeddedHTML(ByVal Value As Boolean): End Property
Attribute ForceInternalConversionOfEmbeddedHTML.VB_UserMemId = -104
Property Get ParseSignatureLocalImages() As Boolean: End Property
Attribute ParseSignatureLocalImages.VB_UserMemId = -105
Property Let ParseSignatureLocalImages(ByVal Value As Boolean): End Property
Attribute ParseSignatureLocalImages.VB_UserMemId = -105
