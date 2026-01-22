Attribute VB_Name = "modSysteem"
Declare Function SHGetFolderPath Lib "SHFOLDER" _
      Alias "SHGetFolderPathA" _
      (ByVal HWND As Long, _
      ByVal CSIDL As Long, _
      ByVal TOKENHANDLE As Long, _
      ByVal FLAGS As Long, _
      ByVal lpPath As String) As Long

'shFolder
Global Const rvDesktop = 0
Global Const rvStartMenuProgrammas = 2
Global Const rvMijnDocumenten = 5
Global Const rvFavorieten = 6
Global Const rvOpstarten = 7
Global Const rvRecent = 8
Global Const rvVerzondenNaar = 9
Global Const rvStartMenu = 11
Global Const rvDesktop2 = 16
Global Const rvNetHood = 19
Global Const rvFonts = 20
Global Const rvShellNew = 21
Global Const rvAllUsersDesktop = 25
Global Const rvProfilesThisUserApplicationData = 26
Global Const rvPrintHood = 27
Global Const rvProfilesThisUserLocalSettingsApplicationData = 28
Global Const rvTemporaryInternetFiles = 32
Global Const rvProfilesThisUserCookies = 33
Global Const rvProfilesThisUserHistory = 34
Global Const rvProfilesAllUsersApplicationData = 35
Global Const rvWindows = 36
Global Const rvWindowsSysteem = 37
Global Const rvProgramFiles = 38
Global Const rvMijnAfbeeldingen = 39
Global Const rvCommonFiles = 43


