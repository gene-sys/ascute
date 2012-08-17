VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinInetConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'WinInet API declarations
'***************************************************************************
'Flags for InternetGetConnectedState and Ex
Private Const INTERNET_CONNECTION_MODEM = &H1
Private Const INTERNET_CONNECTION_LAN = &H2
Private Const INTERNET_CONNECTION_PROXY = &H4
Private Const INTERNET_RAS_INSTALLED = &H10
Private Const INTERNET_CONNECTION_OFFLINE = &H20
Private Const INTERNET_CONNECTION_CONFIGURED = &H40
'Flags for InternetAutodial
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1&
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2&
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4&
'Flags for InternetDial - must not conflict with InternetAutodial
'                         flags as they are valid here also.
Private Const INTERNET_DIAL_FORCE_PROMPT = &H2000
Private Const INTERNET_DIAL_SHOW_OFFLINE = &H4000
Private Const INTERNET_DIAL_UNATTENDED = &H8000
'
Private Const INTERNET_OPTION_CONNECTED_STATE = 50
Private Const INTERNET_STATE_DISCONNECTED_BY_USER = &H10
Private Const ISO_FORCE_DISCONNECTED = &H1
Private Const INTERNET_STATE_CONNECTED = &H1
'
Private Type INTERNET_CONNECTED_INFO
    dwConnectedState    As Long
    dwFlags             As Long
End Type
'
'functions
'
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, lpBuffer As Any, ByVal dwBufferLength As Long) As Long

Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal hwndParent As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Declare Function InternetDial Lib "wininet.dll" Alias "InternetDialA" (ByVal hwndParent As Long, ByVal lpszConnectoid As String, ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetGoOnline Lib "wininet.dll" Alias "InternetGoOnlineA" (ByVal lpszURL As String, ByVal hwndParent As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (lpdwFlags As Long, lpszConnectionName As Long, dwNameLen As Long, ByVal dwReserved As Long) As Long
'***************************************************************************
'Win32API declarations
'***************************************************************************
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'
Private Const ERROR_SUCCESS = 0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const LANG_USER_DEFAULT = &H400&
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
'
'********************************************************************************
'Custom data types
'********************************************************************************
Public Enum AutoDialsFlags
    ADF_FORCE_ONLINE = INTERNET_AUTODIAL_FORCE_ONLINE
    ADF_FORCE_UNATTENDED = INTERNET_AUTODIAL_FORCE_UNATTENDED
End Enum

Public Enum DialsFlags
    DF_FORCE_ONLINE = INTERNET_AUTODIAL_FORCE_ONLINE
    DF_FORCE_UNATTENDED = INTERNET_AUTODIAL_FORCE_UNATTENDED
    DF_DIAL_FORCE_PROMPT = INTERNET_DIAL_FORCE_PROMPT
    DF_DIAL_UNATTENDED = INTERNET_DIAL_UNATTENDED
End Enum
'********************************************************************************
'local variables for the class properties
'********************************************************************************
Private mvarUseModem                As Boolean
Private mvarUseLAN                  As Boolean
Private mvarUseProxy                As Boolean
Private mvarIsRasInstalled          As Boolean
Private mvarIsOffline               As Boolean
Private mvarIsConnectionConfigured  As Boolean
Private mvarConnectionName          As String
Private mvarIsConnected             As Boolean
'
Private m_lConnectionID As Long

Public Function SetGlobalOnline() As Boolean
Attribute SetGlobalOnline.VB_Description = "Turns the local system to global online mode."
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :07.10.99
'Purpose     :Turns the local system to global online mode.
'********************************************************************************
Dim ConInfo As INTERNET_CONNECTED_INFO, lRetValue As Long

On Error GoTo SetGlobalOnline_Err_Handler

ConInfo.dwConnectedState = INTERNET_STATE_CONNECTED

lRetValue = InternetSetOption(0&, INTERNET_OPTION_CONNECTED_STATE, ConInfo, Len(ConInfo))

    If lRetValue <> 0 Then
        SetGlobalOnline = True
    Else
        SetGlobalOnline = False
        Call ProcessError("SetGlobalOnline")
    End If

Exit_Label:
    Exit Function

SetGlobalOnline_Err_Handler:
    Err.Raise vbObjectError + 1000 + Err.Number, "CWinInetConnection.SetGlobalOnline", Err.Description
    GoTo Exit_Label

End Function

Public Function SetGlobalOffline() As Boolean
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :06.10.99
'Purpose     :
'Arguments   :
'********************************************************************************
Dim ConInfo As INTERNET_CONNECTED_INFO, lRetValue As Long

On Error GoTo SetGlobalOffline_Err_Handler

ConInfo.dwConnectedState = INTERNET_STATE_DISCONNECTED_BY_USER
ConInfo.dwFlags = ISO_FORCE_DISCONNECTED

lRetValue = InternetSetOption(0&, INTERNET_OPTION_CONNECTED_STATE, ConInfo, Len(ConInfo))

    If lRetValue <> 0 Then
        SetGlobalOffline = True
    Else
        SetGlobalOffline = False
        Call ProcessError("SetGlobalOffline")
    End If
    
Exit_Label:
    Exit Function

SetGlobalOffline_Err_Handler:
    Err.Raise vbObjectError + 1000 + Err.Number, "CWinInetConnection.SetGlobalOffline", Err.Description
    GoTo Exit_Label

End Function


Public Function GoOnline(strURL As String, hwndParentWindow As Long) As Boolean
Attribute GoOnline.VB_Description = "Prompts the user for permission to initiate connection to a URL."
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :06.10.99
'Purpose     :Prompts the user for permission to initiate connection to a URL.
'Arguments   :
'********************************************************************************
Dim lRetValue As Long

On Error GoTo GoOnline_Err_Handler
    '
    lRetValue = InternetGoOnline(strURL, hwndParentWindow, 0&)
    '
    If lRetValue <> 0 Then
        GoOnline = True
    Else
        GoOnline = False
        Call ProcessError("GoOnline")
    End If
    '
Exit_Label:
    Exit Function

GoOnline_Err_Handler:
    Err.Raise vbObjectError + 1000 + Err.Number, "CWinInetConnection.GoOnline", Err.Description
    GoTo Exit_Label

End Function

Public Function HangUp()
Attribute HangUp.VB_Description = "Instructs the modem to disconnect from the Internet."
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :06.10.99
'Purpose     :Instructs the modem to disconnect from the Internet.
'********************************************************************************
Dim lRetValue As Long

On Error GoTo HangUp_Err_Handler

lRetValue = InternetHangUp(m_lConnectionID, 0&)
HangUp = (lRetValue = ERROR_SUCCESS)

Exit_Label:
    Exit Function

HangUp_Err_Handler:
    Err.Raise vbObjectError + 1000 + Err.Number, "CWinInetConnection.HangUp", Err.Description
    GoTo Exit_Label

End Function

Public Function Dial(hwndParentWindow As Long, strConnectionName As String, lOption As DialsFlags, Optional bShowOfflineButton As Boolean = False) As Boolean
Attribute Dial.VB_Description = "Initiates a connection to the Internet using a modem."
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :06.10.99
'Purpose     :Initiates a connection to the Internet using a modem.
'Arguments   :
'********************************************************************************
Dim lFlags As Long, lRetValue As Long

On Error GoTo Dial_Err_Handler
    '
    If bShowOfflineButton Then
        lFlags = lOption Or INTERNET_DIAL_SHOW_OFFLINE
    Else
        lFlags = lOption
    End If
    '
    lRetValue = InternetDial(hwndParentWindow, strConnectionName, lFlags, m_lConnectionID, 0&)
    '
    If lRetValue = 0 Then
        Dial = True
    Else
        '
        Call ProcessError("Dial")
    End If
    '
Exit_Label:
    Exit Function

Dial_Err_Handler:
    Dial = False
    Err.Raise vbObjectError + 1000 + Err.Number, "Dial", Err.Description
    GoTo Exit_Label

End Function


Public Sub AutodialHangup()
Attribute AutodialHangup.VB_Description = "Disconnects an automatic dial-up connection."
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :06.10.99
'Description :Disconnects an automatic dial-up connection.
'********************************************************************************
On Error GoTo AutodialHangup_Err_Handler

    Call InternetAutodialHangup(0&)

Exit_Label:
    Exit Sub

AutodialHangup_Err_Handler:
    Err.Raise vbObjectError + 1000 + Err.Number, "CWinInetConnection.AutodialHangup", Err.Description
    GoTo Exit_Label

End Sub

Public Function Autodial(hwndParentWindow As Long, lOption As AutoDialsFlags, Optional bFailIfSecurityCheck As Boolean = True) As Boolean
Attribute Autodial.VB_Description = "Causes the modem to automatically dial the default Internet connection."
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :06.10.99
'Return      :Returns TRUE if successful, or FALSE otherwise.
'Description :Causes the modem to automatically dial the default
'             Internet connection.Causes the modem to automatically
'             dial the default Internet connection.
'********************************************************************************
Dim lFlags As Long, lRetValue As Long

On Error GoTo Autodial_Err_Handler
    '
    If bFailIfSecurityCheck Then
        lFlags = lOption Or INTERNET_AUTODIAL_FAILIFSECURITYCHECK
    End If
    '
    lRetValue = InternetAutodial(lFlags, hwndParentWindow)
    '
    If lRetVal <> 0 Then
        Autodial = True
    Else
        Autodial = False
        Call ProcessError("Autodial")
    End If
    '
Exit_Label:
    Exit Function

Autodial_Err_Handler:
'    Err.Raise vbObjectError + 1000 + Err.Number, "CWinInetConnection.Autodial", Err.Description
    GoTo Exit_Label

End Function


Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the class data."
'--------------------------------------------------------------------------------
'Author      :Oleg Gdalevich
'Date/Time   :05.10.99
'Purpose     :
'--------------------------------------------------------------------------------

Dim strConnectionName   As String
Dim lNameLen            As Long
Dim lRetVal             As Long
Dim lConnectionFlags    As Long
Dim lPtr                As Long
Dim lNameLenPtr         As Long

On Error GoTo Refresh_Err_Handler

strConnectionName = Space(256)
lNameLen = 256
lPtr = StrPtr(strConnectionName)
lNameLenPtr = VarPtr(lNameLen)

lRetVal = InternetGetConnectedStateEx(lConnectionFlags, ByVal lPtr, ByVal lNameLen, 0&)

If lRetVal <> 0 Then
    mvarIsConnected = True
Else
    mvarIsConnected = False
    Call ProcessError("Refresh")
End If

mvarUseModem = lConnectionFlags And INTERNET_CONNECTION_MODEM
mvarUseLAN = lConnectionFlags And INTERNET_CONNECTION_LAN
mvarUseProxy = lConnectionFlags And INTERNET_CONNECTION_PROXY
mvarIsRasInstalled = lConnectionFlags And INTERNET_RAS_INSTALLED
mvarIsOffline = lConnectionFlags And INTERNET_CONNECTION_OFFLINE
mvarIsConnectionConfigured = lConnectionFlags And INTERNET_CONNECTION_CONFIGURED
mvarConnectionName = StringFromPointer(lPtr)

Exit_Label:
    Exit Sub

Refresh_Err_Handler:
    Err.Raise vbObjectError + Err.Number, "CWinInetConnection.Refresh", Err.Description
    GoTo Exit_Label
    
End Sub

Public Property Get IsConnected() As Boolean
    IsConnected = mvarIsConnected
End Property

Public Property Get ConnectionName() As String
Attribute ConnectionName.VB_Description = "Name of the active connection."
    ConnectionName = mvarConnectionName
End Property

Public Property Get IsConnectionConfigured() As Boolean
Attribute IsConnectionConfigured.VB_Description = "Local system has a valid connection to the Internet, but it may or may not be currently connected."
    IsConnectionConfigured = mvarIsConnectionConfigured
End Property

Public Property Get IsOffline() As Boolean
Attribute IsOffline.VB_Description = "Local system is in offline mode."
    IsOffline = mvarIsOffline
End Property

Public Property Get IsRasInstalled() As Boolean
Attribute IsRasInstalled.VB_Description = "Local system has RAS installed."
    IsRasInstalled = mvarIsRasInstalled
End Property

Public Property Get UseProxy() As Boolean
Attribute UseProxy.VB_Description = "Local system uses a proxy server to connect to the Internet."
    UseProxy = mvarUseProxy
End Property

Public Property Get UseLAN() As Boolean
Attribute UseLAN.VB_Description = "Local system uses a local area network to connect to the Internet."
    UseLAN = mvarUseLAN
End Property

Public Property Get UseModem() As Boolean
Attribute UseModem.VB_Description = "Local system uses a modem to connect to the Internet."
    UseModem = mvarUseModem
End Property

'***************************************************************************
'Helper functions and procedures:
'       GetWininetErrorDesc
'       GetWinApiDesc
'       StringFromPointer
'       ProcessError
'***************************************************************************

Private Function GetWininetErrorDesc(lErrNumber As Long) As String

    Dim dwLength As Long
    Dim strBuffer As String * 257
    Dim hModule As Long
    Dim dError As Long
    Dim bLoadLib As Boolean
        
    hModule = GetModuleHandle("wininet.dll")
    
    If hModule = 0 Then
        hModule = LoadLibrary("wininet.dll")
        bLoadLib = True
    End If
    

    dwLength = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                             ByVal hModule, lErrNumber, 0&, _
                             ByVal strBuffer, 256&, 0&)
    
    If dwLength > 0 Then
        GetWininetErrorDesc = Left$(strBuffer, dwLength - 2)
    End If

    
    If bLoadLib Then FreeLibrary hModule

End Function

Function GetWinApiDesc(lErrNumber As Long) As String

Dim strBuffer   As String * 257
Dim dwLength    As Long

dwLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS _
           Or FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal 0&, lErrNumber, LANG_USER_DEFAULT, _
            ByVal strBuffer, 256&, 0&)

If dwLength > 0 Then
    GetWinApiDesc = Left$(strBuffer, dwLength)
End If

End Function

Private Function StringFromPointer(ByVal lPointer As Long) As String

   Dim strTemp As String
   Dim lRetVal As Long
   
   strTemp = String$(lstrlen(ByVal lPointer), 0)
   lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
   If lRetVal Then StringFromPointer = strTemp
   
End Function

Private Sub ProcessError(strProcedureName As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :06.10.99
'Description :process Win32API and WinInet API errors
'********************************************************************************
Dim lLastDllError As Long

On Error GoTo ProcessError_Err_Handler

lLastDllError = Err.LastDllError
If lLastDllError > 0 Then
    If lLastDllError > 12000 Then
        Err.Raise vbObjectError + 1000 + lLastDllError, _
                  "CWinInetConnection." & strProcedureName, _
                  "WinInet API Error: " & lLastDllError & " " & _
                  GetWininetErrorDesc(lLastDllError)
    Else
        Err.Raise vbObjectError + 1000 + lLastDllError, _
                  "CWinInetConnection." & strProcedureName, _
                  "Win32API Error: " & lLastDllError & " " & _
                  GetWinApiDesc(lLastDllError)
    End If
End If

Exit_Label:
    Exit Sub

ProcessError_Err_Handler:
    Err.Raise vbObjectError + 1000 + Err.Number, "CWinInetConnection.ProcessError", Err.Description
    GoTo Exit_Label

End Sub
