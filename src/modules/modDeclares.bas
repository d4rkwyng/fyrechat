Attribute VB_Name = "modDeclares"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numBytes As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

'Packet 0x50 Declares
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function GetSystemDefaultLangID Lib "kernel32" () As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

' Local/System Times
Public Declare Sub GetSystemTimeAsFileTime Lib "kernel32.dll" (lpSystemTimeAsFileTime As FILETIME)
Public Type FILETIME
  dwLowDateTime       As Long
  dwHighDateTime      As Long
End Type
Public Type SYSTEMTIME
  wYear               As Integer
  wMonth              As Integer
  wDayOfWeek          As Integer
  wDay                As Integer
  wHour               As Integer
  wMinute             As Integer
  wSecond             As Integer
  wMilliseconds       As Integer
End Type

'BNLS Variables
Public Version As Long              ' (DWORD) Version
Public CheckSum As Long             ' (DWORD) Checksum
Public ClientToken As Long          ' (DWORD) Client Token
Public VerCheckStatString As String ' (STRING) Version check stat string
Public CDKeyLength As Long          ' (DWORD) CD-Key Length
Public CDKeyProdVal As Long         ' (DWORD) CD-Key Product Value
Public CDKeyPubVal As Long          ' (DWORD) CD-Key Public Value
Public CDKeyUnknown As Long              ' (DWORD) Unknown (0)

Public CRC32Table(0 To 255) As Long
Public Hash(2) As String
Public ExeInfo As String
Public CDKeyHash As String
Public CDKey2Hash As String
Public DebugMode As Integer
Public HType As Long
Public cB As Long
Public VerByte As Long
Public AuthTimeStamp As String
Public ServerSignature As String

Public SPass As Boolean
Public CPass As Boolean
Public AttemptedC As Boolean

Public ConnectAfterDL As Boolean
Public FetchNews As Boolean
Public showRealms As Boolean
Public blDownloadFile As Boolean

Public resendLREx As Boolean
Public InitialCon As Boolean


'Battle.net User Flags
Public Const BNFLAGS_BLIZZ = &H1
Public Const BNFLAGS_SYSOP = &H8
Public Const BNFLAGS_OP = &H2
Public Const BNFLAGS_SPKR = &H4
Public Const BNFLAGS_GLASSES = &H40
Public Const BNFLAGS_SQUELCH = &H20
Public Const BNFLAGS_GFOFFICIAL = &H100000
Public Const BNFLAGS_GFPLAYER = &H200000
Public Const BNFLAGS_PLUG = &H10

Public Type ColorData
    Self        As String
    Op          As String
    user        As String
    Message     As String
    Carrot      As String
    PingFlags   As String
    Error       As String
    info        As String
    BotInfo     As String
    BotError    As String
    Join        As String
    Left        As String
    Enc         As String
    WhisperFrom As String
    timestamp   As String
    Bot         As String
    News        As String
    Background  As String
End Type

'Logging Variables
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Idle Variables
Public Enum TimeFormatType
    DaysHoursMinutesSecondsMilliseconds = 0
    DaysHoursMinutesSeconds = 1
    DHMSMColonSeparated = 2
    DaysHoursMinutes = 3
End Enum

'Web List Variables
Public blWL_Enabled     As Integer
Public strWL_FTPAddress As String
Public strWL_FTPLoc     As String
Public strWL_FTPUser    As String
Public strWL_FTPPass    As String
Public intWL_UpTimer    As Long

'Config.ini Variables
Public Type BotData
    username        As String
    TrueUsername    As String
    password        As String
    CDKey           As String
    CDKey2          As String
    CDKeyOwner      As String
    BNCSServer      As String
    BNLSServer      As String
    HomeChannel     As String
    Product         As String
    Character       As String
    varCRealm       As Integer
    CurrentChan     As String
    NewPass         As String
    varARCon        As String
    varEnterLeave   As Integer
    varAwayIdle     As Integer
    varCountIdle    As Integer
    varIdle         As Integer
    Idle            As String
    IdleInt         As String
    varAutoCon      As Integer
    varLagPlug      As String
    varLog          As Integer
    WhispWin        As Integer
    varSysTray      As Integer
    varPingFlag     As Integer
    Realm           As String
    Email           As String
    MaxLength       As String
    FloodProt       As String
    varNews         As Integer
    varConfirmExit  As Integer
    varStatusUpdate As Integer
    varProdSpec     As Integer
    varUDP          As Integer
    Fonts           As String
    FSize           As String
    BGColors        As String
    URL             As String
    varDebugMode    As Integer
    varAdvDebug     As Integer
    varShowSvr     As Integer
End Type

' Server Type Variables
Public svrBNCS As String
Public svrBNLS As String
Public svrFTP As String
Public svrMCP As String
Public svrD2GS As String

'More Variables
Public BNET         As BotData
Public Color        As ColorData
Public sPB          As New clsSBuffer

Public P1           As String
Public P2           As String

' MCP values
Public rCookie      As Long
Public rStatus      As Long
Public rChunk1      As String
Public rChunk2      As String

Public StoredData   As String
Public ServerToken      As Long
Public LastTalk     As Long
Public IdleIO       As Long
Public RunningTime  As Long
Public connecttime  As Long
Public UploadIdle   As Long
Public awayFlag     As Boolean
Public lstUser      As String
Public ProfUser     As String
Public Whisper      As Boolean
Public WReply       As String
Public RWReply       As String
Public varLagPlug   As String
Public vbMax        As Boolean
Public MOverwrite   As Boolean
Public MyPing       As Long
Public MyFlags      As Long
Public MyParsedString As String
Public varRequest   As String
Public RealmName    As String
Public PathFolder   As Boolean
Public showClient   As String
Public IdleSent     As String
Public hmm          As Long
Public PType        As String
Public dots         As String
Public RaceIcon     As String
Public Urealm       As Boolean
Public ChanAdd      As Boolean
Public SendData     As Boolean
Public PVersion     As String
Public lSendDelay   As Long
Public FloodCheck   As Long
Public FloodUser    As String
Public PingFlags    As String
Public ReconCount   As Integer
Public TheVoid      As Boolean
Public Diablo2      As Boolean
Public Warcraft3    As Boolean
Public OnStart      As Boolean
Public OnChan       As Boolean
Public DelPrevent   As String
Public TString      As String
Public Talk         As String
Public OMessage     As String
Public Tusername    As String
Public BnetFileName As String
Public BnetFileTime As String
Public BnetLastFile As String
Public userlist     As String
Public userlist2    As String
Public BnetHomeChan As String
Public tFlag        As Boolean
Public Counter      As String
Public Once         As Boolean
Public ChanProt     As Boolean
Public FirstRun     As Boolean
Public WPause       As Boolean
Public IconsLoaded  As Boolean
Public tmpFlag      As String
Public blEnterChat  As Boolean ' Accepted Connection to Battle.net

' New Variables for Build 100
Public homepage     As String

'Booleans //Started for Type Organizing
Public CloseProf    As Boolean 'For Closing Profile

Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

' System Tray
 Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
 End Type
 
 Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
 Global nid As NOTIFYICONDATA
 
'System Tray Constants
Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
 
'Download
Public Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Public Const ERROR_SUCCESS As Long = 0

Public Declare Function DoFileDownload Lib "shdocvw" _
    (ByVal lpszFile As String) As Long

'AutoURL
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_MOUSEMOVE = &H200
Global Const WM_RBUTTONDBLCLK = &H206
Global Const WM_RBUTTONDOWN = &H204
Global Const WM_RBUTTONUP = &H205
Global Const WM_SETCURSOR = &H20
