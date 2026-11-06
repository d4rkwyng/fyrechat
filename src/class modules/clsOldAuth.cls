VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOldAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
' clsOldAuth
'
' Local (in-process) computation of the classic Battle.net "OLS" password
' hashes - the Broken SHA-1 scheme used by STAR/SEXP/D2DV/D2XP/W2BN-era
' products (LogonType 0x00 in SID_AUTH_INFO) - via bncsutil.dll's
' hashPassword()/doubleHashPassword() exports. See src/bncsutil.dll and
' NOTICE.md at the repo root.
'
' This replaces the historical round trip through a BNLS server
' (BNLS_HASHDATA, see modBNLS.bas's rcvBNLSHash) for these four call sites
' only, so the plaintext account password no longer has to leave the
' machine for old-auth logins. NLS (WAR3/W3XP) logon, CD-key hashing, and
' Warcraft III server-signature verification are untouched and still use
' BNLS - see modBNCS.bas/modBNLS.bas.
'
' --- Why not a plain Declare? ---
' bncsutil exports hashPassword/doubleHashPassword as __cdecl (see its
' src/bncsutil/mutil.h MEXP macro: "__declspec(dllimport) type __cdecl").
' VB6's native `Declare Function ... Lib` calling mechanism only knows how
' to call __stdcall functions: it assumes the CALLEE pops its own
' arguments off the stack before returning. A __cdecl callee does not do
' that (the CALLER is responsible for cleaning the stack), so calling a
' __cdecl export through a plain Declare leaves stale argument bytes on
' the stack after every call and corrupts VB6's own stack frame - this can
' run for a while before it visibly crashes, which makes it an easy trap
' to miss while testing.
'
' (Note: bncsutil's own vb6_example/BNCSutil.bas declares these two
' functions with a plain `Declare Sub ... Lib "bncsutil.dll"`, which is the
' broken pattern described above - that upstream example should not be
' trusted for these two functions specifically.)
'
' The standard, widely-documented workaround is to resolve the export by
' address ourselves (LoadLibrary/GetProcAddress) and invoke it through OLE
' Automation's DispCallFunc, which - unlike a Declare - lets the caller
' specify the calling convention explicitly (CC_CDECL below).
'
' This class is deliberately narrow: it only knows how to call these two
' specific bncsutil exports, not a general-purpose FFI layer.
'
' *** UNVERIFIED - written without a VB6 compiler/runtime available.
' *** See the commit message / PR description for what specifically to
' *** check when building this for the first time.
'------------------------------------------------------------------------------------

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numBytes As Long)

' VarPtr is not a documented VB6 keyword, but resolving it this way (any
' ByRef parameter typed "As Any", aliased straight into msvbvm60's own
' export of the same name) is the standard, widely-used technique for
' getting it back in the IDE - the native-code compiler recognizes VarPtr
' as an intrinsic on its own once compiled, this Declare only matters for
' running/debugging inside the VB6 IDE.
' *** twinBASIC has VarPtr as a true built-in keyword - if building under
' *** twinBASIC and this Declare conflicts with that, delete this line;
' *** the intrinsic should be used instead.
Private Declare Function VarPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr As Any) As Long

' OLE Automation's generic function invoker. Unlike Declare, this lets us
' pick the calling convention (CC_CDECL) instead of always assuming
' __stdcall. pvInstance=0 tells it there is no COM object/vtable involved
' and that oVft is simply the address of the function to call directly.
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal CallConv As Long, _
    ByVal retTYP As Integer, _
    ByVal paCNT As Long, _
    ByRef paTypes As Integer, _
    ByRef paValues As Long, _
    ByRef retVAR As Variant) As Long

' CALLCONV enum (oleauto.h) - only the one we need.
Private Const CC_CDECL As Long = 1

' VARENUM (oleauto.h) - only the ones we need.
Private Const VT_EMPTY As Integer = 0   ' used for retTYP: DispCallFunc's documented way to say "void"
Private Const VT_I4 As Integer = 3      ' 32-bit value - used both for Longs and for raw pointers

Private Const DLL_NAME As String = "bncsutil.dll"
Private Const HASH_LEN As Long = 20     ' Broken SHA-1 digest size (see bncsutil's src/bncsutil/bsha1.h calcHashBuf)

Private hDll As Long
Private pHashPassword As Long
Private pDoubleHashPassword As Long

Private Sub Class_Initialize()
On Error Resume Next
    hDll = LoadLibrary(DLL_NAME)
    If hDll <> 0 Then
        pHashPassword = GetProcAddress(hDll, "hashPassword")
        pDoubleHashPassword = GetProcAddress(hDll, "doubleHashPassword")
    End If
On Error GoTo 0
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    If hDll <> 0 Then
        FreeLibrary hDll
        hDll = 0
    End If
On Error GoTo 0
End Sub

' True once bncsutil.dll (src/bncsutil.dll, next to FyreChat.exe) is loaded
' and both exports were resolved. Callers should check this before relying
' on HashPassword/DoubleHashPassword and fail the logon visibly (AddChat +
' Disconnect) rather than let a Declare-style silent failure through.
Public Function IsAvailable() As Boolean
    IsAvailable = (hDll <> 0) And (pHashPassword <> 0) And (pDoubleHashPassword <> 0)
End Function

' Single Broken SHA-1 hash of the password - used for account creation
' (formerly BNLS_HASHDATA HType=2).
' Returns the raw 20-byte hash as a String, one byte per character - same
' convention clsBuffer/clsSBuffer already use for binary data (see their
' MakeByte/MakeDWORD), NOT null-terminated text. Returns vbNullString on
' failure (check IsAvailable first).
Public Function HashPassword(ByVal password As String) As String
On Error GoTo HashFailed
    If Not IsAvailable() Then Exit Function

    Dim pwBytes() As Byte: pwBytes = ToAnsiZ(password)
    Dim outBytes(0 To HASH_LEN - 1) As Byte

    Dim paTypes(0 To 1) As Integer
    Dim paValues(0 To 1) As Variant
    paTypes(0) = VT_I4: paValues(0) = VarPtr(pwBytes(0))   ' const char* password
    paTypes(1) = VT_I4: paValues(1) = VarPtr(outBytes(0))  ' char* outBuffer (>= 20 bytes)

    If Not CallCdeclVoid(pHashPassword, 2, paTypes, paValues) Then Exit Function

    HashPassword = BytesToRawString(outBytes)
    Exit Function
HashFailed:
    HashPassword = vbNullString
End Function

' Double Broken SHA-1 hash of the password with the client/server tokens
' mixed in - used for logon (formerly BNLS_HASHDATA HType=1):
'   hash1 = BrokenSHA1(password)
'   result = BrokenSHA1(clientToken(4 bytes LE) & serverToken(4 bytes LE) & hash1)
' (bncsutil does both steps internally; see its src/bncsutil/oldauth.cpp.)
' Returns the raw 20-byte hash as a String, same convention as
' HashPassword above. Returns vbNullString on failure (check IsAvailable
' first).
Public Function DoubleHashPassword(ByVal password As String, ByVal clientToken As Long, ByVal serverToken As Long) As String
On Error GoTo HashFailed
    If Not IsAvailable() Then Exit Function

    Dim pwBytes() As Byte: pwBytes = ToAnsiZ(password)
    Dim outBytes(0 To HASH_LEN - 1) As Byte

    Dim paTypes(0 To 3) As Integer
    Dim paValues(0 To 3) As Variant
    paTypes(0) = VT_I4: paValues(0) = VarPtr(pwBytes(0))    ' const char* password
    paTypes(1) = VT_I4: paValues(1) = clientToken           ' uint32_t clientToken
    paTypes(2) = VT_I4: paValues(2) = serverToken           ' uint32_t serverToken
    paTypes(3) = VT_I4: paValues(3) = VarPtr(outBytes(0))   ' char* outBuffer (>= 20 bytes)

    If Not CallCdeclVoid(pDoubleHashPassword, 4, paTypes, paValues) Then Exit Function

    DoubleHashPassword = BytesToRawString(outBytes)
    Exit Function
HashFailed:
    DoubleHashPassword = vbNullString
End Function

' Invokes a cdecl void(...) function through DispCallFunc.
' paValues(i) must already hold the actual argument in a Variant tagged
' with the matching VarType in paTypes(i) - DispCallFunc marshals each one
' onto the call according to that declared type, it does not use whatever
' VarType() the Variant happens to carry on its own.
Private Function CallCdeclVoid(ByVal pFunc As Long, ByVal argCount As Long, paTypes() As Integer, paValues() As Variant) As Boolean
On Error GoTo Failed
    Dim paPtrs() As Long
    ReDim paPtrs(0 To argCount - 1)

    Dim i As Long
    For i = 0 To argCount - 1
        paPtrs(i) = VarPtr(paValues(i))
    Next i

    Dim vRet As Variant
    CallCdeclVoid = (DispCallFunc(0, pFunc, CC_CDECL, VT_EMPTY, argCount, paTypes(0), paPtrs(0), vRet) = 0)
    Exit Function
Failed:
    CallCdeclVoid = False
End Function

' Builds a null-terminated single-byte-per-character buffer from a VB6
' string, byte value = character code point (0-255), matching how the rest
' of this codebase already treats BNET.password (see clsBuffer/clsSBuffer's
' MakeByte/MakeDWORD - a raw byte is carried as one character with the same
' numeric code point). Deliberately not StrConv(..., vbFromUnicode), which
' would reinterpret high code points through the system's default ANSI
' code page instead of passing them through as-is.
Private Function ToAnsiZ(ByVal s As String) As Byte()
    Dim buf() As Byte
    Dim i As Long
    ReDim buf(0 To Len(s))              ' +1 slot for the NUL terminator
    For i = 1 To Len(s)
        buf(i - 1) = Asc(Mid$(s, i, 1)) And &HFF
    Next i
    buf(Len(s)) = 0
    ToAnsiZ = buf
End Function

' Mirror image of ToAnsiZ: packs a Byte array into a String where each
' character's code point is that byte's value, via the same
' CopyMemory-into-a-fixed-length-String trick clsBuffer.MakeByte/MakeDWORD
' already use elsewhere in this codebase.
Private Function BytesToRawString(b() As Byte) As String
    Dim outStr As String * HASH_LEN
    CopyMemory ByVal outStr, b(0), HASH_LEN
    BytesToRawString = outStr
End Function
