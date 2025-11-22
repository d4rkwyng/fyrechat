Attribute VB_Name = "modStorm"
Option Explicit
Private Declare Function SFileCloseArchive Lib "Storm.dll" Alias "#252" (ByVal hMPQ As Long) As Boolean
Private Declare Function SFileCloseFile Lib "Storm.dll" Alias "#253" (ByVal hFile As Long) As Boolean
Private Declare Function SFileDestroy Lib "Storm.dll" Alias "#262" () As Boolean
Private Declare Function SFileGetFileSize Lib "Storm.dll" Alias "#265" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function SFileOpenArchive Lib "Storm.dll" Alias "#266" (ByVal lpFileName As String, ByVal dwPriority As Long, ByVal dwFlags As Long, ByRef hMPQ As Long) As Boolean
Private Declare Function SFileOpenFileEx Lib "Storm.dll" Alias "#268" (ByVal hMPQ As Long, ByVal lpFileName As String, ByVal dwSearchScope As Long, ByRef hFile As Long) As Boolean
Private Declare Function SFileReadFile Lib "Storm.dll" Alias "#269" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBYTEsToRead As Long, ByRef lpNumberOfBYTEsRead As Long, lpOverlapped As Long) As Boolean
Private Declare Function SFileAuthenticateArchive Lib "Storm.dll" Alias "#251" (ByVal hMPQ As Long, ByRef dwAuthenticationStatus As Long) As Boolean
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Public Function ExtractWork()
Dim ExtractMsg
    ExtractMsg = modStorm.ExtractFromMPQ(BnetFileName, Left$(BnetFileName, Len(BnetFileName) - 3) & "dll")
    GetReqWork = False
    
    If BNET.varDebugMode = 1 Then AddChat Color.Carrot, "[Debug] ", Color.Message, ExtractMsg
    Call sndExtraWrk(0, "")
    Kill App.Path & "\" & BnetFileName
End Function


Public Function ExtractFromMPQ(ByVal strArchive As String, ByVal strFile As String) As String
Dim hMPQ       As Long
Dim hFile      As Long
Dim hOutput    As Long
Dim lngSize    As Long
Dim dwBytes    As Long
Dim szBuffer() As Byte
Dim Status     As Long
  On Error GoTo Erred
  SFileDestroy
  SFileOpenArchive App.Path & strArchive, 0, 0, hMPQ
  SFileOpenFileEx hMPQ, strFile, 0, hFile
  If hFile <> 0 Then
    SFileAuthenticateArchive hMPQ, Status
    If Status = &H0 Or Status = &H1 Or Status = &H5 Then
      lngSize = SFileGetFileSize(hFile, 0)
      ReDim szBuffer(0 To lngSize)
      SFileReadFile hFile, szBuffer(0), lngSize, dwBytes, ByVal &H0
      If dwBytes <> lngSize Then
        ExtractFromMPQ = "MPQ Corrupt (" & dwBytes & " / " & lngSize & ")"
        GoTo Failed
      End If
      If LenB(Dir$(App.Path & strFile)) > 0 Then Kill App.Path & strFile
      hOutput = CreateFileA(App.Path & strFile, &H40000000, &H2, ByVal 0&, &H2, ByVal 0&, ByVal 0&)
      WriteFile hOutput, szBuffer(0), lngSize, dwBytes, ByVal &H0
      CloseHandle hOutput
      If dwBytes <> lngSize Then
        ExtractFromMPQ = "DLL Corrupt (" & dwBytes & " / " & lngSize & ")"
        GoTo Failed
      End If
      SFileCloseFile hFile
      SFileCloseArchive hMPQ
      SetFileTimeInfo App.Path & strFile, GetFileTimeInfo(App.Path & strArchive), GetFileTimeInfo(App.Path & strArchive), GetFileTimeInfo(App.Path & strArchive)
      ExtractFromMPQ = "OK"
    Else
      ExtractFromMPQ = "MPQ not authentic"
    End If
  Else
    ExtractFromMPQ = "Could not find " & strFile & " in " & strArchive
  End If
  SFileDestroy
  Exit Function
Failed:
  If hFile <> 0 Then SFileCloseFile hFile
  If hMPQ <> 0 Then SFileCloseArchive hMPQ
Exit Function
Erred:
  ExtractFromMPQ = Err.Description
  GoTo Failed
End Function

Public Function GetFileTimeInfo(FileName As String) As FILETIME
Dim file_handle As Long
Dim create_time As FILETIME
Dim access_time As FILETIME
Dim write_time  As FILETIME
  On Local Error Resume Next
  If LenB(Dir$(FileName)) > 0 Then
    file_handle = CreateFileA(FileName, &H80000000, &H1 Or &H2, 0&, &H3, 0&, 0&)
    GetFileTime file_handle, create_time, access_time, write_time
    GetFileTimeInfo = write_time
    CloseHandle (file_handle)
  Else
    GetFileTimeInfo.dwHighDateTime = 0
    GetFileTimeInfo.dwLowDateTime = 0
  End If
End Function

Public Sub SetFileTimeInfo(FileName As String, Created As FILETIME, Accessed As FILETIME, Written As FILETIME)
Dim file_handle As Long
  On Local Error Resume Next
  If LenB(Dir$(FileName)) > 0 Then
    file_handle = CreateFileA(FileName, &H40000000, &H1 Or &H2, 0&, &H3, 0&, 0&)
    SetFileTime file_handle, Created, Accessed, Written
    CloseHandle (file_handle)
  End If
End Sub
