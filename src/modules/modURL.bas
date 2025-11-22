Attribute VB_Name = "modURL"
Option Explicit
Private Const GWL_WNDPROC = (-4)
Private Const WM_USER = &H400
Private Const WM_NOTIFY = &H4E
Private Const EM_GETEVENTMASK = WM_USER + 59
Private Const EM_GETTEXTRANGE = WM_USER + 75
Private Const EM_AUTOURLDETECT = (WM_USER + 91)
Private Const EM_SETEVENTMASK = WM_USER + 69
Private Const EN_LINK = &H70B
Private Const ENM_LINK = &H4000000
Private Const SW_SHOWNORMAL = 1

Private Type tagNMHDR
    hwndFrom As Long
    idFrom   As Long
    code     As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    NMHDR  As tagNMHDR
    Msg    As Long
    wParam As Long
    lParam As Long
    chrg   As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg      As CHARRANGE
    lpstrText As Long
End Type

Private glnglpOriginalWndProc As Long
Private glngOriginalhWnd As Long

Public Function RichTextBoxSubProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim udtNMHDR               As tagNMHDR
    Dim udtENLINK              As ENLINK
    Dim udtTEXTRANGE           As TEXTRANGE
    Dim strBuffer              As String * 128
    Dim strOperation           As String
    Dim strFileName            As String
    Dim strDefaultDirectory    As String
    Dim lngHInstanceExecutable As Long
    Dim lngWin32apiResultCode  As Long
    Dim hWndNew                As Long
    hWndNew = hWnd

    If uMsg = WM_NOTIFY Then
        RtlMoveMemory udtNMHDR, ByVal lParam, Len(udtNMHDR)
        If udtNMHDR.hwndFrom = frmMain.Chat_Output.hWnd And udtNMHDR.code = EN_LINK Then
            RtlMoveMemory udtENLINK, ByVal lParam, Len(udtENLINK)
            If udtENLINK.Msg = WM_LBUTTONDOWN Then
                strBuffer = ""
                With udtTEXTRANGE
                    .chrg.cpMin = udtENLINK.chrg.cpMin
                    .chrg.cpMax = udtENLINK.chrg.cpMax
                    .lpstrText = StrPtr(strBuffer)
                End With
                With frmMain.Chat_Output
                    lngWin32apiResultCode = SendMessage(.hWnd, EM_GETTEXTRANGE, 0, udtTEXTRANGE)
                End With
                RtlMoveMemory ByVal strBuffer, ByVal udtTEXTRANGE.lpstrText, Len(strBuffer)
                strOperation = "open"
                strFileName = strBuffer
                lngHInstanceExecutable = ShellExecute(frmMain.hWnd, strOperation, strFileName, vbNullString, strDefaultDirectory, SW_SHOWNORMAL)
                hWndNew = frmMain.hWnd
            End If
        ElseIf udtNMHDR.hwndFrom = frmMain.rtbWhisper.hWnd And udtNMHDR.code = EN_LINK Then
            RtlMoveMemory udtENLINK, ByVal lParam, Len(udtENLINK)
            If udtENLINK.Msg = WM_LBUTTONDOWN Then
                strBuffer = ""
                With udtTEXTRANGE
                    .chrg.cpMin = udtENLINK.chrg.cpMin
                    .chrg.cpMax = udtENLINK.chrg.cpMax
                    .lpstrText = StrPtr(strBuffer)
                End With
                With frmMain.rtbWhisper
                    lngWin32apiResultCode = SendMessage(.hWnd, EM_GETTEXTRANGE, 0, udtTEXTRANGE)
                End With
                RtlMoveMemory ByVal strBuffer, ByVal udtTEXTRANGE.lpstrText, Len(strBuffer)
                strOperation = "open"
                strFileName = strBuffer
                lngHInstanceExecutable = ShellExecute(frmMain.hWnd, strOperation, strFileName, vbNullString, strDefaultDirectory, SW_SHOWNORMAL)
                hWndNew = frmMain.hWnd
            End If
        ElseIf udtNMHDR.hwndFrom = frmProfile.rtbDescription.hWnd And udtNMHDR.code = EN_LINK Then
            RtlMoveMemory udtENLINK, ByVal lParam, Len(udtENLINK)
            If udtENLINK.Msg = WM_LBUTTONDOWN Then
                strBuffer = ""
                With udtTEXTRANGE
                    .chrg.cpMin = udtENLINK.chrg.cpMin
                    .chrg.cpMax = udtENLINK.chrg.cpMax
                    .lpstrText = StrPtr(strBuffer)
                End With
                With frmProfile.rtbDescription
                    lngWin32apiResultCode = SendMessage(.hWnd, EM_GETTEXTRANGE, 0, udtTEXTRANGE)
                End With
                RtlMoveMemory ByVal strBuffer, ByVal udtTEXTRANGE.lpstrText, Len(strBuffer)
                strOperation = "open"
                strFileName = strBuffer
                lngHInstanceExecutable = ShellExecute(frmProfile.hWnd, strOperation, strFileName, vbNullString, strDefaultDirectory, SW_SHOWNORMAL)
                hWndNew = frmProfile.hWnd
            End If
        End If
    End If
    RichTextBoxSubProc = CallWindowProc(glnglpOriginalWndProc, hWndNew, uMsg, wParam, lParam)
End Function

Public Sub EnableURLDetect()
    Dim lngEventMask As Long
    Dim lngWin32apiResultCode As Long
    With frmMain.Chat_Output
        lngEventMask = SendMessage(.hWnd, EM_GETEVENTMASK, 0, ByVal CLng(0))
        If lngEventMask Xor ENM_LINK Then
            lngEventMask = lngEventMask Or ENM_LINK
        End If
        lngWin32apiResultCode = SendMessage(.hWnd, EM_SETEVENTMASK, 0, ByVal CLng(lngEventMask))
        lngWin32apiResultCode = SendMessage(.hWnd, EM_AUTOURLDETECT, CLng(1), ByVal CLng(0))
    End With
    With frmMain.rtbWhisper
        lngEventMask = SendMessage(.hWnd, EM_GETEVENTMASK, 0, ByVal CLng(0))
        If lngEventMask Xor ENM_LINK Then
            lngEventMask = lngEventMask Or ENM_LINK
        End If
        lngWin32apiResultCode = SendMessage(.hWnd, EM_SETEVENTMASK, 0, ByVal CLng(lngEventMask))
        lngWin32apiResultCode = SendMessage(.hWnd, EM_AUTOURLDETECT, CLng(1), ByVal CLng(0))
    End With
    With frmProfile.rtbDescription
        lngEventMask = SendMessage(.hWnd, EM_GETEVENTMASK, 0, ByVal CLng(0))
        If lngEventMask Xor ENM_LINK Then
            lngEventMask = lngEventMask Or ENM_LINK
        End If
        lngWin32apiResultCode = SendMessage(.hWnd, EM_SETEVENTMASK, 0, ByVal CLng(lngEventMask))
        lngWin32apiResultCode = SendMessage(.hWnd, EM_AUTOURLDETECT, CLng(1), ByVal CLng(0))
    End With
    glngOriginalhWnd = frmMain.hWnd
    glnglpOriginalWndProc = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, AddressOf RichTextBoxSubProc)
End Sub

Public Sub DisableURLDetect()
    Dim lngWin32apiResultCode As Long
    lngWin32apiResultCode = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, glnglpOriginalWndProc)
End Sub
