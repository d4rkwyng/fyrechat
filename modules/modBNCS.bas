Attribute VB_Name = "modBNCS"
Option Explicit
Public AuthHash As String, MPQName As String
Private bncsPB As New clsBuffer

Private spID As String
Private cpID As String
Private UDPToken As Long

Public Sub ParseBNCS(d As String)
On Local Error Resume Next
Dim pID As Byte, msgLen As Long, Msg As String

    With bncsPB
        ' Clear and Initialize
        .Clear
        .SetBuffer d
        
        ' Protocol Header
        .Skip 1                     '(BYTE) Always 0xFF
        pID = .GetByte              '(BYTE) Message ID
        msgLen = .GetWORD           '(WORD) Message length, including this header
        Msg = Mid(.GetBuffer, 5)    '(VOID) Message Data
    End With
    
    If BNET.varAdvDebug = 1 Then DisplayAdvDebug "BNCS", d, pID
    
    'If (CallPluginMessageNotifications(0, pID, Mid$(d, 5), Len(Mid$(d, 5))) = 1) Then
    If (CallPluginMessageNotifications(0, pID, Msg, Len(Msg))) Then
        SleepEx 0, 1
        Exit Sub
    End If
    
    cpID = ""
    spID = ""
    If BNET.varDebugMode = 1 Then spID = "[S 0x" & Hex(pID) & "] "
    
    Select Case pID
        Case &H0: sPB.SendPacket &H0
        Case &H5: rcvClientID
        Case &H6: rcvStartVer
        Case &H7: rcvReportVer
        'Case &H8 ' SID_STARTADVEX
        Case &H9: rcvGameEX
        Case &HA: rcvEntChat
        Case &HB: rcvChanList
        Case &HF: frmMain.Events.DispatchMessage d
        Case &H13: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, Color.Error, "Flood Detected!"
        Case &H15: rcvAd
        Case &H19: rcvMsgBox
        Case &H1C: rcvStartAdvEx3
        Case &H1D: rcvLogChalEx
        Case &H25: rcvPing
        Case &H26: rcvUsrProf
        Case &H28: rcvLogonChal
        Case &H29: rcvLogonResp
        Case &H2A: rcvCreateAcc
        Case &H2D: rcvIconData
        'Case &H2E ' SID_GETLADDERDATA
        'Case &H2F ' SID_FINDLADDERUSER
        Case &H30: rcvCDKey
        'Case &H31 ' SID_CHANGEPASSWORD
        Case &H33: rcvGetFileTime
        'Case &H35 ' SID_PROFILE
        Case &H3A: rcvLogonResp2
        'Case &H3C ' SID_CHECKDATAFILE2
        Case &H3D: rcvCreateAcc2
        Case &H3E: rcvLogonRealmEX
        Case &H40: rcvQueryRealms2
        'Case &H43 ' SID_WARCRAFTUNKNOWN
        'Case &H41 ' SID_QUERYADURL
        'Case &H44 ' SID_WARCRAFTGENERAL
        Case &H46: rcvNews
        Case &H4A: rcvOptWrk
        Case &H4C: rcvReqWrk
        Case &H50: rcvAuthInfo
        Case &H51: rcvAuthCheck
        Case &H52: rcvAuthAccCreate
        Case &H53: rcvAuthAccLogon
        Case &H54: rcvAuthAccLogonProof
        'Case &H55 ' SID_AUTH_ACCOUNTCHANGE
        'Case &H56 ' SID_AUTH_ACCOUNTCHANGEPROOF
        Case &H59: sndSetEmail
        Case &H5E: rcvWarden
        'Case &H60 ' SID_GAMEPLAYERSEARCH
        Case &H70: rcvClanFindCan
        Case &H71: rcvClanInvMul
        Case &H72: rcvClanCrInv
        Case &H73: rcvClanDisband
        Case &H74: rcvClanMakeChief
        Case &H75: rcvClanInfo
        Case &H76: rcvClanQuit
        Case &H77: rcvClanInv
        Case &H78: rcvClanRemMem
        Case &H79: rcvClanInvRes
        Case &H7A: rcvClanRankChg
        Case &H7C: rcvClanMOTD
        Case &H7D: rcvClanMemList
        Case &H7E: rcvClanMemRem
        Case &H7F: rcvClanMemStatChg
        Case &H81: rcvClanMemRankChg
        Case &H82: rcvClanMemInfo
        Case Else
            If BNET.varDebugMode Then
                AddChat Color.Carrot, "[Debug] Unhandled Packet ", Color.Bot, spID, Color.Message, DebugOutput(Mid(d, 5))
            End If
    End Select
End Sub

'
'
' Server to Client Sub Routines
'
'

' SID_CLIENTID (0x05)
Private Sub rcvClientID()
Dim regVer As Long, regAuth As Long, accNum As Long, regTok As Long
    ' Values no longer used, server returns &H0
    With bncsPB
        regVer = .GetDWORD  ' (DWORD) Registration Version
        regAuth = .GetDWORD ' (DWORD) Registration Authority
        accNum = .GetDWORD  ' (DWORD) Account Number
        regTok = .GetDWORD  ' (DWORD) Registration Token
    End With
End Sub

' SID_STARTVERSIONING (0x06)
Private Sub rcvStartVer()
Dim timestamp As String
    With bncsPB
        timestamp = .GetFileTime   ' (FILETIME) MPQ Filetime
        MPQName = .GetString       ' (STRING) MPQ Filename
        AuthHash = .GetString      ' (STRING) ValueString
    End With

    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Received versioning info!"
    
    'Debug
    If BNET.varDebugMode Then
        AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.BotInfo, "MPQName: ", Color.Message, MPQName
    End If
    
    ' BNLS Version Check
    With sPB
        .InsertDWORD GetBNLSByte()                  ' (DWORD) Product ID
        .InsertDWORD &H0                            ' (DWORD) Flags (None, set to 0)
        .InsertDWORD &H1                            ' (DWORD) Cookie
        .InsertDWORD CLng(Split(timestamp, " ")(1)) ' (FILETIME) Timestamp for version check archive
        .InsertDWORD CLng(Split(timestamp, " ")(0)) '   cont.
        .InsertNTString MPQName                     ' (STRING) Version check archive filename
        .InsertNTString AuthHash                    ' (STRING) Checksum formula
        .SendBNLSPacket &H1A                        ' BNLS_VERSIONCHECKEX2
    End With
End Sub

' SID_REPORTVERSION (0x07)
Private Sub rcvReportVer()
Dim errMsg As String
    Select Case bncsPB.GetDWORD     ' (DWORD) Result
        Case &H0: errMsg = "Failed version check. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H1: errMsg = "Old game version. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H2
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Passed challenge!"
            With sPB
                .SendPacket &H2D
                If BNET.varUDP Then
                    .InsertNonNTString "bnet"
                Else
                    .InsertNonNTString "tenb"
                End If
                '.InsertDWORD UDPToken
                .SendPacket &H14
                
                If BNET.Product = "RTSJ" Then
                    sndCDKEY
                Else
                    HType = 1
                    .InsertDWORD Len(BNET.password)
                    .InsertDWORD &H0
                    .InsertNonNTString BNET.password
                    .SendBNLSPacket &HB
                    SPass = True
                End If
            End With
        Case &H3: errMsg = "Reinstall required. (" & bncsPB.GetString & ")": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
    frmMain.Disconnect
End Sub

' SID_GETADVLISTEX (0x9)
Private Sub rcvGameEX()
Dim n As Long, errMsg As String
    n = bncsPB.GetDWORD
    If n = 0 Then
        Select Case bncsPB.GetDWORD
            Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Game List OK! (No count received)"
            Case &H1: errMsg = "Game does not exist!": GoTo Failed
            Case &H2: errMsg = "Incorrect Password!": GoTo Failed
            Case &H3: errMsg = "Game is currently full!": GoTo Failed
            Case &H4: errMsg = "Game has already started!": GoTo Failed
            Case &H6: errMsg = "Too many server requests!": GoTo Failed
        End Select
    Else
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Found ", Color.Message, n, Color.BotInfo, " games!"
        'For i = 1 To N
        '    Dim gType As Integer, sgType As Integer, langID As Long
        '    Dim addFam As Integer, Port As Integer, IP As String, SZ As Long, gStatus As Long, ET As Long
        '    Dim gName As String, gPass As String, gStat As String
        '    With bncsPB
        '        gType = .GetWORD
        '        sgType = .GetWORD
        '        langID = .GetDWORD
        '        AddFam = .GetWORD
        '        Port = .GetWORD
        '        IP = .GetNonNTString
        '        SZ = .GetDWORD
        '        SZ = .GetDWORD
        '        gStatus = .GetDWORD
        '        ET = .GetDWORD
        '        gName = .GetSTRING
        '        gPass = .GetSTRING
        '        gStat = .GetSTRING
        '    End With
        '    Dim GSS As String, IPAdd As String
        '    IPAdd = Asc(Left$(IP, 1)) & "." & Asc(Mid$(IP, 2, 1)) & "." & Asc(Mid$(IP, 3, 1)) & "." & Asc(Right$(IP, 1))
        '    Select Case GS
        '        Case &H1: GSS = "Doesn't Exist"
        '        Case &H2: GSS = "Incorrect Password"
        '        Case &H3: GSS = "Full"
        '        Case &H4: GSS = "Started Already"
        '        Case &H6: GSS = "Too many server requests"
        '    End Select
        '    AddChat Color.BotInfo, svrBNCS, Color.Bot, GetGameType(gType) & ": " & gName & IIf(LenB(GSS) > 0, " (" & GSS & ")", vbNullString) & _
        '        "   [" & IPAdd & ":" & modWinsock.htons(PR) & "]"
        'Next
    End If
    Exit Sub
Failed:
AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
End Sub

' SID_ENTERCHAT (0x0A)
Private Sub rcvEntChat()
Dim statstring As String, accName As String
    BNET.TrueUsername = bncsPB.GetString
    statstring = bncsPB.GetString
    accName = bncsPB.GetString
    
    If BNET.varDebugMode Then AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Username: ", Color.Message, BNET.TrueUsername, _
        Color.BotInfo, ", Statstring: ", Color.Message, statstring, Color.BotInfo, ", Account: ", Color.Message, accName
    
    frmMain.Caption = BNET.TrueUsername & "@" & BNET.BNCSServer
    ModifyTray frmMain.icon, BNET.TrueUsername & "@" & BNET.BNCSServer, frmMain
    
    LastTalk = GetTickCount()
    frmMain.tmrAway.Enabled = True
    
    ' Close BNLS for successful entering of chat
    'frmMain.wsBNLS.Close
    
    If blWL_Enabled Then frmMain.tmrWL_Upload.Enabled = True
    If BNET.varIdle Then frmMain.tmrIdle.Enabled = True
    If BNET.varCountIdle Then frmMain.tmrIdle.Enabled = True
    If BNET.varAwayIdle Then frmMain.tmrAway.Enabled = True
End Sub

' SID_GETCHANNELLIST (0x0B)
Private Sub rcvChanList()
Dim spltChan() As String, i As Integer
Dim strChan As String
    strChan = Mid(bncsPB.GetBuffer, 5)
    spltChan() = Split(strChan, Chr(0))
    For i = 0 To UBound(spltChan) - 2
        frmMain.lstChanSave.ListItems.Add , , spltChan(i)
    Next i
    Erase spltChan()
End Sub

' SID_CHECKAD (0x15)
Private Sub rcvAd()
Dim adId As Long, fileExt As Long, locFT As Long, FileName As String, linkURL As String
    'Dim spltAd() As String: spltAd() = Split(Mid(d, 21), Chr(0), 3)
    'AddChat Color.BotInfo, svrBNCS & "Ad Banner ", Color.Bot, spID, Color.Message, spltAd(0) & " [url=" & spltAd(1) & "]"
    'Erase spltAd()
    adId = bncsPB.GetDWORD
    fileExt = bncsPB.GetDWORD
    locFT = bncsPB.GetFileTime
    FileName = bncsPB.GetString
    linkURL = bncsPB.GetString
    AddChat Color.BotInfo, svrBNCS & "Ad " & adId & " ", Color.Bot, spID, Color.Message, FileName & "." & fileExt & " - url: " & linkURL
End Sub

' SID_MESSAGEBOX (0x19)
Private Sub rcvMsgBox()
    'Dim spltMsg() As String: spltMsg() = Split(Mid(d, 8), Chr(0))
    'AddChat Color.BotInfo, "[" & spltMsg(2) & "] ", Color.Message, spltMsg(1)
    'MsgBox spltMsg(1), vbOKOnly, spltMsg(2)
    'Erase spltMsg()
    MsgBox bncsPB.GetString, vbOKOnly, bncsPB.GetString
End Sub

' SID_STARTADVEX3 (0x1C)
Private Sub rcvStartAdvEx3()
    Select Case bncsPB.GetDWORD
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Game successfully created!"
            AddChat Color.BotInfo, "Created: ", Color.Bot, frmSTAR.txtGameName.text & " // Password: " & frmSTAR.txtGamePass.text
            frmMain.lstChannel.ListItems.Clear
            frmMain.txtChanName.text = frmSTAR.txtGameName.text
            sPB.SendPacket &H10 ' SID_LEAVECHAT
        Case &H1: AddChat Color.BotInfo, svrBNCS, Color.Error, "Game failed to create!"
    End Select
End Sub

' SID_LOGONCHALLENGEEX (0x1D)
Private Sub rcvLogChalEx()
    'AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Received challenge..."
    UDPToken = bncsPB.GetDWORD      ' (DWORD) UDP Token
    ServerToken = bncsPB.GetDWORD   ' (DWORD) Server Token
End Sub

' SID_PING (0x25)
Private Sub rcvPing()
    Select Case BNET.varLagPlug
        Case "-1ms", "0ms": Exit Sub
        Case Else
            sPB.InsertDWORD bncsPB.GetDWORD ' (DWORD) Ping Value
            sPB.SendPacket &H25
    End Select
End Sub

' SID_READUSERDATA (0x26)
Private Sub rcvUsrProf()
Dim d As String
    d = bncsPB.GetBuffer
'Dim numAcc As Long, numKeys As Long, reqID As Long, reqKeyVal() As String
    If varRequest = "Profile" Then
        Dim ProfileEnd As String
        Dim spltProfile As Variant
        ProfileEnd = Mid(d, 17, Len(d))
        spltProfile = Split(ProfileEnd, Chr(&H0))
        'URL Detect
        With frmProfile
            .Show
            .rtbUsername.text = ""
            .rtbSex.text = ""
            .rtbAge.text = ""
            .rtbLocation.text = ""
            .rtbDescription.text = ""
            frmProfile.AddText Color.Message, ProfUser, frmProfile.rtbUsername
            frmProfile.AddText Color.Message, spltProfile(0), frmProfile.rtbSex
            frmProfile.AddText Color.Message, spltProfile(1), frmProfile.rtbAge
            frmProfile.AddText Color.Message, spltProfile(2), frmProfile.rtbLocation
            frmProfile.AddText Color.Message, spltProfile(3), frmProfile.rtbDescription
        End With
    ElseIf varRequest = "RecordData" Then
        Dim RecordEnd As String
        Dim spltRecord() As String
        RecordEnd = Mid(d, 17, Len(d))
        spltRecord = Split(RecordEnd, Chr(&H0))
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Record Data: "
        AddChat Color.Bot, "System\Account Created: ", Color.Message, spltRecord(0)
        AddChat Color.Bot, "System\Username: ", Color.Message, spltRecord(1)
        AddChat Color.Bot, "System\Last Logon: ", Color.Message, spltRecord(2)
        AddChat Color.Bot, "System\Last Logoff: ", Color.Message, spltRecord(4)
        AddChat Color.Bot, "System\Time Logged: ", Color.Message, spltRecord(5)
        AddChat Color.Bot, "System\Account Expires: ", Color.Message, spltRecord(3)
    End If
End Sub

' SID_LOGONCHALLENGE (0x28)
Private Sub rcvLogonChal()
    ServerToken = bncsPB.GetDWORD   ' (DWORD) Server Token
End Sub

' SID_LOGONRESPONSE (0x29)
Private Sub rcvLogonResp()
    Select Case bncsPB.GetDWORD ' (DWORD) Result
        Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, "Invalid password."
        Case &H1
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Logged on!"
            Call sndEntChat
    End Select
End Sub

' SID_CREATEACCOUNT (0x2A)
Private Sub rcvCreateAcc()
    Select Case bncsPB.GetDWORD
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, "Account creation failed!"
            Call frmMain.Disconnect
        Case &H1
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Account successfully created!"
            frmMain.wsBNLS.Close
            frmMain.wsBNLS.Connect BNET.BNLSServer, 9367
            frmMain.tmrIdle.Enabled = True
            InitialCon = True
    End Select
End Sub

' SID_GETICONDATA (0x2D)
Private Sub rcvIconData()
Dim icondata As String
    bncsPB.Skip 8               ' (FILETIME) Filetime
    icondata = bncsPB.GetString ' (STRING) Filename
    If BNET.varDebugMode Then
        If FileExists(App.Path + "\" + icondata) = True Then
            AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.BotInfo, _
                "Using default icons file: ", Color.Message, icondata
        Else
            'Download icons file
            'frmMain.ConnectBNETFTP icondata
        End If
    End If
End Sub

' SID_CDKEY (0x30)
Private Sub rcvCDKey()
Dim Result As Long, keyOwner As String, errMsg As String
    Result = bncsPB.GetDWORD    ' (DWORD) Result
    keyOwner = bncsPB.GetString ' (STRING) Key owner
    If Not keyOwner = "" Then keyOwner = " (" & keyOwner & ")"
    Select Case Result
        Case &H1
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "CDKey Accepted!" & keyOwner
            
            With sPB
                HType = 1
                .InsertDWORD Len(BNET.password)
                .InsertDWORD &H0
                .InsertNonNTString BNET.password
                .SendBNLSPacket &HB
                SPass = True
            End With
        Case &H2: errMsg = "Invalid CDKey." & keyOwner: GoTo Failed
        Case &H3: errMsg = "Bad Product." & keyOwner: GoTo Failed
        Case &H4: errMsg = "Banned CDKey." & keyOwner: GoTo Failed
        Case &H5: errMsg = "CDKey is in use." & keyOwner: GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
    frmMain.Disconnect
End Sub

' SID_GETFILETIME (0x033)
Private Sub rcvGetFileTime()
Dim ReqID As Long, fName As String, fTime As String
    With bncsPB
        ReqID = .GetDWORD       ' (DWORD) Request ID
        .Skip 4                 ' (DWORD) Unknown
        fTime = .GetFileTime    ' (FILETIME) Last update time
        fName = .GetString      ' (STRING) Filename
    End With
    
    Dim spltTime() As String
    spltTime() = Split(fTime, " ")
    'If BNET.varDebugMode = 1 Then AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "File Last Updated: ", Color.Message, DateAdd("s", spltTime(1), #1/1/1995#)
    Erase spltTime()
    
    If (fName = BnetLastFile) And (blEnterChat = False) Then
        Call sndContAuth ' Continue to 0x53
    ElseIf blDownloadFile Then
        Call ConnectBNETFTP(fName, fTime)
    End If
End Sub

' SID_LOGONRESPONSE2 (0x3A)
Private Sub rcvLogonResp2()
Dim errMsg As String
    Select Case bncsPB.GetDWORD
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Logged on!"
            Select Case BNET.Product
                Case "PX2D", "VD2D": sPB.SendPacket &H40 ' SID_QUERYREALMS2
            End Select
            Call sndEntChat
        Case &H1
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Carrot, "Account doesn't exist! Attempting to create..."
            With sPB
                HType = 2
                .InsertDWORD Len(BNET.password)
                .InsertDWORD &H0
                .InsertNonNTString BNET.password
                .SendBNLSPacket &HB
            End With
            frmMain.tmrIdle.Enabled = False
        Case &H2: errMsg = "Invalid password!": GoTo Failed
        Case &H6: errMsg = "Account Closed! - " & bncsPB.GetString: GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
    Call frmMain.Disconnect
End Sub

' SID_CREATEACCOUNT2 (0x3D)
Private Sub rcvCreateAcc2()
Dim errMsg As String
    Select Case bncsPB.GetDWORD
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, "Account Created! reconnecting.."
            Call frmMain.ConnectProc
        Case &H1: errMsg = "Name is too short! (" & bncsPB.GetString & ")": GoTo Failed
        Case &H2: errMsg = "Name contained invalid characters! (" & bncsPB.GetString & ")": GoTo Failed
        Case &H3: errMsg = "Name contained a banned word! (" & bncsPB.GetString & ")": GoTo Failed
        Case &H4: errMsg = "Account already exists!(" & bncsPB.GetString & ")": GoTo Failed
        Case &H6: errMsg = "Name did not contain enough alphanumeric characters! (" & bncsPB.GetString & ")": GoTo Failed
        Case &H7: errMsg = "Name contained adjacent punctuation characters! (" & bncsPB.GetString & ")": GoTo Failed
        Case &H8: errMsg = "Name contained too many punctuation characters! (" & bncsPB.GetString & ")": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
    Call frmMain.Disconnect
End Sub

'SID_LOGONREALMEX (0x3E)
Private Sub rcvLogonRealmEX()
Dim ip As String, port As Long, errMsg As String
    With bncsPB
        rCookie = .GetDWORD      ' (DWORD) MCP Cookie
        rStatus = .GetDWORD      ' (DWORD) MCP Status
        Select Case rStatus
            Case &H80000001: errMsg = "Realm is down!": GoTo Failed
            Case &H80000002: errMsg = "Realm logon failed!": GoTo Failed
            Case Else
                rChunk1 = .GetRaw(8)        ' (DWORD) [2] MCP Chunk 1
                ip = MakeServer(.GetRaw(4)) ' (DWORD) IP
                port = .GetDWORD            ' (DWORD) Port
                rChunk2 = .GetRaw(48)       ' (DWORD) [12] MCP Chunk 2
                RealmName = .GetString      ' (STRING) Battle.net unique name
                
                AddChat Color.BotInfo, svrMCP, Color.Bot, "Connecting to MCP server ", Color.Message, ip & " (" & BNET.Realm & _
                    ")", Color.Bot, "..."
                    
                frmRealm.wsMCP.Close
                frmRealm.wsMCP.Connect ip, 6112
        End Select
    End With
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
    frmRealm.wsMCP.Close
End Sub

' SID_QUERYREALMS2 (0x40)
Private Sub rcvQueryRealms2()
Dim i As Long, RealmTitle As String, RealmDesc As String
    With bncsPB
        .Skip 4                     '(DWORD) Unknown
   
        'For Each Realm:
        For i = 1 To .GetDWORD      '(DWORD) Count
            .Skip 4                 '(DWORD) Unknown
            RealmTitle = .GetString '(STRING) Realm title
            RealmDesc = .GetString  '(STRING) Realm description
        Next i
    End With
    
    If showRealms Then
        AddChat Color.BotInfo, svrBNCS & "Realm: ", Color.Message, RealmTitle, Color.BotInfo, ", Description: ", Color.Message, RealmDesc
        showRealms = False
    End If
    
    BNET.Realm = RealmTitle
    If BNET.varCRealm = 1 Then sndLogRealmEx BNET.Realm
End Sub

' SID_NEWS_INFO (0x46)
Private Sub rcvNews()
Dim lastLogon As Long, oldNews As Long, newNews As Long
Dim newsTimeStamp As Long, newsMessage As String
Dim numEnt As Long, i As Long
    With bncsPB
        numEnt = .GetByte                   ' (BYTE) Number of entries
        numEnt = 1
        lastLogon = .GetDWORD               ' (DWORD) Last logon timestamp
        oldNews = .GetDWORD                 ' (DWORD) Oldest news timestamp
        newNews = .GetDWORD                 ' (DWORD) Newest news timestamp
        newsTimeStamp = .GetDWORD           ' (DWORD) Timestamp
        newsMessage = .GetString            ' (STRING) News
    End With
    If Len(newsMessage) < 1 Then Exit Sub
    If FetchNews Then
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, _
            DateAdd("s", newsTimeStamp, #1/1/1970#) & ": ", Color.Message, newsMessage
    Else
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, _
            "Last Logon: ", Color.Message, DateAdd("s", lastLogon, #1/1/1970#)
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, _
            "Oldest News: ", Color.Message, DateAdd("s", oldNews, #1/1/1970#)
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, _
            "Newest News: ", Color.Message, DateAdd("s", newNews, #1/1/1970#)
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, _
            DateAdd("s", newsTimeStamp, #1/1/1970#) & ": ", Color.Message, newsMessage
    End If
    FetchNews = True
End Sub

' SID_OPTIONALWORK (0x4A)
Private Sub rcvOptWrk()
    If BNET.varDebugMode Then: AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.BotInfo, _
        "Optional Work: ", Color.Message, bncsPB.GetString
End Sub

' SID_REQUIREDWORK (0x4C)
Private Sub rcvReqWrk()
    BnetFileName = bncsPB.GetString
    'blDownloadFile = True
    
    If BNET.varDebugMode Then: AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.BotInfo, _
        "Ignoring Required Work: ", Color.Message, BnetFileName
    ' Get Files to Emulate Client
    ' sndGetFiles
End Sub

'SID_AUTH_INFO (0x50)
Private Sub rcvAuthInfo()
Dim UDP As Long
Dim LogonType As Long, LogonTypeName As String

    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Received authorization info!"
    
    With bncsPB
        LogonType = .GetDWORD           ' (DWORD) Logon Type
        ServerToken = .GetDWORD         ' (DWORD) Server Token
        UDP = .GetDWORD                 ' (DWORD) UDPValue
        AuthTimeStamp = .GetFileTime    ' (FILETIME) MPQ filetime
        MPQName = .GetString            ' (STRING) IX86ver filename
        AuthHash = .GetString           ' (STRING) ValueString
    End With
    If BNET.Product = "3RAW" Or BNET.Product = "PX3W" Then
        ServerSignature = bncsPB.GetRaw(128)
    End If
    
    Select Case LogonType
        Case 0: LogonTypeName = "Broken SHA-1" ' 0x00: Broken SHA-1 (STAR/SEXP/D2DV/D2XP)
        Case 1: LogonTypeName = "NLS v1" ' 0x01: NLS version 1 (War3Beta)
        Case 2: LogonTypeName = "NLS v2" ' 0x02: NLS Version 2 (WAR3/W3XP)
    End Select

    'Debug
    If BNET.varDebugMode Then
        AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.BotInfo, "LogonType: ", Color.Message, LogonTypeName, Color.BotInfo, ", MPQName: ", Color.Message, MPQName
    End If
    
    
    ' BNLS NLS Version
    If BNET.Product = "3RAW" Or BNET.Product = "PX3W" Then
        sPB.InsertDWORD &H2        ' (DWORD) NLS revision number
        sPB.SendBNLSPacket &HD     ' BNLS_CHOOSENLSREVISION
    Else
        ' BNLS Version Check
        With sPB
            .InsertDWORD GetBNLSByte()                      ' (DWORD) Product ID
            .InsertDWORD &H0                                ' (DWORD) Flags (None, set to 0)
            .InsertDWORD &H1                                ' (DWORD) Cookie
            .InsertDWORD CLng(Split(AuthTimeStamp, " ")(1)) ' (FILETIME) Timestamp for version check archive
            .InsertDWORD CLng(Split(AuthTimeStamp, " ")(0)) '   cont.
            .InsertNTString MPQName                         ' (STRING) Version check archive filename
            .InsertNTString AuthHash                        ' (STRING) Checksum formula
            .SendBNLSPacket &H1A                            ' BNLS_VERSIONCHECKEX2
        End With
    End If
End Sub

 'SID_AUTH_CHECK (0x51)
Private Sub rcvAuthCheck()
Dim errMsg As String
    Select Case bncsPB.GetDWORD
       Case &H0 '0x000: Passed challenge
           AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Passed challenge!"
           With sPB
                'send 0x33 for d2dv, d2xp, war3, w3xp
                If Not Diablo2 Then .SendPacket &H2D
                Call sndContAuth
            End With
        Case &H100: errMsg = "Game version out of date. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H101: errMsg = "Invalid game version. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H102: errMsg = "Game version needs to be downgraded. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H200: errMsg = "CDKey is invalid. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H201: errMsg = "CDKey is in use. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H203: errMsg = "Incorrect CD-Key for this product. (" & bncsPB.GetString & ")": GoTo Failed
        Case &H202: errMsg = "CDKey is banned. (" & bncsPB.GetString & ")": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotError, errMsg
End Sub
 
 ' SID_AUTH_ACCOUNTCREATE (0x52)
Private Sub rcvAuthAccCreate()
Dim errMsg As String
    Select Case bncsPB.GetDWORD
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Successfully created account name!"
            With sPB
                .InsertNTString BNET.username
                .InsertNTString BNET.password
                .SendBNLSPacket &H2
            End With
        Case &H4: errMsg = "Name already exists.": GoTo Failed
        Case &H7: errMsg = "Name is too short/blank.": GoTo Failed
        Case &H8: errMsg = "Name contains an illegal character.": GoTo Failed
        Case &H9: errMsg = "Name contains an illegal word.": GoTo Failed
        Case &H9: errMsg = "Name contains too few alphanumeric characters.": GoTo Failed
        Case &HB: errMsg = "Name contains adjacent punctuation characters.": GoTo Failed
        Case &HC: errMsg = "Name contains too many punctuation characters.": GoTo Failed
        Case Else: errMsg = "Any other: Name already exists.": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotError, errMsg
End Sub

' SID_AUTH_ACCOUNTLOGON (0x53)
Private Sub rcvAuthAccLogon()
Dim errMsg As String
    Select Case bncsPB.GetDWORD
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Logon accepted, requires proof."
            With sPB
                .InsertNonNTString bncsPB.GetRaw(64)    ' (DWORD) [16] Data from SID_AUTH_ACCOUNTLOGON
                .SendBNLSPacket &H3
            End With
        Case &H1
            If AttemptedC = False Then
                AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Carrot, "Account doesn't exist. Attempting to create..."
                With sPB
                    .InsertNTString BNET.username
                    .InsertNTString BNET.password
                    .SendBNLSPacket &H4
                End With
                AttemptedC = True
            End If
        Case &H5: errMsg = "Account requires upgrade."
        Case Else: errMsg = "Unknown (failure)."
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotError, errMsg
    Call frmMain.Disconnect
End Sub

' SID_AUTH_ACCOUNTLOGONPROOF (0x54)
Private Sub rcvAuthAccLogonProof()
Dim errMsg As String
    Select Case bncsPB.GetDWORD
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Logon successful!"
            ' Net Game Port
            If BNET.Product = "3RAW" Or BNET.Product = "PX3W" Then
                sPB.InsertWORD 6112
                sPB.SendPacket &H45
            End If
            'Could verify 54 Proof here -- BNLSWarden doesn't support
            'sPB.InsertNonNTString bncsPB.GetRaw(20)
            'sPB.SendBNLSPacket &HA
            Call sndEntChat
        Case &H2: errMsg = "Incorrect password.": GoTo Failed
        Case &H6: errMsg = "Account closed.": GoTo Failed
        Case &HE
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Logon successful!"
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Carrot, "Battle.net requested an email address! Attempting to add..."
            Call sndSetEmail
            Call sndEntChat
        Case &HF: errMsg = "Error: " & bncsPB.GetString: GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotError, errMsg
    Call frmMain.Disconnect
End Sub

' SID_WARDEN (0x5E)
Private Sub rcvWarden()
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, "Received Warden! (No handler in place)"
End Sub

' SID_CLANFINDCANDIDATES (0x70) NEEDS WORK
Private Sub rcvClanFindCan()
Dim status As Long, numCan As Long, usernames() As String
Dim i As Long
    With bncsPB
        .Skip 4                                 ' (DWORD) Cookie
        status = .GetByte                       ' (BYTsE) Status
        numCan = .GetByte                       ' (BYTE) Number of potential candidates
        usernames() = .GetStringArray(numCan)   ' (STRING) [] Usernames
    End With
        For i = 0 To numCan
            ' Output Candidates
        Next i
    Erase usernames()
End Sub

'SID_CLANINVITEMULTIPLE (0x71) NEEDS WORK
Private Sub rcvClanInvMul()
Dim Result As String, failAcc As String, errMsg As String
    With bncsPB
        .Skip 4                     ' (DWORD) Cookie
        Result = .GetByte           ' (BYTE) Result
        failAcc = .GetString        ' (STRING) [] Failed account names
    End With
    Select Case bncsPB.GetByte
        Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Success!"
        Case &H4: errMsg = "Declined. (" & failAcc & ")": GoTo Failed
        Case &H5: errMsg = "Not available. (not in channel or already in a clan) (" & failAcc & ")": GoTo Failed
        Case &HB: errMsg = "Clan name contains bad word. (" & failAcc & ")": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
End Sub

'SID_CLANCREATIONINVITATION (0x72) NEEDS WORK
Private Sub rcvClanCrInv()
Dim clanTag As Long, clanName As String, inviter As String, numUsers As Long, users() As String
Dim i As Long
    With bncsPB
        .Skip 4                             ' (DWORD) Cookie
        clanTag = .GetDWORD                 ' (DWORD) Clan Tag
        clanName = .GetString               ' (STRING) Clan Name
        inviter = .GetString                ' (STRING) Inviter's username
        numUsers = .GetByte                 ' (BYTE) Number of users being invited
        users() = .GetStringArray(numUsers) ' (STRING) [] List of users being invited
    End With
    
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "You have been invited to " & Color.Message, clanName & " (" & clanTag & ") ", _
        Color.BotInfo, "by " & inviter & "."
End Sub

' SID_CLANDISBAND (0x73)
Private Sub rcvClanDisband()
Dim errMsg As String
    bncsPB.Skip 4 ' (DWORD) Cookie
    Select Case bncsPB.GetByte
        Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Successfully disbanded the clan."
        Case &H2: errMsg = "Cannot quit clan, not 1 week old yet.": GoTo Failed
        Case &H7: errMsg = "Not authorized to disband the clan.": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
End Sub

' SID_CLANMAKECHIEFTAN (0x74)
Private Sub rcvClanMakeChief()
Dim errMsg As String
    bncsPB.Skip 4 ' (DWORD) Cookie
    Select Case bncsPB.GetByte
        Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Success!"
        Case &H2: errMsg = "Cannot change until clan is a week old.": GoTo Failed
        Case &H4: errMsg = "Declined!": GoTo Failed
        Case &H5: errMsg = "Failed!": GoTo Failed
        Case &H7: errMsg = "Not Authorized!": GoTo Failed
        Case &H8: errMsg = "Not Allowed!": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
End Sub

' SID_CLANINFO (0x75)
Private Sub rcvClanInfo()
Dim unknown As Byte, clanTag As Long, rank As Byte, strRank As String
    unknown = bncsPB.GetByte    ' (BYTE) Unknown (0)
    clanTag = bncsPB.GetDWORD   ' (DWORD) Clan tag
    rank = bncsPB.GetByte       ' (BYTE) Rank
    Select Case rank
        Case &H0: strRank = "Initiate (<1 Week)"
        Case &H1: strRank = "Initiate"
        Case &H2: strRank = "Member"
        Case &H3: strRank = "Officer"
        Case &H4: strRank = "Leader"
    End Select
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Clan: ", Color.Bot, CStr(clanTag), _
                Color.BotInfo, ", Rank: ", Color.Bot, strRank
End Sub

' SID_CLANQUITNOTIFY (0x76)
Private Sub rcvClanQuit()
    Select Case bncsPB.GetByte
        Case &H1: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, "Removed from clan!"
    End Select
End Sub

' SID_CLANINVITATION (0x77)
Private Sub rcvClanInv()
Dim errMsg As String
    bncsPB.Skip 4 ' (DWORD) Cookie
    Select Case bncsPB.GetByte
        Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Invitation accepted!"
        Case &H4: errMsg = "Invitation declined!": GoTo Failed
        Case &H5: errMsg = "Failed to invite user!": GoTo Failed
        Case &H9: errMsg = "Clan is full!": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
End Sub

' SID_CLANREMOVEMEMBER (0x78)
Private Sub rcvClanRemMem()
Dim errMsg As String
    bncsPB.Skip 4 ' (DWORD) Cookie
    Select Case bncsPB.GetByte
        Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Removed!"
        Case &H1: errMsg = "Removal failed!": GoTo Failed
        Case &H2: errMsg = "Cannot be removed yet!": GoTo Failed
        Case &H7: errMsg = "Not authorized to remove!": GoTo Failed
        Case &H9: errMsg = "Not allowed to remove!": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
End Sub

' SID_CLANINVITATIONRESPONSE (0x79)
Private Sub rcvClanInvRes()
Dim clanTag As Long, clanName As String, clanInviter As String
    With bncsPB
        .Skip 4                     ' (DWORD) Cookie
        clanTag = .GetDWORD         ' (DWORD) Clan tag
        clanName = .GetString       ' (STRING) Clan name
        clanInviter = .GetString    ' (STRING) Inviter
    End With
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Message, clanInviter, Color.BotInfo, _
        " has invited you to: ", Color.Message, clanName, Color.BotInfo, "[" & CStr(clanTag) & "]", Color.BotInfo, "."
End Sub

' SID_CLANRANKCHANGE (0x7A)
Private Sub rcvClanRankChg()
Dim errMsg As String
    bncsPB.Skip 4 ' (DWORD) Cookie
    Select Case bncsPB.GetByte
        Case &H0: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "Successfully changed rank!"
        Case &H1: errMsg = "Failed to change rank!": GoTo Failed
        Case &H2: errMsg = "Cannot change user's rank yet!": GoTo Failed
        Case &H7: errMsg = "Not authorized to change user rank! (not shaman/chieftan rank)": GoTo Failed
        Case &H9: errMsg = "Not allowed to change user rank! (user is higher rank)": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, errMsg
End Sub

'SID_CLANMOTD (0x7C)
Private Sub rcvClanMOTD()
    bncsPB.Skip 4 ' (DWORD) Cookie
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "MOTD: ", Color.Message, bncsPB.GetString
End Sub

'SID_CLANMEMBERLIST (0x7D)
Private Sub rcvClanMemList()
Dim numMembers As Long, i As Long
    With bncsPB
       .Skip 4                  ' (DWORD) Cookie
        numMembers = .GetByte   ' (BYTE) Number of Members
        ' (STRING) Username
        ' (BYTE) Rank
        ' (BYTE) Online Status
        ' (STRING) Location
        For i = 0 To numMembers
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, "User: ", Color.Message, .GetString, _
                Color.BotInfo, ", Rank: ", Color.Message, .GetByte, Color.BotInfo, ", Status: ", Color.Message, .GetByte, _
                Color.BotInfo, ", Location: ", Color.Message, .GetString
        Next i
    End With
End Sub

'SID_CLANMEMBERREMOVED (0x7E)
Private Sub rcvClanMemRem()
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, bncsPB.GetString & " removed from clan."
End Sub

'SID_CLANMEMBERSTATUSCHANGE (0x7F)
Private Sub rcvClanMemStatChg()
Dim user As String, rank As String, status As String, loc As String
    ' Switch to turn off display of members needed here
    user = bncsPB.GetString         ' (STRING) Username
    rank = getRank(bncsPB.GetByte)  ' (BYTE) Rank
    Select Case bncsPB.GetByte      ' (BYTE) Status
        Case &H0: status = "Offline"
        Case &H1: status = "Online (not in either channel or game)"
        Case &H2: status = "in a channel "
        Case &H3: status = "in a public game "
        Case &H5: status = "in a private game "
    End Select
    loc = bncsPB.GetString          ' (STRING) Location
    AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, user & " (" & rank & ") has changed status: " & status & loc & "."
End Sub

'SID_CLANMEMBERRANKCHANGE (0x81)
Private Sub rcvClanMemRankChg()
Dim oRank As Byte, nRank As Byte, user As String
Dim rankName As String
    With bncsPB
        oRank = .GetByte    ' (BYTE) Old rank
        nRank = .GetByte    ' (BYTE) New rank
        user = .GetString   ' (STRING) Clan member whose rank changed
    End With
    If nRank > oRank Then
        rankName = getRank(nRank)
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, user & " has been promoted to " & rankName & "."
    Else
        rankName = getRank(oRank)
        AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, user & " has been demoted to " & rankName & "."
    End If
End Sub

'SID_CLANMEMBERINFORMATION (0x82)
Private Sub rcvClanMemInfo()
Dim status As Long, clanName As String, rank As String, dateJoined As Long
    With bncsPB
        .Skip 4                     ' (DWORD) Cookie
        status = .GetByte           ' (BYTE) Status code
        clanName = .GetString       ' (STRING) Clan name
        rank = getRank(.GetByte)    ' (BYTE) User's rank
        dateJoined = .GetFileTime   ' (FILETIME) Date joined
    End With
    Select Case status
        Case &H0
            AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.BotInfo, _
                "User found! Clan Name: " & clanName & ", Rank: " & rank & ", Joined: " & DateAdd("s", dateJoined, #1/1/1970#)
        Case &HC: AddChat Color.BotInfo, svrBNCS, Color.Bot, spID, Color.Error, "User not found in clan."
    End Select
End Sub

'
'
' Client to Server Sub Routines
'
'

' SID_CLIENTID C->S (0x05)
Public Sub sndClientID()
    If BNET.varDebugMode Then: cpID = "[C 0x05] "
    If BNET.varDebugMode Then
        AddChat Color.Carrot, "[Debug] ", Color.Bot, cpID, _
            Color.BotInfo, "VerByte: ", Color.Message, Hex(VerByte), _
            Color.BotInfo, ", Product: ", Color.Message, BNET.Product
    End If
    
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending client, locale and versioning info..."
    
    With sPB
        ' Send Client ID
        .InsertDWORD &H1    ' (DWORD) Server Version (0 or 1)
        .InsertDWORD &H0    ' (DWORD) Registration Version
        .InsertDWORD &H0    ' (DWORD) Registration Authority
        .InsertDWORD &H0    ' (DWORD) Account Number
        .InsertDWORD &H0    ' (DWORD) Registration Token
        .InsertNTString ""  ' (STRING) LAN computer name
        .InsertNTString ""  ' (STRING) LAN username
        .SendPacket &H5     ' SID_CLIENTID
        
        sndLocaleInfo       ' SID_LOCALEINFO (0x12)
        
        ' Send System Info (optional)
        ' SSHR JSTR
        
        sndStartVer         ' SID_STARTVERSIONING (0x06)
    
        ' Send SID_PING before Server asks to achieve 0ms.
        If BNET.varLagPlug = "0ms" Then
            .InsertDWORD &H0                     ' (DWORD) Ping Value
            .SendPacket &H25                     ' SID_PING
            If BNET.varDebugMode Then: cpID = "[C 0x25] "
            If BNET.varDebugMode = 1 And blEnterChat = False Then AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending ping reply to achieve 0ms..."
        End If
    End With
End Sub

' SID_REPORTVERSION C->S (0x07)
Public Sub sndReportVer()
    If BNET.varDebugMode Then: cpID = "[C 0x07] "
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending challenge..."
    
    If BNET.varDebugMode = 1 Then
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Platform: ", Color.Message, "68XI" & BNET.Product
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "VerByte: ", Color.Message, Hex(VerByte)
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Version: ", Color.Message, "0x" & Hex(Version)
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "CheckSum: ", Color.Message, "0x" & Hex(CheckSum)
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Exe Info: ", Color.Message, VerCheckStatString
    End If
    
    With sPB
        .InsertNonNTString "68XI"           ' (DWORD) Platform ID
        .InsertNonNTString BNET.Product     ' (DWORD) Platform ID
        .InsertDWORD VerByte                ' (DWORD) Version Byte
        .InsertDWORD Version                ' (DWORD) EXE Version
        .InsertDWORD CheckSum               ' (DWORD) EXE Hash
        .InsertNTString VerCheckStatString  ' (STRING) Exe Information
        .SendPacket &H7                     ' SID_REPORTVERSION
    End With
End Sub

' SID_STARTVERSIONING C->S (0x06)
Private Sub sndStartVer()
    If BNET.varDebugMode Then
        cpID = "[C 0x06] "
        AddChat Color.Carrot, "[Debug] ", Color.Bot, cpID, _
            Color.BotInfo, "VerByte: ", Color.Message, Hex(VerByte), _
            Color.BotInfo, ", Product: ", Color.Message, BNET.Product
    End If
    
    With sPB
        .InsertNonNTString "68XI"       ' (DWORD) Platform ID
        .InsertNonNTString BNET.Product ' (DWORD) Product ID
        .InsertDWORD VerByte            ' (DWORD) Version Byte
        .InsertDWORD &H0                ' (DWORD) Unknown (0)
        .SendPacket &H6                 ' SID_STARTVERSIONING
    End With
End Sub

' SID_LOCALEINFO C->S (0x12)
Private Sub sndLocaleInfo()
Dim sysFT As FILETIME
Dim locFT As FILETIME
    sysFT = GetFTSystem
    locFT = GetFTLocal
    
    With sPB
        .InsertDWORD sysFT.dwLowDateTime    ' (FILETIME) System time
        .InsertDWORD sysFT.dwHighDateTime   '   cont.
        .InsertDWORD locFT.dwLowDateTime    ' (FILETIME) Local time
        .InsertDWORD locFT.dwHighDateTime   '   cont.
        .InsertDWORD &H420                  ' (DWORD) Timezone bias
        .InsertDWORD &H409                  ' (DWORD) SystemDefaultLCID
        .InsertDWORD &H409                  ' (DWORD) UserDefaultLCID
        .InsertDWORD &H409                  ' (DWORD) UserDefaultLangID
        .InsertNTString "ENU"               ' (STRING) Abbreviated language name
        .InsertNTString "1"                 ' (STRING) Country name
        .InsertNTString "USA"               ' (STRING) Abbreviated country name
        .InsertNTString "United States"     ' (STRING) Country name'(STRING) Country (English)
        .SendPacket &H12                    ' SID_LOCALEINFO
    End With
End Sub

' SID_CLIENTID2 C->S (0x1E)
Public Sub sndClientID2()
    If BNET.varDebugMode Then: cpID = "[C 0x1E, 0x12, 0x06] "
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending client, locale and versioning info..."
    
    With sPB
        ' Send Client ID
        .InsertDWORD &H1    ' (DWORD) Server Version (0 or 1)
        .InsertDWORD &H0    ' (DWORD) Registration Version
        .InsertDWORD &H0    ' (DWORD) Registration Authority
        .InsertDWORD &H0    ' (DWORD) Account Number
        .InsertDWORD &H0    ' (DWORD) Registration Token
        .InsertNTString ""  ' (STRING) LAN computer name
        .InsertNTString ""  ' (STRING) LAN usernames
        .SendPacket &H1E    ' SID_CLIENTID2

        sndLocaleInfo       ' SID_LOCALEINFO (0x12)
        sndStartVer         ' SID_STARTVERSIONING (0x06)
        
        ' Send SID_PING before Server asks to achieve 0ms.
        If BNET.varLagPlug = "0ms" Then
            .InsertDWORD &H0                     ' (DWORD) Ping Value
            .SendPacket &H25                     ' SID_PING
        End If
    End With
End Sub

' SID_CDKEY C->S (0x30)
Public Sub sndCDKEY()
    If BNET.varDebugMode Then: cpID = "[C 0x30] "
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending CDKey..."
    With sPB
        .InsertDWORD &H0                ' (DWORD) Spawn (0/1)
        .InsertNTString BNET.CDKey      ' (STRING) CDKey
        .InsertNTString BNET.CDKeyOwner ' (STRING) Key Owner
        .SendPacket &H30
    End With
End Sub

' SID_GETFILETIME C->S (0x33)
' NEEDS WORK
Public Sub sndGetFiles()
Dim FileCount As Integer, GetFile(4) As String, i As Integer
    'Get File Times for various product files
    
    With sPB
        Select Case BNET.Product
            Case "3RAW", "PX3W"
                ' Required Work
                .InsertDWORD &HA0000001         ' (DWORD) Request ID
                .InsertDWORD &H0                ' (DWORD) Unknown
                .InsertNTString BnetFileName    ' (STRING) Filename
                .SendPacket &H33                ' SID_GETFILETIME
                i = i + 1
                
                FileCount = 4
                GetFile(1) = "termsofservices-enUS.txt"
                GetFile(2) = "newaccount-enUS.txt"
                GetFile(3) = "chathelp-war3-enUS.txt"
                GetFile(4) = "bnserver-WAR3.ini"
            Case "RATS", "PXES"
                FileCount = 3
                GetFile(1) = "icons_STAR.bni"
                GetFile(2) = "tos_USA.txt"
                GetFile(3) = "bnserver.ini"
            Case "VD2D", "PX2D"
                FileCount = 1
                GetFile(4) = "bnserver-D2DV.ini"
                i = 4
        End Select
        
        BnetLastFile = GetFile(FileCount)
        
        For i = 1 To FileCount
            .InsertDWORD Val("&H" & i)
            .InsertDWORD &H0
            .InsertNTString GetFile(i)
            .SendPacket &H33
        Next i
    End With
End Sub

' SID_LOGONREALMEX C->S (0x3E)
Public Sub sndLogRealmEx(Realm As String)
    If Not frmMain.wsBNLS.State = sckConnected Then
        frmMain.wsBNLS.Close
        frmMain.wsBNLS.Connect BNET.BNLSServer, 9367
        resendLREx = True
    Else
        With sPB
            HType = 1
            .InsertDWORD Len("password")   ' (DWORD) Size of Data
            .InsertDWORD &H0               ' (DWORD) Flags
            .InsertNonNTString "password"  ' (VOID) Data to be hashed
            .SendBNLSPacket &HB
            .Clear
            SPass = False
        End With
    End If
End Sub

' SID_EXTRAWORK C->S (0x4B)
Public Sub sndExtraWrk(msgLen As Integer, workReturn As String)
Dim gameType As Integer

    If Len(workReturn) > 1 Then
      Select Case BNET.Product
          Case "VD2D", "PX2D": gameType = &H1
          Case "3RAW", "PX3W": gameType = &H2
          Case "RATS", "PXES", "RTSJ": gameType = &H3
          Case Else: gameType = &H0
      End Select
    
      With sPB
          .InsertWORD gameType                ' (WORD) Game type
          .InsertWORD msgLen                  ' (WORD) Length
          .InsertNonNTString workReturn       ' (STRING) Work returned data
          .SendPacket &H4B
      End With
    End If
End Sub

' SID_AUTH_INFO C->S (0x50)
Public Sub sndAuthInfo()
    If BNET.varDebugMode Then: cpID = "[C 0x50] "
    If BNET.varDebugMode Then
        AddChat Color.Carrot, "[Debug] ", Color.Bot, cpID, _
            Color.BotInfo, "VerByte: ", Color.Message, Hex(VerByte), _
            Color.BotInfo, ", Product: ", Color.Message, BNET.Product
    End If
       
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending authorization info..."
    With sPB
        .InsertDWORD &H0                         ' (DWORD) Protocol ID (0)
        .InsertNonNTString "68XI" & BNET.Product ' (DWORD) Platform ID & (DWORD) Product ID
        .InsertDWORD VerByte                     ' (DWORD) Version Byte
        .InsertNonNTString "SUne"                ' (DWORD) Product language (0)
        .InsertDWORD &H0                         ' (DWORD) Local IP for NAT compatibility
        .InsertDWORD 360                         ' (DWORD) Time zone bias
        .InsertDWORD &H409                       ' (DWORD) Locale ID
        .InsertDWORD &H409                       ' (DWORD) Language ID
        .InsertNTString "USA"                    ' (STRING) Country abreviation
        .InsertNTString "United States"          ' (STRING) Country
        .SendPacket &H50                         ' SID_AUTH_INFO
        
        ' Send SID_PING before Server asks to achieve 0ms.
        If BNET.varLagPlug = "0ms" Then
            .InsertDWORD &H0                     ' (DWORD) Ping Value
            .SendPacket &H25                     ' SID_PING
        End If
    End With
End Sub

' SID_AUTH_CHECK C->S (0x51)
Public Sub sndAuthCheck()
Dim numKeys As Long
    
    Select Case BNET.Product
        Case "PX2D", "PX3W": numKeys = &H2
        Case Else: numKeys = &H1
    End Select
    
    If BNET.varDebugMode Then: cpID = "[C 0x51] "
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending challenge..."
    
    If BNET.varDebugMode = 1 Then
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Client Token: ", Color.Message, "0x" & Hex(ClientToken)
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "EXE Version: ", Color.Message, "0x" & Hex(Version)
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "CheckSum: ", Color.Message, "0x" & Hex(CheckSum)
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Number of Keys: ", Color.Message, Hex(numKeys)
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Spawn: ", Color.Message, "0"
        If Not (BNET.Product = "PX2D" Or BNET.Product = "PX3W") Then
            AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "CD-Key Length: ", Color.Message, CDKeyLength
            AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "CD-Key Product: ", Color.Message, CDKeyProdVal
            AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "CD-Key Public: ", Color.Message, "0x" & Hex(CDKeyPubVal)
            AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Unknown: ", Color.Message, CDKeyUnknown
            AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Hashed: ", Color.Message, ToHex(CDKeyHash)
        Else
            AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Hashed: ", Color.Message, ToHex(CDKeyHash)
            AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Hashed 2: ", Color.Message, ToHex(CDKey2Hash)
        End If

        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Exe Info: ", Color.Message, VerCheckStatString
        AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "CD-Key Owner: ", Color.Message, BNET.CDKeyOwner
    End If
    
    With sPB
        .InsertDWORD ClientToken            ' (DWORD) Client Token
        .InsertDWORD Version                ' (DWORD) EXE Version
        .InsertDWORD CheckSum               ' (DWORD) EXE Hash
        .InsertDWORD numKeys                ' (DWORD) Number of CD-keys in this packet
        .InsertDWORD &H0                    ' (BOOLEAN) Spawn CD-key - Actually a DWORD
        If Not (BNET.Product = "PX2D" Or BNET.Product = "PX3W") Then
            .InsertDWORD CDKeyLength            ' (DWORD) Key Length
            .InsertDWORD CDKeyProdVal           ' (DWORD) CD-key's product value
            .InsertDWORD CDKeyPubVal            ' (DWORD) CD-key's public value
            .InsertDWORD CDKeyUnknown           ' (DWORD) Unknown (0)
        End If
        .InsertNonNTString CDKeyHash        ' (DWORD) [x5] Hashed Key Data
        If numKeys = &H2 Then: .InsertNonNTString CDKey2Hash
        .InsertNTString VerCheckStatString  ' (STRING) Exe Information
        .InsertNTString BNET.CDKeyOwner     ' (STRING) CD-Key owner name
        .SendPacket &H51                    ' SID_AUTH_CHECK
    End With
End Sub

Private Sub sndContAuth()
    With sPB
        If BNET.Product = "3RAW" Or BNET.Product = "PX3W" Then
            .InsertNTString BNET.username
            .InsertNTString BNET.password
            .SendBNLSPacket &H2
        Else
            'Can't use 'bnet' with D2DV/D2XP, WAR3/W3XP
            If BNET.varUDP Then
                .InsertNonNTString "bnet" ' UDP - PLUG
            Else
                .InsertNonNTString "tenb"
            End If
            .SendPacket &H14
             
            HType = 1
            .InsertDWORD Len(BNET.password)
            .InsertDWORD &H0
            .InsertNonNTString BNET.password
            .SendBNLSPacket &HB
            SPass = True
        End If
    End With
End Sub

Public Function GetChannelType(ByVal flags As Long) As String
    Dim tmpChannel As String: tmpChannel = ""
    If flags = 0 Then tmpChannel = "Private" & tmpChannel
    If (&H1 And flags) = &H1 Then tmpChannel = "Public" & tmpChannel
    If (&H2 And flags) = &H2 Then tmpChannel = "Moderated " & tmpChannel
    If (&H4 And flags) = &H4 Then tmpChannel = "Restricted " & tmpChannel
    If (&H8 And flags) = &H8 Then tmpChannel = "Silent " & tmpChannel
    If (&H10 And flags) = &H10 Then tmpChannel = "System " & tmpChannel
    If (&H20 And flags) = &H20 Then tmpChannel = "Product-Specific " & tmpChannel
    If (&H1000 And flags) = &H1000 Then tmpChannel = "Globally Accessible " & tmpChannel
    If (&H4000 And flags) = &H4000 Then tmpChannel = "Redirected " & tmpChannel
    GetChannelType = tmpChannel
End Function

Public Function GetIconTier(ByVal IconNum As Long, ByVal Race As String) As String
    Select Case Race
        Case "H"
            Select Case IconNum
                Case 1: GetIconTier = "footman"
                Case 2: GetIconTier = "knight"
                Case 3: GetIconTier = "archmage"
                Case 4: GetIconTier = "medivh"
                Case Else: GetIconTier = "unknown human"
            End Select
        Case "O"
            Select Case IconNum
                Case 1: GetIconTier = "grunt"
                Case 2: GetIconTier = "tauren"
                Case 3: GetIconTier = "far seer"
                Case 4: GetIconTier = "thrall"
                Case Else: GetIconTier = "unknown orc"
            End Select
        Case "N"
            Select Case IconNum
                Case 1: GetIconTier = "archer"
                Case 2: GetIconTier = "druid of the claw"
                Case 3: GetIconTier = "priestess of the moon"
                Case 4: GetIconTier = "furion stomrage"
                Case Else: GetIconTier = "unknown night elf"
            End Select
        Case "U"
            Select Case IconNum
                Case 1: GetIconTier = "ghoul"
                Case 2: GetIconTier = "abomination"
                Case 3: GetIconTier = "lich"
                Case 4: GetIconTier = "tichondrius"
                Case Else: GetIconTier = "unknown undead"
            End Select
        Case "R"
            Select Case IconNum
                Case 1: GetIconTier = "green dragon whelp"
                Case 2: GetIconTier = "blue dragon"
                Case 3: GetIconTier = "red dragon"
                Case 4: GetIconTier = "deathwing"
                Case Else: GetIconTier = "unknown random"
            End Select
        Case Else
            GetIconTier = "unknown race"
    End Select
End Function

Private Sub sndSetEmail()
Dim d As String: d = bncsPB.GetBuffer
Dim e As String
    
    If BNET.varDebugMode = 1 Then: cpID = "[C 0x59] "
    e = Mid(d, 30)
    
    If e = "" Then
        If Len(BNET.Email) > 1 Then
            With sPB
                .InsertNTString BNET.Email
                .SendPacket &H59
            End With
            AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID, Color.BotInfo, "Registered ", Color.Message, BNET.Email, Color.BotInfo, " as e-mail."
        Else
            AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID, Color.Carrot, "No email address present. Please register your account e-mail."
        End If
    End If
End Sub

Public Sub sndEntChat()
Dim username As String, joinFlag As Long, channel As String, chanList As Boolean
    username = BNET.TrueUsername
    channel = BNET.HomeChannel
    joinFlag = 2
    chanList = True
    
    Select Case BNET.Product
        Case "RTSJ"
            joinFlag = 1
            channel = "StarcraftJ"
        Case "RHSS"
            joinFlag = 1
            channel = "Starcraft Shareware"
        Case "LTRD"
            joinFlag = 1
            channel = "Diablo"
        Case "RHSD"
            joinFlag = 1
            channel = "Diablo Shareware"
        Case "RATS", "PXES"
            'joinFlag = 1
            'channel = "Starcraft"
        Case "PX2D", "VD2D"
            chanList = False
        Case "PX3W", "3RAW"
            username = ""
            chanList = False
            'channel = "W3"
    End Select
    
    With sPB
        If username = "" Then
            .InsertByte &H0
        Else
            .InsertNTString username        ' (STRING) Username - Null on WAR3/W3XP
        End If
        .InsertByte &H0                     ' (STRING) Statstring - Null on CDKey Products, except D2DV/D2XP realm characters
        .SendPacket &HA                     ' SID_ENTERCHAT

        If chanList Then
            .InsertNonNTString BNET.Product ' (DWORD) Product ID
            .SendPacket &HB                 ' SID_GETCHANNELLIST
        End If
        
        ' Flags, 0x00: NoCreate join, 0x01: First Join, 0x02: Forced Join, 0x05: D2 first join
        .InsertDWORD joinFlag               ' (DWORD) Flags
        .InsertNTString channel             ' (STRING) Channel
        .SendPacket &HC                     ' SID_JOINCHANNEL
    End With
    
    blEnterChat = True
End Sub

Public Function Rejoin(c As String)
    With sPB
        .SendPacket &H10
        .InsertDWORD 2
        .InsertNTString c
        .SendPacket &HC
    End With
End Function

Public Function sndStartAdvEx3(gName As String, gPass As String, gMap As String, gSpeed As String, gType As String)
    ' Game Types
    '0x02: Melee
    '0x03: Free for All
    '0x04: 1 vs 1
    '0x05: Capture The Flag
    '0x06: Greed (Resources, 0x01: 2500, 0x02: 500, 0x03: 7500, 0x04: 10000)
    '0x07: Slaughter (Minutes, 0x01: 15, 0x02: 30, 0x03: 45, 0x04: 60)
    '0x08: Sudden Death
    '0x09: Ladder (Disconnects, 0x00: Not a loss, 0x01: Counts as a loss)
    '0x0A: Use Map Settings
    '0x0B: Team Melee (Number Of Teams, 0x01: 2 Teams, 0x02: 3 Teams, etc.)
    '0x0C: Team Free For All (Number Of Teams, 0x01: 2 Teams, 0x02: 3 Teams, etc.)
    '0x0D: Team Capture The Flag (Number Of Teams, 0x01: 2 Teams, 0x02: 3 Teams, etc.)
    '0x0F: Top vs. Bottom (Number Of Teams, 1-7 specifies the ratio of players belongin'g to both teams)
    '0x10: Iron Man Ladder (W2BN only)
    '0x20: PGL


    With sPB
        .InsertDWORD 0                                                              ' (DWORD) Game State
        .InsertDWORD 0                                                              ' (DWORD) Game Uptime in seconds
        .InsertWORD 9                                                               ' (WORD) Game Type
        .InsertWORD 1                                                               ' (WORD) Sub Game Type
        .InsertDWORD &H1F                                                           ' (DWORD) Provider Version Constant (0xFF)
        .InsertDWORD 0                                                              ' (DWORD) Ladder Type (0: Normal, 1: Ladder, 3: Iron Man (W2BN))
        .InsertNTString gName                                                       ' (STRING) Game Name
        .InsertNTString gPass                                                       ' (STRING) Game Password
        .InsertNonNTString ",34,12," & gSpeed & ",4," & gType & ",2,ebe6d08d,,,"    ' (STRING) Game Statstring
        .InsertNonNTString BNET.TrueUsername & Chr(&HD)
        .InsertNTString gMap & Chr(&HD)
        .SendPacket &H1C                                                            ' SID_STARTADVEX3
    End With
End Function

Private Function getRank(rank As Byte) As String
    Select Case rank
        Case &H0: getRank = "Initiate"
        Case &H1: getRank = "Peon"
        Case &H2: getRank = "Member"
        Case &H3: getRank = "Officer"
        Case &H4: getRank = "Leader"
    End Select
End Function

Public Function PNBNCS(p As Byte) As String
    Select Case p
        Case &H0: PNBNCS = "SID_NULL"
        Case &H2: PNBNCS = "SID_STOPADV"
        Case &H4: PNBNCS = "SID_SERVERLIST"
        Case &H5: PNBNCS = "SID_CLIENTID"
        Case &H6: PNBNCS = "SID_STARTVERSIONING"
        Case &H7: PNBNCS = "SID_REPORTVERSION"
        Case &H8: PNBNCS = "SID_STARTADVEX"
        Case &H9: PNBNCS = "SID_GETADVLISTEX"
        Case &HA: PNBNCS = "SID_ENTERCHAT"
        Case &HB: PNBNCS = "SID_GETCHANNELLIST"
        Case &HC: PNBNCS = "SID_JOINCHANNEL"
        Case &HE: PNBNCS = "SID_CHATCOMMAND"
        Case &HF: PNBNCS = "SID_CHATEVENT"
        Case &H10: PNBNCS = "SID_LEAVECHAT"
        Case &H12: PNBNCS = "SID_LOCALEINFO"
        Case &H13: PNBNCS = "SID_FLOODDETECTED"
        Case &H14: PNBNCS = "SID_UDPPINGRESPONSE"
        Case &H15: PNBNCS = "SID_CHECKAD"
        Case &H16: PNBNCS = "SID_CLICKAD"
        Case &H17: PNBNCS = "SID_READMEMORY"
        Case &H18: PNBNCS = "SID_REGISTRY"
        Case &H19: PNBNCS = "SID_MESSAGEBOX"
        Case &H1A: PNBNCS = "SID_STARTADVEX2"
        Case &H1B: PNBNCS = "SID_GAMEDATAADDRESS"
        Case &H1C: PNBNCS = "SID_STARTADVEX3"
        Case &H1D: PNBNCS = "SID_LOGONCHALLENGEEX"
        Case &H1E: PNBNCS = "SID_CLIENTID2"
        Case &H1F: PNBNCS = "SID_LEAVEGAME"
        Case &H20: PNBNCS = "SID_ANNOUNCEMENT"
        Case &H21: PNBNCS = "SID_DISPLAYAD"
        Case &H22: PNBNCS = "SID_NOTIFYJOIN"
        Case &H23: PNBNCS = "SID_WRITECOOKIE"
        Case &H24: PNBNCS = "SID_READCOOKIE"
        Case &H25: PNBNCS = "SID_PING"
        Case &H26: PNBNCS = "SID_READUSERDATA"
        Case &H27: PNBNCS = "SID_WRITEUSERDATA"
        Case &H28: PNBNCS = "SID_LOGONCHALLENGE"
        Case &H29: PNBNCS = "SID_LOGONRESPONSE"
        Case &H2A: PNBNCS = "SID_CREATEACCOUNT"
        Case &H2B: PNBNCS = "SID_SYSTEMINFO"
        Case &H2C: PNBNCS = "SID_GAMERESULT"
        Case &H2D: PNBNCS = "SID_GETICONDATA"
        Case &H2E: PNBNCS = "SID_GETLADDERDATA"
        Case &H2F: PNBNCS = "SID_FINDLADDERUSER"
        Case &H30: PNBNCS = "SID_CDKEY"
        Case &H31: PNBNCS = "SID_CHANGEPASSWORD"
        Case &H32: PNBNCS = "SID_CHECKDATAFILE"
        Case &H33: PNBNCS = "SID_GETFILETIME"
        Case &H34: PNBNCS = "SID_QUERYREALMS"
        Case &H35: PNBNCS = "SID_PROFILE"
        Case &H36: PNBNCS = "SID_CDKEY2"
        Case &H3A: PNBNCS = "SID_LOGONRESPONSE2"
        Case &H3C: PNBNCS = "SID_CHECKDATAFILE2"
        Case &H3D: PNBNCS = "SID_CREATEACCOUNT2"
        Case &H3E: PNBNCS = "SID_LOGONREALMEX"
        Case &H3F: PNBNCS = "SID_STARTVERSIONING2"
        Case &H40: PNBNCS = "SID_QUERYREALMS2"
        Case &H41: PNBNCS = "SID_QUERYADURL"
        Case &H43: PNBNCS = "SID_WARCRAFTUNKNOWN"
        Case &H44: PNBNCS = "SID_WARCRAFTGENERAL"
        Case &H45: PNBNCS = "SID_NETGAMEPORT"
        Case &H46: PNBNCS = "SID_NEWS_INFO"
        Case &H4A: PNBNCS = "SID_OPTIONALWORK"
        Case &H4B: PNBNCS = "SID_EXTRAWORK"
        Case &H4C: PNBNCS = "SID_REQUIREDWORK"
        Case &H4E: PNBNCS = "SID_TOURNAMENT"
        Case &H50: PNBNCS = "SID_AUTH_INFO"
        Case &H51: PNBNCS = "SID_AUTH_CHECK"
        Case &H52: PNBNCS = "SID_AUTH_ACCOUNTCREATE"
        Case &H53: PNBNCS = "SID_AUTH_ACCOUNTLOGON"
        Case &H54: PNBNCS = "SID_AUTH_ACCOUNTLOGONPROOF"
        Case &H55: PNBNCS = "SID_AUTH_ACCOUNTCHANGE"
        Case &H56: PNBNCS = "SID_AUTH_ACCOUNTCHANGEPROOF"
        Case &H57: PNBNCS = "SID_AUTH_ACCOUNTUPGRADE"
        Case &H58: PNBNCS = "SID_AUTH_ACCOUNTUPGRADEPROOF"
        Case &H59: PNBNCS = "SID_SETEMAIL"
        Case &H5A: PNBNCS = "SID_RESETPASSWORD"
        Case &H5B: PNBNCS = "SID_CHANGEEMAIL"
        Case &H5C: PNBNCS = "SID_SWITCHPRODUCT"
        Case &H5D: PNBNCS = "SID_REPORTCRASH"
        Case &H5E: PNBNCS = "SID_WARDEN"
        Case &H60: PNBNCS = "SID_GAMEPLAYERSEARCH"
        Case &H65: PNBNCS = "SID_FRIENDSLIST"
        Case &H66: PNBNCS = "SID_FRIENDSUPDATE"
        Case &H67: PNBNCS = "SID_FRIENDSADD"
        Case &H68: PNBNCS = "SID_FRIENDSREMOVE"
        Case &H69: PNBNCS = "SID_FRIENDSPOSITION"
        Case &H70: PNBNCS = "SID_CLANFINDCANDIDATES"
        Case &H71: PNBNCS = "SID_CLANINVITEMULTIPLE"
        Case &H72: PNBNCS = "SID_CLANCREATIONINVITATION"
        Case &H73: PNBNCS = "SID_CLANDISBAND"
        Case &H74: PNBNCS = "SID_CLANMAKECHIEFTAIN"
        Case &H75: PNBNCS = "SID_CLANINFO"
        Case &H76: PNBNCS = "SID_CLANQUITNOTIFY"
        Case &H77: PNBNCS = "SID_CLANINVITATION"
        Case &H78: PNBNCS = "SID_CLANREMOVEMEMBER"
        Case &H79: PNBNCS = "SID_CLANINVITATIONRESPONSE"
        Case &H7A: PNBNCS = "SID_CLANRANKCHANGE"
        Case &H7B: PNBNCS = "SID_CLANSETMOTD"
        Case &H7C: PNBNCS = "SID_CLANMOTD"
        Case &H7D: PNBNCS = "SID_CLANMEMBERLIST"
        Case &H7E: PNBNCS = "SID_CLANMEMBERREMOVED"
        Case &H7F: PNBNCS = "SID_CLANMEMBERSTATUSCHANGE"
        Case &H81: PNBNCS = "SID_CLANMEMBERRANKCHANGE"
        Case &H82: PNBNCS = "SID_CLANMEMBERINFORMATION"
        Case Else: PNBNCS = "UNKNOWN"
    End Select
End Function
