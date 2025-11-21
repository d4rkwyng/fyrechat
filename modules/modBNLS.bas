Attribute VB_Name = "modBNLS"
Option Explicit
' BNLS Variables - Mainly &H1A
Public Cookie As Long               ' (DWORD) Cookie
Private bnlsPB As New clsBuffer
Private spID As String
Private cpID As String

Public Sub ParseBNLS(ByVal d As String)
Dim pID As Byte, msgLen As Long, Msg As String
    
    ' Reset and Initialize
    With bnlsPB
        .Clear
        .SetBuffer d
        
        msgLen = .GetWORD   ' (WORD) Message Length, including this header
        pID = .GetByte      ' (BYTE) Message ID
        Msg = .GetBuffer    ' (VOID) Message Data
    End With
    
    cpID = ""
    spID = ""
    If BNET.varDebugMode = 1 Then spID = "[S 0x" & Hex(pID) & "] "
    
    If BNET.varAdvDebug = 1 Then DisplayAdvDebug "BNLS", d, pID
    
    Select Case pID
        Case &H1: rcvBNLSCDKey
        Case &H2: rcvBNLSLogonChal
        Case &H3: rcvBNLSLogonProof
        Case &H4: rcvBNLSCrAcc
        'Case &H5 ' BNLS_CHANGECHALLENGE
            ' (DWORD) [8] Data for SID_AUTH_ACCOUNTCHANGE (0x55)
        'Case &H6 ' BNLS_CHANGEPROOF
            ' (DWORD) [21] Data for SID_AUTH_ACCOUNTCHANGEPROOF (0x56)
        'Case &H7 ' BNLS_UPGRADECHALLENGE
            ' (BOOLEAN) Success code
            ' If the success code is TRUE, you may send SID_AUTH_ACCOUNTUPGRADE (0x57).
            ' 0x57 is now defunct
        'Case &H8 ' BNLS_UPGRADEPROOF
            ' (DWORD) [22] Data for SID_AUTH_ACCOUNTUPGRADEPROOF (0x58)
            ' 0x58 is now defunct
        Case &H9: rcvBNLSVerCheck
        Case &HA: rcvConfLogon
        Case &HB: rcvBNLSHash
        Case &HC: rcvBNLSCDKeyEx
        Case &HD: rcvBNLSChooseNLSRev
        Case &HE: rcvBNLSAuth ' DEFUNCT
        Case &HF: rcvBNLSAuthProof ' DEFUNCT
        Case &H10: rcvBNLSReqVerByte
        Case &H11: rcvBNLSVerifyServ
        'Case &H12 ' BNLS_RESERVESERVERSLOTS
            ' (DWORD) Number of slots reserved
            ' Not going to use
        'Case &H13 ' BNLS_SERVERLOGONCHALLENGE
            ' (DWORD) Slot index
            ' (DWORD) [16] Data for the server's SID_AUTH_ACCOUNTLOGON (0x53) response
        'Case &H14 ' BNLS_SERVERLOGONPROOF
            ' (DWORD) Slot index
            ' (BOOLEAN) Success (32-bit)
            ' (DWORD) [5] Data for server's SID_AUTH_ACCOUNTLOGONPROOF (0x54) response.
        'Case &H18 ' BNLS_VERSIONCHECKEX
            ' (BOOLEAN) Success
            ' (DWORD) Version
            ' (DWORD) Checksum
            ' (STRING) Version check stat string
            ' (DWORD) Cookie
            ' (DWORD) The latest version code for this product
        Case &H1A: rcvBNLSVerCheckEx2
        Case &HFF
            AddChat Color.BotInfo, svrBNLS, Color.Error, "Your IP address has been restricted temporarily... Try again later."
            Call frmMain.Disconnect
    End Select
End Sub

' BNLS_CDKEY (0x01)
Private Sub rcvBNLSCDKey()
    With bnlsPB
        If Not .GetBoolean Then GoTo Failed
        
        ClientToken = .GetDWORD              ' (DWORD) Client Token
        'CDKeyHash = .GetRaw(36)             ' Old Code - 9 DWORDS (36)
        
        CDKeyLength = .GetDWORD             ' (DWORD) CD-Key Length
        CDKeyProdVal = .GetDWORD            ' (DWORD) CD-Key Product Value
        CDKeyPubVal = .GetDWORD             ' (DWORD) CD-Key Public Value
        CDKeyUnknown = .GetDWORD            ' (DWORD) Unknown (0)
        
        ' (DWORD) [x5] CD-Key Hash (CDKey data for SID_AUTH_CHECK)
        CDKeyHash = .GetRaw(20)
    End With

    Call sndAuthCheck
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNLS, Color.Bot, spID, Color.Error, "Failed Version Check."
    Call frmMain.Disconnect
End Sub

' BNLS_LOGONCHALLENGE (0x02)
Private Sub rcvBNLSLogonChal()
'Dim d As String: d = bnlsPB.GetBuffer
    ' (DWORD)[8] Data for SID_AUTH_ACCOUNTLOGON (0x54)
    If BNET.varDebugMode = 1 Then cpID = "[C 0x53] "
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Initiating logon..."
    With sPB
        '.InsertNonNTString Mid(d, 4)
        .InsertNonNTString bnlsPB.GetRaw(32)    ' (BYTE) [32] Client Key ('A')
        .InsertNTString BNET.username           ' (STRING) Username
        .SendPacket &H53
    End With
End Sub

' BNLS_LOGONPROOF (0x03)
Private Sub rcvBNLSLogonProof()
    ' (DWORD)[5] Data for SID_AUTH_ACCOUNTLOGONPROOF
    If BNET.varDebugMode = 1 Then cpID = "[C 0x54] "
    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Sending logon proof..."
    
    sPB.InsertNonNTString bnlsPB.GetRaw(20) ' (BYTE) [20] Client Password Proof (M1)
    sPB.SendPacket &H54
End Sub

' BNLS_CREATEACCOUNT (0x04)
Private Sub rcvBNLSCrAcc()
Dim d As String: d = bnlsPB.GetBuffer
    ' (DWORD) [16] Data for SID_AUTH_ACCOUNTCREATE (0x52)
    With sPB
        .InsertNonNTString Mid$(d, 4)
        .InsertNTString BNET.username
        .SendPacket &H52
    End With
End Sub

' BNLS_VERSIONCHECK (0x09)
Private Sub rcvBNLSVerCheck()
    ' (BOOLEAN) Success
    ' (DWORD) Version
    ' (DWORD) Checksum
    ' (STRING) Version check stat string
    
    If BNET.Product = "PX2D" Or BNET.Product = "PX3W" Then
        With sPB
            .InsertDWORD &H0
            .InsertByte &H2
            .InsertDWORD &H1
            .InsertDWORD ServerToken
            .InsertNTString BNET.CDKey
            .InsertNTString BNET.CDKey2
            .SendBNLSPacket &HC
        End With
    Else
        With sPB
           .InsertDWORD ServerToken
           .InsertNTString BNET.CDKey
           .SendBNLSPacket &H1
        End With
    End If
End Sub

' BNLS_CONFIRMLOGON (0x0A)
Private Sub rcvConfLogon()
Dim d As String: d = bnlsPB.GetBuffer
    ' (BOOLEAN) Success
    ' Success is TRUE if the server knows your password,
    '  FALSE otherwise. If this is FALSE, the Battle.net connection
    '  should be closed by the client.
    If bnlsPB.GetByte = 1 Then
        If blEnterChat = False Then Call sndEntChat
    Else
        AddChat Color.BotInfo, svrBNLS, Color.Bot, spID, Color.Error, "Failed to confirm proof."
        Call frmMain.DisconnectProc
    End If
End Sub

 ' BNLS_HASHDATA (0x0B)
Private Sub rcvBNLSHash()
Dim d As String: d = bnlsPB.GetBuffer
    '(DWORD) [5] The data hash.
    '(DWORD) Cookie (Optional). Same as the cookie from the request.
    If HType = 1 Then
        'AddChat Color.Carrot, "[Debug] ", Color.Bot, "HType: ", Color.Message, HType
        cB = cB + 1
        'AddChat Color.Carrot, "[Debug] ", Color.Bot, "CB: ", Color.Message, cB
        If cB = 1 Then
            Hash(0) = sPB.MakeDWORD(ClientToken)    ' (DWORD) Client Key (Double Hash Only) - Optional
            Hash(1) = sPB.MakeDWORD(ServerToken)    ' (DWORD) Server Key (Double Hash Only) - Optional
            ' (DWORD) Cookie (Cookie Hash Only) - Optional
            Hash(2) = Mid(d, 4, Len(d) - 3)               ' (VOID) Data to be hashed
            
            With sPB
                .InsertDWORD &H1C                               ' (DWORD) Size of Data
                .InsertDWORD &H1                                ' (DWORD) Flags - unused
                .InsertNonNTString Hash(0) & Hash(1) & Hash(2)  ' Optional Data then Hash Data
                .SendBNLSPacket &HB                             ' BNLS_HASHDATA
            End With
        ElseIf cB = 2 Then
            'AddChat Color.Carrot, "[Debug] ", Color.Bot, "CB: ", Color.Message, cB
            With sPB
                If SPass = True Then
                    If BNET.varDebugMode = 1 Then
                        Select Case BNET.Product
                            Case "RTSJ", "RHSS", "LTRD", "RHSD": cpID = "[C 0x29] "
                            Case Else: cpID = "[C 0x3A] "
                        End Select
                    End If
                    AddChat Color.BotInfo, svrBNCS, Color.Bot, cpID & "Attempting to log on..."
                    .InsertDWORD ClientToken                                    ' (DWORD) Client Token
                    .InsertDWORD ServerToken                                    ' (DWORD) Server Token
                    .InsertNonNTString Mid(d, 4, Len(d) - 3)                    ' (DWORD) [5] Password Hash
                    .InsertNTString BNET.username
                    
                    ' Debug
                    'If BNET.varDebugMode Then
                    '    AddChat Color.Carrot, "[Debug] ", Color.Bot, "Client Token: ", Color.Message, ClientToken
                    '    AddChat Color.Carrot, "[Debug] ", Color.Bot, "Server Token: ", Color.Message, ServerToken
                    'End If
                    
                    Select Case BNET.Product
                        Case "RTSJ", "RHSS", "LTRD", "RHSD": .SendPacket &H29   ' SID_LOGONRESPONSE
                        Case Else: .SendPacket &H3A                             ' SID_LOGONRESPONSE2
                    End Select
                    
                    SPass = False
                    cB = 0
                Else
                    .InsertDWORD ClientToken                        ' (DWORD)      Client key
                    .InsertNonNTString Mid(d, 4, Len(d) - 3)        ' (DWORD[5])   Hashed realm password
                    .InsertNTString BNET.Realm                      ' (STRING)     Realm title
                    .SendPacket &H3E                                ' SID_LOGONREALMEX
                    cB = 0
                End If
            End With
        End If
    ElseIf HType = 2 Then
        'AddChat Color.Carrot, "[Debug] ", Color.Bot, "HType: ", Color.Message, HType
        With sPB
            .InsertNonNTString Mid(d, 4, Len(d) - 3)
            .InsertNTString BNET.username
            .SendPacket &H2A
        End With
    ElseIf HType = 3 Then
        'AddChat Color.Carrot, "[Debug] ", Color.Bot, "HType: ", Color.Message, HType
        Static Hash2 As String
        cB = cB + 1
        If cB = 1 Then
            Hash(0) = sPB.MakeDWORD(ClientToken)            ' (DWORD) Client Key (Double Hash Only)
            Hash(1) = sPB.MakeDWORD(ServerToken)            ' (DWORD) Server Key (Double Hash Only)
            Hash(2) = Mid(d, 4, Len(d) - 3)                 ' (DWORD) Cookie (Cookie Hash Only)
            With sPB
                .InsertDWORD &H1C                               ' (DWORD) Size of Data
                .InsertDWORD &H1                                ' (DWORD) Flags
                .InsertNonNTString Hash(0) & Hash(1) & Hash(2)
                .SendBNLSPacket &HB                             ' BNLS_HASHDATA
            End With
        End If
        If cB = 2 Then
            Hash2 = Mid(d, 4, Len(d) - 3)
            With sPB
                .InsertDWORD "&H" & Len(BNET.NewPass)
                .InsertDWORD &H0
                .InsertNonNTString BNET.NewPass
                .SendBNLSPacket &HB
            End With
        End If
        If cB = 3 Then
            With sPB
                .InsertDWORD ClientToken
                .InsertDWORD ServerToken
                .InsertNonNTString Hash2
                .InsertNonNTString Mid(d, 4, Len(d) - 3)
                .InsertNTString BNET.username
                .SendPacket &H31
            End With
        End If
    End If
End Sub

' BNLS_CDKEY_EX (0x0C)
Private Sub rcvBNLSCDKeyEx()
Dim CDKeyNumReq As Byte, CDKeyNumSuccess As Byte, BitMask As Long
          
     With bnlsPB
         Cookie = .GetDWORD             ' (DWORD) Cookie
         CDKeyNumReq = .GetByte         ' (BYTE) Number of CD-keys requested
         CDKeyNumSuccess = .GetByte     ' (BYTE) Number of successfully encrypted CD-keys
         
         If Not CDKeyNumReq = CDKeyNumSuccess Then GoTo Failed
         
         BitMask = .GetDWORD            ' (DWORD) Bit mask
         ClientToken = .GetDWORD        ' (DWORD) Client session key
         CDKeyHash = .GetRaw(36)        ' (DWORD) [9] CD-key data
    
         Select Case BNET.Product
             Case "PX2D", "PX3W"
                 .Skip 4                    ' (DWORD) Client session key
                 CDKey2Hash = .GetRaw(36)   ' (DWORD) [9] CD-key data
         End Select
     End With

     Call sndAuthCheck
     Exit Sub
Failed:
    AddChat Color.BotInfo, svrBNLS, Color.Error, "Unsuccessful number of encrypted keys."
    Call frmMain.Disconnect
End Sub

' BNLS_CHOOSENLSREVISION (0x0D)
Private Sub rcvBNLSChooseNLSRev()
    ' (BOOLEAN) Success code (32-bit) - Default revision is 1
    ' If the Success code is TRUE, the revision number was recognized by the server
    '   and will be used. If it is FALSE, the revision number was rejected by the server
    '   and this request is ignored.
    If Not bnlsPB.GetByte = 1 Then
        AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.Error, "NLS revision number rejected."
        Call frmMain.Disconnect
    Else
        'AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.BotInfo, "NLS revision number accepted."
        If BNET.Product = "3RAW" Or BNET.Product = "PX3W" Then
            ' Send Server Signature
            Dim serverIP As String, spltServer() As String, i As Integer
            serverIP = frmMain.wsBNET.RemoteHostIP
            spltServer() = Split(serverIP, ".")
            serverIP = ""
            
            For i = 0 To UBound(spltServer())
                serverIP = Hex(spltServer(i)) & serverIP
            Next i
            
            With sPB
                .InsertDWORD Val("&H" & serverIP)   ' (DWORD) Server IP
                .InsertNonNTString ServerSignature  ' (BYTE) [128] Signature
                .SendBNLSPacket &H11                ' BNLS_VERIFYSERVER
            End With
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
    End If
End Sub

' BNLS_AUTHORIZE - Defunct
Private Sub rcvBNLSAuth()
Dim key As Long, key2 As Long
Dim d As String: d = bnlsPB.GetBuffer
    ' (DWORD)[8] Data for SID_AUTH_ACCOUNTLOGON
    key2 = GetDWORD(Mid(d, 4, 4))
    key = BNLSChecksum("FyreChat", key2)
    With sPB
        .InsertDWORD key
        .SendBNLSPacket &HF
    End With
End Sub

' BNLS_AUTHORIZEPROOF - Defunct
Private Sub rcvBNLSAuthProof()
    sPB.InsertDWORD GetBNLSByte()   ' (DWORD) Checksum
    sPB.SendBNLSPacket &H10         ' BNLS_REQUESTVERSIONBYTE
End Sub

' BNLS_VERSIONCHECKEX2 (0x1A)
Private Sub rcvBNLSVerCheckEx2()
Dim VersionCode As Long
    With bnlsPB
        If Not .GetBoolean Then         ' (BOOLEAN) Success
            AddChat Color.BotInfo, svrBNLS, Color.Bot, spID, Color.Error, "Failed Version Check."
            frmMain.Disconnect
            Exit Sub
        End If
        
        Version = .GetDWORD             ' (DWORD) Version
        CheckSum = .GetDWORD            ' (DWORD) Checksum
        VerCheckStatString = .GetString ' (STRING) Version check stat string
        Cookie = .GetDWORD              ' (DWORD) Cookie
        VersionCode = .GetDWORD         ' (DWORD) The latest version code for this product
    End With
    
    If Not VersionCode = VerByte Then VerByte = VersionCode

    Select Case BNET.Product
        Case "PX2D", "PX3W"
            With sPB
                .InsertDWORD &H0            ' (DWORD) Cookie
                .InsertByte &H2             ' (BYTE) Number of CD-keys to encrypt
                .InsertDWORD &H1            ' (DWORD) Flags
                .InsertDWORD ServerToken    ' (DWORD) Server session key(s), depending on flags
                .InsertNTString BNET.CDKey  ' (STRING) CD-keys. No dashes or spaces
                .InsertNTString BNET.CDKey2 '   cont.
                .SendBNLSPacket &HC         ' BNLS_CDKEY_EX (SEXP, D2DV, W2BN, W3XP, STAR, JSTR, D2XP, WAR3)
            End With
        Case "RTSJ", "RHSS", "LTRD", "RHSD": sndReportVer
        Case Else
            With sPB
               .InsertDWORD ServerToken     ' (DWORD) Server Token
               .InsertNTString BNET.CDKey   ' (STRING) CD key
               .SendBNLSPacket &H1          ' BNLS_CDKEY (SSHR, SEXP, DSHR, D2DV, W2BN, W3XP, STAR, JSTR, DRTL, D2XP, WAR3, WOW)
            End With
    End Select
End Sub

 ' BNLS_REQUESTVERSIONBYTE (0x10)
Private Sub rcvBNLSReqVerByte()
'Dim Product As String
    bnlsPB.Skip 4               ' (DWORD) Product
    VerByte = bnlsPB.GetDWORD   ' (DWORD) Version byte
    
    'AddChat Color.Carrot, "[Debug] ", Color.BotInfo, "Product: ", Color.Message, Product, Color.BotInfo, ", VerByte: ", Color.Message, Hex(VerByte)
    
    ' Ininiate Connection
    AddChat Color.BotInfo, svrBNCS, Color.Bot, "Connecting to Battle.net server ", Color.Message, BNET.BNCSServer, Color.Bot, "..."
    frmMain.wsBNET.Close
    frmMain.wsBNET.Connect BNET.BNCSServer, 6112
End Sub

' BNLS_VERIFYSERVER (0x11)
Private Sub rcvBNLSVerifyServ()
    ' (BOOLEAN) Success (32-bit)
    ' If Success is TRUE, the signature matches the server's IP.
    ' If Success is FALSE, the signature does not match.
    If bnlsPB.GetByte Then
        ' BNLS Version Check
        With sPB
            .InsertDWORD GetBNLSByte()                      ' (DWORD) Product ID
            .InsertDWORD &H0                                ' (DWORD) Flags (None, set to 0)
            .InsertDWORD &H1                                ' (DWORD) Cookie
           .InsertDWORD CLng(Split(AuthTimeStamp, " ")(1))  ' (FILETIME) Timestamp for version check archive
            .InsertDWORD CLng(Split(AuthTimeStamp, " ")(0)) '   cont.
            .InsertNTString MPQName                         ' (STRING) Version check archive filename
            .InsertNTString AuthHash                        ' (STRING) Checksum formula
            .SendBNLSPacket &H1A                            ' BNLS_VERSIONCHECKEX2
        End With
    Else
        AddChat Color.BotInfo, svrBNLS, Color.Error, "Server signature could not be verified."
        Call frmMain.Disconnect
    End If
End Sub


Private Function GetDWORD(d As String) As Long
Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal d, 4)
    GetDWORD = lReturn
End Function
    
Private Sub InitCRC32()
Dim i As Long, J As Long, K As Long, XorVal As Long
Dim CRC32_POLYNOMIAL As Long
Static CRC32Initialized As Boolean
    If CRC32Initialized Then Exit Sub
    CRC32Initialized = True
    For i = 0 To 255
        K = i
        For J = 1 To 8
            If K And 1 Then XorVal = CRC32_POLYNOMIAL Else XorVal = 0
            If K < 0 Then K = ((K And &H7FFFFFFF) \ 2) Or &H40000000 Else K = K \ 2
            K = K Xor XorVal
        Next
        CRC32Table(i) = K
    Next
End Sub

Private Function CRC32(ByVal d As String) As Long
    Dim i As Long, J As Long
    
    Call InitCRC32
    
    CRC32 = &HFFFFFFFF
    For i = 1 To Len(d)
        J = CByte(Asc(Mid(d, i, 1))) Xor (CRC32 And &HFF&)
        If CRC32 < 0 Then CRC32 = ((CRC32 And &H7FFFFFFF) \ &H100&) Or &H800000 Else CRC32 = CRC32 \ &H100&
        CRC32 = CRC32 Xor CRC32Table(J)
    Next
    CRC32 = Not CRC32
End Function

Public Function BNLSChecksum(ByVal password As String, ByVal ServerCode As Long) As Long
    BNLSChecksum = CRC32(password & Right("0000000" & Hex(ServerCode), 8))
End Function

Public Function GetBNLSByte() As Long
    Select Case BNET.Product
        Case "RATS": GetBNLSByte = &H1 ' 0x01: StarCraft
        Case "PXES": GetBNLSByte = &H2 ' 0x02: StarCraft Brood War
        Case "NB2W": GetBNLSByte = &H3 ' 0x03: Warcraft II Battle.net Edition
        Case "VD2D": GetBNLSByte = &H4 ' 0x04: Diablo II
        Case "PX2D": GetBNLSByte = &H5 ' 0x05: Diablo II: Lord of Destruction
        Case "RTSJ": GetBNLSByte = &H6 ' 0x06: Starcraft Japanese
        Case "3RAW": GetBNLSByte = &H7 ' 0x07: Warcraft III
        Case "PX3W": GetBNLSByte = &H8 ' 0x08: Warcraft III: The Frozen Throne
        Case "LTRD": GetBNLSByte = &H9 ' 0x09: Diablo Retail
        Case "RHSD": GetBNLSByte = &HA ' 0x0A: Diablo Shareware
        Case "RHSS": GetBNLSByte = &HB ' 0x0B: StarCraft Shareware
        Case "MD3W": GetBNLSByte = &HB ' 0x0C: Warcraft III: Demo
        Case Else: AddChat Color.BotInfo, "[Bot]: ", Color.Error, "Cannot find your game type!"
    End Select
End Function

Public Function PNBNLS(p As Byte) As String
Select Case p
    Case &H0: PNBNLS = "BNLS_NULL"
    Case &H1: PNBNLS = "BNLS_CDKEY"
    Case &H2: PNBNLS = "BNLS_LOGONCHALLENGE"
    Case &H3: PNBNLS = "BNLS_LOGONPROOF"
    Case &H4: PNBNLS = "BNLS_CREATEACCOUNT"
    Case &H5: PNBNLS = "BNLS_CHANGECHALLENGE"
    Case &H6: PNBNLS = "BNLS_CHANGEPROOF"
    Case &H7: PNBNLS = "BNLS_UPGRADECHALLENGE"
    Case &H8: PNBNLS = "BNLS_UPGRADEPROOF"
    Case &H9: PNBNLS = "BNLS_VERSIONCHECK"
    Case &HA: PNBNLS = "BNLS_CONFIRMLOGON"
    Case &HB: PNBNLS = "BNLS_HASHDATA"
    Case &HC: PNBNLS = "BNLS_CDKEY_EX"
    Case &HD: PNBNLS = "BNLS_CHOOSENLSREVISION"
    Case &HE: PNBNLS = "BNLS_AUTHORIZE"
    Case &HF: PNBNLS = "BNLS_AUTHORIZEPROOF"
    Case &H10: PNBNLS = "BNLS_REQUESTVERSIONBYTE"
    Case &H11: PNBNLS = "BNLS_VERIFYSERVER"
    Case &H12: PNBNLS = "BNLS_RESERVESERVERSLOTS"
    Case &H13: PNBNLS = "BNLS_SERVERLOGONCHALLENGE"
    Case &H14: PNBNLS = "BNLS_SERVERLOGONPROOF"
    Case &H18: PNBNLS = "BNLS_VERSIONCHECKEX"
    Case &H1A: PNBNLS = "BNLS_VERSIONCHECKEX2"
    Case &H7D: PNBNLS = "BNLS_WARDEN"
    Case &HFF: PNBNLS = "BNLS_IPBAN"
    Case Else: PNBNLS = "UNKNOWN"
End Select
End Function
