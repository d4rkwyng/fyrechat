Attribute VB_Name = "modD2GS"
Option Explicit
Private d2gsPB As New clsBuffer
Private spID As String

Public Sub ParseD2GS(d As String)
Dim pID As Byte, Msg As String
    With d2gsPB
        .Clear
        .SetBuffer d
    
        pID = .GetByte              '(BYTE) Message ID
        Msg = Mid(.GetBuffer, 2)    '(VOID) Message Data
    End With
    
    If BNET.varAdvDebug = 1 Then DisplayAdvDebug "D2GS", d, pID
    
    spID = ""
    If BNET.varDebugMode = 1 Then: spID = "[S 0x" & Hex(pID) & "] "

    Select Case pID
        Case &HAF: rcvStartLogon
    End Select
End Sub

Private Sub rcvStartLogon()
    sndGameLogon
End Sub

Private Sub sndGameLogon()
    With sPB
        ' (DWORD) D2GS Server Hash
        ' (WORD) D2GS Server Token
        '.InsertDWORD "&H" & d2charClass ' (BYTE) Character ID
        '.InsertDWORD &HB                ' (DWORD) Version byte (Currently 0x0B)
        '.InsertDWORD &HED5DCC50         ' (DWORD) Unknown - Suggested Const (0xED5DCC50)
        '.InsertDWORD &H91A519B6         ' (DWORD) Unknown - Suggested Const (0x91A519B6)
        '.InsertByte &H0                 ' (BYTE) Unknown - Suggested (0x00)
        '.InsertNTString d2char          ' (STRING) Character name
        ' (VOID) *See user-comment below
    End With
End Sub

Public Function PNCD2GS(p As Byte) As String
    Select Case p
        Case &H1: PNCD2GS = "D2GS_WALKTOLOCATION"
        Case &H2: PNCD2GS = "D2GS_WALKTOENTITY"
        Case &H3: PNCD2GS = "D2GS_RUNTOLOCATION"
        Case &H4: PNCD2GS = "D2GS_RUNTOENTITY"
        Case &H5: PNCD2GS = "D2GS_LEFTSKILLONLOCATION"
        Case &H6: PNCD2GS = "D2GS_LEFTSKILLONENTITY"
        Case &H7: PNCD2GS = "D2GS_LEFTSKILLONENTITYEX"
        Case &H8: PNCD2GS = "D2GS_LEFTSKILLONLOCATIONEX"
        Case &H9: PNCD2GS = "D2GS_LEFTSKILLONENTITYEX2"
        Case &HA: PNCD2GS = "D2GS_LEFTSKILLONENTITYEX3"
        Case &HC: PNCD2GS = "D2GS_RIGHTSKILLONLOCATION"
        Case &HD: PNCD2GS = "D2GS_RIGHTSKILLONENTITY"
        Case &HE: PNCD2GS = "D2GS_RIGHTSKILLONENTITYEX"
        Case &HF: PNCD2GS = "D2GS_RIGHTSKILLONLOCATIONEX"
        Case &H10: PNCD2GS = "D2GS_RIGHTSKILLONENTITYEX2"
        Case &H11: PNCD2GS = "D2GS_RIGHTSKILLONENTITYEX3"
        Case &H13: PNCD2GS = "D2GS_INTERACTWITHENTITY"
        Case &H14: PNCD2GS = "D2GS_OVERHEADMESSAGE"
        Case &H15: PNCD2GS = "D2GS_CHATMESSAGE"
        Case &H16: PNCD2GS = "D2GS_PICKUPITEM"
        Case &H17: PNCD2GS = "D2GS_DROPITEM"
        Case &H18: PNCD2GS = "D2GS_ITEMTOBUFFER"
        Case &H19: PNCD2GS = "D2GS_PICKUPBUFFERITEM"
        Case &H1A: PNCD2GS = "D2GS_ITEMTOBODY"
        Case &H1B: PNCD2GS = "D2GS_SWAP2HANDEDITEM"
        Case &H1C: PNCD2GS = "D2GS_PICKUPBODYITEM"
        Case &H1D: PNCD2GS = "D2GS_SWITCHBODYITEM"
        Case &H1F: PNCD2GS = "D2GS_SWITCHINVENTORYITEM"
        Case &H20: PNCD2GS = "D2GS_USEITEM"
        Case &H21: PNCD2GS = "D2GS_STACKITEM"
        Case &H22: PNCD2GS = "D2GS_REMOVESTACKITEM"
        Case &H23: PNCD2GS = "D2GS_ITEMTOBELT"
        Case &H24: PNCD2GS = "D2GS_REMOVEBELTITEM"
        Case &H25: PNCD2GS = "D2GS_SWITCHBELTITEM"
        Case &H26: PNCD2GS = "D2GS_USEBELTITEM"
        Case &H28: PNCD2GS = "D2GS_INSERTSOCKETITEM"
        Case &H29: PNCD2GS = "D2GS_SCROLLTOTOME"
        Case &H2A: PNCD2GS = "D2GS_ITEMTOCUBE"
        Case &H2D: PNCD2GS = "D2GS_UNSELECTOBJ"
        Case &H2F: PNCD2GS = "D2GS_NPCINIT"
        Case &H30: PNCD2GS = "D2GS_NPCCANCEL"
        Case &H32: PNCD2GS = "D2GS_NPCBUY"
        Case &H33: PNCD2GS = "D2GS_NPCSELL"
        Case &H38: PNCD2GS = "D2GS_NPCTRADE"
        Case &H3F: PNCD2GS = "D2GS_CHARACTERPHRASE"
        Case &H49: PNCD2GS = "D2GS_WAYPOINT"
        Case &H4F: PNCD2GS = "D2GS_TRADE"
        Case &H50: PNCD2GS = "D2GS_DROPGOLD"
        Case &H5E: PNCD2GS = "D2GS_PARTY"
        Case &H61: PNCD2GS = "D2GS_POTIONTOMERCENARY"
        Case &H68: PNCD2GS = "D2GS_GAMELOGON"
        Case &H6A: PNCD2GS = "D2GS_ENTERGAMEENVIRONMENT"
        Case &H6D: PNCD2GS = "D2GS_PING"
    End Select
End Function

Public Function PNSD2GS(p As Byte) As String
    Select Case p
        Case &H10: PNSD2GS = "D2GS_CHARTOOBJ"
        Case &H19: PNSD2GS = "D2GS_SMALLGOLDPICKUP"
        Case &H1D: PNSD2GS = "D2GS_SETBYTEATTR"
        Case &H1E: PNSD2GS = "D2GS_SETWORDATTR"
        Case &H1F: PNSD2GS = "D2GS_SETDWORDATTR"
        Case &H51: PNSD2GS = "D2GS_WORLDOBJECT"
        Case &H5C: PNSD2GS = "D2GS_(COMP)STARTGAME"
        Case &H77: PNSD2GS = "D2GS_TRADEACTION"
        Case &H7A: PNSD2GS = "D2GS_LOGONRESPONSE"
        Case &H89: PNSD2GS = "D2GS_UNIQUEEVENTS"
        Case &HAF: PNSD2GS = "D2GS_STARTLOGON"
    End Select
End Function
