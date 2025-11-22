VERSION 5.00
Begin VB.UserControl BNI 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "BNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub SetupChannel(sdata As String)
    ChannelString = sdata
    frmMain.lvwIcons.ListItems.Clear
End Sub
Public Function AddUser(User As String, product As String, flags As Long, ping As Long)
On Local Error Resume Next
    Dim flags2 As String, FColor As String, ParsedString As String
    flags2 = flags
    Call ParseStatString(product, ParsedString)
    product = Left$(product, 4)
    If product = "TAHC" Or product = "CHAT" Then flags2 = "22"
    
    With frmMain.lvwIcons
        If flags = (2 Or 18) Then
            .ListItems.Add 1, , User, , GetBnetIconCode(product, flags2)
            .ListItems(1).ListSubItems.Add 1, , , GetBnetPingCode(ping, flags2), ping
            .ListItems(1).Tag = ParsedString
            .ListItems(1).ToolTipText = "(Moderator) " & ParsedString
            FColor = Color.Op
            If User = Bnet.TrueUsername Then
                'FColor = Color.Self
                MyFlags = flags
                MyPing = ping
                OpMode = True
            End If
            .ListItems(1).ForeColor = FColor
        Else
            .ListItems.Add , , User, , GetBnetIconCode(product, flags2)
            .ListItems(.ListItems.Count).ListSubItems.Add , , , GetBnetPingCode(ping, flags2), ping
            .ListItems(.ListItems.Count).Tag = ParsedString
            .ListItems(.ListItems.Count).ToolTipText = ParsedString
            FColor = Color.User
            If User = Bnet.TrueUsername Then
                FColor = Color.Self
                MyFlags = flags
                MyPing = ping
                OpMode = False
            End If
            .ListItems(.ListItems.Count).ForeColor = FColor
        End If
    End With
    With frmMain
        .txtChannelInfo.text = ChannelString & " (" & .lvwIcons.ListItems.Count & ")"
    End With
End Function
Public Function RemoveUser(User As String)
On Local Error Resume Next
    With frmMain
        .txtChannelInfo.Clear
        .txtChannelInfo.AddItem ChannelString & " (" & .lvwIcons.ListItems.Count - 1 & ")"
        .lvwIcons.ListItems.Remove .lvwIcons.FindItem(User).index
        .txtChannelInfo.text = ChannelString & " (" & .lvwIcons.ListItems.Count & ")"
    End With
End Function
Public Sub RefreshChannelList()
    frmMain.ListRefresh.Enabled = True
End Sub
Public Sub ListRefresh_Timer()
    frmMain.lvwIcons.Refresh
    frmMain.ListRefresh.Enabled = False
End Sub

Private Sub UserControl_Initialize()
   ' UserControl.BackColor = vbBlack
    '********Icon initializers*********
    Call SetIconPaths                '*
    Call CheckBnetIcons              '*
    Call CheckPingIcons              '*
    lvwIcons.ListItems.Clear         '*
    '**********************************
    
End Sub
Private Sub UserControl_Resize()
    lvwIcons.Height = UserControl.Height - 100
    lvwIcons.Width = UserControl.Width
    lvwIcons.ColumnHeaders(1).Width = UserControl.Width - 800
    txtChannelInfo.Width = lvwIcons.Width + 20
    
End Sub
Private Function SetIconPaths()
Dim ap As String
    ap = App.Path
    ICONFile = ap & "\system\bni\icons.bni"
    PINGIconFile(1) = ap & "\system\ping\0.bmp"
    PINGIconFile(2) = ap & "\system\ping\1g.bmp"
    PINGIconFile(3) = ap & "\system\ping\2g.bmp"
    PINGIconFile(4) = ap & "\system\ping\3y.bmp"
    PINGIconFile(5) = ap & "\system\ping\4y.bmp"
    PINGIconFile(6) = ap & "\system\ping\5r.bmp"
    PINGIconFile(7) = ap & "\system\ping\6r.bmp"
    PINGIconFile(8) = ap & "\system\ping\plug.bmp"
    PINGIconFile(9) = ap & "\system\ping\1b.bmp"
    lvwIcons.BackColor = 0
End Function
Private Function InitPingIcons()
Dim i As Long
Dim i2 As Long
    
    BnetPing(7).flag = 22
    
    If NumberOfIcons = 0 Then NumberOfIcons = 1
    
    For i = 1 To UBound(PINGIconFile)
        addPic.Picture = LoadPicture(PINGIconFile(i))
        ImageList1.ListImages.Add NumberOfIcons, , addPic.Image
        BnetPing(i).IListID = NumberOfIcons
        
        NumberOfIcons = NumberOfIcons + 1
        Set addPic = Nothing
        DoEvents
        Dim strta As String
        strta = BnetPing(i).flag & " " & BnetPing(i).High & " " & BnetPing(i).Low
        lvwIcons.ListItems.Add , , strta, , BnetPing(i).IListID 'was ,1 , 1
    Next i

    BnetPing(9).Low = -1: BnetPing(9).High = 0 'negative 1
    BnetPing(1).Low = 0: BnetPing(1).High = 4 '0ms
    BnetPing(2).Low = 5: BnetPing(2).High = 199 '1g
    BnetPing(3).Low = 200: BnetPing(3).High = 300 '2g
    BnetPing(4).Low = 301: BnetPing(4).High = 400 '3y
    BnetPing(5).Low = 401: BnetPing(5).High = 500 '4y
    BnetPing(6).Low = 501: BnetPing(6).High = 600 '5r
    BnetPing(7).Low = 601: BnetPing(7).High = 1200 '6r
    BnetPing(8).Low = &H0: BnetPing(8).High = &H0 'PLUG status


End Function
Private Function CheckBnetIcons()
Dim iconscheck As String
    
    iconscheck = Dir(ICONFile)
    
    If iconscheck = "" Then
        BNIFILEEXISTS = False
        'rtbadd "if you want to see client icons then you will need to find icons.bni"
        Exit Function
    Else
        BNIFILEEXISTS = True
        'rtbadd vbyellow, "icons.bni have been located Initalizing Icon list please wait..."
        Call InitIcons
    End If
End Function
Private Function CheckPingIcons()
Dim iconscheck As String
    Dim i
    For i = 1 To UBound(PINGIconFile)
        iconscheck = Dir(PINGIconFile(i))
        
        If iconscheck = "" Then
            ICOFILEEXISTS = False
            'rtbadd "Ping icon not found: " & PINGIconFile(i) & " in order to use ping icons you must have all the necessary ping icon files"
            Exit Function
        Else
            ICOFILEEXISTS = True
        End If
    Next i
    'rtbadd vbyellow, "All ping icons have been located Initalizing Icon list please wait..."
    Call InitPingIcons
End Function
Private Function InitIcons()
Dim strFileName As String, strBuffer As String
Dim dwCount As Long, cPos As Long

    strFileName = ICONFile
    
    Open strFileName For Binary As #1
        strBuffer = Input(LOF(1), #1)
    Close #1
    'Call DebugOutput(strBuffer)
    
    dwCount = ("&H" & StrToHex(StrReverse(Mid(strBuffer, 9, 4))))
    cPos = 17 '17
    
    ReDim BnetIcon(1 To dwCount) As xBnetIcon
Dim i
    For i = 1 To dwCount
        BnetIcon(i).IconID = ("&H" & StrToHex(StrReverse((Mid(strBuffer, cPos, 4)))))
        cPos = cPos + 4
        BnetIcon(i).X = ("&H" & StrToHex(StrReverse(Mid(strBuffer, cPos, 4))))
        cPos = cPos + 4
        BnetIcon(i).Y = ("&H" & StrToHex(StrReverse(Mid(strBuffer, cPos, 4))))
        cPos = cPos + 4
        BnetIcon(i).Tag = Mid(strBuffer, cPos, 4)
        
        cPos = cPos + 4
        BnetIcon(i).Unknown = ("&H" & StrToHex(StrReverse(Mid(strBuffer, cPos, 4))))
        'cPos = cPos + 4 'Used in icons_star.bni only
        
        If i >= 9 Then      'Do not use this if statement
            cPos = cPos + 4 'for icons_star must add 4 after
        End If              '.unknowen regarless
        
        If InStr(BnetIcon(i).Tag, vbNullChar) Then: BnetIcon(i).Tag = "None"
        
        BnetIcon(i).IListID = i
        ImageList1.MaskColor = vbBlack
    Next i

'*************************LOADING THE TGA AFTER GETTING ALL THE INFOS
Dim tgaFile As New LoadTGA, Scaling As Single
    Scaling = 1
    Picture1.Cls
    tgaFile.Autoscale = False
    tgaFile.LoadTGA cPos, strFileName
    Picture1.Width = tgaFile.TGAWidth
    Picture1.Height = tgaFile.TGAHeight + 67.1
    Call tgaFile.DrawTGA(Picture1)
    DoEvents
'*************************************************
    Call SetImages         'Set the ImageList    *
'*************************************************
    For i = 1 To dwCount
        lvwIcons.ListItems.Add , , BnetIcon(i).Tag & " " & BnetIcon(i).IconID, , i 'was ,1 , 1
    Next i
End Function

Private Function GetBnetIconCode(ByVal statstring As String, ByVal flags As Long) As Integer
Dim strstProduct As String
Dim Values() As String
Dim flId As Long
Dim stproduct
Dim i
    Values() = Split(Mid(statstring, 5, Len(statstring)))
    stproduct = Mid$(statstring, 1, 4)

    If BNIFILEEXISTS = True Then
        Select Case flags
            Case 22, 16:
            Case 0, 22:
                'Do nothing not related to this function
            Case Else:
                For i = 1 To UBound(BnetIcon)
                    flId = BnetIcon(i).IconID
                    If (flId And flags) = flId Then
                        GetBnetIconCode = BnetIcon(i).IListID
                        Exit Function 'You found your flaged icon now exiting the ghey loop
                    End If
                Next i
        End Select
    
        For i = 1 To UBound(BnetIcon)
            If stproduct = BnetIcon(i).Tag Then
                GetBnetIconCode = BnetIcon(i).IListID
                Exit Function 'You found your icon now exiting the ghey loop
            End If
        Next i
    End If
    'NOTE: If it makes it to this point that means either something screwed up cough
    '      or you need to update your icons.bni file so set up the UNKNOWEN Icon here
    '      unless you have implemented file version checking on this file to dl it.
    '      GetBnetIcon = otherIcons.unknowen
End Function
Private Function GetBnetPingCode(ByVal ping As Long, ByVal flags As Long) As Integer
    Dim i
    If ICOFILEEXISTS = True Then
        Select Case flags
            Case 22:
                GetBnetPingCode = BnetPing(8).IListID
                Exit Function 'You found your flaged icon now exiting the ghey loop
            Case Else:
        End Select
    
        For i = 1 To UBound(BnetPing)
            If (ping >= BnetPing(i).Low) And (ping <= BnetPing(i).High) Then
                GetBnetPingCode = BnetPing(i).IListID
                Exit Function
            End If
        Next i
    End If
    GetBnetPingCode = BnetPing(7).IListID
End Function
Private Function StrToHex(ByVal String1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
    For i = 1 To Len(String1)
        strTemp = Hex(Asc(Mid$(String1, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp
    Next i
    StrToHex = strReturn
End Function
Private Function SetImages()
Dim AddThis As Integer, x2 As Integer, x3 As Integer, pos As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer

    pos = 0
    pos2 = 0
    pos3 = 14
    pos4 = 32
    AddThis = 14
Dim X
Dim Y
Dim i
    For X = 1 To UBound(BnetIcon)
        For Y = pos To pos3
            For i = pos2 To pos4
                SetPixel addPic.hdc, x2, x3, GetPixel(Picture1.hdc, i, Y)
                x2 = x2 + 1
            Next i
            x2 = 0
            x3 = x3 + 1
        Next Y
        pos = pos + AddThis
        pos3 = pos3 + AddThis
        x3 = 0
        x2 = x2
        ImageList1.ListImages.Add X, , addPic.Image
        Set addPic = Nothing
    Next X
    NumberOfIcons = X
End Function






