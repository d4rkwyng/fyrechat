VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MsWinSck.ocx"
Begin VB.Form frmMain 
   Caption         =   "FyreChat 2.1"
   ClientHeight    =   7185
   ClientLeft      =   7920
   ClientTop       =   1.04565e5
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsD2GS 
      Left            =   360
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstChanSave 
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "ImgLst"
      SmallIcons      =   "ImgLst"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Plugins"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Timer tmrNews 
      Enabled         =   0   'False
      Left            =   2640
      Top             =   3240
   End
   Begin VB.Timer tmrWL_Upload 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1200
      Top             =   3600
   End
   Begin MSWinsockLib.Winsock wsDL 
      Left            =   0
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsBNLS 
      Left            =   360
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsBNET 
      Left            =   0
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.FileListBox FileChoice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      Pattern         =   "*.bcp;oper.dll"
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComctlLib.ListView lstPlugins 
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "ImgLst"
      SmallIcons      =   "ImgLst"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Plugins"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.ComboBox User_Input 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   5985
   End
   Begin VB.TextBox txtChanName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
   Begin VB.Timer tmrIdle 
      Interval        =   1000
      Left            =   1200
      Top             =   3240
   End
   Begin VB.Timer tmrAway 
      Interval        =   60000
      Left            =   1560
      Top             =   3240
   End
   Begin VB.Timer tmrQueue 
      Interval        =   500
      Left            =   1920
      Top             =   3240
   End
   Begin VB.ListBox lstEnqueue 
      Height          =   270
      Left            =   3720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer tmrRecon 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   2280
      Top             =   3240
   End
   Begin VB.PictureBox addPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3120
      ScaleHeight     =   210
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   455
   End
   Begin MSComctlLib.ImageList BnetIcons 
      Left            =   3120
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   9000
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox Chat_Output 
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":C0C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbWhisper 
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":C139
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstChannel 
      Height          =   2895
      Left            =   6000
      TabIndex        =   7
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "ImgLst"
      SmallIcons      =   "BnetIcons"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User"
         Object.Width           =   3803
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ping"
         Object.Width           =   773
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   210
      Index           =   0
      Left            =   3720
      Picture         =   "frmMain.frx":C1B0
      Tag             =   "0ms"
      Top             =   3960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSetup 
         Caption         =   "C&onfiguration"
         Shortcut        =   {F3}
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
         Begin VB.Menu mnuClearChat 
            Caption         =   "Chat Box"
         End
         Begin VB.Menu mnuClearWhisper 
            Caption         =   "Whisper Box"
         End
         Begin VB.Menu mnuClearSend 
            Caption         =   "Send Box"
         End
         Begin VB.Menu dash3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClearAll 
            Caption         =   "All"
         End
      End
      Begin VB.Menu dash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Today's Log"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuRealms 
         Caption         =   "&Realms"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuCreateGame 
         Caption         =   "&Create Game"
         Begin VB.Menu mnuStarcraft 
            Caption         =   "&Starcraft"
         End
         Begin VB.Menu mnuDiablo2 
            Caption         =   "&Diablo 2"
         End
      End
      Begin VB.Menu dash12 
         Caption         =   "-"
      End
      Begin VB.Menu menuWebList 
         Caption         =   "&Web List"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuLinks 
         Caption         =   "Links"
         Begin VB.Menu mnuWebHelp 
            Caption         =   "Help Documentation"
         End
         Begin VB.Menu mnuAboutPlugins 
            Caption         =   "About Plugins"
         End
         Begin VB.Menu mnuChangeLog 
            Caption         =   "Change Log"
         End
         Begin VB.Menu mnuWebsite 
            Caption         =   "Official Site"
         End
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About FyreChat..."
      End
   End
   Begin VB.Menu mnuOper 
      Caption         =   "&Oper"
      Visible         =   0   'False
      Begin VB.Menu mnuOperCfg 
         Caption         =   "Configuration"
      End
      Begin VB.Menu mnuOperUsr 
         Caption         =   "Database"
      End
      Begin VB.Menu mnuOperAlias 
         Caption         =   "Aliases"
      End
      Begin VB.Menu mnuOperDoc 
         Caption         =   "Documentation"
      End
      Begin VB.Menu dash11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOperLog 
         Caption         =   "Command Log"
      End
      Begin VB.Menu mnuOperSecur 
         Caption         =   "Security Log"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuProfile 
         Caption         =   "&Profile"
      End
      Begin VB.Menu dash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBasic 
         Caption         =   "Basic"
         Begin VB.Menu mnuKick 
            Caption         =   "&Kick"
         End
         Begin VB.Menu mnuBan 
            Caption         =   "&Ban"
         End
         Begin VB.Menu dash7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDesignate 
            Caption         =   "D&esignate"
         End
         Begin VB.Menu dash8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSquelch 
            Caption         =   "&Ignore"
         End
         Begin VB.Menu mnuUnSquelch 
            Caption         =   "&UnIgnore"
         End
      End
      Begin VB.Menu mnuStats 
         Caption         =   "Stats"
         Begin VB.Menu mnuStar 
            Caption         =   "StarCraft"
         End
         Begin VB.Menu mnuSEXP 
            Caption         =   "StarCraft: Broodwar"
         End
         Begin VB.Menu mnuW2BN 
            Caption         =   "Warcraft II: Battle.net Edition"
         End
         Begin VB.Menu mnuWAR3 
            Caption         =   "Warcraft III"
         End
         Begin VB.Menu mnuW3XP 
            Caption         =   "Warcraft III: The Frozen Throne"
         End
      End
      Begin VB.Menu mnuCopyName 
         Caption         =   "Copy Name"
         Begin VB.Menu mnuSendBox 
            Caption         =   "To Send Box"
         End
         Begin VB.Menu mnuClipBoard 
            Caption         =   "To Clip Board"
         End
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayPref 
         Caption         =   "Preferences..."
      End
      Begin VB.Menu dash10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTraySignOff 
         Caption         =   "Signoff"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DownloadTotal As Long
Public WithEvents Events    As clsEvents      'Events
Attribute Events.VB_VarHelpID = -1
Private mclsToolTip         As New clsToolTip 'ToolTip
Private Reg                 As clsRegistry    'Registry
Private blDLHeader As Boolean
Private DLfileSize As Long

Private Sub Form_Load()
On Local Error Resume Next
    RunningTime = GetTickCount 'Start Application Timer
    
    LoadConfig      'Load Configuration
    LoadVariables   'Load Variables
    EnableURLDetect 'Enable URL Detection
    Display         'Format Appearance
    ToolTip         'Enable ToolTip
    DisplayAbout    'Display About Information
    CheckFirst      'Check For First Time Loading
    
    AddChat Color.Bot, "[Execution Time " & FormatCount(GetTickCount - RunningTime) & "]"
    StartLoadPlugins 'Load BCP Plugins
    
    If BNET.varAdvDebug Then
        AddChat Color.Carrot, "Advanced Debug Mode is enabled. (Note: Ignores BNCS packet 0xF for display purposes. It is also normal to experience higher latency due to Debugging of Output process.)"
    End If

    If BNET.varNews Then FC_DownloadNews
    If BNET.varAutoCon = 1 Then ConnectProc
End Sub

Private Sub StartLoadPlugins()
On Local Error Resume Next
    Dim i As Integer, FolderCount As Integer
    Dim procPlugin As String
    Dim ResPlug As Boolean: ResPlug = 0
    Dim errorMsg As Boolean: errorMsg = False
    
    AddChat Color.BotInfo, "Loading Plugins..."
    'Load all plugins into listbox
    FileChoice.Path = App.Path & "\plugins"
    FolderCount = FileChoice.ListCount
    If FolderCount = 0 Then
        AddChat Color.Error, "No Plugins Found."
        Exit Sub
    Else
        For i = 0 To FolderCount - 1
            lstPlugins.ListItems.Add , , App.Path & "\plugins\" & FileChoice.List(i), , 0
        Next i
    End If
    
    'Process all loaded plugins
    Dim fileExt As String, FileName As String
    For i = 1 To lstPlugins.ListItems.Count
        FileName = lstPlugins.ListItems(i).text
        fileExt = Mid(StrReverse(FileName), 1, 4)
        'AddChat Color.Message, Mid(StrReverse(lstPlugins.ListItems(i).text), 1, 4)
        Select Case StrReverse(fileExt)
            Case ".bcp"
                If InStr(1, FileName, "oper.bcp") Then frmMain.mnuOper.Visible = True
                ResPlug = AddPlugin(FileName)
                If ResPlug Then
                    'Plugin Successfully Loaded
                    'AddChat Color.BotInfo, "Loaded plugin: " & FileName
                Else
                    AddChat Color.BotError, "Failed to load plugin: " & FileName
                    FolderCount = FolderCount - 1
                    errorMsg = True
                End If
                ResPlug = 0 'Reset
            Case "bcpx"
                AddChat Color.BotError, "BCPX plugins are not supported."
                FolderCount = FolderCount - 1
                'ResPlug = AddPluginEx(FileName)
                'If ResPlug Then
                '    'Plugin Successfully Loaded
                '    AddChat Color.BotInfo, "Loaded plugin: " & FileName
                'Else
                '    AddChat Color.BotError, "Failed to load plugin: " & FileName
                '    FolderCount = FolderCount - 1
                '     errorMsg = True
                'End If
                'resPlug = 0 'Reset
            Case Else
                'AddChat Color.BotInfo, "File: " & FileName
                FolderCount = FolderCount - 1
        End Select
    Next i
    If errorMsg Then AddChat Color.BotInfo, "The bot has sensed that one or more plugins have failed to load. " & _
        "This could be caused by not having the C++ Runtime files in your Windows/System32 OR Windows/SysWOW64 (64-bit) directory. " & _
        "Download them here: " & "http://openfyre.net/?p=required"
    AddChat Color.BotInfo, FolderCount & " Plugins Loaded!"
End Sub

Private Sub StartUnLoadPlugins()
On Local Error Resume Next
    Dim i As Integer, ResPlug As Boolean
    ResPlug = 0
    For i = 1 To lstPlugins.ListItems.Count
        ResPlug = FreePlugin(lstPlugins.ListItems(i).text)
        If ResPlug Then
            'Plugin Successfully Unloaded
        End If
        ResPlug = 0 'Reset
    Next i
End Sub

Private Sub ToolTip()
On Local Error Resume Next
    Dim ctrl As Control
    With mclsToolTip
        Call .Create(Me)
        .MaxTipWidth = 240
        .DelayTime(ttDelayShow) = 20000
        .ToolTipHeader = "FyreChat ToolTip"
    
        For Each ctrl In Controls
          Call mclsToolTip.AddTool(ctrl)
        Next ctrl
    End With
End Sub

Public Sub CheckFirst()
    If (FirstRun) Then
        AddChat Color.Bot, "Thank you for downloading " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & "!"
        AddChat Color.Bot, "Don't forget to setup the config!"
    End If
End Sub

Public Sub LoadVariables()
On Local Error Resume Next
    Set Events = New clsEvents
    Call AddToTray(Me.icon, App.Title & " - Offline", Me)
    Me.Caption = App.Title & " " & PVersion
    
    'Strings
    PType = " (Beta)"
    PVersion = App.Major & "." & App.Minor & " Build " & App.Revision
    homepage = "http://fyrechat.openfyre.net/"
    
    'Booleans
    IdleIO = BNET.IdleInt
    awayFlag = False
    tFlag = False
    MOverwrite = True
    SendData = True
    Urealm = True
    Whisper = False
    CloseProf = True
    WPause = False
    IconsLoaded = False
    UploadIdle = 0
    ConnectAfterDL = False
    FetchNews = False
    blDLHeader = False
    resendLREx = False
    blDownloadFile = False
End Sub

Private Sub DisplayAbout()
    Dim isnot As String
    isnot = "is not"
    If modDeclares.PType = " (Beta)" Then isnot = "is not"
    AddChat Color.Bot, App.Title & " - (C) Copyright 2003-2012 Fyre - " & Format(App.Major, "#0") _
        & "." & Format(App.Minor, "#0") & " build " & Format(App.Revision, "00")
    AddChat Color.Bot, "Distribution of any kind " & isnot & " permitted by the author!"
End Sub

Public Sub Display()
On Local Error Resume Next
    If BNET.WhispWin = 0 Then
        rtbWhisper.Visible = False
        Form_Resize
    Else
        rtbWhisper.Visible = True
        Form_Resize
    End If

    Chat_Output.BackColor = Color.Background
    Chat_Output.Font = BNET.Fonts
    Chat_Output.Font.size = BNET.FSize
    
    rtbWhisper.BackColor = Color.Background
    rtbWhisper.Font = BNET.Fonts
    rtbWhisper.Font.size = BNET.FSize
    
    User_Input.BackColor = Color.Background
    User_Input.ForeColor = Color.Message
    
    lstChannel.BackColor = Color.Background
    lstChannel.Font = BNET.Fonts
    lstChannel.Font.size = BNET.FSize
    
    txtChanName.BackColor = Color.Background
    txtChanName.Font = BNET.Fonts
    txtChanName.Font.size = BNET.FSize
    txtChanName.ForeColor = Color.Message
End Sub

Public Sub DownloadLink(sDownload As String): Call DoFileDownload(sDownload): End Sub

Public Sub Form_Resize()
    Select Case WindowState
        Case 0: vbMax = False
        Case vbMaximized: vbMax = True
        Case vbMinimized
            If BNET.varSysTray Then
                App.TaskVisible = False
                Me.Hide
                Exit Sub
            End If
        Case Else: vbMax = False
    End Select
    
    If BNET.WhispWin = 0 Then
        Chat_Output.Height = Me.ScaleHeight - 450
        Chat_Output.Width = Me.Width - 3250
        Chat_Output.Left = 50
        
        lstChannel.Height = (Chat_Output.Height + txtChanName.Height) - 325
        
        Chat_Output.Top = 50
        
        User_Input.Width = Chat_Output.Width + 5
        User_Input.Top = Me.ScaleHeight - 350
        User_Input.Left = Chat_Output.Left
        User_Input.Top = Me.ScaleHeight - 375
    
        lstChannel.Top = Chat_Output.Top + 350
        lstChannel.Left = Me.Width - 3150
        lstChannel.Width = 3000
        
        txtChanName.Top = Chat_Output.Top
        txtChanName.Width = lstChannel.Width
        txtChanName.Left = lstChannel.Left
    Else
        Chat_Output.Height = Me.ScaleHeight - 1825
        Chat_Output.Width = Me.Width - 3250
        Chat_Output.Left = 50
        
        rtbWhisper.Height = 1300
        rtbWhisper.Width = Chat_Output.Width + 5
        rtbWhisper.Top = Chat_Output.Height + 100
        rtbWhisper.Left = Chat_Output.Left
        
        lstChannel.Height = (Chat_Output.Height + rtbWhisper.Height + txtChanName.Height) - 235
        
        Chat_Output.Top = 50
        
        User_Input.Width = Chat_Output.Width + 5
        User_Input.Top = Me.ScaleHeight - 350
        User_Input.Left = Chat_Output.Left
        User_Input.Top = Me.ScaleHeight - 375
        
        lstChannel.Top = Chat_Output.Top + 350
        lstChannel.Left = Me.Width - 3150
        lstChannel.Width = 3000
        
        txtChanName.Top = Chat_Output.Top
        txtChanName.Width = lstChannel.Width
        txtChanName.Left = lstChannel.Left
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If wsBNET.State = sckConnected Then Call ExitProgram(Cancel) Else: Call ExitFinal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If RespondToTray(x) <> 0 Then Call ShowFormAgain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Urealm = False
    CloseProf = False
    If wsBNET.State Then: DisconnectProc
    
    'Unload All Forms
    Dim Form As Form
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next Form
    'End
End Sub

Private Sub ExitProgram(Cancel As Integer)
    Dim Msg
    Msg = "Are you sure you want to close while connected?"
    If BNET.varConfirmExit = 0 Then
        Cancel = False
    Else
        If MsgBox(Msg, vbQuestion + vbYesNo, "FyreChat - Confirm") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    Set Msg = Nothing
    ExitFinal
End Sub

Public Sub ExitFinal()
    If wsBNET.State Then: DisconnectProc
    
    StartUnLoadPlugins
    DisableURLDetect
    RemoveFromTray
End Sub

Public Sub ConnectProc()
On Local Error Resume Next
    Disconnect
    tmrRecon.Enabled = False
    ReconCount = 0
    BnetHomeChan = BNET.HomeChannel
    InitialCon = True
    If Not FileExists(App.Path + "\icons.bni") = True Then
        AddChat Color.BotInfo, "[Bot] ", Color.Message, "icons.bni", Color.Bot, " is missing, attempting to download..."
        ConnectBNETFTP "icons.bni"
        ConnectAfterDL = True
    Else
        ContinueConnect
    End If
End Sub

Private Sub ContinueConnect()
On Local Error Resume Next
    If IconsLoaded = False Then
        SetIconPaths
        CheckBnetIcons
        CheckPingIcons
        lstChannel.ListItems.Clear
        IconsLoaded = True
    End If
    AddChat Color.BotInfo, svrBNLS, Color.Bot, "Connecting to BNLS server ", Color.Message, BNET.BNLSServer, Color.Bot, "..."
    wsBNLS.Close
    wsBNLS.Connect BNET.BNLSServer, 9367
End Sub

Private Sub menuWebList_Click()
    frmWebList.Show
End Sub

Private Sub mnuDiablo2_Click()
    If wsBNET.State = 0 Then
        MsgBox "You are not connected to Battle.net", vbOKOnly, "Bot"
        Exit Sub
    End If
    
    If frmRealm.wsMCP.State = 0 Then
        MsgBox "You are not connected to MCP", vbOKOnly, "Bot"
        Exit Sub
    End If
    
    If BNET.Product = "VD2D" Or BNET.Product = "PX2D" Then
        frmD2DV.Show
    Else
        MsgBox "You are not using Diablo 2 or Lord of Destruction!", vbOKOnly, "Bot"
    End If
End Sub

Private Sub mnuStarcraft_Click()
    If wsBNET.State = 0 Then
        MsgBox "You are not connected to Battle.net", vbOKOnly, "Bot"
        Exit Sub
    End If
    
    If BNET.Product = "RATS" Or BNET.Product = "PXES" Then
        frmSTAR.Show
    Else
        MsgBox "You are not using Starcraft or Starcraft Brood War!", vbOKOnly, "Bot"
    End If
End Sub

Public Sub TrayPop(): PopupMenu mnuTray: End Sub

Private Sub mnuOperCfg_Click(): ShellExecute hWnd, "open", App.Path & "\oper\oper.cfg", vbNullString, App.Path, 1: End Sub

Private Sub mnuOperUsr_Click(): ShellExecute hWnd, "open", App.Path & "\oper\users.txt", vbNullString, App.Path, 1: End Sub

Private Sub mnuOperAlias_Click(): ShellExecute hWnd, "open", App.Path & "\oper\Aliases.txt", vbNullString, App.Path, 1: End Sub

Private Sub mnuOperDoc_Click(): ShellExecute hWnd, "open", App.Path & "\oper\OperDoc.txt", vbNullString, App.Path, 1: End Sub

Private Sub mnuOperLog_Click(): ShellExecute hWnd, "open", App.Path & "\oper\" & Format(Now, "mmyyyy") & ".log", vbNullString, App.Path, 1: End Sub

Private Sub mnuOperSecur_Click(): ShellExecute hWnd, "open", App.Path & "\oper\security.log", vbNullString, App.Path, 1: End Sub

Private Sub mnuBotNews_Click(): FC_DownloadNews: End Sub

Private Sub mnuExit_Click(): Unload Me: End Sub

Private Sub mnuClearSendBox_Click(): User_Input.Clear: End Sub

Private Sub mnuSetup_Click(): frmConfig.Show: End Sub

Private Sub mnuTrayExit_Click(): Unload Me: End Sub

Private Sub mnuTraySignOff_Click(): DisconnectProc: End Sub

Private Sub mnuTrayPref_Click(): frmConfig.Show: End Sub

Private Sub mnuSTAR_Click(): Call CheckStats(lstUser, "STAR"): End Sub

Private Sub mnuSEXP_Click(): Call CheckStats(lstUser, "SEXP"): End Sub

Private Sub mnuW2BN_Click(): Call CheckStats(lstUser, "W2BN"): End Sub

Private Sub mnuWAR3_Click(): Call CheckStats(lstUser, "WAR3"): End Sub

Private Sub mnuW3XP_Click(): Call CheckStats(lstUser, "W3XP"): End Sub

Private Sub CheckStats(user As String, Product As String)
    If Diablo2 And (Mid(user, 1, 1) = "*") Then user = Mid(user, 2)
    Send "/stats " & user & " " & Product, wsBNET
End Sub

Private Sub mnuSendbox_Click(): User_Input.text = User_Input.text & lstUser: End Sub

Private Sub mnuClipboard_Click(): Clipboard.SetText (lstUser): End Sub

Private Sub mnuProfile_Click(): Call RequestProfile(lstChannel.SelectedItem.text): End Sub

Private Sub mnuAbout_Click(): frmAbout.Show: End Sub

Private Sub mnuChangeLog_Click(): ShellExecute Me.hWnd, "Open", homepage & "docs/ChangeLog.htm", 0&, 0&, 0&: End Sub

Private Sub mnuWebsite_Click(): ShellExecute Me.hWnd, "Open", homepage, 0&, 0&, 0&: End Sub

Private Sub mnuWebHelp_Click(): ShellExecute Me.hWnd, "Open", homepage & "docs/Help.htm", 0&, 0&, 0&: End Sub

Private Sub mnuAboutPlugins_Click(): ShellExecute Me.hWnd, "Open", homepage & "docs/AboutPlugins.htm", 0&, 0&, 0&: End Sub

Private Sub lstChannel_DblClick()
On Local Error Resume Next
    Call RequestProfile(lstChannel.SelectedItem.text)
End Sub

Private Sub mnuClearChat_Click(): Chat_Output.text = "": End Sub

Private Sub mnuClearWhisper_Click(): rtbWhisper.text = "": End Sub

Private Sub mnuClearAll_Click(): ClearBuffers: End Sub

Private Sub mnuConnect_Click(): ConnectProc: End Sub

Private Sub mnuDisconnect_Click(): DisconnectProc: End Sub

Private Sub mnuLog_Click(): ShellExecute hWnd, "open", App.Path & "\logs\" & Format(Now, "mm-dd-yyyy") & _
    ".txt", vbNullString, App.Path, 1: End Sub
    
Private Sub mnuRealms_Click()
    If wsBNET.State = 0 Then
        MsgBox "You are not connected to Battle.net", vbOKOnly, "Bot"
        Exit Sub
    End If
    If Diablo2 Then
        frmRealm.Show
    Else
        MsgBox "You are not on Diablo 2 or Lord of Destruction!", vbOKOnly, "Bot"
    End If
End Sub

Private Sub lstChannel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
    lstUser = lstChannel.SelectedItem.text
    If Diablo2 Then lstUser = "*" & lstUser
    If Button = 2 Then PopupMenu mnuList
End Sub

Private Sub mnuWho_Click(): Send "/whois " & lstUser, wsBNET: End Sub

Private Sub mnuBan_Click(): Send "/ban " & lstUser, wsBNET: End Sub

Private Sub mnuKick_Click(): Send "/kick " & lstUser, wsBNET: End Sub

Private Sub mnuDesignate_Click(): Send "/designate " & lstUser, wsBNET: End Sub

Private Sub mnuSquelch_Click(): Send "/squelch " & lstUser, wsBNET: End Sub

Private Sub mnuUnsquelch_Click(): Send "/unsquelch " & lstUser, wsBNET: End Sub

Private Sub txtChanName_KeyPress(KeyAscii As Integer)
     User_Input.SetFocus
     SendKeys Chr(KeyAscii)
     KeyAscii = 0
End Sub
Private Sub Chat_Output_KeyPress(KeyAscii As Integer)
     User_Input.SetFocus
     SendKeys Chr(KeyAscii)
     KeyAscii = 0
End Sub
Private Sub lstChannel_KeyPress(KeyAscii As Integer)
     User_Input.SetFocus
     SendKeys Chr(KeyAscii)
     KeyAscii = 0
End Sub
Private Sub rtbWhisper_KeyPress(KeyAscii As Integer)
     User_Input.SetFocus
     SendKeys Chr(KeyAscii)
     KeyAscii = 0
End Sub

Public Sub User_Input_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
    Dim i As Integer, tNet As Long, PluginLoaded As Boolean
    Dim verInfo As String
    verInfo = App.Title & " - (C) Copyright 2003-2012 Fyre - " & App.Title & " " & _
        Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & " build " & _
        Format(App.Revision, "00") & " - (fyrechat.openfyre.net)"
    tNet = 0
    'If wsBnet.State = 7 Then tNet = 1
    If (KeyAscii = 13) And (Len(User_Input.text) >= 1) Then
        If (Mid(User_Input.text, 1, 1) = "/") Then
            PluginLoaded = CallPluginCommandNotifications(Mid(User_Input.text, 2), vbNullString, tNet)
            If PluginLoaded Then
                SleepEx 0, 1
            Else
                Dim CmdData() As String, EventText As String
                EventText = Mid(User_Input.text, 2)
                If InStr(EventText, " ") Then
                    CmdData = Split(EventText, " ", 2)
                    EventText = CmdData(0)
                End If
                Select Case EventText
                    Case "battlenet"
                        Send Mid(User_Input.text, 12), wsBNET
                        User_Input.text = ""
                        KeyAscii = "0"
                        Exit Sub
                    Case "about", "version"
                        AddChat Color.Bot, verInfo
                    Case "versionout"
                        Dim toUser() As String
                        toUser() = Split(User_Input.text, " ")
                        Send "/m " & toUser(1) & " " & verInfo, wsBNET
                        Erase toUser()
                    Case "profile"
                        If InStr(User_Input.text, "*") Then
                            Call RequestProfile(WildCard(Mid(User_Input.text, 10), lstChannel))
                        Else
                            Call RequestProfile(Mid(User_Input.text, 10))
                        End If
                    Case "recorddata"
                        Call RequestRecordData(BNET.TrueUsername)
                    Case "uptime"
                        If wsBNET.State = 7 Then
                            AddChat Color.BotInfo, "[Bot] ", Color.Bot, "Connected for: ", Color.Message, FormatCount(GetTickCount - connecttime)
                        End If
                        AddChat Color.BotInfo, "[Bot] ", Color.Bot, "Application Uptime: ", Color.Message, FormatCount(GetTickCount - RunningTime)
                        AddChat Color.BotInfo, "[Bot] ", Color.Bot, "Computer Uptime:   ", Color.Message, FormatCount(GetTickCount)
                    Case "loadplugin"
                        Dim Plugin As String, ResPlug As Boolean
                        Plugin = Mid(User_Input.text, 13)
                        ResPlug = AddPlugin(Plugin)
                        If ResPlug Then
                            AddChat Color.BotInfo, "Successfully loaded plugin " & Plugin
                            lstPlugins.ListItems.Add , , Plugin, , i
                        Else
                            AddChat Color.BotError, "Failed to load plugin " & Plugin
                        End If
                    Case "unloadplugin"
                        Plugin = Mid(User_Input.text, 15)
                        ResPlug = RemovePlugin(Plugin)
                        If ResPlug Then
                            AddChat Color.BotInfo, "Successfully unloaded plugin " & Plugin
                            For i = 1 To lstPlugins.ListItems.Count
                                If InStr(LCase(lstPlugins.ListItems(i).text), LCase(Plugin)) Then lstPlugins.ListItems.Remove (i)
                            Next i
                        Else
                            AddChat Color.BotError, "Failed to unload plugin " & Plugin
                        End If
                    Case "listplugins"
                        Dim tmpPlugin As String
                        AddChat Color.BotInfo, lstPlugins.ListItems.Count & " Loaded Plugin(s)"
                        For i = 0 To lstPlugins.ListItems.Count
                            AddChat Color.info, lstPlugins.ListItems(i).text
                        Next i
                    Case "setcharactername"
                        Dim tmpSetChar() As String
                        tmpSetChar() = Split(User_Input.text, " ")
                        If Len(tmpSetChar(1)) > 1 Then
                            frmConfig.txtCharacter.text = tmpSetChar(1)
                            frmConfig.cmdSave_Click
                        End If
                        Erase tmpSetChar()
                    Case "realmconnect"
                        If Not BNET.Product = "VD2D" Or BNET.Product = "PX2D" Then
                            AddChat Color.BotInfo, "[Bot] ", Color.Error, "Not using Diablo 2 or Lord of Destruction."
                        Else
                            If frmRealm.wsMCP.State = sckConnected Then
                                AddChat Color.BotInfo, "[Bot] ", Color.Error, "Already connected to Realms."
                            Else
                                sndLogRealmEx BNET.Realm
                            End If
                        End If
                    Case "setusername"
                        Dim tmpSetUser() As String
                        tmpSetUser() = Split(User_Input.text, " ")
                        If Len(tmpSetUser(1)) > 1 Then
                            frmConfig.txtUsername.text = tmpSetUser(1)
                            frmConfig.cmdSave_Click
                        End If
                        Erase tmpSetUser()
                    Case "setpassword"
                        Dim tmpSetPass() As String
                        tmpSetPass() = Split(User_Input.text, " ")
                        If Len(tmpSetPass(1)) > 1 Then
                            frmConfig.txtPassword.text = tmpSetPass(1)
                            frmConfig.cmdSave_Click
                        End If
                        Erase tmpSetPass()
                    Case "setserver"
                        Dim tmpSetServ() As String
                        tmpSetServ() = Split(User_Input.text, " ")
                        If Len(tmpSetServ(1)) > 1 Then
                            frmConfig.cboBattlenet.text = tmpSetServ(1)
                            frmConfig.cmdSave_Click
                        End If
                        Erase tmpSetServ()
                    Case "setproduct"
                        Dim tmpSetProduct() As String
                        tmpSetProduct() = Split(User_Input.text, " ")
                        If UBound(tmpSetProduct()) = 1 Then
                            Select Case tmpSetProduct(1)
                                Case "STAR": frmConfig.cboProduct.text = "StarCraft"
                                Case "SEXP": frmConfig.cboProduct.text = "StarCraft Brood War"
                                Case "JSTR": frmConfig.cboProduct.text = "StarCraft Japan"
                                Case "SSHR": frmConfig.cboProduct.text = "StarCraft Shareware"
                                Case "DRTL": frmConfig.cboProduct.text = "Diablo"
                                Case "DSHR": frmConfig.cboProduct.text = "Diablo Shareware"
                                Case "D2DV": frmConfig.cboProduct.text = "Diablo II"
                                Case "D2XP": frmConfig.cboProduct.text = "Diablo II Lord of Destruction"
                                Case "W2BN": frmConfig.cboProduct.text = "Warcraft II Battle.net Edition"
                                Case "WAR3": frmConfig.cboProduct.text = "Warcraft III"
                                Case "W3XP": frmConfig.cboProduct.text = "Warcraft III The Frozen Throne"
                                Case Else: AddChat Color.BotError, "Not a supported product."
                            End Select
                            frmConfig.cmdSave_Click
                            Erase tmpSetProduct()
                        Else
                            AddChat Color.BotInfo, "[Bot] ", Color.Bot, "Product Options: ", Color.Message, "STAR, SEXP, JSTR, SSHR, " & _
                            "DRTL, D2DV, D2XP, W2BN, WAR3, W3XP"
                        End If
                    Case "sethome"
                        Dim tmpSetHome() As String
                        tmpSetHome() = Split(User_Input.text, " ")
                        If Len(tmpSetHome(1)) > 1 Then
                            frmConfig.txtHomeChan.text = tmpSetHome(1)
                            frmConfig.cmdSave_Click
                        End If
                        Erase tmpSetHome()
                    Case "getlag"
                        Dim getLag As String
                        getLag = ""
                        If InStr(User_Input.text, " ") Then
                            If InStr(User_Input.text, "*") Then
                                Dim gLag() As String, UserG As String
                                gLag() = Split(User_Input.text, " ")
                                getLag = DisplayFound(WildCard(gLag(1), lstChannel), "getlag")
                                AddChat Color.BotInfo, "Ping (" & Counter & "): " & vbNewLine, Color.Bot, getLag
                            Else
                                UserG = lstChannel.FindItem(WildCard(Mid(User_Input.text, 9), lstChannel))
                                getLag = lstChannel.FindItem(WildCard(Mid(User_Input.text, 9), lstChannel)).ListSubItems(1).ToolTipText
                                If Not getLag = "" Then AddChat Color.BotInfo, "Ping: " & UserG & ":" & getLag
                            End If
                        Else
                            getLag = lstChannel.FindItem(BNET.TrueUsername).ListSubItems(1).ToolTipText
                            If Not getLag = "" Then AddChat Color.BotInfo, "Your ping is: " & getLag
                        End If
                    Case "getinfo"
                        Dim getInfo As String
                        getInfo = ""
                        If InStr(User_Input.text, " ") Then
                            If InStr(User_Input.text, "*") Then
                                Dim GInfo() As String
                                GInfo() = Split(User_Input.text, " ")
                                getInfo = DisplayFound(WildCard(GInfo(1), lstChannel), "getinfo")
                                AddChat Color.BotInfo, "(" & Counter & ") Users found: " & vbNewLine, Color.Bot, getInfo
                            Else
                                getInfo = ""
                                UserG = lstChannel.FindItem(WildCard(Mid(User_Input.text, 10), lstChannel))
                                getInfo = lstChannel.FindItem(WildCard(Mid(User_Input.text, 10), lstChannel)).Tag
                                If Not getInfo = "" Then AddChat Color.BotInfo, UserG & " is using " & getInfo & "."
                            End If
                        Else
                            getInfo = lstChannel.FindItem(BNET.TrueUsername).Tag
                            AddChat Color.BotInfo, "You are using " & getInfo & "."
                        End If
                    Case "news"
                        sPB.InsertDWORD &H0
                        sPB.SendPacket &H46
                    Case "realmserverlist"
                        showRealms = True
                        sPB.SendPacket &H40
                    Case "realmgamelist"
                        If Not frmRealm.wsMCP.State = 7 Then
                            AddChat Color.BotInfo, "[Bot] ", Color.BotError, "You are not connect to MCP."
                        Else
                            With sPB
                                .InsertWORD &H2
                                .InsertDWORD &H0
                                .InsertNTString ""
                                .SendRPacket &H5
                            End With
                        End If
                    Case "realmjoingame"
                        If Not frmRealm.wsMCP.State = 7 Then
                            AddChat Color.BotInfo, "[Bot] ", Color.BotError, "You are not connect to MCP."
                        Else
                            Dim spltJGame() As String
                            spltJGame() = Split(User_Input.text, " ")
                            With sPB
                                .InsertWORD Val("&H" & 2)   ' (WORD) Request ID
                                .InsertNTString spltJGame(1)    ' (STRING) Game name
                                .InsertNTString spltJGame(2)    ' (STRING) Game Password
                                .SendRPacket &H4                ' MCP_JOINGAME
                            End With
                        End If
                    Case "botnews"
                        Call FC_DownloadNews
                    Case Else
                        Dim username As String
                        username = BNET.username
                        If Len(BNET.TrueUsername) >= 1 Then username = BNET.TrueUsername
                        ProcessBotCommand User_Input.text, username
                End Select
                Erase CmdData()
            End If
        Else
            If wsBNET.State = 7 Then
                Send User_Input.text, wsBNET
            Else
                AddChat Color.BotInfo, "[Bot] ", Color.Error, "You are not connected to Battle.net."
            End If
        End If
        DoAddToSendList User_Input.text
        User_Input.text = ""
        KeyAscii = "0"
    ElseIf (KeyAscii = 96) Then
        Dim splitFind() As String
        With lstChannel
            splitFind() = Split(StrReverse(User_Input.text), " ")
            splitFind(0) = StrReverse(splitFind(0))
            For i = 1 To .ListItems.Count
                If LCase(splitFind(0)) = Mid(LCase(.ListItems(i).text), 1, Len(splitFind(0))) Then
                    User_Input.text = ""
                    User_Input.text = Replace(User_Input.text, Left(User_Input, Len(splitFind(0))), "") & _
                        .ListItems(i).text
                    Exit For
                End If
            Next i
        End With
        Erase splitFind()
    End If
End Sub

Public Sub tmrAway_Timer()
On Local Error Resume Next
    If Not wsBNET.State = sckConnected Then Exit Sub
    If BNET.varAwayIdle = 1 Then
        If GetTickCount - LastTalk >= 60000 Then
            tFlag = False
            awayFlag = True
            Send "/away Bot has been idle for " & FormatCount(GetTickCount - LastTalk, 3) & ".", wsBNET
        End If
    Else
       If GetTickCount - LastTalk >= 300000 Then sPB.SendPacket &H0
    End If
End Sub

Private Sub tmrIdle_Timer()
    If Not blEnterChat Then Exit Sub
    If BNET.varIdle Then
        If GetTickCount - IdleIO >= (BNET.IdleInt * 1000) Then
            IdleIO = GetTickCount()
            IdleSent = BNET.Idle
            If InStr(BNET.Idle, "%") Then
                IdleSent = Replace(IdleSent, "%ver", Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & " Build " & Format(App.Revision, "00"))
                IdleSent = Replace(IdleSent, "%uptime", FormatCount(GetTickCount))
                IdleSent = Replace(IdleSent, "%botuptime", FormatCount(GetTickCount - RunningTime))
                IdleSent = Replace(IdleSent, "%connected", FormatCount(GetTickCount - connecttime))
                IdleSent = Replace(IdleSent, "%idle", FormatCount(GetTickCount - LastTalk, 4))
                IdleSent = Replace(IdleSent, "%self", BNET.TrueUsername)
                IdleSent = Replace(IdleSent, "%chan", BNET.CurrentChan)
            End If
            Send IdleSent, wsBNET
        End If
    End If
    If BNET.varCountIdle Then
        If hmm = 35 Then
            hmm = 0
            Send "Hmm", wsBNET
        End If
    End If
End Sub

Private Sub tmrQueue_Timer()
On Local Error Resume Next
If Not blEnterChat Then Exit Sub
    Dim Message As String, strText As String
    If lstEnqueue.List(0) <> "" Then
        If Message = "" Then Message = lstEnqueue.List(0)
        If InStr(lstEnqueue.List(0), Chr(&H1)) Then
            Message = Mid(lstEnqueue.List(0), 1, InStr(lstEnqueue.List(0), Chr(1)) - 1)
        End If
        If strText = "" Then strText = lstEnqueue.List(0)
        If strText = "" Then Exit Sub
        sPB.InsertNTString Message
        sPB.SendPacket &HE
        Call modFunctions.PrintUser(Message)
        lstEnqueue.RemoveItem (0)
     End If
 End Sub

Private Sub tmrRecon_Timer()
On Local Error Resume Next
    If ReconCount <= 10 Then
        If wsBNET.State = 7 Then wsBNET.Close
        ConnectProc
        ReconCount = ReconCount + 1
    Else
        AddChat Color.BotError, "Bot has halted reconnecting after " & ReconCount & " times."
        tmrRecon.Enabled = False
    End If
End Sub

Private Sub tmrNews_Timer()
On Local Error Resume Next
    Call FC_Download(homepage & "motd.txt", "download_motd.txt")
    Call FC_Download(homepage & "ad.txt", "download_ad.txt")
    Call FC_Download(homepage & "current_version.txt", "current_version.txt")
    tmrNews.Enabled = False
End Sub

Private Sub LoadConVariables()
On Local Error Resume Next
    Select Case BNET.Product
        Case "VD2D", "PX2D": Diablo2 = True
        Case "3RAW", "PX3W": Warcraft3 = True
    End Select
    connecttime = GetTickCount()
    Urealm = True
    OnStart = True
    OnChan = True
    ChanProt = False
End Sub

Private Sub wsBNET_Connect()
    AddChat Color.BotInfo, svrBNCS & "Connected to Battle.net!"
    LoadConVariables
    wsBNET.SendData Chr$(&H1)
    Select Case BNET.Product
        Case "LTRD", "RHSD": sndClientID2
        Case "RTSJ", "RHSS": sndClientID2
        Case Else: sndAuthInfo ' SID_AUTHINFO (0x50)
    End Select
End Sub

Private Sub wsBNET_DataArrival(ByVal bytesTotal As Long)
Static strBuffer As String
Dim strTemp As String, lngLen As Long
    wsBNET.GetData strTemp, vbString
    strBuffer = strBuffer & strTemp
    While Len(strBuffer) > 4
        lngLen = Val("&H" & StrToHex(StrReverse(Mid(strBuffer, 3, 2))))
        If Len(strBuffer) < lngLen Then: Exit Sub
        ParseBNCS (Left(strBuffer, lngLen))
        strBuffer = Mid(strBuffer, lngLen + 1)
    Wend
End Sub

Private Sub wsBNET_Close()
    AddChat Color.BotInfo, svrBNCS, Color.BotError, "Connection to Battle.net has been lost."
    DisconnectProc
    If BNET.varARCon = 1 Then tmrRecon.Enabled = True
End Sub

Private Sub wsBNET_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat Color.BotInfo, svrBNCS, Color.Error, Number & ": " & Description
    Disconnect
    tmrRecon.Enabled = False
End Sub

Private Sub wsBNLS_Connect()
    If InitialCon Then
        AddChat Color.BotInfo, svrBNLS & "Connected to BNLS!"
        
        ' Need to Bypass this for cached Version Bytes
        sPB.InsertDWORD GetBNLSByte()   ' (DWORD) Checksum
        sPB.SendBNLSPacket &H10         ' BNLS_REQUESTVERSIONBYTE
        
        ' Defunct Authentication
        'sPB.InsertNTString "FyreChat"
        'sPB.SendBNLSPacket &HE
        
        InitialCon = False
    End If
    
    If resendLREx Then
        AddChat Color.BotInfo, svrBNLS & "Re-Connected to BNLS!"
        sndLogRealmEx BNET.Realm
    End If
End Sub

Private Sub wsBNLS_DataArrival(ByVal bytesTotal As Long)
Dim TempData As String
    wsBNLS.GetData TempData, vbString
    ParseBNLS TempData
End Sub

Private Sub wsBNLS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat Color.BotInfo, svrBNLS, Color.Error, Number & ": " & Description
    wsBNLS.Close
End Sub

Private Sub wsD2GS_Connect()
    wsD2GS.SendData Chr(1)
    AddChat Color.BotInfo, "[D2GS] Connected to D2GS!"
End Sub

Private Sub wsD2GS_DataArrival(ByVal bytesTotal As Long)
Dim TempData As String
    wsD2GS.GetData TempData, vbString
    ParseD2GS TempData
End Sub

Private Sub wsD2GS_Close()
    AddChat Color.BotInfo, svrBNCS, Color.BotError, "Connection to D2GS has been lost!"
    wsD2GS.Close
End Sub

Private Sub wsD2GS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat Color.BotInfo, svrD2GS, Color.Error, "Error [" & Number & "] " & Description
    wsD2GS.Close
End Sub

Private Sub wsDL_Connect()
    If BNET.varDebugMode = 1 Then AddChat Color.BotInfo, "[FTP] Connected to Battle.net FTP!"
    DownloadTotal = 0
    BNETFTP (BnetFileName)
End Sub

Private Sub wsDL_DataArrival(ByVal bytesTotal As Long)
'Dim headerLen As Long, unknown As Long
Dim BannerID As Long, BannerEXT As Long, fTime As Long, fName As String
Dim fileData As String, dlfilePB As New clsBuffer
Dim strTemp As String
    On Local Error Resume Next
    
    wsDL.GetData strTemp, vbString, bytesTotal
    
    'If BNET.varAdvDebug = 1 Then DisplayAdvDebug "BNCS", strTemp, 0
    
    If Not blDLHeader Then
        With dlfilePB
            .Clear
            .SetBuffer strTemp
            ' (WORD) Header length - Does not include the length of the file.
            ' (WORD) Unknown (probably padding)
            .Skip 4
            DLfileSize = .GetDWORD          ' (DWORD) Filesize
            BannerID = .GetDWORD            ' (DWORD) Banners ID*
            BannerEXT = .GetDWORD           ' (DWORD) Banners File Extension*
            fTime = .GetFileTime            ' (FILETIME) Remote Filetime
            fName = .GetString              ' (STRING) Filename
            fileData = Mid(.GetBuffer, 24 + Len(fName) + 2)   ' (VOID) File data
        End With
    
        If fName = BnetFileName Then
            AddChat Color.BotInfo, "[Bot] ", Color.Bot, "Downloading ", Color.Message, BnetFileName, Color.Bot, "..."
        Else
            Exit Sub
        End If
        
        blDLHeader = True
    Else
        fileData = strTemp    ' (VOID) File data
    End If
    
    'fileData = Mid(strData, InStr(strData, BnetFileName) + Len(BnetFileName) + 1, Len(strData))
    'fileSize = Val("&H" & ToHex(Mid(strData, 8, 1)) & ToHex(Mid(strData, 7, 1)) & ToHex(Mid(strData, 6, 1)) & ToHex(Mid(strData, 5, 1)))
    
    
    Open App.Path & "\" & fName For Binary As #98
    Put #98, , fileData
    
    DownloadTotal = DownloadTotal + Len(fileData)
    
    If DownloadTotal = DLfileSize Then
        Close #98
        AddChat Color.BotInfo, "[Bot] ", Color.Bot, "Downloaded ", Color.Message, BnetFileName
        wsDL.Close
        blDLHeader = False
        If ConnectAfterDL Then ContinueConnect
    End If
End Sub

Private Sub wsDL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat Color.BotInfo, "[Bot] ", Color.Error, "Error [" & Number & "] " & Description
    wsDL.Close
End Sub

Public Sub DisconnectProc()
On Local Error Resume Next
    AddChat Color.BotError, "Disconnected!"
    Disconnect
End Sub

Public Sub Disconnect()
    If tmrIdle.Enabled Then tmrIdle.Enabled = False
    If tmrWL_Upload.Enabled Then tmrWL_Upload.Enabled = False
    If tmrAway.Enabled Then tmrAway.Enabled = False
    If tmrRecon.Enabled Then tmrRecon.Enabled = False
    
    lstChannel.ListItems.Clear
    mclsToolTip.RemoveTool txtChanName
    txtChanName.text = ""
    ModifyTray Me.icon, App.Title & " - Offline", Me
    Me.Caption = App.Title & " " & PVersion & " - By Fyre"
    
    awayFlag = False
    tFlag = False
    blEnterChat = False
    Urealm = False
    LastTalk = GetTickCount()
    Unload frmRealm
    
    
    wsDL.Close
    wsD2GS.Close
    frmRealm.wsMCP.Close
    wsBNET.Close
    wsBNLS.Close
End Sub

Private Sub tmrWL_Upload_Timer()
    If Not wsBNET.State = sckConnected Then Exit Sub
    If Not blWL_Enabled = 1 Then Exit Sub
    If GetTickCount - UploadIdle >= (intWL_UpTimer * 1000) Then
        If frmWebList.lstWebChan.ListItems.Count >= 1 Then
            frmWebList.WebListSave
            UploadIdle = GetTickCount()
            If frmWebList.UploadFile(strWL_FTPAddress, strWL_FTPUser, strWL_FTPPass, 21, _
                App.Path & "\channel_list.txt", strWL_FTPLoc & "channel_list.txt") Then
            Else
                AddChat Color.BotInfo, "[Bot] ", Color.Error, "Upload Error!"
            End If
        End If
    End If
End Sub

Private Sub FC_DownloadNews()
    AddChat Color.BotInfo, "Downloading News..."
    tmrNews.Enabled = True
    tmrNews.Interval = 400
End Sub

Private Sub FC_Download(ByVal location As String, ByVal SaveFile As String)
On Local Error Resume Next
    Dim data As String, hFile As Long
    SaveFile = App.Path & "/" & SaveFile
    If DownloadFile(location, SaveFile) Then
        hFile = FreeFile
        Open SaveFile For Input As #hFile
            data = Input$(LOF(hFile), hFile)
            If InStr(SaveFile, "current_version.txt") Then
                If Not data = "$null" Then
                    'If PType = " (Beta)" Then
                    'Else
                        Dim pA() As String, tA() As String
                        pA = Split(data, " ")
                        tA = Split(pA(0), ".")
                        'If App.Revision < tA(2) Then
                        '    If MsgBox("Would you like to download the update?", vbYesNo, "Build " & tA(1) & " Available!") = vbYes Then
                        '        DownloadLink StrConv(homepage & "download/FyreChat.exe", vbUnicode)
                        '    End If
                        'Else
                            AddChat Color.Bot, "Current Public Version: ", Color.Message, pA(0)
                        'End If
                    'End If
                End If
            ElseIf InStr(SaveFile, "download_motd.txt") Then
                If Not data = "$null" Then
                    AddChat Color.News, App.Title & " News: ", Color.Message, data
                End If
            ElseIf InStr(SaveFile, "download_ad.txt") Then
                If Not data = "$null" Then
                    AddChat Color.News, "Check out... ", Color.Message, data
                End If
            End If
        Close #hFile
        hFile = ""
        data = ""
        Kill (SaveFile)
    End If
End Sub

Private Sub Events_OnTalk(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
On Local Error Resume Next
    Dim eData, mData As String
    Select Case Mid$(Message, 1, 1)
        Case Chr(166)
            eData = " (Octal)"
            mData = DecodeBase8(Mid$(Message, 2, Len(Message)))
        Case "£"
            eData = " (Hex)"
            mData = hexEncrypt(Message)
        Case "·"
            eData = " (Reversed)"
            mData = StrReverse(Mid(Message, 2))
        Case Else
            eData = ""
            mData = Message
    End Select
    
    PingFlags = ""
    If BNET.varPingFlag = 1 Then PingFlags = "[" & ping & ":" & flags & "] "
    If (flags And BNFLAGS_OP) Then
        AddChat Color.PingFlags, PingFlags, Color.Carrot, "<", Color.Op, _
            username, Color.Enc, eData, Color.Carrot, "> ", Color.Message, mData
    Else
        AddChat Color.PingFlags, PingFlags, Color.Carrot, "<", Color.user, _
            username, Color.Enc, eData, Color.Carrot, "> ", Color.Message, mData
    End If

    If (BNET.varCountIdle = 1) Then hmm = hmm + 1
End Sub

Private Sub Events_OnEmote(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
On Local Error Resume Next
    PingFlags = ""
    If BNET.varPingFlag = 1 Then PingFlags = "[" & MyPing & ":" & MyFlags & "] "
    If (flags And BNFLAGS_OP) Then
        AddChat Color.PingFlags, PingFlags, Color.Carrot, "<", Color.Op, username & " " & _
            Message, Color.Carrot, ">"
    Else
        AddChat Color.PingFlags, PingFlags, Color.Carrot, "<", Color.user, username & " " & _
            Message, Color.Carrot, ">"
    End If
    
    If BNET.varCountIdle Then hmm = hmm + 1
End Sub

Private Sub Events_OnChannel(ByVal ChannelName As String, ByVal flags As Long)
On Local Error Resume Next
    lstChannel.ListItems.Clear
    frmWebList.lstWebChan.ListItems.Clear
    
    txtChanName.text = ChannelName
    BNET.CurrentChan = ChannelName
  
    OnChan = False
    AddChat Color.Message, "Joining ", Color.Carrot, "(", Color.Message, GetChannelType(flags), _
        Color.Carrot, ")", Color.Message, " Channel: ", Color.Carrot, ChannelName
    
    SetupChannel (BNET.CurrentChan)
    If LCase(BNET.CurrentChan) = "the void" Then
        Dim tmpUser As String
        tmpUser = BNET.TrueUsername
        If Diablo2 Then tmpUser = "*" & BNET.TrueUsername
        Send "/unsquelch " & tmpUser, frmMain.wsBNET
    End If
End Sub

Private Sub Events_OnUser(ByVal user As String, ByVal flags As Long, ByVal Product As String, ByVal ping As Long)
On Local Error Resume Next
    Dim i As Integer
    Dim Found As Boolean
    Dim ParsedString As String
    Dim spltuser() As String, tmpUser As String

    Call ParseStatString(Product, ParsedString)
    With frmMain.lstChannel
        For i = 1 To .ListItems.Count
            If LCase(user) = LCase(.ListItems(i).text) Then
                Found = True
                Exit For
            Else
                Found = False
            End If
        Next i
        If Found Then
            If Not (ParsedString = .FindItem(user).Tag) Then
                If (ParsedString = .FindItem(user).Tag) Then
                Else
                    .FindItem(user).Tag = ParsedString
                    .FindItem(user).ToolTipText = ParsedString
                    If BNET.varStatusUpdate Then AddChat Color.BotInfo, "Stats Update: ", Color.Bot, user & " is now here using " & ParsedString
                    Exit Sub
                End If
            End If
        Else
            AddUser user, Product, flags, ping
            frmWebList.WebListAdd user, Product, flags, ping
        End If
    End With
End Sub

Private Sub Events_OnJoin(ByVal user As String, ByVal flags As Long, ByVal Product As String, ByVal ping As Long)
On Local Error Resume Next
    Dim ParsedString As String
    Call ParseStatString(Product, ParsedString)
    
    AddUser user, Product, flags, ping
    frmWebList.WebListAdd user, Product, flags, ping
    
    If BNET.varEnterLeave = 1 Then
        If (flags And BNFLAGS_OP) Then
            AddChat Color.Op, user, Color.Join, " joined the channel using " & ParsedString & _
                "; with a ping of " & ping & "ms."
        Else
            AddChat Color.Join, user & " joined the channel using " & ParsedString & _
                "; with a ping of " & ping & "ms."
        End If
    End If
End Sub

Private Sub Events_OnLeave(ByVal user As String, ByVal flags As Long, ByVal ping As Long)
On Local Error Resume Next
    RemoveUser user
    frmWebList.WebListDel user
    
    If BNET.varEnterLeave = 1 Then
        If (flags And BNFLAGS_OP) Then
            AddChat Color.Op, user, Color.Left, " left the channel."
        Else
            AddChat Color.Left, user & " left the channel."
        End If
    End If
End Sub

Private Sub Events_OnWhisperTo(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
On Local Error Resume Next
    PingFlags = ""
    If BNET.varPingFlag = 1 Then PingFlags = "[" & MyPing & ":" & MyFlags & "] "
    AddWhisper Color.PingFlags, PingFlags, Color.Self, "<" & "To: " & username & "> ", _
        Color.Message, Message
    RWReply = username
    If Diablo2 Then RWReply = "*" & username
End Sub

Private Sub Events_OnWhisperFrom(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
On Local Error Resume Next
    PingFlags = ""
    If BNET.varPingFlag = 1 Then PingFlags = "[" & ping & ":" & flags & "] "
    AddWhisper Color.PingFlags, PingFlags, Color.WhisperFrom, "<" & "From: " & username & "> ", _
        Color.Message, Message
    WReply = username
    If Diablo2 Then WReply = "*" & WReply
End Sub

Private Sub Events_OnInfo(ByVal Message As String)
On Local Error Resume Next
    If InStr(Message, "now marked as being away") And awayFlag Then Exit Sub
    If InStr(Message, "still marked as being away") And awayFlag Then Exit Sub
    If InStr(Message, "no longer marked") And awayFlag Then
        awayFlag = False
        Exit Sub
    End If
    AddChat Color.info, Message
End Sub

Private Sub Events_OnFlags(ByVal user As String, ByVal flags As Long, ByVal Product As String, ByVal ping As Long)
On Local Error Resume Next
    If (LCase(BNET.CurrentChan) = "the void") Then AddUser user, Product, flags, ping
    If (flags And BNFLAGS_OP) Then
        tmpFlag = frmMain.lstChannel.FindItem(user).Tag
        RemoveUser user
        AddUser user, Product, flags, ping
        frmWebList.WebListAdd user, Product, flags, ping
    End If
    RefreshChannelList user, Product, flags
    frmWebList.RefreshWebChanList user, Product, flags, ping
End Sub

Private Sub Events_OnUnknown(ByVal Message As String)
    AddChat Color.BotInfo, "[Bot] Unknown: ", Color.Error, Message
End Sub

Private Sub Events_OnError(ByVal Message As String)
    AddChat Color.Error, Message
End Sub
