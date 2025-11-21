VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSTAR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create a StarCraft Game"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSTAR.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Join Game"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   4455
      Begin VB.CommandButton cmdJoinGame 
         Caption         =   "&Join"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtJoinPass 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   19
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtJoinName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create Game"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdCreate 
         Caption         =   "&Create"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2640
         Width           =   1095
      End
      Begin VB.ComboBox txtGameTeams 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSTAR.frx":000C
         Left            =   2400
         List            =   "frmSTAR.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox txtGameType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSTAR.frx":0061
         Left            =   240
         List            =   "frmSTAR.frx":0077
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1099
      End
      Begin VB.TextBox txtGameMap 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtGamePass 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtGameName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdLeave 
         Caption         =   "Leave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   2640
         Width           =   1095
      End
      Begin MSComctlLib.Slider sldGameSpeed 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   1
         Max             =   6
         SelStart        =   3
         TickStyle       =   3
         Value           =   3
         TextPosition    =   1
      End
      Begin VB.Label lblGameSpeed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Game Type"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Map:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Clos&e"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog SelectFile 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSTAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.icon = frmMain.icon
    txtGameMap.text = "Challenger"
    lblGameSpeed.Caption = "Normal"
    txtGameType = "Use Map Settings"
End Sub

Private Sub cmdBrowse_Click()
    SelectFile.DialogTitle = "Choose Map"
    SelectFile.MaxFileSize = 16384
    SelectFile.FileName = ""
    SelectFile.FileName = "StarCraft Maps|*.SMP"
    'SelectFile.ShowOpen
End Sub

Private Sub sldGameSpeed_Change()
    Select Case sldGameSpeed.value
        Case "0": lblGameSpeed.Caption = "Slowest"
        Case "1": lblGameSpeed.Caption = "Slower"
        Case "2": lblGameSpeed.Caption = "Slow"
        Case "3": lblGameSpeed.Caption = "Normal"
        Case "4": lblGameSpeed.Caption = "Fast"
        Case "5": lblGameSpeed.Caption = "Faster"
    Case Else: lblGameSpeed.Caption = "Fastest"
    End Select
End Sub

Private Sub cmdCreate_Click()
    Dim gType As String
    Select Case gType
        Case "Melee": gType = "2"
        Case "Free for All": gType = "3"
        Case "1 on 1": gType = "4"
        Case "Ladder": gType = "9"
        Case "Use Map Settings": gType = "A"
        Case "Top vs. Bottom": gType = "F"
        Case Else: gType = "2"
    End Select
    sndStartAdvEx3 txtGameName.text, txtGamePass.text, txtGameMap.text, sldGameSpeed.value, gType
    gType = ""
End Sub

Private Sub cmdJoinGame_Click()
    With sPB
        .InsertDWORD &HC7                           ' Product ID (any valid)
        .InsertNTString frmSTAR.txtJoinName.text    ' Game Name
        .InsertNTString frmSTAR.txtJoinPass.text    ' Game Password
        .SendPacket &H22                            ' SID_NOTIFYJOIN
    End With
End Sub

Private Sub cmdLeave_Click()
On Local Error Resume Next
    With sPB
        .InsertDWORD 2
        .InsertNTString BNET.CurrentChan
        .SendPacket &HC
    End With
End Sub

Private Sub cmdStart_Click(): sPB.SendPacket &H10: End Sub

Private Sub cmdCancel_Click(): Unload Me: End Sub
