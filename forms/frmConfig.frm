VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   5280
   ClientLeft      =   6075
   ClientTop       =   8190
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9340
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Connection"
      TabPicture(0)   =   "frmConfig.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grpPing"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grpLogin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grpClient"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "grpDiablo2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "frmConfig.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmIdlePrev"
      Tab(1).Control(1)=   "frChanOptions"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "frmIdleMesSet"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Appearance"
      TabPicture(2)   =   "frmConfig.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblCNote"
      Tab(2).Control(1)=   "lblRGB"
      Tab(2).Control(2)=   "frBG"
      Tab(2).Control(3)=   "frColors"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame5 
         Caption         =   "Update Version byte"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   82
         Top             =   2640
         Width           =   2775
         Begin VB.ComboBox GameVersions 
            Height          =   330
            ItemData        =   "frmConfig.frx":0060
            Left            =   120
            List            =   "frmConfig.frx":0079
            TabIndex        =   87
            Text            =   "Select Product"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton UpdateVersion 
            Caption         =   "Update Version byte"
            Height          =   375
            Left            =   120
            TabIndex        =   86
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox NewByte 
            Height          =   285
            Left            =   1920
            TabIndex        =   84
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Current Version bytes"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   71
         Top             =   360
         Width           =   2775
         Begin VB.Label D2XP 
            Caption         =   "Diablo LOD -"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label D2DV 
            Caption         =   "Diablo II -"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label WAR3 
            Caption         =   "Warcraft III"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label W2BN 
            Caption         =   "Warcraft II -"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label SEXP 
            Caption         =   "Broodwars -"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label STAR 
            Caption         =   "Starcraft -"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label RTSJ 
            Caption         =   "Starcraft Japan -"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Controls"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   69
         Top             =   3720
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame frColors 
         Caption         =   "Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   3495
         Begin VB.TextBox txtColor 
            Height          =   255
            Left            =   600
            TabIndex        =   93
            Text            =   "00"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.PictureBox picPad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   705
            Left            =   2520
            ScaleHeight     =   675
            ScaleWidth      =   825
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox lstColors 
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
            ItemData        =   "frmConfig.frx":00D6
            Left            =   120
            List            =   "frmConfig.frx":00D8
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   600
            Width           =   2295
         End
         Begin VB.HScrollBar rScroll 
            Height          =   200
            LargeChange     =   10
            Left            =   600
            Max             =   255
            TabIndex        =   68
            Top             =   1080
            Width           =   1815
         End
         Begin VB.HScrollBar gScroll 
            Height          =   200
            LargeChange     =   10
            Left            =   600
            Max             =   255
            TabIndex        =   70
            Top             =   1320
            Width           =   1815
         End
         Begin VB.HScrollBar bScroll 
            Height          =   200
            LargeChange     =   10
            Left            =   600
            Max             =   255
            TabIndex        =   72
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtHEX 
            Appearance      =   0  'Flat
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
            Left            =   600
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtRGB 
            Appearance      =   0  'Flat
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
            Left            =   600
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdAppend 
            Caption         =   "&Append"
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
            Left            =   2400
            TabIndex        =   79
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblBGColor 
            Caption         =   "Color to modify:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   1575
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   210
            Left            =   120
            Top             =   1560
            Width           =   375
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   210
            Left            =   120
            Top             =   1320
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   210
            Left            =   120
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblHex 
            Alignment       =   2  'Center
            Caption         =   "HEX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   65
            Top             =   1920
            Width           =   405
         End
         Begin VB.Label lblRGB2 
            Alignment       =   2  'Center
            Caption         =   "RGB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   64
            Top             =   2280
            Width           =   405
         End
         Begin VB.Label lblCNote2 
            Caption         =   "Color Profiling can be done by using the key from the registry path below."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   63
            Top             =   3240
            Width           =   3255
         End
      End
      Begin VB.Frame frBG 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -71160
         TabIndex        =   57
         Top             =   480
         Width           =   3495
         Begin VB.TextBox lstFonts 
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
            Left            =   1440
            TabIndex        =   81
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtFSize 
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
            Left            =   2880
            MaxLength       =   31
            TabIndex        =   83
            Text            =   "8"
            Top             =   720
            Width           =   495
         End
         Begin VB.ComboBox lstVariable 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            ItemData        =   "frmConfig.frx":00DA
            Left            =   120
            List            =   "frmConfig.frx":00DC
            Style           =   1  'Simple Combo
            TabIndex        =   85
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label lblFont 
            Caption         =   "Font:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblSize 
            Caption         =   "Font Size:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame grpDiablo2 
         Caption         =   "Diablo 2 Logon Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3840
         TabIndex        =   52
         Top             =   3120
         Width           =   3495
         Begin VB.TextBox txtCharacter 
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
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkRealms 
            Caption         =   "Auto Logon Character"
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
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblChar 
            Caption         =   "Character:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   400
            Width           =   975
         End
      End
      Begin VB.Frame grpClient 
         Caption         =   "Client Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   3840
         TabIndex        =   45
         Top             =   480
         Width           =   3495
         Begin VB.TextBox txtEmail 
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
            Left            =   1560
            MaxLength       =   24
            TabIndex        =   16
            Top             =   1800
            Width           =   1815
         End
         Begin VB.ComboBox cboProduct 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmConfig.frx":00DE
            Left            =   960
            List            =   "frmConfig.frx":0103
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtCDKey1 
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
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   13
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtCDKey2 
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
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   14
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtCDKeyOwner 
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
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   15
            Text            =   "FyreChat"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblEmail 
            Caption         =   "E-mail:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblCDkey 
            Caption         =   "CD Key:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblLOD 
            Caption         =   "Expansion Key:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblOwner 
            Caption         =   "CD Key Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblProduct 
            Caption         =   "Product:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame grpLogin 
         Caption         =   "Logon Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   3495
         Begin VB.ComboBox cboBNLSServer 
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
            ItemData        =   "frmConfig.frx":01DF
            Left            =   1200
            List            =   "frmConfig.frx":01FB
            TabIndex        =   92
            Text            =   "bnls.openfyre.net"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox txtUsername 
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
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "FyreChat"
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   12
            PasswordChar    =   "*"
            TabIndex        =   7
            Text            =   "123456"
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdPasswordShow 
            Height          =   360
            Left            =   3240
            TabIndex        =   36
            Top             =   720
            Width           =   135
         End
         Begin VB.ComboBox cboBattlenet 
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
            ItemData        =   "frmConfig.frx":0288
            Left            =   1200
            List            =   "frmConfig.frx":0340
            TabIndex        =   8
            Text            =   "useast.battle.net"
            Top             =   1080
            Width           =   2190
         End
         Begin VB.TextBox txtHomeChan 
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
            Left            =   1200
            MaxLength       =   31
            TabIndex        =   9
            Text            =   "Op FyreChat"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblBNLS 
            Caption         =   "BNLS:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblServer 
            Caption         =   "Server:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblPassword 
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
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblUsername 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblChannel 
            Caption         =   "Channel:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.Frame grpPing 
         Caption         =   "Ping Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         TabIndex        =   31
         Top             =   3120
         Width           =   3495
         Begin VB.ComboBox lstPing 
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
            ItemData        =   "frmConfig.frx":06CD
            Left            =   1200
            List            =   "frmConfig.frx":06DA
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkUDP 
            Caption         =   "UDP Plug:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblPingSpoof 
            Caption         =   "Ping Spoof:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   365
            Left            =   120
            TabIndex        =   33
            Top             =   400
            Width           =   1320
         End
         Begin VB.Image imgPlug 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2880
            Picture         =   "frmConfig.frx":06EF
            Stretch         =   -1  'True
            Top             =   840
            Width           =   375
         End
         Begin VB.Image imgNegOne 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2880
            Picture         =   "frmConfig.frx":0BC9
            Stretch         =   -1  'True
            Top             =   360
            Width           =   375
         End
         Begin VB.Image imgZero 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2880
            Picture         =   "frmConfig.frx":16D3
            Stretch         =   -1  'True
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame frmIdlePrev 
         Caption         =   "Idle Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -71160
         TabIndex        =   28
         Top             =   2760
         Width           =   3495
         Begin RichTextLib.RichTextBox rtbIdle 
            Height          =   1575
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2778
            _Version        =   393217
            BackColor       =   0
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmConfig.frx":21DD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frChanOptions 
         Caption         =   "Display Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   3495
         Begin VB.CheckBox chkShowSvr 
            Caption         =   "Show Server Types"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CheckBox chkDebugMode 
            Caption         =   "Debug Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chkEnterLeave 
            Caption         =   "Join/Leave Confirmations"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkStatusUpdate 
            Caption         =   "Diablo 2 Status Updates"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox chkWhisper 
            Caption         =   "Whisper Window"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkPingFlag 
            Caption         =   "Ping:Flags on Events"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtMaxLength 
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
            Left            =   2520
            MaxLength       =   255
            TabIndex        =   34
            Text            =   "10000"
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label lblMaxLength 
            Caption         =   "Max Chat Length:"
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
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Misc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -71160
         TabIndex        =   21
         Top             =   480
         Width           =   3495
         Begin VB.CheckBox chkLog 
            Caption         =   "Log Chat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox chkAutoCon 
            Caption         =   "Automatically Connect"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox chkNews 
            Caption         =   "Show News on Start"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   2295
         End
         Begin VB.CheckBox chkSysTray 
            Caption         =   "Minimize to System Tray"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkConfirmExit 
            Caption         =   "Confirm on Exit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox chkARCon 
            Caption         =   "Reconnect on Disconnect"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkFloodProt 
            Caption         =   "Flood Protection"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1440
            Width           =   1935
         End
      End
      Begin VB.CheckBox chkProdSpec 
         Caption         =   "Join Product-Specific First"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   4320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame frmIdleMesSet 
         Caption         =   "Idle Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74760
         TabIndex        =   11
         Top             =   2760
         Width           =   3495
         Begin VB.CheckBox chkIdle 
            Caption         =   "Normal:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtIdle 
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
            Left            =   1200
            MaxLength       =   255
            TabIndex        =   55
            Text            =   "/me %ver - %uptime"
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtInterval 
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
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   56
            Text            =   "300"
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox chkCountIdle 
            Caption         =   "Hmm Idle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkAwayIdle 
            Caption         =   "Away Idle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblInterval 
            Caption         =   "Idle Interval (seconds)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Label lblRGB 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "RGB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   -73560
         TabIndex        =   89
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblCNote 
         Caption         =   "HKLM/Software/OpenFyre/Fyrechat/<Path>/Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   88
         Top             =   4320
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mclsToolTip As New clsToolTip
Dim Reg As clsRegistry
Dim tRED, tGREEN, tBLUE As Integer  'to hold the RGB color values
Dim hred, hgreen, hblue As String  'to hold the HEX color values

Dim pBool As Boolean


Private Sub cmdAppend_Click()
On Local Error Resume Next
    Dim tmpHex As String
    tmpHex = txtHEX.text
    tmpHex = Replace(tmpHex, "#", "")
    If Val("&H" & Mid(tmpHex, 1, 2)) <= 255 Then hred = Val("&H" & Mid(tmpHex, 1, 2))
    If Val("&H" & Mid(tmpHex, 3, 2)) <= 255 Then hgreen = Val("&H" & Mid(tmpHex, 3, 2))
    If Val("&H" & Mid(tmpHex, 5, 2)) <= 255 Then hblue = Val("&H" & Mid(tmpHex, 5, 2))
    
    update_RGB
    update_Scrolls
    update_Color
    If lstColors.text = "Window Background" Then Color.Background = txtColor.text
    If lstColors.text = "Self" Then Color.Self = txtColor.text
    If lstColors.text = "User" Then Color.user = txtColor.text
    If lstColors.text = "Operator" Then Color.Op = txtColor.text
    If lstColors.text = "Spoken Text" Then Color.Message = txtColor.text
    If lstColors.text = "Ping & Flags" Then Color.PingFlags = txtColor.text
    If lstColors.text = "Server Error" Then Color.Error = txtColor.text
    If lstColors.text = "Server Info" Then Color.info = txtColor.text
    If lstColors.text = "Bot Error" Then Color.BotError = txtColor.text
    If lstColors.text = "Bot Info" Then Color.BotInfo = txtColor.text
    If lstColors.text = "User Leaves" Then Color.Left = txtColor.text
    If lstColors.text = "User Joins" Then Color.Join = txtColor.text
    If lstColors.text = "Encrypted" Then Color.Enc = txtColor.text
    If lstColors.text = "Whisper From" Then Color.WhisperFrom = txtColor.text
    If lstColors.text = "Server News" Then Color.News = txtColor.text
    If lstColors.text = "TimeStamp" Then Color.timestamp = txtColor.text
    If lstColors.text = "Bot" Then Color.Bot = txtColor.text
    frmMain.Display
End Sub

Private Sub cmdCancel_Click()
On Local Error Resume Next
    Unload frmConfig
    Set frmConfig = Nothing
End Sub

Private Sub cmdPasswordShow_Click()
    If pBool = True Then
        txtPassword.PasswordChar = ""
        pBool = False
    Else
        txtPassword.PasswordChar = "*"
        pBool = True
    End If
End Sub

Private Sub Form_Load()
On Local Error Resume Next
    Me.icon = frmMain.icon
    pBool = True
    
    txtUsername.text = BNET.username
    txtPassword.text = BNET.password
    cboBattlenet.text = BNET.BNCSServer
    txtHomeChan.text = BNET.HomeChannel
    cboBNLSServer.text = BNET.BNLSServer
    
    txtCDKey1.text = UCase(BNET.CDKey)
    txtCDKey2.text = UCase(BNET.CDKey2)
    txtCDKeyOwner.text = BNET.CDKeyOwner
    cboProduct.text = ProductName(BNET.Product)
    txtEmail.text = BNET.Email
    
    'Realm Stuff
    txtCharacter.text = BNET.Character
    chkRealms.Value = BNET.varCRealm
    
    chkAutoCon.Value = BNET.varAutoCon
    chkARCon.Value = BNET.varARCon
    lstPing.text = BNET.varLagPlug
    chkUDP.Value = BNET.varUDP
    chkLog.Value = BNET.varLog
    chkWhisper.Value = BNET.WhispWin
    chkSysTray.Value = BNET.varSysTray
    chkNews.Value = BNET.varNews
    chkConfirmExit.Value = BNET.varConfirmExit
    chkPingFlag.Value = BNET.varPingFlag
    chkShowSvr.Value = BNET.varShowSvr
    chkFloodProt.Value = BNET.FloodProt
    chkDebugMode.Value = BNET.varDebugMode
    
    txtIdle.text = BNET.Idle
    txtInterval = BNET.IdleInt
    chkProdSpec.Value = BNET.varProdSpec
    chkEnterLeave.Value = BNET.varEnterLeave
    chkStatusUpdate.Value = BNET.varStatusUpdate
    chkCountIdle.Value = BNET.varCountIdle
    chkAwayIdle.Value = BNET.varAwayIdle
    chkIdle.Value = BNET.varIdle
    txtMaxLength.text = BNET.MaxLength
    
    lstFonts.text = BNET.Fonts
    txtFSize.text = BNET.FSize

    If lstPing.text = "-1ms" Then
        imgNegOne.Visible = True
        imgZero.Visible = False
    End If
    If lstPing.text = "0ms" Then
        imgNegOne.Visible = False
        imgZero.Visible = True
    End If
    If lstPing.text = "None" Then
        imgNegOne.Visible = False
        imgZero.Visible = False
    End If
    
    txtIdle_Change
    
    lstColors.AddItem "Window Background"
    lstColors.AddItem "Self"
    lstColors.AddItem "User"
    lstColors.AddItem "Operator"
    lstColors.AddItem "Spoken Text"
    lstColors.AddItem "Ping & Flags"
    lstColors.AddItem "Server Error"
    lstColors.AddItem "Server Info"
    lstColors.AddItem "Bot Error"
    lstColors.AddItem "Bot Info"
    lstColors.AddItem "User Leaves"
    lstColors.AddItem "User Joins"
    lstColors.AddItem "Encrypted"
    lstColors.AddItem "Whisper From"
    lstColors.AddItem "Server News"
    lstColors.AddItem "TimeStamp"
    lstColors.AddItem "Bot"
    lstColors.text = "Window Background"
    
    
    Dim ctrl As Control
    With mclsToolTip
        Call .Create(Me)
        .MaxTipWidth = 240
        .DelayTime(ttDelayShow) = 20000
        .ToolTipHeader = "FyreChat ToolTip" & vbNewLine & "_______________"
    
        For Each ctrl In Controls
          Call .AddTool(ctrl)
        Next ctrl
    
        .ToolTipHeaderShow = False
        .ToolText(txtUsername) = "Enter your battle.net account name here." & vbCrLf & _
                                 "If the account doesn't exist, the bot will" & vbCrLf & _
                                 "try to create it."
        .ToolText(txtPassword) = "Enter your battle.net account password here." & vbCrLf & _
                                 "Clicking the button on the right will reveal" & vbCrLf & _
                                 "your password."
        .ToolText(cboBattlenet) = "Select or enter in a Battle.net Server."
        .ToolText(cboBNLSServer) = "Select or enter a BNLS/JBLS Server."
        .ToolText(txtHomeChan) = "The channel the bot will enter when it" & vbCrLf & _
                                 "connects to battle.net."
        .ToolText(lstPing) = "(Optional)" & vbCrLf & _
                            "Spoof your ping with one of the" & vbCrLf & _
                            "following."
        .ToolText(chkUDP) = "(Optional)" & vbCrLf & _
                            "UDP Plug, Not Compatible with Diablo 2" & vbCrLf & _
                            "or WarCraft III."
        .ToolText(txtCDKey1) = "Your CD-Key (With no dashes or spaces)" & vbCrLf & _
                               "goes here."
        .ToolText(txtCDKey2) = "Your Expansion CD-Key (With no dashes or" & vbCrLf & _
                               "spaces) goes here."
        .ToolText(txtCDKeyOwner) = "(Optional)" & vbCrLf & _
                                   "The owner of your CD-Key goes here."
        .ToolText(txtCharacter) = "(Optional)" & vbCrLf & _
                                  "The Diablo 2 character you wish to log on" & vbCrLf & _
                                  "with."
        .ToolText(txtEmail) = "(Optional)" & vbCrLf & _
                              "If this field is not empty, then the bot will" & vbCrLf & _
                              "register this e-mail to the account the next" & vbCrLf & _
                              "you log on."
        .ToolText(chkRealms) = "(Optional)" & vbCrLf & _
                                "Logs on to selected realm server when bot" & vbCrLf & _
                                "connects to Battle.net."
        'Options Tab
        .ToolText(chkEnterLeave) = "(Optional)" & vbCrLf & _
                                   "Show when users join or leave the channel."
        .ToolText(chkStatusUpdate) = "(Optional)" & vbCrLf & _
                                     "Show when users change their Diablo II" & vbCrLf & _
                                     "character."
        .ToolText(chkWhisper) = "(Optional)" & vbCrLf & _
                                "Shows a whisper box below the chat box." & vbCrLf & _
                                "All whispers will be forwarded to the whisper box."
        .ToolText(chkPingFlag) = "(Optional)" & vbCrLf & _
                                "Shows the user's ping and flags when they" & vbCrLf & _
                                "speak (ex: [00:00:00] [125:2] <Fyre> meh)."
        .ToolText(chkShowSvr) = "(Optional)" & vbCrLf & _
                            "Displays server types before each line." & vbCrLf & _
                            "[BNCS] Connected to Battle.net!"
        .ToolText(chkLog) = "(Optional)" & vbCrLf & _
                            "Logs all chat events into log folder" & vbCrLf & _
                            "under date of logging."
        .ToolText(chkDebugMode) = "(Optional)" & vbCrLf & _
                            "Displays useful packet data."
        .ToolText(txtMaxLength) = "(Optional)" & vbCrLf & _
                                  "Sets a max to the ammount of characters" & vbCrLf & _
                                  "in the chat window(s)."
        .ToolText(chkAutoCon) = "(Optional)" & vbCrLf & _
                                "The bot will automatically connect to" & vbCrLf & _
                                "battle.net when started."
        .ToolText(chkNews) = "(Optional)" & vbCrLf & _
                             "Shows the bot news on application start."
        .ToolText(chkSysTray) = "(Optional)" & vbCrLf & _
                                "When minimized the bot will go to the" & vbCrLf & _
                                "system tray."
        .ToolText(chkConfirmExit) = "(Optional)" & vbCrLf & _
                                    "The bot will confirm your exit if it's" & vbCrLf & _
                                    "connected to battle.net."
        .ToolText(chkARCon) = "(Optional)" & vbCrLf & _
                              "The bot will attempt to connect to battle.net" & vbCrLf & _
                              "if connection is lost."
        .ToolText(chkFloodProt) = "(Optional)" & vbCrLf & _
                                  "Text sent to battle.net will be queued" & vbCrLf & _
                                  "to prevent self flooding."
        .ToolText(chkProdSpec) = "(Optional)" & vbCrLf & _
                                 "The bot will join a product-specific channel" & vbCrLf & _
                                 "before it joins the home channel."
        .ToolText(txtIdle) = "(Optional)" & vbCrLf & _
                             "The idle for the bot."
        .ToolText(txtInterval) = "(Optional)" & vbCrLf & _
                             "The idle's interval for how often the" & vbCrLf & _
                             "message is sent."
        .ToolText(chkCountIdle) = "(Optional)" & vbCrLf & _
                                  "The Hmm Idle" & vbCrLf & _
                                  "Outputs 'hmm' every 35 talks."
        .ToolText(chkAwayIdle) = "(Optional)" & vbCrLf & _
                                 "Sends how long you've been idle as your" & vbCrLf & _
                                 "away message."
        'Appearance Tab
        .ToolText(lstFonts) = "The font for the program."
        .ToolText(txtFSize) = "The size for the font."

    End With
    
    Dim K As Integer
    For K = 1 To Screen.FontCount
        lstVariable.AddItem Screen.Fonts(K - 1)
    Next

    Dim hColor As String
    hColor = Color.Background
    hColor = hColor & String(6 - Len(hColor), "0")
    txtHEX.text = "#" & hColor
    'txtHEX = "#" & Color.Background
    
    picPad.AutoRedraw = True
    Call ToRGB(txtColor.text)
    update_RGB
    update_Scrolls
    update_Color
End Sub

Public Sub cmdSave_Click()
On Local Error Resume Next
    BNET.username = txtUsername.text
    BNET.password = txtPassword.text
    BNET.CDKey = UCase(txtCDKey1.text)
    BNET.CDKey2 = (txtCDKey2.text)
    BNET.CDKeyOwner = txtCDKeyOwner.text
    BNET.Product = ProductID(cboProduct.text)
    BNET.BNCSServer = cboBattlenet.text
    BNET.BNLSServer = cboBNLSServer.text
    
    'Realm Stuff
    BNET.Character = txtCharacter.text
    BNET.varCRealm = chkRealms.Value
    BNET.Email = txtEmail.text
    
    BNET.varAutoCon = chkAutoCon.Value
    BNET.varARCon = chkARCon.Value
    BNET.varLagPlug = lstPing.text
    BNET.varUDP = chkUDP.Value
    BNET.varPingFlag = chkPingFlag.Value
    BNET.varShowSvr = chkShowSvr.Value
    BNET.varLog = chkLog.Value
    BNET.varDebugMode = chkDebugMode.Value
    BNET.WhispWin = chkWhisper.Value
    BNET.varNews = chkNews.Value
    BNET.varSysTray = chkSysTray.Value
    BNET.varConfirmExit = chkConfirmExit.Value
    BNET.FloodProt = chkFloodProt
    BNET.Idle = txtIdle.text
    BNET.IdleInt = txtInterval
    BNET.HomeChannel = txtHomeChan.text
    BNET.varEnterLeave = chkEnterLeave.Value
    BNET.varStatusUpdate = chkStatusUpdate.Value
    BNET.varCountIdle = chkCountIdle.Value
    BNET.varAwayIdle = chkAwayIdle.Value
    BNET.varIdle = chkIdle.Value
    BNET.MaxLength = txtMaxLength.text
    BNET.varProdSpec = chkProdSpec.Value

    BNET.Fonts = lstFonts.text
    BNET.FSize = txtFSize.text
    
    frmMain.Display
    modFunctions.SaveConfig
    modFunctions.LoadConfig
    Me.Hide
End Sub

Private Sub lstColors_Click()
On Local Error Resume Next
    If lstColors.text = "Window Background" Then txtColor.text = Color.Background
    If lstColors.text = "Self" Then txtColor.text = Color.Self
    If lstColors.text = "User" Then txtColor.text = Color.user
    If lstColors.text = "Operator" Then txtColor.text = Color.Op
    If lstColors.text = "Spoken Text" Then txtColor.text = Color.Message
    If lstColors.text = "Ping & Flags" Then txtColor.text = Color.PingFlags
    If lstColors.text = "Server Error" Then txtColor.text = Color.Error
    If lstColors.text = "Server Info" Then txtColor.text = Color.info
    If lstColors.text = "Bot Error" Then txtColor.text = Color.BotError
    If lstColors.text = "Bot Info" Then txtColor.text = Color.BotInfo
    If lstColors.text = "User Leaves" Then txtColor.text = Color.Left
    If lstColors.text = "User Joins" Then txtColor.text = Color.Join
    If lstColors.text = "Encrypted" Then txtColor.text = Color.Enc
    If lstColors.text = "Whisper From" Then txtColor.text = Color.WhisperFrom
    If lstColors.text = "Server News" Then txtColor.text = Color.News
    If lstColors.text = "TimeStamp" Then txtColor.text = Color.timestamp
    If lstColors.text = "Bot" Then txtColor.text = Color.Bot
End Sub

Private Sub lstPing_Click()
    If lstPing.text = "-1ms" Then
        imgNegOne.Visible = True
        imgZero.Visible = False
    End If
    If lstPing.text = "0ms" Then
        imgNegOne.Visible = False
        imgZero.Visible = True
    End If
    If lstPing.text = "None" Then
        imgNegOne.Visible = False
        imgZero.Visible = False
    End If
End Sub

Private Sub lstVariable_Click()
    lstFonts.text = lstVariable.text
End Sub

Private Sub rScroll_Change()
On Local Error Resume Next
    tRED = rScroll.Value
    hred = Hex(rScroll.Value)
    update_RGB
    update_HEX
End Sub

Private Sub gScroll_Change()
On Local Error Resume Next
    tGREEN = gScroll.Value
    hgreen = Hex(gScroll.Value)
    update_RGB
    update_HEX
End Sub

Private Sub bScroll_Change()
On Local Error Resume Next
    tBLUE = bScroll.Value
    hblue = Hex(bScroll.Value)
    update_RGB
    update_HEX
End Sub

Private Sub update_RGB()
On Local Error Resume Next
    txtRGB.text = "(" & tRED & "," & tGREEN & "," & tBLUE & ")"
    picPad.BackColor = RGB(tRED, tGREEN, tBLUE)
End Sub

Private Sub update_Scrolls()
On Local Error Resume Next
    rScroll.Value = tRED
    gScroll.Value = tGREEN
    bScroll.Value = tBLUE
End Sub

Private Sub update_HEX()
    Dim hColor As String
    
    hColor = Hex(RGB(tRED, tGREEN, tBLUE))
    hColor = hColor & String(6 - Len(hColor), "0")
    txtHEX.text = "#" & hColor
End Sub

Private Sub update_Color()
    Dim tmpHex As String
    tmpHex = txtHEX.text
    tmpHex = Replace(tmpHex, "#", "")
    txtColor.text = CLng("&H" & tmpHex)
End Sub

Public Function ToRGB(lColor As Long) As String
On Local Error Resume Next
    tRED = lColor Mod 256
    tGREEN = ((lColor And &HFF00) / 256&) Mod 256&
    tBLUE = (lColor And &HFF0000) / 65536
End Function

Private Sub txtColor_Change()
    ToRGB txtColor.text
    update_RGB
    update_Scrolls
End Sub

Private Sub txtHEX_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
    If KeyAscii = 13 Then
        Dim tmpHex As String
        tmpHex = txtHEX.text
        tmpHex = Replace(tmpHex, "#", "")
        If Val("&H" & Mid(tmpHex, 1, 2)) <= 255 Then
            tRED = Val("&H" & Mid(tmpHex, 1, 2))
        End If
        If Val("&H" & Mid(tmpHex, 3, 2)) <= 255 Then
            tGREEN = Val("&H" & Mid(tmpHex, 3, 2))
        End If
        If Val("&H" & Mid(tmpHex, 5, 2)) <= 255 Then
            tBLUE = Val("&H" & Mid(tmpHex, 5, 2))
        End If
        update_RGB
        update_Scrolls
        update_Color
    End If
End Sub

Private Sub txtRGB_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
    If KeyAscii = 13 Then
        Dim tmpRGB As String, tmpRGB2() As String
        tmpRGB = txtRGB.text
        tmpRGB = Replace(tmpRGB, "(", "")
        tmpRGB = Replace(tmpRGB, ")", "")
        tmpRGB = Replace(tmpRGB, ",", " ")
        tmpRGB2() = Split(tmpRGB, " ")
        If tmpRGB2(0) <= 255 Then tRED = tmpRGB2(0)
        If tmpRGB2(1) <= 255 Then tGREEN = tmpRGB2(1)
        If tmpRGB2(2) <= 255 Then tBLUE = tmpRGB2(2)
        update_RGB
        update_Scrolls
        update_Color
    End If
End Sub

Public Sub txtIdle_Change()
On Local Error Resume Next
    Dim data As String, timestamp As String
    timestamp = "[" & Format(Time, "hh:mm:ss") & "] "
    rtbIdle.text = ""
    data = txtIdle.text
    If InStr(data, "%") Then
        data = Replace(data, "%ver", Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & " Build " & Format(App.Revision, "00"))
        data = Replace(data, "%uptime", FormatCount(GetTickCount))
        data = Replace(data, "%botuptime", FormatCount(GetTickCount - RunningTime))
        data = Replace(data, "%connected", FormatCount(GetTickCount - connecttime))
        data = Replace(data, "%idle", FormatCount(GetTickCount - LastTalk, 4))
        data = Replace(data, "%self", BNET.TrueUsername)
        data = Replace(data, "%chan", BNET.CurrentChan)
    End If
    If InStr(data, "/me") Then
        data = Replace(data, "/me ", "")
        With rtbIdle
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = Color.timestamp
            .SelText = timestamp
            .SelStart = Len(.text)
            
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = Color.Carrot
            .SelText = "<" & BNET.username & " " & data
            .SelColor = Color.Carrot
            .SelText = ">"
            .SelStart = Len(.text)
        End With
    Else
        With rtbIdle
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = Color.timestamp
            .SelText = timestamp
            .SelStart = Len(.text)
            
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = Color.Self
            .SelText = "<" & BNET.username & "> "
            .SelStart = Len(.text)
            
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = Color.Message
            .SelText = data
            .SelStart = Len(.text)
        End With
    End If
End Sub
