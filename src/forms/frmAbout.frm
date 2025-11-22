VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About FyreChat"
   ClientHeight    =   4320
   ClientLeft      =   75
   ClientTop       =   240
   ClientWidth     =   5430
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4320
   ScaleWidth      =   5430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3840
      Width           =   945
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   5340
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5340
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Label lblVbpjUrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Web Site"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblComments 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5235
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5340
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   5340
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   525
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4920
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FyreChat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Left            =   1440
      TabIndex        =   1
      Top             =   300
      Width           =   2100
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private m_Splash As Boolean

Private Sub cmdClose_Click(): Unload Me: End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub

Private Sub cmdClose_GotFocus()
On Local Error Resume Next
   Dim r As RECT
   Static BeenThereDoneThat As Boolean

   If Not BeenThereDoneThat Then
      Call GetWindowRect((cmdClose.hWnd), r)
      Call SetCursorPos(r.Left + (r.Right - r.Left) \ 2, _
                        r.Top + (r.Bottom - r.Top) \ 2)
      BeenThereDoneThat = True
   End If
End Sub

Private Sub Form_Load()
On Local Error Resume Next
    Me.icon = frmMain.icon
    lblCopyright.Caption = "Copyright ©2003 - 2012, OpenFyre Software"
    lblVersion.Caption = "Version " & _
       Format(App.Major, "#0") & "." & _
       Format(App.Minor, "#0") & "." & _
       Format(App.Revision, "00")
     lblComments.Caption = vbCrLf & _
         " Special Thanks:" & vbCrLf & _
         "   Programming Help: Zonker[RC], Raihan[xL], and UserLoser." & vbCrLf & _
         "   Bot Testing Help: WoLF`FaLLeN[RC], ExOrCizT[RC], WoLF`Damian[RC]" & vbCrLf & _
         "                                thAw[RC], LeVaRiS[RC], CrYpTiC, EyE-Deep-EyE" & vbCrLf & _
         "                                idiot[RC], Murder[sL], Goffy59, Marshall" & vbCrLf & _
         "   DebugOutput Code: Grok[vL]" & vbCrLf & _
         "   Moral Support: Feanor[xL]"

    Line1.x2 = Me.ScaleWidth - Line1.x1
    Line2.x1 = Line1.x1
    Line2.x2 = Line1.x2
    Line2.y1 = Line1.y1 + Screen.TwipsPerPixelY
    Line2.Y2 = Line1.Y2 + Screen.TwipsPerPixelY
    
    imgIcon.Picture = Me.icon
    
    cmdClose.Left = Line1.x2 - cmdClose.Width
    '
    ' Setup links portions
    '
    lblVbpjUrl.Caption = "http://fyrechat.openfyre.net"
    lblVbpjUrl.MousePointer = 99 'Custom
    
    Line3.x2 = Me.ScaleWidth - Line3.x1
    Line4.x1 = Line3.x1
    Line4.x2 = Line3.x2
    Line4.y1 = Line3.y1 + Screen.TwipsPerPixelY
    Line4.Y2 = Line3.Y2 + Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub
Private Sub imgVBPJ_Click(): Call ShellExecute(0&, vbNullString, lblVbpjUrl.Caption, vbNullString, vbNullString, vbNormalFocus): End Sub
Private Sub imgVBPJ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, True: End Sub
Private Sub imgIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub
Private Sub lblActiveX1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub
Private Sub lblActiveX2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub

Private Sub lblVbpjUrl_Click(): Call ShellExecute(0&, vbNullString, lblVbpjUrl.Caption, vbNullString, vbNullString, vbNormalFocus): End Sub
Private Sub lblVbpjUrl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, True: End Sub
Private Sub lblComments_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub
Private Sub lblCopyright_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub
Private Sub lblPrep_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub
Private Sub lblVersion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): ToggleLink lblVbpjUrl, False: End Sub

Private Sub ToggleLink(lbl As Label, LightUp As Boolean)
On Local Error Resume Next
   With lbl
      If LightUp Then
         If .ForeColor = &HFFFFFF Then
            .Font.Underline = True
            .ForeColor = &HEED8C8
         End If
      Else
         If .ForeColor = &HEED8C8 Then
            .Font.Underline = False
            .ForeColor = &HFFFFFF
         End If
      End If
   End With
End Sub

Private Sub RemoveCancelMenuItem(frm As Form)
On Local Error Resume Next
   Dim hSysMenu As Long
   hSysMenu = GetSystemMenu(frm.hWnd, 0)
   Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
   Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub
