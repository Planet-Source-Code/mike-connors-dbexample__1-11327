VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMainMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Example"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12090
   Icon            =   "mdiMainMenu.frx":0000
   Picture         =   "mdiMainMenu.frx":08CA
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1335
      Top             =   1515
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1535
      ButtonWidth     =   2593
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Browse..."
            Key             =   "BROWSE"
            Object.ToolTipText     =   "Browse Student Records Screen..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search / Reports..."
            Key             =   "SEARCH"
            Object.ToolTipText     =   "Search Student Records and Print Reports..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit User..."
            Key             =   "EDITUSER"
            Object.ToolTipText     =   "Modify Your User Access..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Colors..."
            Key             =   "COLORS"
            Object.ToolTipText     =   "Set Program Colors..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "EXIT"
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox picUserInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   8640
         Picture         =   "mdiMainMenu.frx":24090C
         ScaleHeight     =   780
         ScaleWidth      =   3285
         TabIndex        =   2
         Top             =   30
         Width           =   3285
         Begin VB.Label lblLoginTime 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   1110
            TabIndex        =   6
            Tag             =   "ButtonLabel"
            Top             =   420
            Width           =   2085
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login Time:"
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   420
            Width           =   825
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   1110
            TabIndex        =   4
            Tag             =   "ButtonLabel"
            Top             =   120
            Width           =   2085
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User:"
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   375
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   240
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":248F5E
            Key             =   "CONSULTANTS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24983A
            Key             =   "CLIENTS"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24A116
            Key             =   "JOBORDERS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24A9F2
            Key             =   "ENGAGEMENTS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24B2CE
            Key             =   "LOGIN"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24BBB2
            Key             =   "ACCOUNTS"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24C496
            Key             =   "COLORS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24C7B2
            Key             =   "EXIT"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   825
      Top             =   1455
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Lunec.mdb"
      InitDir         =   "app.Path"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBrowse 
         Caption         =   "&Browse..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuQKeys 
         Caption         =   "&Quick Keys"
         Shortcut        =   {F3}
      End
      Begin VB.Menu h981826 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinterSetup 
         Caption         =   "&Printer Setup..."
         Shortcut        =   ^P
      End
      Begin VB.Menu H83762 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUser 
         Caption         =   "&Edit User..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuColorSettings 
         Caption         =   "&Colors..."
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "mdiMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Set Reg
Dim WithEvents Reg As Registry
Attribute Reg.VB_VarHelpID = -1

Private Sub MDIForm_Load()

'set reg & check
Set Reg = New Registry
Call CheckReg

On Local Error Resume Next

'Set program colors...
Call SetColors(Me)

'Load the main menu's form settings...
Call LoadINISettings

'Set visible to true so that the main menu will be visible with the login dialog box in front of it...
Me.Visible = True
DoEvents

'Show the login screen...
frmLogin.Show vbModal

Timer1.Enabled = True

End Sub
Sub LoadINISettings()

'Form properties...
If Trim$(ReadINI(Me.Name, "Caption")) <> "" Then
    Me.Caption = ReadINI(Me.Name, "Caption")
End If

'Form Coordinates...
Me.WindowState = val(ReadINI(Me.Name, "WindowState"))
If Me.WindowState = vbMaximized Then Exit Sub
Me.Left = val(ReadINI(Me.Name, "Left"))
Me.Top = val(ReadINI(Me.Name, "Top"))
Me.Height = val(ReadINI(Me.Name, "Height"))
Me.Width = val(ReadINI(Me.Name, "Width"))

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)

'Save this form's settings...
Call SaveINISettings

Call appTerminate

End Sub
Sub SaveINISettings()

'WindowState...
Call WriteINI(Me.Name, "WindowState", Me.WindowState)

'If windowstate is maximized, exit...
If Me.WindowState = vbMaximized Then Exit Sub

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)
Call WriteINI(Me.Name, "Height", Me.Height)
Call WriteINI(Me.Name, "Width", Me.Width)

End Sub

Private Sub mnuAccounts_Click()

frmAccounts.Show
frmAccounts.ZOrder

End Sub

Private Sub mnuAbout_Click()

frmAbout.Show

End Sub

Private Sub mnuBrowse_Click()

frmExampledb.Show
frmExampledb.ZOrder

End Sub

Private Sub mnuCascade_Click()

Call ArrangeIcons(vbCascade)

End Sub

Private Sub mnuColorSettings_Click()

frmColors.Show
frmColors.ZOrder

End Sub


Private Sub mnuEditUser_Click()

frmAccounts.Show
frmAccounts.ZOrder

End Sub

Private Sub mnuExit_Click()

'Unload the help window...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

If frmExampledb.Visible Then Unload frmExampledb
If frmSearch.Visible Then Unload frmSearch
If frmColors.Visible Then Unload frmColors
If frmAccounts.Visible Then Unload frmAccounts
If frmHelper.Visible Then Unload frmHelper
If frmAbout.Visible Then Unload frmAbout
Unload Me

End Sub

Private Sub mnuPrinterSetup_Click()

On Local Error Resume Next

Dialog.ShowPrinter

End Sub

Private Sub mnuQKeys_Click()
    
    frmQKeys.Show vbModal
    
End Sub

Private Sub mnuSearch_Click()

frmSearch.Show

End Sub

Private Sub mnuTileHorizontally_Click()

Call ArrangeIcons(vbHorizontal)

End Sub
Private Sub mnuTileVertically_Click()

Call ArrangeIcons(vbVertical)

End Sub





Private Sub Timer1_Timer()

On Local Error Resume Next

Dim X As Long

'Show the main menu toolbar...
mdiMainMenu.Toolbar1.Visible = True

'Set Colors...
If QuickRef.UpdateColors = True Then
    Call LoadProgramColors
    QuickRef.UpdateColors = False
    For X = 0 To Forms.Count - 1
        Call SetColors(Forms(X))
    Next X
End If

'Audible Help...
If Help.HelpIsLoaded = True Then
    If frmHelper.txtHelper.Text <> Help.HelpText Then
        frmHelper.txtHelper.Text = Help.HelpText
    End If
End If

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Local Error GoTo ToolBar1_ButtonClickError

'Toolbar Buttons...
Select Case Button.Key

    'Browse...
    Case "BROWSE"

        mnuBrowse_Click
        
    'Search...
    Case "SEARCH"

        mnuSearch_Click

    'User Edit...
    Case "EDITUSER"
        mnuEditUser_Click

    'Colors...
    Case "COLORS"
        mnuColorSettings_Click

    'Exit...
    Case "EXIT"
        If frmExampledb.Visible Then Unload frmExampledb
        If frmSearch.Visible Then Unload frmSearch
        If frmColors.Visible Then Unload frmColors
        If frmAccounts.Visible Then Unload frmAccounts
        If frmHelper.Visible Then Unload frmHelper
        If frmAbout.Visible Then Unload frmAbout
        Unload Me

End Select

Exit Sub



ToolBar1_ButtonClickError:
    Call WriteToErrorLog(Me.Name, "ToolBar1_ButtonClickError", Error, Err, False)
    Exit Sub

End Sub


Private Sub Reg_onFailed()
 If Reg.Data = "-1" Then
    End
 Else
    Registered.Show vbModal
    Call CheckReg
 End If
 
End Sub

Sub CheckReg()
 Reg.hKey = HKEY_LOCAL_MACHINE
 Reg.SubKey = "Software\Example" 'example of rooted sub keys
 Reg.ValueName = "RegKey"
 Reg.DataType = REG_SZ
 Reg.GetSetting
 
    Dim regCompare As Long
    regCompare = val(Reg.Data)
    If val("qzsjh8") <> regCompare Then End
End Sub

