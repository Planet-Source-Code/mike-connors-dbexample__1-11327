VERSION 5.00
Begin VB.Form Registered 
   BorderStyle     =   0  'None
   Caption         =   "Initial Startup Information"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   315
      TabIndex        =   0
      Top             =   1845
      Width           =   4050
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      Height          =   315
      Left            =   1710
      TabIndex        =   1
      Top             =   2385
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1230
      Left            =   585
      TabIndex        =   3
      Top             =   495
      Width           =   3480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   4635
   End
   Begin VB.Image imgPanel 
      Height          =   2985
      Index           =   0
      Left            =   0
      Picture         =   "Registered.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "Registered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Reg As Registry
Attribute Reg.VB_VarHelpID = -1

Private Sub Command1_Click()

'If Text1.Text = "" Then
 Reg.hKey = HKEY_LOCAL_MACHINE
 Reg.DataType = REG_SZ
 Reg.SubKey = "Software\Example"
 Reg.ValueName = "RegKey"

 Reg.Data = "qzsjh8"

'Else
' Reg.hKey = HKEY_LOCAL_MACHINE
' Reg.DataType = REG_SZ
' Reg.SubKey = "Software\Example"
' Reg.ValueName = "RegKey"

 'Reg.Data = "-1"

'End If
Reg.SaveSetting
'MDIForm.Load
Unload Me
End Sub

Private Sub Form_Load()

    Me.Height = 2985
    Me.Width = 4770
    Set Reg = New Registry
    Label2.Caption = "Install:" & vbCrLf & vbCrLf & "Press enter to proceed." & vbCrLf & "You can specify instructions here if you want"
End Sub

