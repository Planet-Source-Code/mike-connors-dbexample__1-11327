VERSION 5.00
Begin VB.Form frmQKeys 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&QKeys Exit"
      Height          =   255
      Left            =   2790
      TabIndex        =   2
      Top             =   3195
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2685
      HideSelection   =   0   'False
      Left            =   210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmQKeys.frx":0000
      Top             =   420
      Width           =   3765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALT + Q  =  Exit / Enter"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   210
      TabIndex        =   1
      Top             =   3180
      Width           =   2340
   End
   Begin VB.Image imgPanel 
      Height          =   3555
      Index           =   0
      Left            =   0
      Picture         =   "frmQKeys.frx":0006
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4275
   End
End
Attribute VB_Name = "frmQKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    frmQKeys.Height = 3555
    frmQKeys.Width = 4260
    
    Text1.Text = "Quick Keys have been installed for the Buttons on this form" & vbCrLf & vbCrLf & _
        "ALT + R  -  Reload" & vbCrLf & _
        "ALT + S  -  Save Current" & vbCrLf & _
        "ALT + D  -  Delete Current" & vbCrLf & _
        "ALT + N  -  New Record" & vbCrLf & _
        "ALT + C  -  Change Colors" & vbCrLf & _
        "ALT + P  -  Print Current" & vbCrLf & _
        "ALT + X  -  Exit Browse" & vbCrLf & _
        "ALT + O  -  View Notes"
End Sub

