VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFindDB 
   BorderStyle     =   0  'None
   Caption         =   "Lunec Database Found"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   345
      Left            =   1770
      TabIndex        =   0
      Top             =   1740
      Width           =   1545
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   990
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Path ..."
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Left            =   330
      TabIndex        =   1
      Top             =   420
      Width           =   1395
   End
   Begin VB.Image imgPanel 
      Height          =   2355
      Index           =   0
      Left            =   0
      Picture         =   "frmFindDB.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4995
   End
End
Attribute VB_Name = "frmFindDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()

  Me.Height = 2355
  Me.Width = 4995
  
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  
  CommonDialog1.DialogTitle = "Please Find 'Example.mdb' Database to Continue ..."
  CommonDialog1.FileName = "Example.mdb"
  CommonDialog1.InitDir = App.Path

  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "All Files (*.*)|*.*|MS Access 2000 Files " & _
        "(*.mdb)|*.mdb|Example Database File (Example.mdb)|Example.mdb"
  ' Specify default filter
  CommonDialog1.FilterIndex = 3
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  'CommonDialog1.FileName
  QuickRef.DBFileName = CommonDialog1.FileName
  
  Dim num As Integer
  
  'finds the position of '\' starting from right side of string
  num = InStrRev(QuickRef.DBFileName, "\", -1, vbBinaryCompare)
  QuickRef.DBFileName = Left(QuickRef.DBFileName, num - 1)
  
  Call WriteINI("Database", "DatabaseLocation", QuickRef.DBFileName)
  
  Label2.Caption = QuickRef.DBFileName
  Exit Sub
  
ErrHandler:
  Unload Me
  Call appTerminate
  Exit Sub

End Sub
