VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmExampledb 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Browse Students"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   -225
   ClientWidth     =   11250
   ControlBox      =   0   'False
   Icon            =   "frmExampledb.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmExampledb.frx":08CA
   ScaleHeight     =   6810
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnQKeys 
      Caption         =   "&Quick Keys"
      Height          =   315
      Left            =   9870
      TabIndex        =   88
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      Height          =   285
      Left            =   9870
      TabIndex        =   86
      Top             =   4980
      Width           =   1065
   End
   Begin VB.CommandButton btnNotes 
      Caption         =   "N&otes"
      Height          =   285
      Left            =   9870
      TabIndex        =   85
      Top             =   4680
      Width           =   1035
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0E0FF&
      Height          =   645
      Left            =   240
      TabIndex        =   83
      Top             =   4830
      Width           =   2565
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   3570
      Left            =   240
      TabIndex        =   82
      Top             =   870
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   6297
      _Version        =   393216
      BackColor       =   12640511
   End
   Begin VB.TextBox txtUpdate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   6930
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   585
      Width           =   2265
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtFirst_Name 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   5940
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1710
      Width           =   3255
   End
   Begin VB.TextBox txtLast_Name 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   5940
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1305
      Width           =   3255
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   285
      Left            =   9870
      TabIndex        =   44
      Top             =   4380
      Width           =   1065
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "&Print"
      Height          =   285
      Left            =   9870
      TabIndex        =   43
      Top             =   4080
      Width           =   1065
   End
   Begin VB.CommandButton btnColors 
      Caption         =   "&Colors"
      Height          =   285
      Left            =   9870
      TabIndex        =   42
      Top             =   3780
      Width           =   1065
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
      Height          =   285
      Left            =   9870
      TabIndex        =   41
      Top             =   3480
      Width           =   1065
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   285
      Left            =   9870
      TabIndex        =   40
      Top             =   3180
      Width           =   1065
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   285
      Left            =   9870
      TabIndex        =   39
      Top             =   2880
      Width           =   1065
   End
   Begin VB.CommandButton btnReload 
      Caption         =   "&Reload"
      Height          =   285
      Left            =   9870
      TabIndex        =   38
      Top             =   2580
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9900
      Top             =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Post Graduation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2325
      Index           =   3
      Left            =   4410
      TabIndex        =   72
      Tag             =   "ButtonLabel"
      Top             =   2250
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtEmployer 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   40
         TabIndex        =   26
         Top             =   1755
         Width           =   3240
      End
      Begin VB.TextBox txtJob_Title 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   25
         Top             =   1395
         Width           =   3240
      End
      Begin VB.TextBox txtInstitution 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   35
         TabIndex        =   24
         Top             =   720
         Width           =   3240
      End
      Begin VB.TextBox txtGrad_Studies 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   23
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Employer"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   615
         TabIndex        =   76
         Tag             =   "Label"
         Top             =   1770
         Width           =   810
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Job Title"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Tag             =   "Label"
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Institution"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   225
         TabIndex        =   74
         Tag             =   "Label"
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   210
         TabIndex        =   73
         Tag             =   "Label"
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Education"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2325
      Index           =   2
      Left            =   4410
      TabIndex        =   64
      Tag             =   "ButtonLabel"
      Top             =   2250
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtProgram 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   16
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtProgram_Length 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   5
         TabIndex        =   17
         Top             =   720
         Width           =   360
      End
      Begin VB.TextBox txtLast_Course 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2940
         MaxLength       =   30
         TabIndex        =   18
         Top             =   750
         Width           =   1845
      End
      Begin VB.TextBox txtFull_Part 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4440
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1110
         Width           =   345
      End
      Begin VB.TextBox txtCurrent_Past 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1080
         Width           =   360
      End
      Begin VB.TextBox txtEnroll 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1035
         MaxLength       =   30
         TabIndex        =   21
         Top             =   1800
         Width           =   1305
      End
      Begin VB.TextBox txtGrad 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   225
         TabIndex        =   71
         Tag             =   "Label"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Program Length"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   225
         TabIndex        =   70
         Tag             =   "Label"
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Course"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   1995
         TabIndex        =   69
         Tag             =   "Label"
         Top             =   750
         Width           =   840
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full or Part Time"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   3075
         TabIndex        =   68
         Tag             =   "Label"
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current or Past"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Tag             =   "Label"
         Top             =   1110
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Enroll Date"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Tag             =   "Label"
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Completion"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2535
         TabIndex        =   65
         Tag             =   "Label"
         Top             =   1830
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Permanent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2325
      Index           =   1
      Left            =   4410
      TabIndex        =   58
      Tag             =   "ButtonLabel"
      Top             =   2250
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtP_Address 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtP_City 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtP_Province 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3570
         MaxLength       =   20
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtP_PostalCode 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3810
         MaxLength       =   8
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtP_Telephone 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Perm. Address"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   210
         TabIndex        =   63
         Tag             =   "Label"
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Perm. City"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   210
         TabIndex        =   62
         Tag             =   "Label"
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2850
         TabIndex        =   61
         Tag             =   "Label"
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2850
         TabIndex        =   60
         Tag             =   "Label"
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Perm. Tel."
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   690
         TabIndex        =   59
         Tag             =   "Label"
         Top             =   1110
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2325
      Index           =   0
      Left            =   4410
      TabIndex        =   45
      Tag             =   "ButtonLabel"
      Top             =   2250
      Width           =   4935
      Begin VB.TextBox txtFirstNation_Telephone 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3570
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtFirstNation_Contact 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtFirstNation 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   35
         TabIndex        =   8
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtC_Telephone 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtC_PostalCode 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3810
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtC_Province 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3570
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtC_City 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtC_Address 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tel."
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   3210
         TabIndex        =   53
         Tag             =   "Label"
         Top             =   1830
         Width           =   255
      End
      Begin VB.Label lblFirstnation_Contact 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.N. Conact"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   210
         TabIndex        =   52
         Tag             =   "Label"
         Top             =   1830
         Width           =   1215
      End
      Begin VB.Label lblFirstNation 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "First Nation"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   330
         TabIndex        =   51
         Tag             =   "Label"
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label lblC_Telephone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tel."
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   690
         TabIndex        =   50
         Tag             =   "Label"
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2850
         TabIndex        =   49
         Tag             =   "Label"
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2850
         TabIndex        =   48
         Tag             =   "Label"
         Top             =   750
         Width           =   615
      End
      Begin VB.Label lblC_City 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current City"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   210
         TabIndex        =   47
         Tag             =   "Label"
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label lblC_Address 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Address"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   210
         TabIndex        =   46
         Tag             =   "Label"
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   1935
      TabIndex        =   89
      Tag             =   "Label"
      Top             =   585
      Width           =   825
   End
   Begin VB.Label lblQKeys 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Keys"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5250
      TabIndex        =   87
      Tag             =   "ButtonLabel"
      Top             =   4890
      UseMnemonic     =   0   'False
      Width           =   825
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Students by ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   300
      TabIndex        =   84
      Tag             =   "Label"
      Top             =   4560
      Width           =   1350
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   300
      TabIndex        =   81
      Tag             =   "Label"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblFrame 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Post Grad"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   3
      Left            =   3105
      TabIndex        =   80
      Tag             =   "ButtonLabel"
      Top             =   3585
      Width           =   1200
   End
   Begin VB.Label lblFrame 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Education"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   2
      Left            =   3105
      TabIndex        =   79
      Tag             =   "ButtonLabel"
      Top             =   3210
      Width           =   1200
   End
   Begin VB.Label lblFrame 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   1
      Left            =   3105
      TabIndex        =   78
      Tag             =   "ButtonLabel"
      Top             =   2850
      Width           =   1200
   End
   Begin VB.Label lblFrame 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   0
      Left            =   3105
      TabIndex        =   77
      Tag             =   "ButtonLabel"
      Top             =   2475
      Width           =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   4455
      X2              =   9360
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label lblUpdate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Update"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5820
      TabIndex        =   57
      Tag             =   "Label"
      Top             =   630
      Width           =   975
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4560
      TabIndex        =   56
      Tag             =   "Label"
      Top             =   630
      Width           =   255
   End
   Begin VB.Label lblFirstName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4620
      TabIndex        =   55
      Tag             =   "Label"
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label lblLastName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4620
      TabIndex        =   54
      Tag             =   "Label"
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   7
      Left            =   3075
      Picture         =   "frmExampledb.frx":B3CDC
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   1410
   End
   Begin VB.Label lblNotes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4380
      TabIndex        =   36
      Tag             =   "ButtonLabel"
      Top             =   4890
      UseMnemonic     =   0   'False
      Width           =   435
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8700
      TabIndex        =   28
      Tag             =   "ButtonLabel"
      Top             =   4890
      Width           =   495
   End
   Begin VB.Label lblReload 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reload"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6465
      TabIndex        =   29
      Tag             =   "ButtonLabel"
      Top             =   4890
      UseMnemonic     =   0   'False
      Width           =   510
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4380
      TabIndex        =   30
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   405
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5400
      TabIndex        =   31
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6585
      TabIndex        =   32
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   345
   End
   Begin VB.Label lblColors 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   7605
      TabIndex        =   33
      Tag             =   "ButtonLabel"
      Top             =   4890
      UseMnemonic     =   0   'False
      Width           =   465
   End
   Begin VB.Label lblPrint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   7620
      TabIndex        =   34
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   345
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   9870
      Picture         =   "frmExampledb.frx":B530E
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   1
      Left            =   9870
      Picture         =   "frmExampledb.frx":B6D30
      Top             =   510
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8745
      TabIndex        =   35
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   285
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse Student Records"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   360
      TabIndex        =   37
      Top             =   60
      Width           =   2325
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   8370
      Picture         =   "frmExampledb.frx":B841A
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   8
      Left            =   3075
      Picture         =   "frmExampledb.frx":B9E3C
      Stretch         =   -1  'True
      Top             =   2805
      Width           =   1410
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   9
      Left            =   3075
      Picture         =   "frmExampledb.frx":BB46E
      Stretch         =   -1  'True
      Top             =   3165
      Width           =   1410
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   10
      Left            =   3090
      Picture         =   "frmExampledb.frx":BCAA0
      Stretch         =   -1  'True
      Top             =   3525
      Width           =   1410
   End
   Begin VB.Image imgPrint 
      Height          =   375
      Left            =   7275
      Picture         =   "frmExampledb.frx":BE0D2
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgColors 
      Height          =   375
      Left            =   7275
      Picture         =   "frmExampledb.frx":BFAF4
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Image imgNew 
      Height          =   375
      Left            =   6195
      Picture         =   "frmExampledb.frx":C1516
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Image imgDelete 
      Height          =   375
      Left            =   5130
      Picture         =   "frmExampledb.frx":C2F38
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1065
   End
   Begin VB.Image imgSave 
      Height          =   375
      Left            =   4050
      Picture         =   "frmExampledb.frx":C495A
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Image imgReload 
      Height          =   375
      Left            =   6195
      Picture         =   "frmExampledb.frx":C637C
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   8370
      Picture         =   "frmExampledb.frx":C7D9E
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Image imgNotes 
      Height          =   375
      Left            =   4050
      Picture         =   "frmExampledb.frx":C97C0
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Image imgQKeys 
      Height          =   375
      Left            =   5130
      Picture         =   "frmExampledb.frx":CB1E2
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1065
   End
End
Attribute VB_Name = "frmExampledb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rep As New ADODB.Recordset
Dim rec As New ADODB.Recordset

Dim dontWatchText As Boolean
Dim iDirty As Boolean
Dim iStudentNameHasChanged As Boolean

Sub ClearAllFields()

txtLast_Name = ""
txtFirst_Name = ""
txtID = ""
txtUpdate = ""
txtC_Address = ""
txtC_City = ""
txtC_Province = ""
txtC_PostalCode = ""
txtC_Telephone = ""
txtFirstNation = ""
txtFirstNation_Contact = ""
txtFirstNation_Telephone = ""
txtP_Address = ""
txtP_City = ""
txtP_Province = ""
txtP_PostalCode = ""
txtP_Telephone = ""

txtProgram = ""
txtProgram_Length = ""
txtLast_Course = ""
txtEnroll = ""
txtGrad = ""
txtCurrent_Past = ""
txtFull_Part = ""

txtGrad_Studies = ""
txtInstitution = ""
txtJob_Title = ""
txtEmployer = ""

'frmNotes.txtComments = ""

iDirty = False
iStudentNameHasChanged = False

End Sub


Function SaveChanges() As Boolean

On Local Error GoTo SaveChangesError

'New code starts here -------------------
Dim fullnam As String
fullnam = txtLast_Name.Text & ", " & txtFirst_Name.Text
Dim IDnum As String
IDnum = Trim$(txtID.Text)

' Check for Current Past or Future studnet info correct
Dim tempCheck As String
txtCurrent_Past.Text = Trim$(UCase(txtCurrent_Past.Text))
tempCheck = txtCurrent_Past.Text
If tempCheck <> "F" And tempCheck <> "C" And tempCheck <> "P" Then
    If tempCheck = "" Then
        txtCurrent_Past.Text = "C"
    Else
        Dim FCPInput As String
        Dim FCP As Boolean
        FCP = False
        While (FCP = False)
            FCPInput = UCase(InputBox$("Please Specify" & vbCrLf & vbCrLf & "    F - Future Student" & vbCrLf & "    C - Current Student" & vbCrLf & "    P - Past Student", "Future, Current Or Past Student"))
            If (FCPInput = "P" Or FCPInput = "C" Or FCPInput = "F") Then FCP = True
        Wend
        txtCurrent_Past.Text = FCPInput
    End If
End If

' Check for full or part time student info correct
txtFull_Part.Text = Trim$(UCase(txtFull_Part.Text))
tempCheck = txtFull_Part.Text
If tempCheck <> "F" And tempCheck <> "P" Then
    If tempCheck = "" Then
        txtFull_Part.Text = "F"
    Else
        FCP = False
        While (FCP = False)
            FCPInput = UCase(InputBox$("Please Specify" & vbCrLf & vbCrLf & "    F - Full Time Student" & vbCrLf & "    P - Part Time Student", "Full or Part Time Student"))
            If (FCPInput = "P" Or FCPInput = "F") Then FCP = True
        Wend
        txtFull_Part.Text = FCPInput
    End If
End If

'IDnum = Trim$(Left$(DataList1.Text, InStr(1, DataList1.Text, " ", vbTextCompare)))
dbcn.BeginTrans
dbcn.Execute "UPDATE data SET " & _
    "Last_Name = '" & Trim$(txtLast_Name.Text) & _
    "', First_Name = '" & Trim$(txtFirst_Name.Text) & _
    "', Full_Name = '" & Trim$(fullnam) & _
    "', C_Address = '" & Trim$(txtC_Address.Text) & _
    "', C_City = '" & Trim$(txtC_City.Text) & _
    "', C_Province = '" & Trim$(txtC_Province.Text) & _
    "', C_PostalCode = '" & Trim$(txtC_PostalCode.Text) & _
    "', C_Telephone = '" & Trim$(txtC_Telephone.Text) & "' WHERE ID = " & IDnum
dbcn.Execute "UPDATE data SET " & _
    "FirstNation = '" & Trim$(txtFirstNation.Text) & _
    "', FirstNation_Contact = '" & Trim$(txtFirstNation_Contact.Text) & _
    "', FirstNation_Telephone = '" & Trim$(txtFirstNation_Telephone.Text) & _
    "', P_Address = '" & Trim$(txtP_Address.Text) & _
    "', P_City = '" & Trim$(txtP_City.Text) & _
    "', P_Province = '" & Trim$(txtP_Province.Text) & "' WHERE ID = " & IDnum
dbcn.Execute "UPDATE data SET " & _
    "P_PostalCode = '" & Trim$(txtP_PostalCode.Text) & _
    "', P_Telephone = '" & Trim$(txtP_Telephone.Text) & _
    "', Program = '" & Trim$(txtProgram.Text) & _
    "', Program_Length = '" & Trim$(txtProgram_Length.Text) & _
    "', Last_Course = '" & Trim$(txtLast_Course.Text) & _
    "', Enroll = '" & Trim$(txtEnroll.Text) & _
    "', Grad = '" & Trim$(txtGrad.Text) & "' WHERE ID = " & IDnum
dbcn.Execute "UPDATE data SET " & _
    "Current_Past = '" & Trim$(txtCurrent_Past.Text) & _
    "', Full_Part = '" & Trim$(txtFull_Part.Text) & _
    "', Grad_Studies = '" & Trim$(txtGrad_Studies.Text) & _
    "', Institution = '" & Trim$(txtInstitution.Text) & _
    "', Job_Title = '" & Trim$(txtJob_Title.Text) & _
    "', Employer = '" & Trim$(txtEmployer.Text) & "' WHERE ID = " & IDnum

dbcn.Execute "UPDATE data SET Updat = '" & Now & "' WHERE ID = " & IDnum
dbcn.CommitTrans
    
txtUpdate.Text = FormatDateTime(Now, vbLongDate)

SaveChanges = True
iDirty = False
QuickRef.UpdateNotes = QuickRef.NotesHaveChanged

Exit Function


SaveChangesError:
    'DB.Close
    MsgBox (" Error occurred while saving ... Partial or all data was updated")
    'Call WriteToErrorLog(Me.Name, "SaveChanges", Error, Err, True)
    Exit Function
    Resume Next

End Function
Private Sub btnColors_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Colors...
If lblColors.Enabled Then
    lblColors_Click
End If

End Sub
Private Sub btnDelete_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Delete...
If lblDelete.Enabled Then
    lblDelete_Click
End If

End Sub

Private Sub btnExit_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Exit...
If lblExit.Enabled Then
    lblExit_Click
End If

End Sub

Private Sub btnHelp_Click()
'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Help...
If lblHelp.Enabled Then
    lblHelp_Click
End If
End Sub

Private Sub btnNew_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'New...
If lblNew.Enabled Then
    lblNew_Click
End If

End Sub

Private Sub btnNotes_Click()
'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Exit...
If lblNotes.Enabled Then
    lblNotes_Click
End If
End Sub

Private Sub btnPrint_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Print...
If lblPrint.Enabled Then
    lblPrint_Click
End If

End Sub

Private Sub btnQKeys_Click()
    lblQKeys_Click
End Sub

Private Sub btnReload_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Reload...
If lblReload.Enabled Then
    lblReload_Click
End If

End Sub
Private Sub btnSave_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Save...
If lblSave.Enabled Then
    lblSave_Click
End If

End Sub

Private Sub DataList1_Click()

On Error GoTo DataListClickError
   
dontWatchText = True
Dim IDnum As String
Dim IDTemp As Integer

IDTemp = InStr(1, DataList1.Text, "(", vbTextCompare)
If IDTemp = 0 Then Exit Sub
IDnum = Trim$(Right$(DataList1.Text, Len(DataList1.Text) - IDTemp))
IDnum = Trim$(Left$(IDnum, Len(IDnum) - 1))

If IDnum <> "" Then

If rep.State <> adStateClosed Then rep.Close
rep.Source = "Select * From data where ID = " & IDnum
rep.Open
    
Call ClearAllFields
If Not IsNull(rep.Fields(0).Value) Then
    txtID.Text = rep.Fields(0).Value
End If
If Not IsNull(rep.Fields(2).Value) Then
    txtLast_Name.Text = rep.Fields(2).Value
End If
If Not IsNull(rep.Fields(3).Value) Then
    txtFirst_Name = rep.Fields(3).Value
End If

If Not IsNull(rep.Fields(4).Value) Then
    txtC_Address = rep.Fields(4).Value
End If
If Not IsNull(rep.Fields(5).Value) Then
    txtC_City.Text = rep.Fields(5).Value
End If
If Not IsNull(rep.Fields(6).Value) Then
    txtC_Province.Text = rep.Fields(6).Value
End If
If Not IsNull(rep.Fields(7).Value) Then
    txtC_PostalCode.Text = rep.Fields(7).Value
End If
If Not IsNull(rep.Fields(8).Value) Then
    txtC_Telephone.Text = rep.Fields(8).Value
End If

If Not IsNull(rep.Fields(9).Value) Then
    txtP_Address.Text = rep.Fields(9).Value
End If
If Not IsNull(rep.Fields(10).Value) Then
    txtP_City.Text = rep.Fields(10).Value
End If
If Not IsNull(rep.Fields(11).Value) Then
    txtP_Province.Text = rep.Fields(11).Value
End If
If Not IsNull(rep.Fields(12).Value) Then
    txtP_PostalCode.Text = rep.Fields(12).Value
End If
If Not IsNull(rep.Fields(13).Value) Then
    txtP_Telephone.Text = rep.Fields(13).Value
End If

If Not IsNull(rep.Fields(14).Value) Then
    txtFirstNation.Text = rep.Fields(14).Value
End If
If Not IsNull(rep.Fields(15).Value) Then
    txtFirstNation_Telephone.Text = rep.Fields(15).Value
End If
If Not IsNull(rep.Fields(16).Value) Then
    txtFirstNation_Contact.Text = rep.Fields(16).Value
End If

If Not IsNull(rep.Fields(17).Value) Then
    txtProgram.Text = rep.Fields(17).Value
End If
If Not IsNull(rep.Fields(18).Value) Then
    txtProgram_Length.Text = rep.Fields(18).Value
End If
If Not IsNull(rep.Fields(19).Value) Then
    txtLast_Course.Text = rep.Fields(19).Value
End If

If Not IsNull(rep.Fields(20).Value) Then
    txtEnroll.Text = rep.Fields(20).Value
End If
If Not IsNull(rep.Fields(21).Value) Then
    txtGrad.Text = rep.Fields(21).Value
End If
If Not IsNull(rep.Fields(22).Value) Then
    txtGrad_Studies.Text = rep.Fields(22).Value
End If
If Not IsNull(rep.Fields(23).Value) Then
    txtJob_Title.Text = rep.Fields(23).Value
End If
If Not IsNull(rep.Fields(24).Value) Then
    txtCurrent_Past.Text = rep.Fields(24).Value
End If
If Not IsNull(rep.Fields(25).Value) Then
    txtFull_Part.Text = rep.Fields(25).Value
End If
If Not IsNull(rep.Fields(26).Value) Then
    txtInstitution.Text = rep.Fields(26).Value
End If
If Not IsNull(rep.Fields(27).Value) Then
    txtEmployer.Text = rep.Fields(27).Value
End If
If Not IsNull(rep.Fields(28).Value) Then
    txtUpdate.Text = FormatDateTime(rep.Fields(28).Value, vbLongDate)
End If

dontWatchText = False
iDirty = False
iStudentNameHasChanged = False
End If
Exit Sub

DataListClickError:
    MsgBox "Error while processing Data List Event. Call Technical Support", vbCritical, "Error..."
    Exit Sub

    
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)

'Delete key...
If KeyCode = vbKeyDelete And lblDelete.Enabled = True Then
    lblDelete_Click
End If

End Sub


Private Sub DataList1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to view information on a student."

End Sub

Private Sub Form_DblClick()

On Local Error Resume Next

'Change the windowstate...
If Me.WindowState = vbNormal Then
    Me.WindowState = vbMinimized
ElseIf Me.WindowState = vbMinimized Then
    Me.WindowState = vbNormal
End If

End Sub
Private Sub Form_Load()

List1.AddItem "Future"
List1.AddItem "Current"
List1.AddItem "Past"
List1.ListIndex = 1

rec.ActiveConnection = dbcn
' change this  by closing and changing then reopening
rec.CursorLocation = adUseClient
'persisted in memory
rec.CursorType = adOpenStatic
rec.LockType = adLockBatchOptimistic

' open record set
Call popList

rep.ActiveConnection = dbcn
' change this  by closing and changing then reopening
rep.CursorLocation = adUseClient
'persisted in memory
rep.CursorType = adOpenStatic
rep.LockType = adLockBatchOptimistic


On Local Error Resume Next
'Load INI Settings...
Call LoadINISettings

'Set Colors...
Call SetColors(Me)

'Set form width and height...
Me.Height = QuickRef.LargeMenuHeight
Me.Width = QuickRef.LargeMenuWidth

iDirty = False
iStudentNameHasChanged = False
QuickRef.NotesHaveChanged = False
QuickRef.UpdateNotes = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = val(ReadINI(Me.Name, "Left"))
Me.Top = val(ReadINI(Me.Name, "Top"))

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = ""

'Move the form...
If Button = vbLeftButton And Me.WindowState = vbNormal Then
    Call DragForm(Me)
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
If rec.State <> adStateClosed Then rec.Close
Set rec = Nothing
    
If rep.State <> adStateClosed Then rep.Close
Set rep = Nothing

'Save INI Settings...
Call SaveINISettings

'Call appTerminate

End Sub

Private Sub imgColors_Click()

lblColors_Click

End Sub
Private Sub imgColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgColors.Picture = imgButton(1).Picture
    lblColors.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgColors.Picture = imgButton(0).Picture
lblColors.ForeColor = lButtonForeColor

End Sub
Private Sub imgDelete_Click()

lblDelete_Click

End Sub
Private Sub imgDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblDelete.Enabled = True Then
    imgDelete.Picture = imgButton(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgDelete.Picture = imgButton(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub
Private Sub imgExit_Click()

lblExit_Click

End Sub
Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgButton(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgExit.Picture = imgButton(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub

Private Sub imgHelp_Click()

lblHelp_Click

End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgButton(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgHelp.Picture = imgButton(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub
Private Sub imgLabelHolder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form...
If Button = vbLeftButton And Me.WindowState = vbNormal Then
    Call DragForm(Me)
End If

End Sub

Private Sub imgNew_Click()

lblNew_Click

End Sub
Private Sub imgNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgButton(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgNew.Picture = imgButton(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub

Private Sub imgNotes_Click()

If lblNotes.Enabled = True Then
    lblNotes_Click
End If

End Sub
Private Sub imgNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblNotes.Enabled = True Then
    imgNotes.Picture = imgButton(1).Picture
    lblNotes.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgNotes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgNotes.Picture = imgButton(0).Picture
lblNotes.ForeColor = lButtonForeColor

End Sub
Private Sub imgPrint_Click()

If lblPrint.Enabled = True Then
    lblPrint_Click
End If

End Sub
Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblPrint.Enabled = True Then
    imgPrint.Picture = imgButton(1).Picture
    lblPrint.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgPrint.Picture = imgButton(0).Picture
lblPrint.ForeColor = lButtonForeColor

End Sub

Private Sub imgQKeys_Click()

    lblQKeys_Click
    
End Sub

Private Sub imgQKeys_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgQKeys.Picture = imgButton(1).Picture
    lblQKeys.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgQKeys_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgQKeys.Picture = imgButton(0).Picture
lblQKeys.ForeColor = lButtonForeColor

End Sub

Private Sub imgReload_Click()
If lblReload.Enabled = True Then
    lblReload_Click
End If

End Sub
Private Sub imgReload_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblReload.Enabled = True Then
    imgReload.Picture = imgButton(1).Picture
    lblReload.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgReload_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgReload.Picture = imgButton(0).Picture
lblReload.ForeColor = lButtonForeColor

End Sub
Private Sub imgSave_Click()

If lblSave.Enabled = True Then
    lblSave_Click
End If

End Sub
Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblSave.Enabled = True Then
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgSave.Picture = imgButton(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub lblCaptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form...
If Button = vbLeftButton And Me.WindowState = vbNormal Then
    Call DragForm(Me)
End If

End Sub

Private Sub lblColors_Click()

frmColors.Show
frmColors.ZOrder

End Sub
Private Sub lblColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgColors.Picture = imgButton(1).Picture
    lblColors.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to go to the colors window."

End Sub
Private Sub lblColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgColors.Picture = imgButton(0).Picture
lblColors.ForeColor = lButtonForeColor

End Sub


Private Sub lblDelete_Click()

    If lblDelete.Enabled = True Then
        'MsgBox QuickRef.ID
        QuickRef.ID = txtID.Text
        If (DeleteStudent(DataList1.Text)) Then
            ClearAllFields
            List1_Click
        End If
    End If
    
End Sub

Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblDelete.Enabled = True Then
    imgDelete.Picture = imgButton(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to delete this student. Once a student has been deleted from the 'Past' list the results are permanent."

End Sub
Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgDelete.Picture = imgButton(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub
Private Sub lblExit_Click()

Unload frmExampledb

End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgButton(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to exit this window."

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgExit.Picture = imgButton(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub

Private Sub lblFrame_Click(Index As Integer)

    Call ShowFrame(Index)
    
End Sub

Private Sub lblFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then Help.HelpText = "Click this to see a student's personal information and First Nation information."
If Index = 1 Then Help.HelpText = "Click this to see a student's permanent address information."
If Index = 2 Then Help.HelpText = "Click this to see a student's education information."
If Index = 3 Then Help.HelpText = "Click this to see a student's post graduation information (Job title and Employer)."
End Sub

Private Sub lblHelp_Click()

Help.HelpCallingForm = Me.Name

frmHelper.Show
frmHelper.ZOrder

End Sub

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgButton(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to show the help window."

End Sub
Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgHelp.Picture = imgButton(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub


Private Sub lblNew_Click()

On Local Error GoTo NewError

'Dim x As Long
Dim sInput As String
Dim sInput2 As String

'Clear all fields first...
dontWatchText = True
Call ClearAllFields
dontWatchText = False

'Get new Student Last name...
Do
sInput = Trim$(InputBox$("Enter the LAST name of the new student." & vbCrLf & "(must be less than 30 characters)", "New..."))
'Nothing entered...
If sInput = "" Then Exit Sub
Loop Until (Len(sInput) < 30)

'Strip out invalid characters in the student last name...
If InStr(sInput, "'") > 0 Then
    Mid$(sInput, InStr(sInput, "'"), 1) = Chr$(32)
ElseIf InStr(sInput, "#") > 0 Then
    Mid$(sInput, InStr(sInput, "#"), 1) = Chr$(32)
ElseIf InStr(sInput, "$") > 0 Then
    Mid$(sInput, InStr(sInput, "$"), 1) = Chr$(32)
ElseIf InStr(sInput, "%") > 0 Then
    Mid$(sInput, InStr(sInput, "%"), 1) = Chr$(32)
ElseIf InStr(sInput, "&") > 0 Then
    Mid$(sInput, InStr(sInput, "&"), 1) = Chr$(32)
ElseIf InStr(sInput, "*") > 0 Then
    Mid$(sInput, InStr(sInput, "*"), 1) = Chr$(32)
ElseIf InStr(sInput, "(") > 0 Then
    Mid$(sInput, InStr(sInput, "("), 1) = Chr$(32)
ElseIf InStr(sInput, ")") > 0 Then
    Mid$(sInput, InStr(sInput, ")"), 1) = Chr$(32)
End If

Do
'Get new Student First name...
sInput2 = Trim$(InputBox$("Enter the FIRST name of the new student." & vbCrLf & "(must be less than 20 characters)", "New..."))
'Nothing entered...
If sInput2 = "" Then Exit Sub
Loop Until (Len(sInput2) < 20)

'Strip out invalid characters in the student last name...
If InStr(sInput2, "'") > 0 Then
    Mid$(sInput2, InStr(sInput2, "'"), 1) = Chr$(32)
ElseIf InStr(sInput2, "#") > 0 Then
    Mid$(sInput2, InStr(sInput2, "#"), 1) = Chr$(32)
ElseIf InStr(sInput2, "$") > 0 Then
    Mid$(sInput2, InStr(sInput2, "$"), 1) = Chr$(32)
ElseIf InStr(sInput2, "%") > 0 Then
    Mid$(sInput2, InStr(sInput2, "%"), 1) = Chr$(32)
ElseIf InStr(sInput2, "&") > 0 Then
    Mid$(sInput2, InStr(sInput2, "&"), 1) = Chr$(32)
ElseIf InStr(sInput2, "*") > 0 Then
    Mid$(sInput2, InStr(sInput2, "*"), 1) = Chr$(32)
ElseIf InStr(sInput2, "(") > 0 Then
    Mid$(sInput2, InStr(sInput2, "("), 1) = Chr$(32)
ElseIf InStr(sInput2, ")") > 0 Then
    Mid$(sInput2, InStr(sInput2, ")"), 1) = Chr$(32)
End If

'Create the new student...
'Convert first name and last name to full name
Dim fulln As String
Dim cur_past As String
Dim fullprt As String

sInput = UCase(Mid(sInput, 1, 1)) + LCase(Mid(sInput, 2))
sInput2 = UCase(Mid(sInput2, 1, 1)) + LCase(Mid(sInput2, 2))
fulln = sInput & ", " & sInput2
txtLast_Name.Text = sInput
txtFirst_Name.Text = sInput2
' determine if Future Current Past
cur_past = Left$(List1.Text, 1)
' init as default Full Time Student
fullprt = "F"

Dim xdate As Date
xdate = Now
dbcn.BeginTrans
dbcn.Execute "INSERT INTO data (Full_Name, First_Name, Last_Name, " & _
    "Current_Past, Full_Part, Updat) " & _
    "Values ('" & fulln & "', '" & sInput2 & "', '" & sInput & "', '" & _
    cur_past & " ', '" & fullprt & "', '" & xdate & "')"
dbcn.CommitTrans
    
'get the ID Num for the new user
If rep.State <> adStateClosed Then rep.Close
rep.Source = "Select * From data Where (Full_Name = '" & fulln & "') AND (Updat LIKE '" & xdate & "')"
'rep.Source = "Select * From data Where (Full_Name = '" & fulln & "')" ' AND (" & _
    "Updat = '" & xdate & "')"
rep.Open
If rep.RecordCount <> 1 Then GoTo NewError
txtID = rep.Fields(0).Value
rep.Close
    
txtCurrent_Past.Text = cur_past
txtUpdate.Text = FormatDateTime(xdate, vbLongDate)
txtFull_Part = fullprt

iDirty = False
iStudentNameHasChanged = False
Exit Sub

NewError:
    MsgBox (" Error occurred while creating new record ... Contact Technical Support.")
    Exit Sub
    Resume Next

End Sub
Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgButton(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to create a new student."

End Sub
Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgNew.Picture = imgButton(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub

Private Sub lblNotes_Click()

If lblNotes.Enabled = True Then
    QuickRef.Full_Name = DataList1.Text
    frmNotes.Show
End If

End Sub
Private Sub lblNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblNotes.Enabled = True Then
    imgNotes.Picture = imgButton(1).Picture
    lblNotes.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblNotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to view the notes in a larger window."

End Sub
Private Sub lblNotes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgNotes.Picture = imgButton(0).Picture
lblNotes.ForeColor = lButtonForeColor

End Sub

Private Sub lblPrint_Click()

If lblPrint.Enabled = True Then
    Set SingleStudent.DataSource = rep
    SingleStudent.Orientation = rptOrientPortrait
    SingleStudent.Show vbModal
End If

End Sub
Private Sub lblPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblPrint.Enabled = True Then
    imgPrint.Picture = imgButton(1).Picture
    lblPrint.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to print this student's information to the printer."

End Sub
Private Sub lblPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgPrint.Picture = imgButton(0).Picture
lblPrint.ForeColor = lButtonForeColor

End Sub

Private Sub lblQKeys_Click()

    'Load frmQKeys
    Timer1.Enabled = False
    frmQKeys.Show vbModal
    Timer1.Enabled = True
    
End Sub

Private Sub lblQKeys_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgQKeys.Picture = imgButton(1).Picture
    lblQKeys.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblQKeys_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to display the Quick Keys for easier navigation."

End Sub

Private Sub lblQKeys_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgQKeys.Picture = imgButton(0).Picture
lblQKeys.ForeColor = lButtonForeColor

End Sub

Private Sub lblReload_Click()

If lblReload.Enabled = True Then
    List1_Click
End If

End Sub
Private Sub lblReload_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblReload.Enabled = True Then
    imgReload.Picture = imgButton(1).Picture
    lblReload.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblReload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to reload this student, losing any changes made."

End Sub
Private Sub lblReload_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgReload.Picture = imgButton(0).Picture
lblReload.ForeColor = lButtonForeColor

End Sub
Private Sub lblSave_Click()
    
If lblSave.Enabled = True Then

    Call SaveChanges
End If

End Sub


Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And lblSave.Enabled = True Then
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to save any changes you have made to this student."

End Sub
Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgSave.Picture = imgButton(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub List1_Click()
    If frmExampledb.Visible = True Then
        Call popList
    End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to view the Future, Current, or Past Student list above."

End Sub

Function ShowFrame(fFrame As Integer)
Dim X As Integer

For X = 0 To 3
    If X <> fFrame Then
        Frame1(X).Visible = False
        lblFrame(X).ForeColor = lButtonForeColor
        lblFrame(X).Font.Bold = False
    Else
        Frame1(X).Visible = True
        lblFrame(X).ForeColor = lLabelForeColor
        lblFrame(X).Font.Bold = True
    End If
Next X
    
End Function


Private Sub lstStudents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to view information on a student."

End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Timer1_Timer()

On Local Error Resume Next

Dim X As Long
Dim iTempDirty As Boolean

'Set Colors...
If QuickRef.UpdateColors = True Then
    Call LoadProgramColors
    QuickRef.UpdateColors = False
    For X = 0 To Forms.Count - 1
        Call SetColors(Forms(X))
    Next X
End If

'Notes...
If txtID.Text = "" Or iDirty = True Then
    lblNotes.Enabled = False
    imgNotes.Enabled = False
ElseIf Not txtID.Text And iDirty = False And lblNotes.Enabled = False Then
    lblNotes.Enabled = True
    imgNotes.Enabled = True
End If

'Students Listbox...
If DataList1.Enabled = True And iDirty = True Then
    DataList1.Enabled = False
    List1.Enabled = False
ElseIf DataList1.Enabled = False And iDirty = False Then
    DataList1.Enabled = True
    List1.Enabled = True
End If

'Delete...
'MsgBox DataList1.Text
If lblDelete.Enabled = True And iDirty = True Then
    lblDelete.Enabled = False
    imgDelete.Enabled = False
ElseIf lblDelete.Enabled = False And txtID.Text <> "" And iDirty = False Then
    lblDelete.Enabled = True
    imgDelete.Enabled = True
End If

'New...
If lblNew.Enabled = True And iDirty = True Then
    lblNew.Enabled = False
    imgNew.Enabled = False
ElseIf lblNew.Enabled = False And iDirty = False Then
    lblNew.Enabled = True
    imgNew.Enabled = True
End If

'Save...
If lblSave.Enabled = True And iDirty = False Then
    lblSave.Enabled = False
    imgSave.Enabled = False
ElseIf lblSave.Enabled = False And iDirty = True Then
    lblSave.Enabled = True
    imgSave.Enabled = True
End If

'Reload...
If lblReload.Enabled = True And iDirty = False Then
    lblReload.Enabled = False
    imgReload.Enabled = False
ElseIf lblReload.Enabled = False And iDirty = True Then
    lblReload.Enabled = True
    imgReload.Enabled = True
End If

'Print...
If iDirty = True Or txtID.Text = "" Then
    lblPrint.Enabled = False
    imgPrint.Enabled = False
ElseIf iDirty = False And txtID.Text <> "" Then
    lblPrint.Enabled = True
    imgPrint.Enabled = True
End If

End Sub

Private Sub txtC_Address_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtC_Address_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtC_Address_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtC_City_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtC_City_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtC_City_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtC_PostalCode_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtC_PostalCode_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtC_PostalCode_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtC_Province_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtC_Province_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtC_Province_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtC_Telephone_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtC_Telephone_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtC_Telephone_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub


Private Sub txtCurrent_Past_Change()

    If dontWatchText = False Then iDirty = True
   
End Sub

Private Sub txtCurrent_Past_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtCurrent_Past_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub


Private Sub txtEmployer_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtEmployer_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtEmployer_KeyPress(KeyAscii As Integer)
'Enter or Tab Key...
'MsgBox "Asccii = " & KeyAscii, vbInformation, "Keypress"

If KeyAscii = vbKeyReturn Then
    Call ShowFrame(0)
    txtC_Address.SetFocus
    KeyAscii = 0
End If

End Sub


Private Sub txtEnroll_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtEnroll_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtEnroll_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub


Private Sub txtFirst_Name_Change()

If dontWatchText = False Then iDirty = True
If dontWatchText = False Then iStudentNameHasChanged = True

End Sub

Private Sub txtFirst_Name_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtFirst_Name_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtFirstNation_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtFirstNation_Contact_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtFirstNation_Contact_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtFirstNation_Contact_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtFirstNation_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtFirstNation_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtFirstNation_Telephone_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtFirstNation_Telephone_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    SendKeys "{TAB}", True
    Call ShowFrame(1)
    txtP_Address.SetFocus
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtFirstNation_Telephone_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    Call ShowFrame(1)
    txtP_Address.SetFocus
    'SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtFull_Part_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtFull_Part_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtFull_Part_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtGrad_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtGrad_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    Call ShowFrame(3)
    txtGrad_Studies.SetFocus
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call ShowFrame(3)
    txtGrad_Studies.SetFocus
End If

End Sub
Private Sub txtGrad_Studies_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtGrad_Studies_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtGrad_Studies_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Help.HelpText = "This field cannot be edited. The system automatically assigns a new record number when a new record is created."
End Sub

Private Sub txtInstitution_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtInstitution_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtInstitution_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub



Private Sub txtJob_Title_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtJob_Title_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtJob_Title_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub
Private Sub txtLast_Course_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtLast_Course_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtLast_Course_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtLast_Name_Change()

If dontWatchText = False Then iDirty = True
If dontWatchText = False Then iStudentNameHasChanged = True

End Sub

Private Sub txtLast_Name_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtLast_Name_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtP_Address_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtP_Address_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtP_Address_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtP_City_Change()

If dontWatchText = False Then iDirty = True
    
End Sub

Private Sub txtP_City_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtP_City_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtP_PostalCode_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtP_PostalCode_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    SendKeys "{TAB}", True
    Call ShowFrame(2)
    txtProgram.SetFocus
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtP_PostalCode_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    Call ShowFrame(2)
    txtProgram.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtP_Province_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtP_Province_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtP_Province_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtP_Telephone_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtP_Telephone_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtP_Telephone_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtProgram_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtProgram_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtProgram_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub txtProgram_Length_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtProgram_Length_KeyDown(KeyCode As Integer, Shift As Integer)
'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If
End Sub

Private Sub txtProgram_Length_KeyPress(KeyAscii As Integer)
'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub


Private Sub txtUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "This field cannot be edited. Records are automatically assigned today's date when edited."

End Sub

Private Sub popList()
    Dim spce As String
    If rec.State <> adStateClosed Then rec.Close
    rec.Source = "Select Full_Name & '  (' & ID & ')' As firstCol " & _
    ", Full_Name as Sortd From data Where Current_Past = '" & Left$(List1.Text, 1) & "' ORDER BY 2"
    rec.Open
   
    Set DataList1.RowSource = rec
    lblCount = "Total: " & rec.RecordCount
    DataList1.ListField = "firstCol"
    Call ClearAllFields
   
End Sub
