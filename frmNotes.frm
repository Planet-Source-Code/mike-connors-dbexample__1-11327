VERSION 5.00
Begin VB.Form frmNotes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Notes"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmNotes.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Exit&o"
      Height          =   345
      Left            =   7650
      TabIndex        =   6
      Top             =   1770
      Width           =   1095
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H00C0E0FF&
      Height          =   3045
      Left            =   240
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      Width           =   6525
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7530
      Top             =   960
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   495
      TabIndex        =   1
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click for Help..."
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblClear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4140
      TabIndex        =   2
      Tag             =   "ButtonLabel"
      Top             =   3825
      Width           =   360
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   1
      Left            =   7530
      Picture         =   "frmNotes.frx":62BC2
      Top             =   510
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   7530
      Picture         =   "frmNotes.frx":642AC
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblSave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5100
      TabIndex        =   3
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6180
      TabIndex        =   5
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   5790
      Picture         =   "frmNotes.frx":65CCE
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Image imgSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmNotes.frx":676F0
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Label lblNotes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   6855
   End
   Begin VB.Image imgClear 
      Height          =   375
      Left            =   3810
      Picture         =   "frmNotes.frx":69112
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   240
      Picture         =   "frmNotes.frx":6AB34
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ignoreChange As Boolean
Dim iDirty As Boolean
Dim rep As New ADODB.Recordset

Private Sub Command1_Click()
    lblExit_Click
End Sub

Private Sub Form_Load()

frmExampledb.Timer1.Enabled = False
frmExampledb.Enabled = False

'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Width and Height...
Me.Width = QuickRef.MediumMenuWidth
Me.Height = QuickRef.MediumMenuHeight

lblNotes.Caption = "Notes for " & QuickRef.Full_Name

rep.ActiveConnection = dbcn
' change this  by closing and changing then reopening
rep.CursorLocation = adUseClient
'persisted in memory
rep.CursorType = adOpenStatic
rep.LockType = adLockBatchOptimistic

If rep.State <> adStateClosed Then rep.Close
rep.Source = "Select Comments From data where ID = " & frmExampledb.txtID.Text
rep.Open

If Not IsNull(rep.Fields(0).Value) Then
    txtComments.Text = rep.Fields(0).Value
End If

iDirty = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = val(ReadINI(Me.Name, "Left"))
Me.Top = val(ReadINI(Me.Name, "Top"))

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = ""

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

On Local Error Resume Next

Dim X As Long

'Save changes...
If iDirty Then
    If MsgBox("Changes were not saved. Do you want to exit anyway?", vbYesNo + vbQuestion, "Exit without saving...") = vbNo Then
        Cancel = True
        'iDirty = False
        Exit Sub
    End If
End If

'Save INI Settings...
Call SaveINISettings

frmExampledb.Enabled = True
frmExampledb.Timer1.Enabled = True
rep.Close

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

End Sub
Private Sub imgClear_Click()

lblClear_Click

End Sub

Private Sub imgClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClear.Picture = imgButton(1).Picture
    lblClear.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgClear.Picture = imgButton(0).Picture
lblClear.ForeColor = lButtonForeColor

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

Private Sub imgSave_Click()

lblSave_Click

End Sub
Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblCategories_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub

Private Sub lblClear_Click()
'Confirm...
If MsgBox("Are you sure you want to clear?", vbYesNo + vbQuestion + vbDefaultButton2, "Clear Notes...") = vbNo Then
    Exit Sub
End If
txtComments = ""

End Sub
Private Sub lblClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClear.Picture = imgButton(1).Picture
    lblClear.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to clear the notes."

End Sub
Private Sub lblClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgClear.Picture = imgButton(0).Picture
lblClear.ForeColor = lButtonForeColor

End Sub
Private Sub lblExit_Click()

'Unload the help window...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If
Unload Me


End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgButton(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to exit the notes window."

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgExit.Picture = imgButton(0).Picture
lblExit.ForeColor = lButtonForeColor

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

Private Sub lblNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub


Private Sub lblSave_Click()

Call SaveChanges2

End Sub
Function SaveChanges2() As Boolean

On Local Error GoTo SaveChanges2Error

Dim IDnum As String
IDnum = Trim$(frmExampledb.txtID.Text)
dbcn.BeginTrans
dbcn.Execute "UPDATE data SET " & _
    "Comments = '" & Trim$(txtComments.Text) & "', " & _
    "Updat = '" & Now & "' WHERE ID = " & IDnum
dbcn.CommitTrans

frmExampledb.txtUpdate.Text = FormatDateTime(Now, vbLongDate)

'QuickRef.UpdateNotes = QuickRef.NotesHaveChanged
iDirty = False
SaveChanges2 = True

Exit Function

SaveChanges2Error:
    'DB.Close
    'Call WriteToErrorLog(Me.Name, "SaveChangesError", Error, Err, True)
    MsgBox "Error saving Notes... Exit and seek Technical Support", vbCritical, "Error"
    Exit Function

End Function
Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to save any changes that you have made."

End Sub
Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgSave.Picture = imgButton(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub Timer1_Timer()

On Local Error Resume Next

Dim iTempDirty As Boolean

'Notes...
'If QuickRef.UpdateNotes And QuickRef.CallingForm <> "Notes" Then
'    iTempDirty = iDirty
    'Call GetNotes(txtComments)
'    iDirty = iTempDirty
'End If

'Save...
If imgSave.Enabled = False And iDirty = True Then
    imgSave.Enabled = True
    lblSave.Enabled = True
ElseIf imgSave.Enabled = True And iDirty = False Then
    imgSave.Enabled = False
    lblSave.Enabled = False
End If

'Clear...
If imgClear.Enabled = False And txtComments <> "" Then
    imgClear.Enabled = True
    lblClear.Enabled = True
ElseIf imgClear.Enabled = True And txtComments = "" Then
    imgClear.Enabled = False
    lblClear.Enabled = False
End If

End Sub
Private Sub txtComments_Change()

'QuickRef.NotesHaveChanged = True
'QuickRef.CallingForm = "Notes"
iDirty = True

End Sub

Private Sub txtComments_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Help.HelpText = "Type any " & lblNotes.Caption & " in this area, then click SAVE."
End Sub
