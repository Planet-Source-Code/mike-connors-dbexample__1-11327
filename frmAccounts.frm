VERSION 5.00
Begin VB.Form frmAccounts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "System Accounts"
   ClientHeight    =   4770
   ClientLeft      =   420
   ClientTop       =   0
   ClientWidth     =   8745
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAccounts.frx":0000
   ScaleHeight     =   4770
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkUserSecurity 
      Height          =   195
      Left            =   4260
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtLastLogin 
      Height          =   285
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Users Last Login"
      Top             =   2340
      Width           =   2445
   End
   Begin VB.TextBox txtPassWord 
      Height          =   285
      Left            =   4260
      TabIndex        =   3
      ToolTipText     =   "Users Password"
      Top             =   1920
      Width           =   2445
   End
   Begin VB.TextBox txtFullName 
      Height          =   285
      Left            =   4260
      TabIndex        =   2
      ToolTipText     =   "Users Full Name"
      Top             =   1500
      Width           =   2445
   End
   Begin VB.TextBox txtLoginName 
      Height          =   285
      Left            =   4260
      TabIndex        =   1
      ToolTipText     =   "Users Login Name"
      Top             =   1080
      Width           =   2445
   End
   Begin VB.ListBox lstUsers 
      Height          =   1620
      ItemData        =   "frmAccounts.frx":62BC2
      Left            =   300
      List            =   "frmAccounts.frx":62BC4
      TabIndex        =   0
      Top             =   1020
      Width           =   2565
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7230
      Top             =   840
   End
   Begin VB.Label lblUserSecurity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Access"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   3150
      TabIndex        =   17
      Tag             =   "Label"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblLastlogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Login"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   3
      Left            =   3150
      TabIndex        =   15
      Tag             =   "Label"
      Top             =   2370
      Width           =   735
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4890
      TabIndex        =   13
      Tag             =   "ButtonLabel"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   2610
      TabIndex        =   12
      Tag             =   "ButtonLabel"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   2
      Left            =   3150
      TabIndex        =   11
      Tag             =   "Label"
      Top             =   1950
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   1
      Left            =   3150
      TabIndex        =   10
      Tag             =   "Label"
      Top             =   1530
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   0
      Left            =   3150
      TabIndex        =   9
      Tag             =   "Label"
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label lblUsers 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   330
      TabIndex        =   8
      Tag             =   "Label"
      Top             =   750
      Width           =   405
   End
   Begin VB.Label lblSystemAccounts 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit User Access"
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
      Height          =   240
      Left            =   390
      TabIndex        =   7
      Top             =   60
      Width           =   1560
   End
   Begin VB.Image imgOKPicture 
      Height          =   375
      Index           =   1
      Left            =   7230
      Picture         =   "frmAccounts.frx":62BC6
      Top             =   450
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgOKPicture 
      Height          =   360
      Index           =   0
      Left            =   7230
      Picture         =   "frmAccounts.frx":642B0
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   780
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   1740
      TabIndex        =   5
      Tag             =   "ButtonLabel"
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6090
      TabIndex        =   4
      Tag             =   "ButtonLabel"
      Top             =   3360
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   5730
      Picture         =   "frmAccounts.frx":65CD2
      Stretch         =   -1  'True
      Top             =   3270
      Width           =   975
   End
   Begin VB.Image imgNew 
      Height          =   375
      Left            =   450
      Picture         =   "frmAccounts.frx":676F4
      Stretch         =   -1  'True
      Top             =   3270
      Width           =   975
   End
   Begin VB.Image imgSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1410
      Picture         =   "frmAccounts.frx":69116
      Stretch         =   -1  'True
      Top             =   3270
      Width           =   975
   End
   Begin VB.Image imgDelete 
      Height          =   375
      Left            =   2370
      Picture         =   "frmAccounts.frx":6AB38
      Stretch         =   -1  'True
      Top             =   3270
      Width           =   975
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   4650
      Picture         =   "frmAccounts.frx":6C55A
      Stretch         =   -1  'True
      Top             =   3270
      Width           =   975
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As New ADODB.Recordset
Dim flds As New ADODB.Recordset
Dim iDirty As Boolean
Dim dontWatchText As Boolean

Sub ClearAllControls()

On Local Error Resume Next

txtLoginName = ""
txtFullName = ""
txtPassWord = ""
txtLastLogin = ""
chkUserSecurity = False

iDirty = False

End Sub
Sub DeleteAccount()

On Local Error GoTo DeleteAccountError

If UserSecurity = False Then
    MsgBox "You do not have authorization to delete your account. Contact your System Administrator", vbInformation, "Authorization..."
    Exit Sub
End If

' Check for last Admin. If so, then disable save if Admin Access is off
If flds.State <> adStateClosed Then flds.Close
flds.Source = "Select * From Usernames where Admin = True"
flds.Open

If flds.RecordCount = 1 And chkUserSecurity = 1 Then
    MsgBox "You cannot remove the only remaining Admin User.", vbCritical, "Admin Access"
    flds.Close
    Exit Sub
End If
flds.Close

'Confirm...
If MsgBox("Are you sure you want to delete " & UCase(txtFullName) & "s' account?", vbYesNo + vbQuestion, "Delete Account...") = vbNo Then
    Exit Sub
End If

If flds.State <> adStateClosed Then flds.Close
flds.Source = "Select * From Usernames"
flds.Open
If flds.RecordCount = 1 Then
    MsgBox "Deleting the last account would prevent anyone from accessing the software. You must make another account before deleting " & UCase(txtFullName), vbInformation, "Delete"
    flds.Close
    Exit Sub
End If
flds.Close

Dim IDnum As String
Dim IDTemp As Integer

IDTemp = InStr(1, lstUsers.Text, "-", vbTextCompare)
IDnum = Trim$(Left$(lstUsers.Text, IDTemp - 1))

dbcn.BeginTrans
dbcn.Execute "DELETE * FROM Usernames WHERE ID = " & IDnum
dbcn.CommitTrans
Call popList
Exit Sub

DeleteAccountError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "DeleteAccountError", Error, Err, True)
    Exit Sub

End Sub

Function SaveChanges() As Boolean

On Local Error GoTo SaveChangesError

'New code starts here -------------------
Dim IDnum As String
Dim IDTemp As Integer

IDTemp = InStr(1, lstUsers.Text, "-", vbTextCompare)
IDnum = Trim$(Left$(lstUsers.Text, IDTemp - 1))

' Check for info correct
If (Trim$(txtLoginName.Text) = "" Or Trim$(txtFullName.Text) = "" Or Trim$(txtPassWord.Text) = "") Then
    MsgBox "Login Name, Full Name and Password fields must be filled before saving", vbInformation, "Error Saving..."
    Exit Function
End If

' Check for last Admin. If so, then disable save if Admin Access is off
If flds.State <> adStateClosed Then flds.Close
flds.Source = "Select * From Usernames where Admin = True"
flds.Open

If flds.RecordCount = 1 And chkUserSecurity.Value = 0 Then
    MsgBox "You cannot remove Admin Access to the only remaining Admin User.", vbCritical, "Admin Access"
    flds.Close
    Exit Function
End If
flds.Close

Dim xtemp As Boolean
xtemp = False
If chkUserSecurity.Value = 1 Then xtemp = True

'update user
dbcn.BeginTrans
dbcn.Execute "UPDATE Usernames SET LoginName = '" & Trim$(txtLoginName.Text) & _
    "', FullName = '" & Trim$(txtFullName.Text) & _
    "', Pwd = '" & Trim$(txtPassWord.Text) & _
    "', Admin = " & xtemp & " WHERE ID = " & IDnum
dbcn.CommitTrans
    
iDirty = False
Call popList
Exit Function

SaveChangesError:
    MsgBox "Error while trying to save username info. Contact technical support.", vbCritical, "Error"
    Exit Function

End Function

Private Sub chkUserSecurity_Click()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub chkUserSecurity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "States whether the user has Administration privledges."

End Sub

Private Sub Form_Load()

'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Form Coordinates...
Me.Width = QuickRef.MediumMenuWidth
Me.Height = QuickRef.MediumMenuHeight

rec.ActiveConnection = dbcn
' change this  by closing and changing then reopening
rec.CursorLocation = adUseClient
'persisted in memory
rec.CursorType = adOpenStatic
rec.LockType = adLockBatchOptimistic

flds.ActiveConnection = dbcn
' change this  by closing and changing then reopening
flds.CursorLocation = adUseClient
'persisted in memory
flds.CursorType = adOpenStatic
flds.LockType = adLockBatchOptimistic

' open record set
Call popList

iDirty = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = val(ReadINI(Me.Name, "Left"))
Me.Top = val(ReadINI(Me.Name, "Top"))

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

'Prompt to save first...
If iDirty Then
    If MsgBox("Changes were not saved. Do you still want to exit anyway?", vbYesNo + vbQuestion, "Save Changes...") = vbNo Then
        Cancel = True
        Exit Sub
    End If
End If

'Save INI Settings...
Call SaveINISettings

If rec.State <> adStateClosed Then rec.Close
Set rec = Nothing
    
If flds.State <> adStateClosed Then flds.Close
Set flds = Nothing

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

End Sub

Private Sub imgDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgOKPicture(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgDelete.Picture = imgOKPicture(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub

Private Sub imgHelp_Click()

lblHelp_Click

End Sub
Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgOKPicture(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgHelp.Picture = imgOKPicture(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub
Private Sub imgNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgOKPicture(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgNew.Picture = imgOKPicture(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub
Private Sub imgExit_Click()

lblExit_Click

End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgOKPicture(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgExit.Picture = imgOKPicture(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub
Private Sub imgSave_Click()

lblSave_Click

End Sub

Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgOKPicture(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgSave.Picture = imgOKPicture(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub lblDelete_Click()

'Delete account...
Call DeleteAccount

End Sub
Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgOKPicture(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Deletes this user from the system."

End Sub
Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgDelete.Picture = imgOKPicture(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Exit this screen."

End Sub

Private Sub lblHelp_Click()

Help.HelpCallingForm = Me.Name

frmHelper.Show
frmHelper.ZOrder

End Sub
Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgOKPicture(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Shows the Help Window."

End Sub
Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgHelp.Picture = imgOKPicture(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub
Private Sub lblNew_Click()

On Local Error GoTo lblNew_ClickError

If UserSecurity = False Then
    MsgBox "You do not have authorization to create new accounts. Contact your System Administrator", vbInformation, "Authorization..."
    Exit Sub
End If

Dim sInput As String
Dim sInput2 As String
Dim sInput3 As String

'Enter a new user name...
sInput2 = Trim$(InputBox$("Enter the user's FULL NAME for this new account.", "New Account..."))
If sInput2 = "" Then Exit Sub

'Enter a new user name...
sInput = Trim$(InputBox$("Enter the user's LOGIN NAME for this new account.", "New Account..."))
If sInput = "" Then Exit Sub

'Enter a new user name...
sInput3 = Trim$(InputBox$("Enter the user's PASSWORD for this new account.", "New Account..."))
If sInput3 = "" Then Exit Sub

'Clear out all controls...
Call ClearAllControls

txtFullName = sInput2
txtLoginName = sInput
txtPassWord = sInput3
'txtLoginName = sInput2

dbcn.BeginTrans
dbcn.Execute "INSERT INTO Usernames ( LoginName , FullName, Pwd) " & _
    "Values ('" & sInput & "', '" & sInput2 & "', '" & sInput3 & "')"
dbcn.CommitTrans
    
iDirty = False
Call popList
Exit Sub

lblNew_ClickError:
    'DB.Close
    'Call WriteToErrorLog(Me.Name, "lblNew_ClickError", Error$, Err, True)
    MsgBox "Error occured while trying to create a new user account. Contact Technical Support"
    Exit Sub
    Resume Next

End Sub

Private Sub lblNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to create a new user."

End Sub
Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Saves any changes you have made to this user."

End Sub
Private Sub lblSystemAccounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgOKPicture(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgNew.Picture = imgOKPicture(0).Picture
lblNew.ForeColor = lButtonForeColor

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
    imgExit.Picture = imgOKPicture(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgExit.Picture = imgOKPicture(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub
Private Sub lblSave_Click()

Call SaveChanges

End Sub
Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgOKPicture(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgSave.Picture = imgOKPicture(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub
Private Sub lstUsers_Click()

On Local Error Resume Next

Dim iLocalIPHasChanged As Boolean


'Clear out all of the fields...
Call ClearAllControls

    Dim IDnum As String
    Dim IDTemp As Integer

    IDTemp = InStr(1, lstUsers.Text, "-", vbTextCompare)
    IDnum = Trim$(Left$(lstUsers.Text, IDTemp - 1))

    
    
    If IDnum <> "" Then
    'MsgBox ("active record selected... filling text fields")
    If flds.State <> adStateClosed Then flds.Close
    flds.Source = "Select * From Usernames where ID = " & IDnum
    flds.Open
    
    Call ClearAllControls
    dontWatchText = True
    
    If Not IsNull(flds.Fields(1).Value) Then
        txtLoginName.Text = flds.Fields(1).Value
    End If
    If Not IsNull(flds.Fields(2).Value) Then
        txtFullName.Text = flds.Fields(2).Value
    End If
    If Not IsNull(flds.Fields(3).Value) Then
        txtPassWord.Text = flds.Fields(3).Value
    End If
    If Not IsNull(flds.Fields(4).Value) Then
        txtLastLogin.Text = FormatDateTime(flds.Fields(4).Value, vbLongDate)
    End If
    'MsgBox "Check box is " & chkUserSecurity & vbCrLf & "Field value is " & flds.Fields(5).Value
    If flds.Fields(5).Value = True Then chkUserSecurity = 1
     
    dontWatchText = False
    iDirty = (iLocalIPHasChanged = True)
    End If

End Sub

Private Sub lstUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Listing of all users currently set up in the system."

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

'Users Listbox...
lstUsers.Enabled = iDirty = False

'New...
If imgNew.Enabled = False And iDirty = False Then
    imgNew.Enabled = True
    lblNew.Enabled = True
ElseIf imgNew.Enabled = True And iDirty = True Then
    imgNew.Enabled = False
    lblNew.Enabled = False
End If

'Save...
If imgSave.Enabled = False And iDirty = True Then
    imgSave.Enabled = True
    lblSave.Enabled = True
ElseIf imgSave.Enabled = True And iDirty = False Then
    imgSave.Enabled = False
    lblSave.Enabled = False
End If

'Delete...
If imgDelete.Enabled = True And lstUsers.List(lstUsers.ListIndex) = "Administrator" And iDirty = False Then
    imgDelete.Enabled = False
    lblDelete.Enabled = False
ElseIf imgDelete.Enabled = False And lstUsers.List(lstUsers.ListIndex) <> "Administrator" And iDirty = False Then
    imgDelete.Enabled = True
    lblDelete.Enabled = True
End If

End Sub

Private Sub txtFullName_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtFullName_GotFocus()

txtFullName.SelStart = 0
txtFullName.SelLength = Len(txtFullName)

End Sub

Private Sub txtFullName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Type in this users full name here."

End Sub

Private Sub txtLastLogin_Change()

If dontWatchText = False Then iDirty = True
End Sub

Private Sub txtLastLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "The last time the selected user accessed the database."

End Sub

Private Sub txtLoginName_Change()

If dontWatchText = False Then iDirty = True

End Sub
Private Sub txtLoginName_GotFocus()

txtLoginName.SelStart = 0
txtLoginName.SelLength = Len(txtLoginName)

End Sub

Private Sub txtLoginName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Type in the users login name here."

End Sub
Private Sub txtPassWord_Change()

If dontWatchText = False Then iDirty = True

End Sub

Private Sub txtPassWord_GotFocus()

txtPassWord.SelStart = 0
txtPassWord.SelLength = Len(txtPassWord)

End Sub

Private Sub txtPassWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Type in a password for this user to log in with."

End Sub

Private Sub popList()
    
    lstUsers.Clear
    
    Dim t As Integer
    Dim listRow As String
    t = 0
    If rec.State <> adStateClosed Then rec.Close
    If (UserSecurity = True) Then
        rec.Source = "Select * From Usernames"
        chkUserSecurity.Visible = True
        lblUserSecurity.Visible = True
    Else
        rec.Source = "Select * From Usernames where ID = " & CurrentUser
        chkUserSecurity.Visible = False
        lblUserSecurity.Visible = False
    End If
    rec.Open
    rec.MoveFirst
       
    While Not rec.EOF
        'add to listRow string and poplulate listbox
        listRow = rec.Fields(0).Value & " - " & rec.Fields(2).Value
        lstUsers.AddItem listRow, t
        
        t = t + 1
        rec.MoveNext
    Wend
    rec.Close
    lstUsers.ListIndex = 0
   
End Sub


