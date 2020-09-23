VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRememberMyLoginName 
      Height          =   195
      Left            =   990
      TabIndex        =   4
      Top             =   1380
      Width           =   195
   End
   Begin VB.TextBox txtPassWord 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   990
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "jkdd"
      Top             =   900
      Width           =   2175
   End
   Begin VB.TextBox txtLoginName 
      Height          =   285
      Left            =   990
      TabIndex        =   0
      Top             =   450
      Width           =   2175
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   2970
      TabIndex        =   7
      Tag             =   "ButtonLabel"
      Top             =   1875
      Width           =   285
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   1815
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      Top             =   1875
      Width           =   405
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   1
      Left            =   4440
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   810
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   4440
      Picture         =   "frmLogin.frx":1A22
      Top             =   390
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblRememberMyLoginName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remember my login name"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   1260
      TabIndex        =   5
      Tag             =   "Label"
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Tag             =   "Label"
      Top             =   930
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Tag             =   "Label"
      Top             =   480
      Width           =   405
   End
   Begin VB.Image imgCancel 
      Height          =   405
      Left            =   2550
      Picture         =   "frmLogin.frx":3444
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   1125
   End
   Begin VB.Image imgLogin 
      Height          =   405
      Left            =   1440
      Picture         =   "frmLogin.frx":4E66
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   1125
   End
   Begin VB.Image imgPanel 
      Height          =   2295
      Index           =   0
      Left            =   0
      Picture         =   "frmLogin.frx":6888
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3825
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset


Function LogUserIn() As Boolean

On Local Error GoTo LogUserInError

Dim lgn As String
lgn = Trim$(txtLoginName.Text)
Dim pwd As String
pwd = Trim$(txtPassWord)

rs.Source = "Select * From Usernames Where LoginName = '" & _
    lgn & "' AND Pwd = '" & pwd & "'"
rs.Open

'No info found for contact ???...
If rs.RecordCount = 0 Then
    If rs.State <> adStateClosed Then rs.Close
    MsgBox "Your login name or password is incorrect. If you can not log in, contact your system administrator and have them set up an account for you on the system.", vbInformation, "Login..."
    txtPassWord.SelStart = 0
    txtPassWord.SelLength = Len(txtPassWord)
    Exit Function
End If

'Account Name...
If Not IsNull(rs.Fields(1).Value) Then
    Login.LoginName = rs.Fields(1).Value
End If

'Fullname...
If Not IsNull(rs.Fields(2).Value) Then
    Login.FullName = rs.Fields(2).Value
End If

'Security Access...
UserSecurity = rs.Fields(5).Value
CurrentUser = rs.Fields(0).Value

'Login Date and Time...
'Login.LoginDateTime = Format(Now, "H:MM AMPM")

Dim Dat As String
Dat = Now
dbcn.BeginTrans
dbcn.Execute "UPDATE Usernames SET LoginDateTime = '" & Now & "' Where ID = " & rs.Fields(0).Value
dbcn.CommitTrans

'Main Menu Panel...
mdiMainMenu.lblLoginTime.Caption = Format$(Now, "H:MM AMPM")
mdiMainMenu.lblUser.Caption = Login.FullName

'Close the recordset...
If rs.State <> adStateClosed Then rs.Close
'rs.Close
'DB.Close

LogUserIn = True
Unload Me
Exit Function



LogUserInError:
    If rs.State <> adStateClosed Then rs.Close
    MsgBox ("Error in Login ... Exit and Contact technical support")
    'Call WriteToErrorLog(Me.Name, "LogUserIn", Error, Err, True)
    Exit Function
    Resume Next

End Function
Private Sub Form_Activate()

On Local Error Resume Next

'Set focus to password if loginname is already in the login textbox...
If txtLoginName <> "" Then
    txtPassWord.SetFocus
End If

End Sub
Private Sub Form_Load()

On Local Error Resume Next

'Load INI Settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Form Coordinates...
Me.Width = 3825
Me.Height = 2295

rs.ActiveConnection = dbcn
' change this  by closing and changing then reopening
rs.CursorLocation = adUseClient
'persisted in memory
rs.CursorType = adOpenStatic
rs.LockType = adLockBatchOptimistic

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = val(ReadINI(Me.Name, "Left"))
Me.Top = val(ReadINI(Me.Name, "Top"))

'Remember my login name...
chkRememberMyLoginName.Value = val(ReadINI(Me.Name, "RememberMyLoginName"))
If chkRememberMyLoginName.Value = 1 Then
    txtLoginName = ReadINI(Me.Name, "MyLoginName")
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

'Save this form's settings...
Call SaveINISettings

'Call appTerminate

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

'Remember my login name...
Call WriteINI(Me.Name, "RememberMyLoginName", chkRememberMyLoginName.Value)
Call WriteINI(Me.Name, "MyLoginName", txtLoginName)

End Sub

Private Sub imgCancel_Click()

lblCancel_Click

End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgLogin.Picture = imgButton(1).Picture
    lblLogin.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgCancel.Picture = imgButton(1).Picture
    lblCancel.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgCancel.Picture = imgButton(0).Picture
lblCancel.ForeColor = lButtonForeColor

End Sub
Private Sub imgLogin_Click()

lblLogin_Click

End Sub
Private Sub imgLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgLogin.Picture = imgButton(0).Picture
lblLogin.ForeColor = lButtonForeColor

End Sub

Private Sub imgPanel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = ""

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub lblCancel_Click()

    Unload mdiMainMenu
    Unload Me
    Call appTerminate
End Sub
Private Sub lblCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgCancel.Picture = imgButton(1).Picture
    lblCancel.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to cancel the log in process."

End Sub
Private Sub lblCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgCancel.Picture = imgButton(0).Picture
lblCancel.ForeColor = lButtonForeColor

End Sub

Private Sub lblLogin_Click()

'Log user in...
Call LogUserIn

End Sub
Private Sub lblLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgLogin.Picture = imgButton(1).Picture
    lblLogin.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to log into the system."

End Sub
Private Sub lblLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgLogin.Picture = imgButton(0).Picture
lblLogin.ForeColor = lButtonForeColor

End Sub
Private Sub lblRememberMyLoginName_Click()

'Swap values...
If chkRememberMyLoginName.Value = 1 Then
    chkRememberMyLoginName.Value = 0
ElseIf chkRememberMyLoginName.Value = 0 Then
    chkRememberMyLoginName.Value = 1
End If

End Sub

Private Sub lblRememberMyLoginName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here if you want the login screen to remember your login name."

End Sub
Private Sub txtLoginName_GotFocus()

txtLoginName.SelStart = 0
txtLoginName.SelLength = Len(txtLoginName)

End Sub
Private Sub txtLoginName_KeyDown(KeyCode As Integer, Shift As Integer)

On Local Error Resume Next

'Down key...
If KeyCode = vbKeyDown Then
    KeyCode = 0
    txtPassWord.SetFocus
End If

End Sub
Private Sub txtLoginName_KeyPress(KeyAscii As Integer)

On Local Error Resume Next

'Enter key...
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    txtPassWord.SetFocus
End If

End Sub

Private Sub txtLoginName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Enter your login name here."

End Sub

Private Sub txtPassWord_GotFocus()

txtPassWord.SelStart = 0
txtPassWord.SelLength = Len(txtPassWord)

End Sub
Private Sub txtPassWord_KeyDown(KeyCode As Integer, Shift As Integer)

On Local Error Resume Next

'Up key...
If KeyCode = vbKeyUp Then
    KeyCode = 0
    txtLoginName.SetFocus
End If

End Sub
Private Sub txtPassWord_KeyPress(KeyAscii As Integer)

On Local Error Resume Next

'Enter key...
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    'Log user in...
    If LogUserIn() = True Then
        Unload Me
        Exit Sub
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub txtPassWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Enter your password here."

End Sub
