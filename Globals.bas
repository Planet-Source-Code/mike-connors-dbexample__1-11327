Attribute VB_Name = "Globals"
Option Explicit

'database connection
Public dbcn As New ADODB.Connection

'Misc...
Global NotesTitle As String
Global DB As Database
Global rs As Recordset
Global SQL As String
Global Help As tHelp
Global Login As tLogin
Global QuickRef As tQuickRef
Global iActiveStatus As String
Global sString As String
Global CurrentUser As Integer
Global UserSecurity As Boolean


'Color variables...
Global lButtonForeColor As Long
Global lLabelForeColor As Long
Global lTextBoxBackColor As Long
Global lTextBoxForeColor As Long
Global lListBoxBackColor As Long

'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'INI File Functions...
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Always on top...
Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

'Quick Reference Array...
Type tQuickRef
    CallingForm As String
    Full_Name As String
    ID As Long
    DBFileName As String
    DBPassWord As String
    DBTimeOut As Long
    INIFileName As String
    LargeMenuHeight As Long
    LargeMenuWidth As Long
    MediumMenuHeight As Long
    MediumMenuWidth As Long
    NotesHaveChanged As Boolean
    PassNotes As Boolean
    'ReLoggingIn As Boolean
    UpdateColors As Boolean
    'UpdateInternetSites As Boolean
    UpdateNotes As Boolean
End Type

'Login...
Type tLogin
    FullName As String
    LoginDateTime As String
    LoginName As String
End Type

'Technical Support Type...
Type tHelp
    TechnicalSupportCompany As String
    TechnicalSupportPhone As String
    HelpText As String
    HelpIsLoaded As Boolean             'Tells all forms that the help form is loaded...
    HelpIsAligned As Boolean            'Tells all forms that the help form is loaded...
    HelpCallingForm As String           'Tells the helper form what form loaded it...
End Type

Public Sub AlwaysOnTop(Who As Form, iPosition As Boolean)

Dim lFlag As Long

'On top or not on top...
If iPosition Then
    lFlag = -1
Else
    lFlag = -2
End If

'Call the API to make the form on or not on top...
Call SetWindowPos(Who.hwnd, lFlag, Who.Left / Screen.TwipsPerPixelX, Who.Top / Screen.TwipsPerPixelY, Who.Width / Screen.TwipsPerPixelX, Who.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW)

End Sub


Sub AlignHelpToForm()

On Local Error Resume Next

Dim X As Long

'Align the helper form to the form it is associated with...
If Help.HelpIsAligned = True Then
    For X = 0 To Forms.Count - 1
        If Forms(X).Name = Help.HelpCallingForm Then
            frmHelper.Left = Forms(X).Left + Forms(X).Width + 20
            frmHelper.Top = Forms(X).Top
            Exit For
        End If
    Next X
End If

'Make sure the helper form is visible...
If frmHelper.Left >= mdiMainMenu.Width - 600 Then
    frmHelper.Left = mdiMainMenu.Width - frmHelper.Width - 180
    frmHelper.ZOrder
End If

End Sub

Sub CloseAllOpenWindows()

On Local Error Resume Next

Dim X As Long

'Close all currently open forms...
For X = 1 To Forms.Count - 1
    If Forms(X).Name <> "mdiMainMenu" Then
        Unload Forms(X)
    End If
Next X

End Sub
Sub ArrangeIcons(iArrangeType As Integer)

mdiMainMenu.Arrange iArrangeType

End Sub
Function DeleteStudent(sStudent As String) As Boolean

On Local Error GoTo DeleteStudentError

Dim tempFCP As String
If frmExampledb.txtCurrent_Past.Text = "P" Then
    If MsgBox("PERMANENTLY delete " & UCase$(sStudent) & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete...") = vbNo Then
        Exit Function
    End If
    dbcn.BeginTrans
    dbcn.Execute "DELETE * FROM data WHERE ID = " & QuickRef.ID
    dbcn.CommitTrans
Else

    If frmExampledb.txtCurrent_Past.Text = "C" Then
        'Confirm...
        If MsgBox("Are you sure you want to delete the student " & UCase$(sStudent) & " ?" & Chr(10) & Chr(13) & _
            "Please note that once deleted the student will be moved to the PAST STUDENTS List", vbYesNo + vbQuestion + vbDefaultButton2, "Delete...") = vbNo Then
            Exit Function
        End If
        tempFCP = "P"
    Else
        'Confirm...
        If MsgBox("Are you sure you want to delete the student " & UCase$(sStudent) & " ?" & Chr(10) & Chr(13) & _
            "Please note that once deleted the student will be moved to the CURRENT STUDENTS List", vbYesNo + vbQuestion + vbDefaultButton2, "Delete...") = vbNo Then
            Exit Function
        End If
        tempFCP = "C"
    End If
    dbcn.BeginTrans
    dbcn.Execute "UPDATE data SET Current_Past = '" & tempFCP & "' WHERE ID = " & QuickRef.ID
    dbcn.CommitTrans
End If


DeleteStudent = True
Exit Function

DeleteStudentError:

    MsgBox "Error while trying to Delete. Call for Technical support"
    Call appTerminate
    Exit Function

End Function

Sub LoadProgramColors()

On Local Error Resume Next

lLabelForeColor = val(ReadINI("Colors", "lLabelForeColor"))
lButtonForeColor = val(ReadINI("Colors", "lButtonForeColor"))
lTextBoxBackColor = val(ReadINI("Colors", "lTextBoxBackColor"))
lTextBoxForeColor = val(ReadINI("Colors", "lTextBoxForeColor"))
lListBoxBackColor = val(ReadINI("Colors", "lListBoxBackColor"))

End Sub

Sub SetColors(Who As Form)

On Local Error Resume Next

Dim X As Long

'Set label and button label fore colors...
For X = 0 To Who.Controls.Count - 1
    If InStr(LCase$(Who.Controls(X).Tag), "nocolorchange") = 0 Then
        'Label Colors...
        If Who.Controls(X).Tag = "Label" Then
            Who.Controls(X).ForeColor = lLabelForeColor
        'Button Label Colors...
        ElseIf Who.Controls(X).Tag = "ButtonLabel" Then
            Who.Controls(X).ForeColor = lButtonForeColor
        'Textbox ForeGround and BackGround Colors...
        ElseIf TypeOf Who.Controls(X) Is TextBox Then
            Who.Controls(X).ForeColor = lTextBoxForeColor
            Who.Controls(X).BackColor = lTextBoxBackColor
        'List and combo box BackGround Colors...
        ElseIf TypeOf Who.Controls(X) Is ListBox Or TypeOf Who.Controls(X) Is ComboBox Then
            Who.Controls(X).BackColor = lListBoxBackColor
            Who.Controls(X).ForeColor = lTextBoxForeColor
        'datalist backgraound colors...
        ElseIf TypeOf Who.Controls(X) Is DataList Then
            Who.Controls(X).BackColor = lListBoxBackColor
            Who.Controls(X).ForeColor = lTextBoxForeColor
        End If
    End If
Next X

End Sub

Public Sub DragForm(frm As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

'Align the help window to the form that loaded it...
If Help.HelpIsLoaded And Help.HelpIsAligned Then
    Call AlignHelpToForm
End If

End Sub
Sub Main()

On Local Error Resume Next

'DB Password...
QuickRef.DBPassWord = "admin"

'Window Heights and Widths...
QuickRef.MediumMenuHeight = 4320
QuickRef.MediumMenuWidth = 7020
QuickRef.LargeMenuHeight = 5705
QuickRef.LargeMenuWidth = 9665

'INI Filename...
If Dir$(App.Path & "\Example.Ini") <> "" Then
    QuickRef.INIFileName = App.Path & "\Example.Ini"
Else
    MsgBox "Can't find the 'Example.Ini' file. This file is mandatory for this program to run correctly. If you can find this file elsewhere on your computer, safely copy it to " & UCase$(App.Path) & "." & _
    Chr(13) & Chr(10) & Chr(13) & Chr(10) & "If this does not work, you will have to reinstall the program.", vbCritical, "File not found..."
    End
End If

'Database Filename...
If Trim(ReadINI("Database", "DatabaseLocation")) <> "" Then
    QuickRef.DBFileName = ReadINI("Database", "DatabaseLocation") & "\Example.Mdb"
Else
    QuickRef.DBFileName = App.Path & "\Example.Mdb"
End If
If Dir$(QuickRef.DBFileName) = "" Then
    Load frmFindDB
    frmFindDB.Show vbModal
    QuickRef.DBFileName = QuickRef.DBFileName & "\Example.Mdb"
End If
QuickRef.DBTimeOut = 500

'Load color settings...
Call LoadProgramColors

dbcn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & QuickRef.DBFileName & ";Jet OLEDB:Database Password=admin"
'This cursor is on the client machine instead of the server
dbcn.CursorLocation = adUseClient
dbcn.Open

'Main Menu...
mdiMainMenu.Show
    
End Sub
Function WriteINI(sSection As String, sKeyName As String, sNewString As String) As Boolean

On Local Error Resume Next

Call WritePrivateProfileString(sSection, sKeyName, sNewString, QuickRef.INIFileName)

WriteINI = (Err = 0)

End Function
Function ReadINI(sSection As String, sKeyName As String) As String

On Local Error Resume Next

Dim sRet As String

sRet = String(255, Chr(0))

ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), QuickRef.INIFileName))

End Function
Sub WriteToErrorLog(sFormName As String, sRoutineName As String, sError As String, iErrorNumber As Integer, iDisplayMsgBox As Boolean)

On Local Error Resume Next

Dim FileFree As Integer

FileFree = FreeFile
Open App.Path & "\ErrorLog.Txt" For Append As #FileFree
    Print #FileFree, sFormName, sRoutineName, sError, iErrorNumber
Close #FileFree

'Display the error that occured...
If iDisplayMsgBox = True Then
    MsgBox "The following error has occured in your program: " & vbCrLf & vbCrLf & sError & vbCrLf & vbCrLf & "Error Number: " & iErrorNumber, vbInformation, "Error..."
End If

End Sub


Public Sub appTerminate()
    If dbcn.State <> adStateClosed Then dbcn.Close

    'closes Connection
    Set dbcn = Nothing
    End

End Sub
