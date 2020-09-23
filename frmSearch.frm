VERSION 5.00
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Color Settings"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&F"
      Height          =   345
      Left            =   7170
      TabIndex        =   13
      Top             =   1020
      Width           =   1125
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   270
      TabIndex        =   12
      Top             =   2430
      Width           =   6465
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   1260
      Width           =   1725
   End
   Begin VB.ComboBox SearchFlds2 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1275
      Width           =   2445
   End
   Begin VB.ComboBox SearchFlds 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   765
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   765
      Width           =   1725
   End
   Begin VB.Label lblFind 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Search"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   3045
      TabIndex        =   4
      Tag             =   "ButtonLabel"
      Top             =   1905
      Width           =   915
   End
   Begin VB.Image imgSearch 
      Height          =   375
      Left            =   2700
      Picture         =   "frmSearch.frx":62BC2
      Stretch         =   -1  'True
      Top             =   1815
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "And :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1020
      TabIndex        =   11
      Top             =   1305
      Width           =   465
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3525
      TabIndex        =   10
      Top             =   1305
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3525
      TabIndex        =   6
      Top             =   795
      Width           =   195
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Tag             =   "ButtonLabel"
      Top             =   3810
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   2
      Top             =   795
      Width           =   885
   End
   Begin VB.Label lblCategories 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Student Records"
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
      TabIndex        =   1
      Top             =   60
      Width           =   2325
   End
   Begin VB.Image imgOKPicture 
      Height          =   375
      Index           =   1
      Left            =   7170
      Picture         =   "frmSearch.frx":645E4
      Top             =   510
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgOKPicture 
      Height          =   360
      Index           =   0
      Left            =   7170
      Picture         =   "frmSearch.frx":65CCE
      Stretch         =   -1  'True
      Top             =   120
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
      Left            =   5730
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      Top             =   3810
      Width           =   285
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   5250
      Picture         =   "frmSearch.frx":676F0
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image imgOptionsPanel 
      Height          =   300
      Index           =   1
      Left            =   540
      Picture         =   "frmSearch.frx":69112
      Stretch         =   -1  'True
      Top             =   765
      Width           =   960
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   480
      Picture         =   "frmSearch.frx":70BD4
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image imgOptionsPanel 
      Height          =   300
      Index           =   0
      Left            =   3390
      Picture         =   "frmSearch.frx":725F6
      Stretch         =   -1  'True
      Top             =   765
      Width           =   480
   End
   Begin VB.Image imgOptionsPanel 
      Height          =   300
      Index           =   2
      Left            =   540
      Picture         =   "frmSearch.frx":7A0B8
      Stretch         =   -1  'True
      Top             =   1275
      Width           =   960
   End
   Begin VB.Image imgOptionsPanel 
      Height          =   300
      Index           =   3
      Left            =   3390
      Picture         =   "frmSearch.frx":81B7A
      Stretch         =   -1  'True
      Top             =   1275
      Width           =   480
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean
Dim srch As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim lst As New ADODB.Recordset
Dim srchTemp As String
Dim srchTemp2 As String
Dim LstTotal As Integer

Private Sub Command1_Click()

    lblExit_Click
End Sub

Private Sub Form_Load()
 
'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Form Coordinates...
Me.Width = QuickRef.MediumMenuWidth
Me.Height = QuickRef.MediumMenuHeight

Dim i As Integer
Dim fld As Field
    
iDirty = False

rec.ActiveConnection = dbcn
' change this  by closing and changing then reopening
rec.CursorLocation = adUseClient
'persisted in memory
rec.CursorType = adOpenStatic
rec.LockType = adLockBatchOptimistic

lst.ActiveConnection = dbcn
' change this  by closing and changing then reopening
lst.CursorLocation = adUseClient
'persisted in memory
lst.CursorType = adOpenStatic
lst.LockType = adLockBatchOptimistic

srch.ActiveConnection = dbcn
' change this  by closing and changing then reopening
srch.CursorLocation = adUseClient
'persisted in memory
srch.CursorType = adOpenStatic
srch.LockType = adLockBatchOptimistic

' open record set
Call popCombos
Call popList


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

If rec.State <> adStateClosed Then rec.Close
Set rec = Nothing
    
If lst.State <> adStateClosed Then lst.Close
Set lst = Nothing

If srch.State <> adStateClosed Then srch.Close
Set srch = Nothing

'Save INI Settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

End Sub

Private Sub imgFind_Click()

lblFind_Click
    
End Sub

Private Sub imgFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub imgFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSearch.Picture = imgOKPicture(0).Picture
lblFind.ForeColor = lButtonForeColor
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


Private Sub imgSearch_Click()

lblFind_Click

End Sub

Private Sub imgSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    imgSearch.Picture = imgOKPicture(1).Picture
    lblFind.ForeColor = QBColor(0)
End If
End Sub

Private Sub imgSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSearch.Picture = imgOKPicture(0).Picture
lblFind.ForeColor = lButtonForeColor

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here close the Search Window."

End Sub

Private Sub lblFind_Click()

On Error GoTo SearchError

If Text1.Text <> "" Then
    sString = "(" & SearchFlds.Text & " Like '" & Text1.Text & "%" & "')"
    StudentFull.Title = "Full Student Information - Search For '" & Text1.Text & "' in Field '" & SearchFlds.Text & "'"
    
    If Text2.Text <> "" Then
        sString = sString & " AND (" & SearchFlds2.Text & " Like '" & Text2.Text & "%" & "') "
        StudentFull.Title = StudentFull.Title & " & '" & Text2.Text & "' in Field '" & SearchFlds2.Text & "'"
    End If

    If rec.State <> adStateClosed Then rec.Close
    rec.Source = "SELECT * FROM data WHERE " & sString & " Order by ID"
    rec.Open
        
    If rec.RecordCount = 0 Then
        MsgBox "No Records found ...", vbInformation, "Search"
        rec.Close
        Exit Sub
    Else
        Set StudentFull.DataSource = rec
        StudentFull.Orientation = rptOrientLandscape
        'ChngPrinterOrientationLandscape Me
        StudentFull.Title = StudentFull.Title & Space(25) & "Total Records Found: " & rec.RecordCount
        StudentFull.Show vbModal
    End If
    rec.Close
    
    'Saving search data
    If lst.State <> adStateClosed Then lst.Close
    lst.Source = "Select * from Search Where ID = 1"
    lst.Open
    
    Dim xtemp As String
    xtemp = lst.Fields(3).Value
    If xtemp = "Empty" Then xtemp = ""
    
    If (lst.Fields(1).Value <> Text1.Text Or val(lst.Fields(2).Value) <> SearchFlds.ListIndex Or xtemp <> Text2.Text Or val(lst.Fields(4).Value) <> SearchFlds2.ListIndex) Then
        Dim num As Integer
        For num = 5 To 2 Step -1
            If lst.State <> adStateClosed Then lst.Close
            lst.Source = "Select * from Search Where ID = " & (num - 1)
            lst.Open
            dbcn.BeginTrans
            'If (lst.Fields(1).Value <> Trim$(Text1.Text) Or val(lst.Fields(2).Value) <> Str(SearchFlds.ListIndex)) Then
            dbcn.Execute "Update Search SET text1 = '" & lst.Fields(1).Value & "', Field1 = '" & val(lst.Fields(2).Value) & _
                "', text2 = '" & lst.Fields(3).Value & "', Field2 = '" & val(lst.Fields(4).Value) & "' Where ID = " & num
            'End If
            dbcn.CommitTrans
        Next num
        dbcn.BeginTrans
        dbcn.Execute "Update Search SET Text1 = '" & Text1.Text & "', Field1 = '" & SearchFlds.ListIndex & "' Where ID = 1"
        dbcn.CommitTrans
        If Trim$(Text2.Text) = "" Then
            dbcn.BeginTrans
            dbcn.Execute "Update Search SET Text2 = 'Empty', Field2 = '" & SearchFlds2.ListIndex & "' Where ID = 1"
            dbcn.CommitTrans
        Else
            dbcn.BeginTrans
            dbcn.Execute "Update Search SET Text2 = '" & Text2.Text & "', Field2 = '" & SearchFlds2.ListIndex & "' Where ID = 1"
            dbcn.CommitTrans
        End If
        lst.Close
        Call popList
    End If
    
Else
    MsgBox ("Please specify a search.")
End If
Exit Sub

SearchError:
    MsgBox "Error occured while searching ... Contact Technical Support", vbCritical, "Error"

End Sub

Private Sub lblFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    imgSearch.Picture = imgOKPicture(1).Picture
    lblFind.ForeColor = QBColor(0)
End If
End Sub

Private Sub lblFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Click here to start the search."

End Sub

Private Sub lblFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSearch.Picture = imgOKPicture(0).Picture
lblFind.ForeColor = lButtonForeColor
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



Private Sub lblCategories_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

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

Private Sub List1_Click()
    Dim spce As String
    If rec.State <> adStateClosed Then rec.Close
    rec.Source = "Select * From Search where ID = " & Left$(List1.Text, 1)
    rec.Open
   
    'Set DataList1.RowSource = rec
    Text1.Text = rec.Fields(1).Value
    SearchFlds.ListIndex = val(rec.Fields(2).Value)
    
    If (rec.Fields(3).Value <> "Empty") Then
        Text2.Text = rec.Fields(3).Value
        SearchFlds2.ListIndex = val(rec.Fields(4).Value)
    Else
        Text2.Text = ""
    End If
    
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Help.HelpText = "Click on any of these to recall the your last 5 searches."
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Type a word or first part of a word. Click the Field and then click START SEARCH." & _
vbCrLf & vbCrLf & "If you wanted to search for a Student with 'mic' in thier Full Name, You can use a '%'. ie. '%mic'."

End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "This text area narrows the search. It is applied to the first text search criteria above."

End Sub

Private Sub popList()
 
    List1.Clear
    Dim t As Integer
    Dim listRow As String
    For t = 1 To 5
        If rec.State <> adStateClosed Then rec.Close
        rec.Source = "Select * From Search WHERE ID = " & t
        rec.Open
    
        'check for data in search record then add to listRow string
        If (rec.Fields(1).Value <> "Empty") Then  'val(rec.Fields(2).Value
        'MsgBox SearchFlds.ListIndex
            listRow = rec.Fields(0).Value & "  -  Find:  " & UCase(rec.Fields(1).Value) & _
                "  -in-  " & UCase(SearchFlds.List(val(rec.Fields(2).Value)))
            If (rec.Fields(3).Value <> "Empty") Then
                listRow = listRow & "  -and-    " & UCase(rec.Fields(3).Value) & _
                   "  -in-  " & UCase(SearchFlds.List(val(rec.Fields(4).Value)))
            End If
            List1.AddItem listRow, (t - 1)
        End If
    Next t
   
End Sub

Private Sub popCombos()

    'clear the list
    SearchFlds.Clear

    'clear the list for combo2
    SearchFlds2.Clear
    
    srch.Source = "Select * from data"
    srch.Open
    
    Dim t As Integer
    For t = 0 To 28
        SearchFlds.AddItem srch.Fields(t).Name
        SearchFlds2.AddItem srch.Fields(t).Name
    Next t
    srch.Close
    
    SearchFlds.ListIndex = 2
 
    SearchFlds2.ListIndex = 3
    
End Sub


