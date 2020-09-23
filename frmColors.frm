VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColors 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Color Settings"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11370
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmColors.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnColors 
      Caption         =   "&C"
      Height          =   240
      Left            =   7425
      TabIndex        =   20
      Top             =   2160
      Width           =   915
   End
   Begin VB.CheckBox chkAutoApply 
      Height          =   195
      Left            =   3540
      TabIndex        =   17
      Top             =   3390
      Value           =   1  'Checked
      Width           =   195
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   9570
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CheckBox chkShowHideSolidColors 
      Height          =   195
      Left            =   270
      TabIndex        =   13
      Top             =   3850
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.PictureBox picColorPalette 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3000
      Left            =   270
      MouseIcon       =   "frmColors.frx":62BC2
      Picture         =   "frmColors.frx":62D14
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   3
      Top             =   630
      Width           =   3000
      Begin VB.PictureBox picColorsSquare 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1605
         Left            =   420
         MouseIcon       =   "frmColors.frx":7EF86
         Picture         =   "frmColors.frx":7F0D8
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   12
         Top             =   660
         Width           =   2115
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9030
      Top             =   180
   End
   Begin VB.PictureBox picOptionsPanel 
      BackColor       =   &H00404040&
      Height          =   2325
      Left            =   3540
      ScaleHeight     =   2265
      ScaleWidth      =   3075
      TabIndex        =   4
      ToolTipText     =   "Click on an item to select it"
      Top             =   930
      Width           =   3135
      Begin VB.ComboBox lstListBoxBackGroundColor 
         Height          =   315
         ItemData        =   "frmColors.frx":896DE
         Left            =   570
         List            =   "frmColors.frx":896F1
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Listbox Background Color"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtTextBoxForeGroundColor 
         Height          =   285
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Textbox Foreground Color"
         Top             =   1380
         Width           =   2295
      End
      Begin VB.TextBox txtTextBoxBackGroundColor 
         Height          =   285
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Textbox Background Color"
         Top             =   960
         Width           =   2295
      End
      Begin VB.Image imgLights 
         Height          =   210
         Index           =   4
         Left            =   120
         Picture         =   "frmColors.frx":89777
         Top             =   1830
         Width           =   150
      End
      Begin VB.Image imgLights 
         Height          =   210
         Index           =   3
         Left            =   120
         Picture         =   "frmColors.frx":89979
         Top             =   1410
         Width           =   150
      End
      Begin VB.Image imgLights 
         Height          =   210
         Index           =   2
         Left            =   120
         Picture         =   "frmColors.frx":89B7B
         Top             =   990
         Width           =   150
      End
      Begin VB.Image imgLights 
         Height          =   210
         Index           =   0
         Left            =   120
         Picture         =   "frmColors.frx":89D7D
         Top             =   180
         Width           =   150
      End
      Begin VB.Image imgLights 
         Height          =   210
         Index           =   1
         Left            =   120
         Picture         =   "frmColors.frx":89F7F
         Top             =   570
         Width           =   150
      End
      Begin VB.Label lblLabelColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label Color"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Tag             =   "Label"
         Top             =   165
         Width           =   795
      End
      Begin VB.Label lblButtonLabelColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Button Text Color"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Tag             =   "ButtonLabel"
         Top             =   570
         Width           =   1230
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   570
         Picture         =   "frmColors.frx":8A181
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1515
      End
      Begin VB.Image imgOptionsPanel 
         Height          =   2250
         Index           =   0
         Left            =   0
         Picture         =   "frmColors.frx":8BBA3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6090
      TabIndex        =   19
      Tag             =   "ButtonLabel"
      Top             =   3510
      Width           =   495
   End
   Begin VB.Label lblAutoApply 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Apply"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   3810
      TabIndex        =   18
      Top             =   3390
      Width           =   765
   End
   Begin VB.Image imgRedLight 
      Height          =   210
      Index           =   1
      Left            =   8970
      Picture         =   "frmColors.frx":93665
      Top             =   1530
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgRedLight 
      Height          =   210
      Index           =   0
      Left            =   9210
      Picture         =   "frmColors.frx":93867
      Top             =   1530
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblWindowsColors 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Colors..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   2580
      TabIndex        =   15
      Tag             =   "ButtonLabel"
      Top             =   3870
      Width           =   1275
   End
   Begin VB.Label lblHideSolidColors 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Show / Hide Solid Colors"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   540
      TabIndex        =   14
      Top             =   3850
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
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
      Left            =   3630
      TabIndex        =   11
      Top             =   675
      Width           =   435
   End
   Begin VB.Label lblControlSelected 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label Color"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4110
      TabIndex        =   10
      Top             =   675
      Width           =   795
   End
   Begin VB.Label lblApply 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6150
      TabIndex        =   9
      Tag             =   "ButtonLabel"
      Top             =   3870
      Width           =   390
   End
   Begin VB.Label lblCategories 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Settings"
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
      Left            =   330
      TabIndex        =   2
      Top             =   60
      Width           =   1365
   End
   Begin VB.Image imgOKPicture 
      Height          =   375
      Index           =   1
      Left            =   8970
      Picture         =   "frmColors.frx":93A69
      Top             =   1080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgOKPicture 
      Height          =   360
      Index           =   0
      Left            =   8970
      Picture         =   "frmColors.frx":95153
      Stretch         =   -1  'True
      Top             =   690
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4350
      TabIndex        =   1
      Tag             =   "ButtonLabel"
      Top             =   3870
      Width           =   225
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5160
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      Top             =   3870
      Width           =   495
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   4920
      Picture         =   "frmColors.frx":96B75
      Stretch         =   -1  'True
      Top             =   3780
      Width           =   945
   End
   Begin VB.Image imgSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3990
      Picture         =   "frmColors.frx":98597
      Stretch         =   -1  'True
      Top             =   3780
      Width           =   945
   End
   Begin VB.Image imgApply 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5850
      Picture         =   "frmColors.frx":99FB9
      Stretch         =   -1  'True
      Top             =   3780
      Width           =   945
   End
   Begin VB.Image imgWindowsColors 
      Height          =   375
      Left            =   2430
      Picture         =   "frmColors.frx":9B9DB
      Stretch         =   -1  'True
      Top             =   3780
      Width           =   1575
   End
   Begin VB.Image imgOptionsPanel 
      Height          =   300
      Index           =   1
      Left            =   3540
      Picture         =   "frmColors.frx":9D3FD
      Stretch         =   -1  'True
      Top             =   630
      Width           =   3150
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   5850
      Picture         =   "frmColors.frx":A4EBF
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   945
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean
Sub SaveChanges()

On Local Error Resume Next

'Save the color settings to the ini file...
Call WriteINI("Colors", "lLabelForeColor", lblLabelColor.ForeColor)
Call WriteINI("Colors", "lButtonForeColor", lblButtonLabelColor.ForeColor)
Call WriteINI("Colors", "lTextBoxBackColor", txtTextBoxBackGroundColor.BackColor)
Call WriteINI("Colors", "lTextBoxForeColor", txtTextBoxForeGroundColor.ForeColor)
Call WriteINI("Colors", "lListBoxBackColor", lstListBoxBackGroundColor.BackColor)

iDirty = False
QuickRef.UpdateColors = True
lblExit.Caption = "Close"

End Sub

Private Sub btnColors_Click()
    lblSave_Click
End Sub

Private Sub chkAutoApply_Click()

'Auto Apply...
If chkAutoApply.Value = 1 Then
    lblApply_Click
End If

End Sub
Private Sub chkShowHideSolidColors_Click()

picColorsSquare.Visible = chkShowHideSolidColors.Value = 1

End Sub
Private Sub Form_Load()

'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Form Coordinates...
Me.Width = QuickRef.MediumMenuWidth
Me.Height = QuickRef.MediumMenuHeight

'Display the control selected...
lblControlSelected.Caption = lblLabelColor.Caption

iDirty = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = val(ReadINI(Me.Name, "Left"))
Me.Top = val(ReadINI(Me.Name, "Top"))

'Auto Apply...
chkAutoApply = val(ReadINI(Me.Name, "AutoApply"))

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = ""

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

Dim x As Long

'Prompt to save first...
If iDirty Then
    x = MsgBox("Save changes before exiting?", vbYesNoCancel + vbQuestion, "Save Changes...")
    Select Case x
        Case vbYes
            Call SaveChanges
        Case vbCancel
            Cancel = True
            Exit Sub
    End Select
End If

'Save INI Settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

'Auto Apply...
Call WriteINI(Me.Name, "AutoApply", chkAutoApply.Value)

End Sub
Private Sub Image1_Click()

'Display the control selected...
lblControlSelected.Caption = lblButtonLabelColor.Caption

End Sub

Private Sub imgApply_Click()

Call SaveChanges

End Sub

Private Sub imgApply_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgApply.Picture = imgOKPicture(1).Picture
    lblApply.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgApply_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgApply.Picture = imgOKPicture(0).Picture
lblApply.ForeColor = lButtonForeColor

End Sub

Private Sub imgHelp_Click()

lblHelp_Click

End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgOKPicture(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgHelp.Picture = imgOKPicture(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub
Private Sub imgLights_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

On Local Error Resume Next

'Change colors...
If Button = vbLeftButton Then
    Select Case Index
        Case 0
            lblLabelColor_Click
        Case 1
            lblButtonLabelColor_Click
        Case 2
            txtTextBoxBackGroundColor_Click
        Case 3
            txtTextBoxForeGroundColor_Click
        Case 4
            lstListBoxBackGroundColor_GotFocus
    End Select
End If

End Sub

Private Sub imgOptionsPanel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub imgWindowsColors_Click()

lblWindowsColors_Click

End Sub
Private Sub imgWindowsColors_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgWindowsColors.Picture = imgOKPicture(1).Picture
    lblWindowsColors.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgWindowsColors_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgWindowsColors.Picture = imgOKPicture(0).Picture
lblWindowsColors.ForeColor = lButtonForeColor

End Sub

Private Sub lblApply_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to apply your new color settings."

End Sub
Private Sub lblAutoApply_Click()

'Show / Hide Solid Colors...
If chkAutoApply.Value = 1 Then
    chkAutoApply.Value = 0
Else
    chkAutoApply.Value = 1
End If

End Sub

Private Sub lblAutoApply_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Check this box to automatically apply the color changes to all open windows immediately."

End Sub
Private Sub lblButtonLabelColor_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to change the button colors."

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to cancel any color settings you have made."

End Sub
Private Sub lblHelp_Click()

Help.HelpCallingForm = Me.Name

frmHelper.Show
frmHelper.ZOrder

End Sub

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgOKPicture(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Shows the Help Window."

End Sub
Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgHelp.Picture = imgOKPicture(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub

Private Sub lblHideSolidColors_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to show and hide the square colors."

End Sub
Private Sub lblLabelColor_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to change the label colors."

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to apply the new color settings."

End Sub
Private Sub lblWindowsColors_Click()

On Local Error Resume Next

'Set the dialog boxes color to the color of the control that is selected...
If lblControlSelected.Caption = lblLabelColor.Caption Then
    Dialog.Color = lblLabelColor.ForeColor
ElseIf lblControlSelected.Caption = lblButtonLabelColor.Caption Then
    Dialog.Color = lblButtonLabelColor.ForeColor
ElseIf lblControlSelected.Caption = txtTextBoxBackGroundColor.Text Then
    Dialog.Color = txtTextBoxBackGroundColor.BackColor
ElseIf lblControlSelected.Caption = txtTextBoxForeGroundColor.Text Then
    Dialog.Color = txtTextBoxForeGroundColor.ForeColor
End If

'Show the color dialog box...
Dialog.Flags = cdlCCFullOpen Or cdlCCRGBInit
Dialog.ShowColor
If Err > 0 Then Exit Sub

'Set the color to the control currently selected...
If lblControlSelected.Caption = lblLabelColor.Caption Then
    lblLabelColor.ForeColor = Dialog.Color
ElseIf lblControlSelected.Caption = lblButtonLabelColor.Caption Then
    lblButtonLabelColor.ForeColor = Dialog.Color
ElseIf lblControlSelected.Caption = txtTextBoxBackGroundColor.Text Then
    txtTextBoxBackGroundColor.BackColor = Dialog.Color
    txtTextBoxForeGroundColor.BackColor = Dialog.Color
ElseIf lblControlSelected.Caption = txtTextBoxForeGroundColor.Text Then
    txtTextBoxForeGroundColor.ForeColor = Dialog.Color
    txtTextBoxBackGroundColor.ForeColor = Dialog.Color
End If

'Set the dirty flag to true...
iDirty = True

End Sub
Private Sub lblWindowsColors_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgWindowsColors.Picture = imgOKPicture(1).Picture
    lblWindowsColors.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblApply_Click()

Call SaveChanges

End Sub
Private Sub lblApply_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgApply.Picture = imgOKPicture(1).Picture
    lblApply.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblApply_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgApply.Picture = imgOKPicture(0).Picture
lblApply.ForeColor = lButtonForeColor

End Sub
Private Sub lblHideSolidColors_Click()

'Show / Hide Solid Colors...
If chkShowHideSolidColors.Value = 1 Then
    chkShowHideSolidColors.Value = 0
Else
    chkShowHideSolidColors.Value = 1
End If

End Sub

Private Sub lblWindowsColors_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to see the default colors in Windows."

End Sub
Private Sub lblWindowsColors_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgWindowsColors.Picture = imgOKPicture(0).Picture
lblWindowsColors.ForeColor = lButtonForeColor

End Sub

Private Sub lstListBoxBackGroundColor_Click()

Help.HelpText = "Click here to change the list box background colors."

End Sub
Private Sub lstListBoxBackGroundColor_GotFocus()

'Display the control selected...
lblControlSelected.Caption = lstListBoxBackGroundColor.Text

'Set the red lights picture...
imgLights(0).Picture = imgRedLight(0).Picture
imgLights(1).Picture = imgRedLight(0).Picture
imgLights(2).Picture = imgRedLight(0).Picture
imgLights(3).Picture = imgRedLight(0).Picture
imgLights(4).Picture = imgRedLight(1).Picture

End Sub
Private Sub picColorPalette_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

On Local Error Resume Next

'Change the label color...
If Button = vbLeftButton Then
    If lblControlSelected.Caption = lblLabelColor.Caption Then
        lblLabelColor.ForeColor = picColorPalette.Point(x, Y)
    ElseIf lblControlSelected.Caption = lblButtonLabelColor.Caption Then
        lblButtonLabelColor.ForeColor = picColorPalette.Point(x, Y)
    ElseIf lblControlSelected.Caption = txtTextBoxBackGroundColor.Text Then
        txtTextBoxBackGroundColor.BackColor = picColorPalette.Point(x, Y)
        txtTextBoxForeGroundColor.BackColor = picColorPalette.Point(x, Y)
    ElseIf lblControlSelected.Caption = txtTextBoxForeGroundColor.Text Then
        txtTextBoxForeGroundColor.ForeColor = picColorPalette.Point(x, Y)
        txtTextBoxBackGroundColor.ForeColor = picColorPalette.Point(x, Y)
        lstListBoxBackGroundColor.ForeColor = picColorPalette.Point(x, Y)
    ElseIf lblControlSelected.Caption = lstListBoxBackGroundColor.Text Then
        lstListBoxBackGroundColor.BackColor = picColorPalette.Point(x, Y)
    End If
    iDirty = True
End If

End Sub
Private Sub picColorPalette_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click on any color here to change the color for the selected item."

End Sub
Private Sub imgExit_Click()

lblExit_Click

End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgOKPicture(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgExit.Picture = imgOKPicture(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub
Private Sub imgSave_Click()

lblSave_Click

End Sub

Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgOKPicture(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgOKPicture(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub
Private Sub lblButtonLabelColor_Click()

'Display the control selected...
lblControlSelected.Caption = lblButtonLabelColor.Caption

'Set the red lights picture...
imgLights(0).Picture = imgRedLight(0).Picture
imgLights(1).Picture = imgRedLight(1).Picture
imgLights(2).Picture = imgRedLight(0).Picture
imgLights(3).Picture = imgRedLight(0).Picture
imgLights(4).Picture = imgRedLight(0).Picture

End Sub
Private Sub lblCategories_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

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
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgOKPicture(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgExit.Picture = imgOKPicture(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub
Private Sub lblLabelColor_Click()

'Display the control selected...
lblControlSelected.Caption = lblLabelColor.Caption

'Set the red lights picture...
imgLights(0).Picture = imgRedLight(1).Picture
imgLights(1).Picture = imgRedLight(0).Picture
imgLights(2).Picture = imgRedLight(0).Picture
imgLights(3).Picture = imgRedLight(0).Picture
imgLights(4).Picture = imgRedLight(0).Picture

End Sub
Private Sub lblSave_Click()

Call SaveChanges

'Unload the help window...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

Unload Me

End Sub
Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgOKPicture(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgOKPicture(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub picColorPalette_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Auto Apply...
If chkAutoApply.Value = 1 Then
    Call lblApply_Click
End If

End Sub
Private Sub picColorsSquare_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

On Local Error Resume Next

'Change the label color...
If Button = vbLeftButton Then
    If lblControlSelected.Caption = lblLabelColor.Caption Then
        lblLabelColor.ForeColor = picColorsSquare.Point(x, Y)
    ElseIf lblControlSelected.Caption = lblButtonLabelColor.Caption Then
        lblButtonLabelColor.ForeColor = picColorsSquare.Point(x, Y)
    ElseIf lblControlSelected.Caption = txtTextBoxBackGroundColor.Text Then
        txtTextBoxBackGroundColor.BackColor = picColorsSquare.Point(x, Y)
        txtTextBoxForeGroundColor.BackColor = picColorsSquare.Point(x, Y)
    ElseIf lblControlSelected.Caption = txtTextBoxForeGroundColor.Text Then
        txtTextBoxForeGroundColor.ForeColor = picColorsSquare.Point(x, Y)
        txtTextBoxBackGroundColor.ForeColor = picColorsSquare.Point(x, Y)
        lstListBoxBackGroundColor.ForeColor = picColorsSquare.Point(x, Y)
    ElseIf lblControlSelected.Caption = lstListBoxBackGroundColor.Text Then
        lstListBoxBackGroundColor.BackColor = picColorsSquare.Point(x, Y)
    End If
    iDirty = True
End If

End Sub
Private Sub picColorsSquare_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click on any color here to change the color for the selected item."

End Sub
Private Sub picColorsSquare_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Auto Apply...
If chkAutoApply.Value = 1 Then
    Call lblApply_Click
End If

End Sub
Private Sub picOptionsPanel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

'Save...
If imgSave.Enabled = False And iDirty = True Then
    imgSave.Enabled = True
    lblSave.Enabled = True
ElseIf imgSave.Enabled = True And iDirty = False Then
    imgSave.Enabled = False
    lblSave.Enabled = False
End If

'Apply...
If imgApply.Enabled = False And iDirty = True Then
    imgApply.Enabled = True
    lblApply.Enabled = True
ElseIf imgApply.Enabled = True And iDirty = False Then
    imgApply.Enabled = False
    lblApply.Enabled = False
End If

End Sub
Private Sub txtTextBoxBackGroundColor_Click()

'Display the control selected...
lblControlSelected.Caption = txtTextBoxBackGroundColor.Text

'Set the red lights picture...
imgLights(0).Picture = imgRedLight(0).Picture
imgLights(1).Picture = imgRedLight(0).Picture
imgLights(2).Picture = imgRedLight(1).Picture
imgLights(3).Picture = imgRedLight(0).Picture
imgLights(4).Picture = imgRedLight(0).Picture

End Sub
Private Sub txtTextBoxBackGroundColor_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to change the text box background colors."

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub txtTextBoxForeGroundColor_Click()

'Display the control selected...
lblControlSelected.Caption = txtTextBoxForeGroundColor.Text

'Set the red lights picture...
imgLights(0).Picture = imgRedLight(0).Picture
imgLights(1).Picture = imgRedLight(0).Picture
imgLights(2).Picture = imgRedLight(0).Picture
imgLights(3).Picture = imgRedLight(1).Picture
imgLights(4).Picture = imgRedLight(0).Picture

End Sub

Private Sub txtTextBoxForeGroundColor_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to change the text box foreground colors."

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
