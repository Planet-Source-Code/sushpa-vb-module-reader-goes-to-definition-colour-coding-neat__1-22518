VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Visual Basic Module Reader"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7230
   Icon            =   "vbread.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlFun 
      Left            =   6975
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":030A
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":0466
            Key             =   "function"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":071E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":09D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":0B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":0C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":0DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":0F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":10A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbread.frx":11FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6885
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "bas"
      DialogTitle     =   "Open module"
      Filter          =   "Visual Basic modules & classes (*.bas, *.cls)|*.bas; *.cls|Visual Basic forms and projects (*.frm, *.vbp)|*.frm; *.vbp"
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   635
      ButtonWidth     =   1429
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlFun"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New "
            Object.ToolTipText     =   "Create a new file"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load "
            Object.ToolTipText     =   "Open a file"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save "
            Object.ToolTipText     =   "Save this file"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageCombo imSubs 
         Height          =   330
         Left            =   4455
         TabIndex        =   4
         Top             =   0
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Text            =   "0 Function(s), 0 Sub(s)."
         ImageList       =   "imlFun"
      End
      Begin MSComctlLib.Toolbar TB2 
         Height          =   330
         Left            =   2565
         TabIndex        =   5
         Top             =   0
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlFun"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cut selected text"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy selected text"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Paste text from clipboard"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Undo last action"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Goto Definition"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      ScaleHeight     =   3825
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   375
      Width           =   7215
      Begin RichTextLib.RichTextBox RTF1 
         Height          =   3795
         Left            =   315
         TabIndex        =   2
         Top             =   0
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   6694
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"vbread.frx":135A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Ln 
         X1              =   300
         X2              =   300
         Y1              =   0
         Y2              =   3780
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4260
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7990
            Picture         =   "vbread.frx":141B
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New file"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open file"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveS 
         Caption         =   "&Save"
         Begin VB.Menu mnuFileSave 
            Caption         =   "&Save file"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuFileSaveAs 
            Caption         =   "S&ave as...   "
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "&File Info..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditText 
         Caption         =   "T&ext"
         Begin VB.Menu mnuEditCut 
            Caption         =   "C&ut"
            Shortcut        =   ^X
         End
         Begin VB.Menu mnuFileCopy 
            Caption         =   "&Copy"
            Shortcut        =   ^C
         End
         Begin VB.Menu mnuFilePaste 
            Caption         =   "&Paste"
            Shortcut        =   ^V
         End
         Begin VB.Menu mnuEditDelete 
            Caption         =   "D&elete       "
            Shortcut        =   {DEL}
         End
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditEntireThing 
         Caption         =   "Entirety"
         Begin VB.Menu mnuEdittSelect 
            Caption         =   "Se&lect"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuEditClear 
            Caption         =   "&Clear"
            Shortcut        =   +{DEL}
         End
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActions 
         Caption         =   "A&ctions"
         Begin VB.Menu mnuViewDefine 
            Caption         =   "D&efine "
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuLastPos 
            Caption         =   "Last position"
            Shortcut        =   ^P
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Options"
      Begin VB.Menu mnuViewMargin 
         Caption         =   "&Margin Bar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuViewSeparator 
         Caption         =   "&Separator"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
         Shortcut        =   ^{F2}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "H&elp"
      Begin VB.Menu mnnuHelpAbout 
         Caption         =   "A&bout"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================
'ModuleReader Application For Visual Basic classes
'====================================
'This application will read all VB modules and class
'modules and derive the names of subs in them.
'It will list all these items just like the VB IDE.
'====================================
'By Sushant Pandurangi (sushant@phreaker.net)
'====================================
'Visit http://sushant.iscool.net for more source,
'files, tutorials, and the massive VB6LIB with over
'100 functions for your daily programming needs
'API, strings, network, registry, INI, menus, ...
'====================================
Option Explicit
Dim pos As Long
Dim pos2 As Long
'====================================

Private Sub Form_Load()
pM.BackColor = ReadValue("MarginBack", &H8000000F)
Ln.Visible = ReadValue("MarginLine", True)
CD1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNOverwritePrompt
End Sub

Private Sub Form_Resize()
Arrange
End Sub

Private Sub imSubs_Click()
On Error Resume Next
'to hold the ID value
Dim pTYPE As String
pTYPE = ""
    'Get the type, based on the image.
    If imSubs.SelectedItem.Image = 11 Then
    pTYPE = "Sub "
    ElseIf imSubs.SelectedItem.Image = 12 Then
    pTYPE = "Function "
    End If
'find the ID as well as the 'sub' or 'function' keyword before it
StrFind imSubs.Text, pTYPE
RTF1.SetFocus
End Sub

Private Sub mnnuHelpAbout_Click()
If MsgBox("ModuleReader by Sushant Pandurangi, (c) 2000." & vbNewLine & "http://sushant.iscool.net | sushant@phreaker.net" & vbNewLine & vbNewLine & "Click OK to go to my visual basic site on the Net.", vbOKCancel + vbInformation, "About") = vbOK Then
ShellExecute Me.hwnd, "open", "http://sushant.iscool.net", "", "", 10
End If
End Sub

Private Sub mnuEditClear_Click()
RTF1.Text = ""
End Sub

Private Sub mnuEditCut_Click()
SendMessage RTF1.hwnd, WM_CUT, 0, 0&
End Sub

Private Sub mnuEditDelete_Click()
If RTF1.SelLength = 0 Then RTF1.SelLength = 1
RTF1.SelText = ""
End Sub

Private Sub mnuEdittSelect_Click()
RTF1.SelStart = 0
RTF1.SelLength = Len(RTF1.Text)
End Sub

Private Sub mnuEditUndo_Click()
SendMessage RTF1.hwnd, EM_UNDO, 0, 0&
End Sub

Private Sub mnuFileCopy_Click()
SendMessage RTF1.hwnd, WM_COPY, 0, 0&
End Sub

Private Sub mnuFileExit_Click()
Terminate
End Sub

Private Sub mnuFileInfo_Click()
If SB.Panels(1).Text = "" Then
If MsgBox("You must save this file before you" & vbNewLine & "can get information. Save it now?", vbYesNo + vbQuestion, "Save") = vbYes Then mnuFileSave_Click
Else
frmInfo.Show vbModal
End If
End Sub

Private Sub mnuFileNew_Click()
If MsgBox("Do you want to create a new file?" & vbNewLine & vbNewLine & "Any changes to this file may be lost" & vbNewLine & "if they have not been saved yet.", vbYesNoCancel + vbInformation, "Module") = vbYes Then RTF1.Text = "": SB.Panels(1).Text = ""
AddItems
End Sub

Private Sub mnuFileOpen_Click()
OpenFile
 ColorCode Keywords, vbBlue
 ColorComments

End Sub

Private Sub mnuFilePaste_Click()
SendMessage RTF1.hwnd, WM_PASTE, 0, 0&
End Sub

Private Sub mnuFileSave_Click()
On Error Resume Next
If Mid(SB.Panels(1).Text, 2, 1) <> ":" Then mnuFileSaveAs_Click: Exit Sub
Open SB.Panels(1).Text For Output As #1
Print #1, RTF1.Text
Close #1
End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo hell
CD1.DialogTitle = "Save file"
CD1.ShowSave
Open CD1.FileName For Output As #1
Print #1, RTF1.Text
Close #1
SB.Panels(1).Text = CD1.FileName
hell:
End Sub

Private Sub mnuLastPos_Click()
'TODO
RTF1.SelStart = LastItem
End Sub

Private Sub mnuViewDefine_Click()
If DefineWhat = "" Then Exit Sub
Define RTF1
End Sub

Private Sub mnuViewMargin_Click()
mnuViewMargin.Checked = Not mnuViewMargin.Checked
'Select or deselect the margin menu
If mnuViewMargin.Checked = False Then
RTF1.Left = 0
'no margin
Else
RTF1.Left = 270
'yes margin
End If
Arrange
'resize controls
SaveValue "Margin", mnuViewMargin.Checked
'write to INI
End Sub

Private Sub mnuView_Click()
'appropriate checking to the margin menu
mnuViewMargin.Checked = (RTF1.Left > 0)
'appropriate checking to the separator menu
mnuViewSeparator.Checked = Ln.Visible
End Sub

Private Sub mnuViewOptions_Click()
'show options dialog
frmOpts.Show vbModal
End Sub

Private Sub mnuViewSeparator_Click()
mnuViewSeparator.Checked = Not mnuViewSeparator.Checked
'check/uncheck
Ln.Visible = mnuViewSeparator.Checked
'Make LN visible
Arrange
'resize controls
SaveValue "MarginLine", Ln.Visible
'write to INI
End Sub

Private Sub RTF1_KeyPress(KeyAscii As Integer)
SB.Panels(3).Text = KeyAscii
If KeyAscii = 13 Then
ColorCode Keywords, vbBlue
ColorComments
RTF1.SetFocus
RTF1.SelStart = Len(RTF1.Text)
RTF1.SelColor = 0
End If
End Sub

Private Sub RTF1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuEdit
DefineWhat = RichWordOver(RTF1, x, y)
'Get the word to define
If DefineWhat = "" Or DefineWhat = "Sub " Then Exit Sub
TB2.Buttons(6).ToolTipText = "Define '" & DefineWhat & "'"
'put in the tooltiptext
mnuViewDefine.Caption = "De&fine '" & DefineWhat & "'"
'put in the menu caption
LastItem = RTF1.SelStart
End Sub

Private Sub RTF1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Now we will attempt to give the user a description of the sub
'If SB.Panels(1).Text = "" Then Exit Sub
'Since module isn't loaded
Dim TEMP As String, pos As Integer, TEMP2 As String
TEMP = RichWordOver(RTF1, x, y)
For pos = 1 To imSubs.ComboItems.Count
    If TEMP = imSubs.ComboItems.Item(pos).Text Then TEMP2 = GetDesc(TEMP)
Next pos
If Trim(TEMP2) <> "" Then
SB.Panels(1).Text = TEMP & ": " & TEMP2
RTF1.ToolTipText = TEMP & ": " & TEMP2
Else
SB.Panels(1).Text = CD1.FileName
RTF1.ToolTipText = vbNullChar
End If
End Sub

Private Sub RTF1_SelChange()
On Error Resume Next
SB.Panels(3).Text = Asc(RTF1.SelText)
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
mnuFileNew_Click
Case 2
mnuFileOpen_Click
Case 3
mnuFileSave_Click
End Select
End Sub

Private Sub TB2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
SendMessage RTF1.hwnd, WM_CUT, 0, 0&
Case 2
SendMessage RTF1.hwnd, WM_COPY, 0, 0&
Case 3
SendMessage RTF1.hwnd, WM_PASTE, 0, 0&
Case 4
SendMessage RTF1.hwnd, EM_UNDO, 0, 0&
Case 6
Define RTF1
End Select
End Sub

Function StrFind(vData As String, vType As String) As Long
On Error Resume Next
Dim pos As Long
pos = InStr(1, RTF1.Text, vType & vData)
If pos = 0 Then Exit Function
RTF1.SelStart = pos - 1 + Len(vType)
RTF1.SelLength = Len(vData)
RTF1.SetFocus
StrFind = pos
pos = 0
End Function
