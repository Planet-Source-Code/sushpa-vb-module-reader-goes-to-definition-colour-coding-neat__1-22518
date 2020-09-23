VERSION 5.00
Begin VB.Form frmOpts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Options"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chLine 
      Caption         =   "Sepa&rate using line"
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   1215
      Width           =   1680
   End
   Begin VB.CheckBox chMargin 
      Caption         =   "S&how Editing margin"
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   945
      Width           =   1770
   End
   Begin VB.PictureBox pTxCol 
      Height          =   285
      Left            =   180
      ScaleHeight     =   225
      ScaleWidth      =   1575
      TabIndex        =   4
      Top             =   630
      Width           =   1635
   End
   Begin VB.Frame FR 
      Caption         =   "User options"
      Height          =   125
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1950
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "Cl&ose"
      Default         =   -1  'True
      Height          =   375
      Left            =   1395
      TabIndex        =   0
      Top             =   1575
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TextBox margin colour:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   405
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User options"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   90
      Width           =   900
   End
End
Attribute VB_Name = "frmOpts"
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
'====================================

Private Sub chLine_Click()
SaveValue "MarginLine", CBool(chLine.Value)
fMainForm.Ln.Visible = CBool(chLine.Value)
End Sub

Private Sub chMargin_Click()
SaveValue "Margin", CBool(chMargin.Value)
End Sub

Private Sub cmOK_Click()
Unload Me
'done with it
End Sub

Private Sub Form_Load()
pTxCol.BackColor = ReadValue("MarginBack", &H8000000F) ' Menu bar colour is default
chMargin.Value = CBin(ReadValue("Margin", True))
chLine.Value = CBin(ReadValue("MarginLine", True))
End Sub

Private Sub pTxCol_Click()
On Error Resume Next
Dim PQRS As Long
'PQRS: holds colour
PQRS = BrowseColor(pTxCol)
If PQRS <> -1 Then
'user not cancelled
pTxCol.BackColor = PQRS
End If
SaveValue "MarginBack", pTxCol.BackColor
fMainForm.pM.BackColor = PQRS
End Sub
