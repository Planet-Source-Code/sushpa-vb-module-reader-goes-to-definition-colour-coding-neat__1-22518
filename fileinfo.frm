VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Information"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fileinfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2205
      Width           =   2850
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "&Copy Info"
      Height          =   420
      Left            =   1170
      TabIndex        =   9
      Top             =   2610
      Width           =   960
   End
   Begin VB.PictureBox pcMD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   135
      Picture         =   "fileinfo.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   2610
      Width           =   480
   End
   Begin VB.PictureBox pcCM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   135
      Picture         =   "fileinfo.frx":0614
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   2610
      Width           =   480
   End
   Begin VB.TextBox txName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1890
      Width           =   2850
   End
   Begin VB.TextBox txFileLen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1575
      Width           =   2175
   End
   Begin VB.TextBox txLines 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1260
      Width           =   2130
   End
   Begin VB.TextBox txCLSBAS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   2850
   End
   Begin VB.CommandButton cmClose 
      Caption         =   "C&lose"
      Default         =   -1  'True
      Height          =   420
      Left            =   2160
      TabIndex        =   0
      Top             =   2610
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2430
      Picture         =   "fileinfo.frx":091E
      Top             =   1305
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1575
      TabIndex        =   3
      Top             =   90
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "Visual Basic                    Module Reader 1.2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   2790
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmClose_Click()
Unload Me
End Sub

Private Sub cmCopy_Click()
Dim TEMP As String
TEMP = txCLSBAS.Text & vbNewLine & txFileLen.Text & vbNewLine & txLines.Text & vbNewLine & txName.Text & vbNewLine & txFile.Text
Clipboard.Clear
Clipboard.SetText TEMP
End Sub

Private Sub Form_Load()
'First get file type
Dim sTYPE As String
Select Case LCase(Right(fMainForm.SB.Panels(1).Text, 3))
Case "cls"
sTYPE = "Class module"
pcCM.ZOrder
Case "bas"
sTYPE = "Module"
pcMD.ZOrder
Case Else
sTYPE = "(unknown)"
End Select
'then get name
txName.Text = "Module name: " & ReadCustomVal("Attribute VB_Name", "(unknown)")
txLines.Text = "Number of lines: " & SendMessage(fMainForm.RTF1.hwnd, &HBA, 0, 0&)
txFileLen.Text = "Filesize (bytes): " & FileLen(fMainForm.SB.Panels(1).Text)
txCLSBAS.Text = "Module type: " & sTYPE
txFile.Text = fMainForm.SB.Panels(1).Text
End Sub
