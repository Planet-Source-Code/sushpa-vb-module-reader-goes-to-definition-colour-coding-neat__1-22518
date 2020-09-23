Attribute VB_Name = "Declarations"
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
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const Keywords = "Public ,Private ,Friend ,Dim ,Sub ,Function , WithEvents , Not , And , Or , Xor ,If , Then ,Do,Loop,Next ,For ,GoTo ,GoSub , As , Long, String, Integer, Variant, Object, Nothing,Else,End,Optional ,UBound,LBound,Mid,CBool,CByte,CStr,CInt,Const , Declare , Lib , Alias ,Set ,On , Error , Resume ,Option , Explicit, Compare ,Option Base 0,Option Base 1, Binary ,Open ,Input , Output ,Print ,Close , Shared , Append ,Read ,Write ,Line ,Type ,Enum ,Attribute ,Let ,Property ,Get ,False,True,ReDim ,ByVal ,End Enum,End Type,End If,Exit ,End Sub,End Function,End With,With ,End Select,Select Case ,Case , Is ,End Property, New "
Public Const EM_UNDO = &HC7
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302
Public SubCount As Integer
Public FunCount As Integer
Public DefineWhat As String
Public LastItem As Long
Public fMainForm As frmMain
Private Const EM_CHARFROMPOS& = &HD7
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'====================================
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Sub Terminate()
Dim pForm As Form
'To iterate over
    Unload fMainForm
    'The main form
        For Each pForm In Forms
            Unload pForm
            'Kill it, finish it...
            Set pForm = Nothing
        Next pForm
        'same thing
    '-------------------
    'VB's BEST COMMAND
            End
    'VB's BEST COMMAND
    '-------------------
End Sub

Function OpenFile() As Boolean
Dim pF As String
'FileName to open
    On Error GoTo hell
    'CommonDialog will put up an error on cancelling
        fMainForm.CD1.DialogTitle = "Open file"
        fMainForm.CD1.ShowOpen
        'Show open dialog
        pF = fMainForm.CD1.FileName
        'Load the file
      fMainForm.SB.Panels(2).Text = "Loading...  "
        If fMainForm.CD1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNOverwritePrompt + 1 Then
        '6147 I derived as I have put in a few flags on load, they total up
        'to 6146. The ReadOnly is 1, therefore it becomes 6146 + 1= 6147.
            fMainForm.RTF1.Locked = True
            MsgBox "You have chosen to open as read only." & vbNewLine & _
            "You will not be able to edit this module.", vbInformation, "Readonly"
            fMainForm.SB.Panels(3).Text = "READONLY"
        Else
        'If not readonly then don't lock textbox
            fMainForm.RTF1.Locked = False
            fMainForm.SB.Panels(3).Text = ""
        End If
    'Add the functions and subs
      fMainForm.RTF1.LoadFile pF, rtfText
      DoEvents
        With fMainForm.RTF1
        .Visible = False
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = 0
        .SelLength = 0
        .Visible = True
        End With
        AddItems
fMainForm.SB.Panels(1).Text = fMainForm.CD1.FileName
OpenFile = True
Exit Function
hell:
OpenFile = False
'Error handling
End Function

Public Sub AddItems()
Dim pos As Long, TEMP As String, BRPOS As Long, FINAL As String
'POS    = position of word "function"
'TEMP  = TEMPorary string to hold ID
'BRPOS = position of ( symbol from POS
'FINAL  = FINAL string to add if qualifies
    fMainForm.imSubs.ComboItems.Clear: SubCount = 0: FunCount = 0
    'Clear all things, count should be made 0 now at this time.
Do 'Get stuck in a loop
    pos = InStr(pos + 1, fMainForm.RTF1.Text, "Function ")
    'Where is the word "function"
        If pos = 0 Then Exit Do
        'no functions, word not found
            BRPOS = InStr(pos, fMainForm.RTF1.Text, "(")
            'where bracket symbol is after function name
        If BRPOS = 0 Then GoTo looper1
        'no bracket; syntax error
            TEMP = Mid$(fMainForm.RTF1.Text, pos + 9, BRPOS - pos - 9)
            'POS+9 as we dont want to add the word 'function' and
            If Mid(fMainForm.RTF1.Text, pos - 5, 5) = "Exit " Then GoTo looper1
            'BRPOS-POS will give us the identifier name, but we have
            'added 9 here and hence should substract it there.
            'the word 'function ' is 9 characters long, and 'sub ' is 4.
        If InStr(1, TEMP, "Lib") = 0 Then FINAL = TEMP
        'no LIB keyword found, so it is not a declaration
    If FINAL <> "" Then fMainForm.imSubs.ComboItems.Add , LCase(FINAL), FINAL, 12: FunCount = FunCount + 1
    'Unless LIB is found, FINAL will be TEMP, else FINAL will be "". FunCount is # of functions; increase.
looper1:
Loop
pos = 0: TEMP = "": BRPOS = 0: FINAL = ""
'OK, now functions have been dealt with.
Do
    'Now the same procedure is followed here, replacing the word
    '"function" with "sub". It looks for subs, adds them, that's it.
    pos = InStr(pos + 1, fMainForm.RTF1.Text, "Sub ")
    'Position of word "Sub"
        If pos = 0 Then Exit Do
        '"Sub" word doesn't exist
            BRPOS = InStr(pos, fMainForm.RTF1.Text, "(")
            'Position of ( symbol
        If BRPOS = 0 Then GoTo looper2
        '"(" doesn't exist
        TEMP = Mid$(fMainForm.RTF1.Text, pos + 4, BRPOS - pos - 4)
            'adjust 4 characters (length of "Sub ")
            If Mid(fMainForm.RTF1.Text, pos - 5, 5) = "Exit " Then GoTo looper2
        If InStr(1, TEMP, "Lib") = 0 Then FINAL = TEMP
        'No LIB keyword is there
    If FINAL <> "" Then fMainForm.imSubs.ComboItems.Add , LCase(FINAL), FINAL, 11: SubCount = SubCount + 1
    'FINAL is not empty; add it to the list and increase the number of subs
looper2:
Loop
pos = 0: TEMP = "": BRPOS = 0: FINAL = ""
'Clear up the variables
fMainForm.imSubs.Text = SubCount & " Sub(s), " & FunCount & " Function(s)."
'Show how many functions, and how many subs
fMainForm.SB.Panels(2).Text = " Loaded " & ReadCustomVal("Attribute VB_Name", "module") & ".  "
ListProcs fMainForm.imSubs
'Regulate the list
End Sub

Sub Arrange()
'Resize controls appropriately
On Error Resume Next
    fMainForm.pM.Width = fMainForm.ScaleWidth
    fMainForm.pM.Height = fMainForm.ScaleHeight - fMainForm.SB.Height - fMainForm.pM.Top
    fMainForm.RTF1.Width = fMainForm.pM.ScaleWidth - fMainForm.RTF1.Left
    fMainForm.RTF1.Height = fMainForm.pM.ScaleHeight
    fMainForm.Ln.Y2 = fMainForm.RTF1.Height
    fMainForm.imSubs.Left = fMainForm.TB.Width - fMainForm.imSubs.Width - ((fMainForm.ScaleWidth - fMainForm.pM.Width) / 2)
End Sub

Sub Main()
'Start up sub
    Set fMainForm = New frmMain
    Load fMainForm
    fMainForm.Show
End Sub

Function BrowseColor(pObject As Object) As Long
    'Get the colour from a user
    On Error GoTo hell
        With fMainForm.CD1
            'CommonDialog object
            .Color = pObject.BackColor
            .ShowColor 'Show it
            BrowseColor = .Color
        End With
    Exit Function
hell:
    'return -1 so we can find
    'out that user cancelled.
    BrowseColor = -1
End Function

Function ReadValue(Name As String, Optional Default As String, Optional Section As String = "Settings")
    'Function to read values from default filename
    ReadValue = ReadINI(Section, Name, App.Path & "\settings.ini", Default)
End Function

Sub SaveValue(Name As String, Value As String, Optional Section As String = "Settings")
    'Function to save values to default filename
    SaveINI Section, Name, Value, App.Path & "\settings.ini"
End Sub

Public Function ReadINI(Section As String, Key As String, FileName As String, Optional Default As String)
    'Read from INI file
    Dim sReturn As String
    sReturn = String(255, Chr(0))
    ReadINI = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), FileName))
End Function

Public Sub SaveINI(Section As String, Key As String, Value As String, FileName As String)
    'Write to INI file
    WritePrivateProfileString Section, Key, Value, FileName
End Sub

Function CBin(Expression As Boolean) As Integer
'Converts boolean to 0 or 1 in binary
If Expression = True Then CBin = 1 Else CBin = 0
End Function

Sub Define(TextBox As Object)
On Error Resume Next
If DefineWhat = "" Then GoTo ERRORS
'This function should attempt to get the selected
'word, then go to its definition if the same exists
'in the curent file which is being viewed or edited.
DefineWhat = "Sub " & DefineWhat
'add sub to the beginning
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it function instead of sub
DefineWhat = "Function " & Right(DefineWhat, Len(DefineWhat) - Len("Sub "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it public
DefineWhat = "Public " & Right(DefineWhat, Len(DefineWhat) - Len("Function "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it private
DefineWhat = "Private " & Right(DefineWhat, Len(DefineWhat) - Len("Public "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it dim
DefineWhat = "Dim " & Right(DefineWhat, Len(DefineWhat) - Len("Private "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it withevents
DefineWhat = "WithEvents " & Right(DefineWhat, Len(DefineWhat) - Len("Dim "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it static
DefineWhat = "Static " & Right(DefineWhat, Len(DefineWhat) - Len("WithEvents "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it property
DefineWhat = "Property " & Right(DefineWhat, Len(DefineWhat) - Len("Static "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it byval
DefineWhat = "ByVal " & Right(DefineWhat, Len(DefineWhat) - Len("Property "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it ,
DefineWhat = ", " & Right(DefineWhat, Len(DefineWhat) - Len("ByVal "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it const
DefineWhat = "Const " & Right(DefineWhat, Len(DefineWhat) - Len(", "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it optional
DefineWhat = "Optional " & Right(DefineWhat, Len(DefineWhat) - Len("Const "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it =
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Optional ")) & " = "
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF
'make it (
DefineWhat = "(" & Left(DefineWhat, Len(DefineWhat) - Len(" = "))
If fMainForm.RTF1.Find(DefineWhat) > 0 Then GoTo SF Else GoTo ERRORS
'if its found then good else tell user that its not there
SF:
'got it
fMainForm.RTF1.SetFocus
Exit Sub
ERRORS:
'remove function or sub or whatever
If Left(DefineWhat, 9) = "Function " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Function "))
ElseIf Left(DefineWhat, 4) = "Sub " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Sub "))
ElseIf Left(DefineWhat, 4) = "Dim " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Dim "))
ElseIf Left(DefineWhat, 11) = "WithEvents " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("WithEvents "))
ElseIf Left(DefineWhat, 7) = "Static " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Static "))
ElseIf Left(DefineWhat, 9) = "Property " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Property "))
ElseIf Left(DefineWhat, 7) = "Public " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Public "))
ElseIf Left(DefineWhat, 8) = "Private " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Private "))
ElseIf Left(DefineWhat, 1) = "(" Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("("))
ElseIf Left(DefineWhat, 6) = "ByVal " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("ByVal "))
ElseIf Left(DefineWhat, 2) = ", " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len(", "))
ElseIf Left(DefineWhat, 6) = "Const " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Const "))
ElseIf Left(DefineWhat, 9) = "Optional " Then
DefineWhat = Right(DefineWhat, Len(DefineWhat) - Len("Optional "))
ElseIf Right(DefineWhat, 3) = " = " Then
DefineWhat = Left(DefineWhat, Len(DefineWhat) - Len(" = "))
End If
'tell the user
MsgBox "The identifier '" & DefineWhat & "' is unrecognized." & vbNewLine & "Make sure such declaration exists.", vbExclamation, "Not found"
End Sub

Function ListProcs(List As ImageCombo)
Dim F As Integer, TEMP As Integer, Strin As String
'For loop
For F = 1 To List.ComboItems.Count
'variable
Strin = List.ComboItems.Item(F).Text
Do 'and keep doing
    TEMP = InStr(1, List.ComboItems.Item(F).Text, "Private ")
    'position of private
    If TEMP > 0 Then
    Mid$(Strin, TEMP, 8) = Space(8)
    'replace with space so we can trim later,8 is len of 'private '
    Else
    Exit Do
    'it isnt there, proceed
    End If
        TEMP = InStr(1, List.ComboItems.Item(F).Text, "Public ")
        'position of public
        If TEMP > 0 Then
        Mid$(Strin, TEMP, 7) = Space(7)
        'replace with spaces, 7 is len of 'Public '
        Else
        Exit Do
        'not found, proceed
        End If
            TEMP = InStr(1, List.ComboItems.Item(F).Text, "Sub ")
            'find the word 'sub '
            If TEMP > 0 Then
            Mid$(Strin, TEMP, 4) = Space(4)
            'replace spaces, again 4 is len of 'Sub '
            Else
            Exit Do
            'not found then proceed
            End If
        TEMP = InStr(1, List.ComboItems.Item(F).Text, "Function ")
        'word 'Function '
        If TEMP > 0 Then
        Mid$(Strin, TEMP, 9) = Space(9)
        'equal no. of spaces
        Else
        Exit Do
        'not found
        End If
    TEMP = InStr(1, List.ComboItems.Item(F).Text, "Friend ")
    'generally unused friend keyword
    If TEMP > 0 Then
    Mid$(Strin, TEMP, 7) = Space(7)
    'replace spaces
    Else
    Exit Do
    'get out
    End If
Loop
'keep doing until some Exit Do works
If InStr(1, List.ComboItems(F).Text, Chr(13)) > 0 Then List.ComboItems.Remove F: If List.ComboItems.Item(F).Image = 11 Then SubCount = SubCount - 1 Else FunCount = FunCount - 1
'contains carriage return; remove this junk item
If InStr(1, List.ComboItems(F).Text, Chr(10)) > 0 Then List.ComboItems.Remove F: If List.ComboItems.Item(F).Image = 11 Then SubCount = SubCount - 1 Else FunCount = FunCount - 1
'contains linefeed; remove this junk item
If InStr(1, List.ComboItems(F).Text, vbCrLf) > 0 Then List.ComboItems.Remove F: If List.ComboItems.Item(F).Image = 11 Then SubCount = SubCount - 1 Else FunCount = FunCount - 1
'contains CrLf; remove this junk item
If InStr(1, List.ComboItems(F).Text, ",") > 0 Then List.ComboItems.Remove F: If List.ComboItems.Item(F).Image = 11 Then SubCount = SubCount - 1 Else FunCount = FunCount - 1
'contains comma; remove this junk item
If InStr(1, List.ComboItems(F).Text, "=") > 0 Then List.ComboItems.Remove F: If List.ComboItems.Item(F).Image = 11 Then SubCount = SubCount - 1 Else FunCount = FunCount - 1
'contains comma; remove this junk item
If InStr(1, List.ComboItems(F).Text, vbNewLine) > 0 Then List.ComboItems.Remove F: If List.ComboItems.Item(F).Image = 11 Then SubCount = SubCount - 1 Else FunCount = FunCount - 1
'contains comma; remove this junk item
'STRIN is the thing that has been spruced up
List.ComboItems.Item(F).Text = Strin
'trim it, we have replaced certain words with spaces
List.ComboItems.Item(F).Text = Trim(List.ComboItems.Item(F).Text)
Next F
List.Text = SubCount & " Sub(s), " & FunCount & " Function(s)."
End Function

Public Function RichWordOver(rch As RichTextBox, x As Single, y As Single) As String
Dim pt As POINTAPI
Dim pos As Long
Dim start_pos As Long
Dim end_pos As Long
Dim ch As String
Dim txt As String
Dim txtlen As Long
    ' Convert the position to pixels.
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY
    ' Get the character number
    pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function
    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores, $ and %.
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or ch = "_" Or ch = "$" Or ch = "%") Then Exit For
    Next start_pos
    start_pos = start_pos + 1
    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "0" And ch <= "9") Or _
            (ch >= "a" And ch <= "z") Or _
            (ch >= "A" And ch <= "Z") Or _
            ch = "_" _
        ) Then Exit For
    Next end_pos
    end_pos = end_pos - 1
    If start_pos <= end_pos Then _
        RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
End Function

Function ReadCustomVal(sValueName As String, Optional sDefault As String)
On Error Resume Next
Dim pos, pos1, pos2
With fMainForm.RTF1
pos = InStr(1, .Text, sValueName) - Len(sValueName)
If pos = 0 Then ReadCustomVal = sDefault: Exit Function
pos = pos + Len(sValueName)
'where is svaluename, also we need to go ahead of it
pos1 = InStr(pos, .Text, "=") + 3
If pos = 0 Then ReadCustomVal = sDefault: Exit Function
'where is = symbol; 3 as there is a " " and quote after it, skip them
pos2 = InStr(pos1, .Text, vbNewLine)
If pos = 0 Then ReadCustomVal = sDefault: Exit Function
'where is end of line
ReadCustomVal = Mid(.Text, pos1, pos2 - pos1 - 1)
'we start from pos1, go upto pos2-pos1 and there's a chr(34)
'at the end which should be removed
End With
End Function

Function GetDesc(Text As String) As String
Dim TEMP As String
TEMP = ReadCustomVal("Attribute " & Text & ".VB_Description", "")
GetDesc = TEMP
End Function

Sub ColorCode(StringList As String, Colour As Long)
fMainForm.RTF1.Visible = False
Screen.MousePointer = 11
fMainForm.SB.Panels(2).Text = " Loading..."
Dim StrSplitted() As String, pos As Long, i As Long
SplitStr StringList, StrSplitted, ","
For i = 0 To UBound(StrSplitted)
Do
pos = InStr(pos + 1, fMainForm.RTF1.Text, StrSplitted(i))
If pos = 0 Then Exit Do
fMainForm.RTF1.SelStart = pos - 1
fMainForm.RTF1.SelLength = Len(StrSplitted(i))
fMainForm.RTF1.SelColor = Colour
Loop
Next i
fMainForm.RTF1.SelLength = 0
fMainForm.RTF1.Visible = True
fMainForm.RTF1.SelColor = 0
Screen.MousePointer = 0
fMainForm.SB.Panels(2).Text = " Loaded " & ReadCustomVal("Attribute VB_Name", "(module)") & ".  "
End Sub

Private Sub SplitStr(strMessage As String, StrLines() As String, Character As String)
'FUNCTION TO SPLIT STRINGS BY CHARS. FROM PSC.
'http://www.planet-source-code.com/vb
Dim intAccs As Long
Dim i
Dim lngSpacePos As Long, lngStart As Long
    lngSpacePos = 1
    lngSpacePos = InStr(lngSpacePos, strMessage, Character)
    Do While lngSpacePos
        intAccs = intAccs + 1
        lngSpacePos = InStr(lngSpacePos + 1, strMessage, Character)
    Loop
    ReDim StrLines(intAccs)
    lngStart = 1
    For i = 0 To intAccs
        lngSpacePos = InStr(lngStart, strMessage, Character)
        If lngSpacePos Then
            StrLines(i) = Mid(strMessage, lngStart, lngSpacePos - lngStart)
            lngStart = lngSpacePos + Len(Character)
        Else
            StrLines(i) = Right(strMessage, Len(strMessage) - lngStart + 1)
        End If
    Next
End Sub

Sub ColorComments(Optional Colour As Long = &H80&)
fMainForm.RTF1.Visible = False
fMainForm.SB.Panels(2).Text = "Loading...  "
Screen.MousePointer = 11
Dim Sentences() As String, i As Integer, pos As Long, pos2 As Long
SplitStr fMainForm.RTF1.Text, Sentences, vbNewLine
For i = 0 To UBound(Sentences)
    If Left(Trim(Sentences(i)), 1) = "'" Then
        Do
            pos = InStr(pos + 1, fMainForm.RTF1.Text, Sentences(i))
            If pos = 0 Then Exit Do
            fMainForm.RTF1.SelStart = pos - 1
            fMainForm.RTF1.SelLength = Len(Sentences(i))
            fMainForm.RTF1.SelColor = Colour
        Loop
    End If
        Do
            pos = InStr(1, Trim(Sentences(i)), "'")
            If pos = 0 Then Exit Do
            pos2 = InStr(pos + 1, Sentences(i), Chr(13))
            If pos2 = 0 Then Exit Do
            fMainForm.RTF1.SelStart = pos2
            fMainForm.RTF1.SelLength = Len(Sentences(i)) - pos
            fMainForm.RTF1.SelColor = Colour
        Loop
Next i
fMainForm.RTF1.SelLength = 0
fMainForm.RTF1.SelColor = 0
fMainForm.RTF1.Visible = True
fMainForm.SB.Panels(2).Text = " Loaded " & ReadCustomVal("Attribute VB_Name", "") & ".  "
Screen.MousePointer = 0
End Sub

