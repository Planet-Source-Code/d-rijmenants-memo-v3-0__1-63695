Attribute VB_Name = "ModMain"

'-------------------------------------
'                                    '
'     Memo v3.0 memo organizer       '
'                                    '
'  (c) RDDS - RD Data Systems 2002   '
'                                    '
'-------------------------------------
Option Explicit

Public gstrMemo         As String
Public gstrCurrentFile As String
Public gstrOldMemo      As String
Public gstrPrevSearch   As String

Public gstrUserId       As String
Public gstrUserName     As String
Public gblnIgnoreLOC    As Boolean
Public gblnreadLOC      As Boolean

Public gblnMemoChanged  As Boolean
Public gblnReadOnly     As Boolean
Public gblnSearchBusy   As Boolean
Public gblnEditNewMemo  As Boolean
Public gblnNewSearch    As Boolean
Public gblnSearchLock   As Boolean
Public gblnSearchHit    As Boolean
Public gblnLock         As Boolean
Public gblnOldLock      As Boolean
Public gblnAlarm        As Boolean
Public gblnOldAlarm     As Boolean
Public gcontinueFind    As Boolean

Public gintTextCol      As Integer
Public gintTextFont     As Integer
Public gintTextSize     As Integer

Public gblnChangeMarges As Boolean
Public gblnPrintBusy    As Boolean
Public gblnCancelPrint  As Boolean

Public gintArguments    As Integer
Public gstrArg1         As String
Public gstrArg2         As String
Public gstrArg3         As String

Public curPos           As Long
Public curStart         As Long
Public curStop          As Long
Public oldStart         As Long
Public oldStop          As Long
Public lastFound        As Long

Public glngLmPrint      As Long
Public glngRmPrint      As Long
Public glngTmPrint      As Long
Public glngBmPrint      As Long

Public gstrLOC As String * 20
Public Enum LOCstatus
    Free
    Busy
End Enum

Public Enum Direction
    Up
    Down
End Enum

Public Const COL_WHITE = &H8000000E
Public Const COL_GRAY = &H8000000B

Sub Main()
Dim tmp     As String
Dim x       As Integer
Load frmMain
Load frmFile
Load frmPage
Load frmInfo
Load frmSplash
Load frmUser
'get user id info
Call CheckID
'settings menu
On Error Resume Next
gintTextCol = Val(GetSetting(App.EXEName, "CONFIG", "TextCol", "1"))
gintTextSize = Val(GetSetting(App.EXEName, "CONFIG", "TextSize", "3"))
gintTextFont = Val(GetSetting(App.EXEName, "CONFIG", "TextFont", "1"))
If gintTextCol < 1 Or gintTextCol > 7 Then gintTextCol = 1
If gintTextSize < 1 Or gintTextSize > 6 Then gintTextSize = 3
If gintTextFont < 1 Or gintTextFont > 4 Then gintTextFont = 1
frmMain.mniCol(gintTextCol).Checked = True
frmMain.mniSize(gintTextSize).Checked = True
frmMain.mniFont(gintTextFont).Checked = True
'view textfield
Call SetColor
With frmMain
    .txtBox.Font = Mid(.mniFont(gintTextFont).Caption, 2)
    .txtBox.FontSize = Val(.mniSize(gintTextSize).Caption)
    .txtBox.FontBold = False
    .txtBox.FontItalic = False
End With
'printer check
tmp = Printer.DeviceName
If Err Or tmp = "" Then
    frmMain.Toolbar1.Buttons("print").Enabled = False
    frmMain.mniPrint.Enabled = False
    Err.Clear
    Else
    frmMain.Toolbar1.Buttons("print").Enabled = True
    frmMain.mniPrint.Enabled = True
    End If
' marges printer
glngTmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterTop", "5"))
glngBmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterBottom", "5"))
glngLmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterLeft", "5"))
glngRmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterRight", "5"))
frmPage.txtPrintHead = GetSetting(App.EXEName, "CONFIG", "Headtext", "")
frmPage.chkPrint.Value = Val(GetSetting(App.EXEName, "CONFIG", "Chkhead", "false"))
Call GetWindowPos
frmMain.Show
' current memo file
Screen.MousePointer = 11
gstrCurrentFile = GetSetting(App.EXEName, "CONFIG", "File", "")
If Command <> "" Then
    If Right(Command, 4) = ".mem" Then
        gstrCurrentFile = Command
        End If
    End If
'check file
getMemoFile:
If gstrCurrentFile = "" Or Dir(gstrCurrentFile) = "" Then
    gstrCurrentFile = ""
    'open-file window if no file
    Screen.MousePointer = 0
    frmFile.Show (vbModal)
    End If
Screen.MousePointer = 0
If gstrCurrentFile = "" Then
    'no file, exit prg
    Unload frmMain
    Exit Sub
    End If
frmMain.txtBox.Text = vbCrLf & "   OPENING MEMO FILE..."
frmMain.Refresh
'load file
x = OpenMemoFile(gstrCurrentFile)
'if error, get again file
If x <> 0 Then gstrCurrentFile = "": GoTo getMemoFile
'show mainform
frmMain.Show
'show last memo
If gblnSearchLock = False Then
    curPos = Len(gstrMemo)
    Call FindMemo(Down)
    Call GetWindowPos
    End If
End Sub

Public Sub GetMemo(UpDn As Direction)
Dim sta1 As Long
Dim stp1 As Long
Dim revPos As Long
Dim allArg As String
If UpDn = Direction.Up Then
    'search up
    sta1 = InStr(curPos + 1, gstrMemo, "[~")
    If sta1 > 0 Then stp1 = InStr(sta1 + 1, gstrMemo, "[~")
    Else
    'search down
    For revPos = curPos - 1 To 1 Step -1
        If revPos < 1 Then revPos = 0: Exit For
        If Mid(gstrMemo, revPos, 2) = "[~" Then Exit For
        Next revPos
    sta1 = revPos
    If sta1 < 0 Then sta1 = 0
    stp1 = InStr(sta1 + 1, gstrMemo, "[~")
End If
If sta1 > 0 Then
    curStart = sta1
    curStop = stp1
    curPos = sta1
    Else
    curPos = 0
    End If
Call SetPosBar(curPos)
End Sub


Public Sub FindMemo(UpDn As Direction)
Dim k           As Integer
Dim tmp         As String
Dim blnA        As Boolean
Dim allArg      As String
Dim blnFound    As Boolean
oldStart = curStart
oldStop = curStop
If frmMain.Combo1.Text = "!" Then frmMain.Combo1.Text = "Important Memos"
'continue same search?
If gstrPrevSearch = frmMain.Combo1.Text Then
    gblnNewSearch = False
    Else
    gblnNewSearch = True
    gstrPrevSearch = frmMain.Combo1.Text
    End If
' check for limit begin or end
If UpDn = Direction.Down And curStop > Len(gstrMemo) + 1 Then Beep: Exit Sub
If UpDn = Direction.Up And curStart < 1 Then Beep: Exit Sub
If Not gblnSearchBusy Then
    Call GetMemo(UpDn) '<<<<<<
    ShowCurrentMemo
    Exit Sub
    End If
If gblnNewSearch Then
    'add new search to combo list
    For k = 0 To frmMain.Combo1.ListCount
    If frmMain.Combo1.Text = frmMain.Combo1.List(k) Then blnA = True
    Next k
    If Not blnA Then frmMain.Combo1.AddItem (frmMain.Combo1.Text)
    End If
'if new search, start at begin or end
If gblnNewSearch = True And UpDn = Direction.Up Then curPos = 0: gblnNewSearch = False
If gblnNewSearch = True And UpDn = Direction.Down Then curPos = Len(gstrMemo): gblnNewSearch = False
Call SetArguments
Screen.MousePointer = 11
frmMain.Combo1.SetFocus
gcontinueFind = True
'block all menu's during search
With frmMain
.mnuFile.Enabled = False
.mnuEdit.Enabled = False
.mnuView.Enabled = False
.mnuExtra.Enabled = False
.mnuInfo.Enabled = False
.Toolbar1.Enabled = False
.Toolbar2.Enabled = False
.txtBox.Enabled = False
End With
Do
'get next memo
Call GetMemo(UpDn)
If curPos = 0 Then gcontinueFind = False: Exit Do
'check if match
If curStop <> 0 Then
    tmp = UCase(Mid(gstrMemo, curStart, (curStop - curStart)))
    Else
    tmp = UCase(Mid(gstrMemo, curStart))
    End If
blnFound = FindMatch(tmp)
'make interrupt possible when searching next hit
DoEvents
If gcontinueFind = False Then blnFound = True
Loop While Not blnFound
With frmMain
.mnuFile.Enabled = True
.mnuEdit.Enabled = True
.mnuView.Enabled = True
.mnuExtra.Enabled = True
.mnuInfo.Enabled = True
.Toolbar1.Enabled = True
.Toolbar2.Enabled = True
.txtBox.Enabled = True
End With
Screen.MousePointer = 0
If gintArguments = 0 Then blnFound = True
If blnFound Then
    'show hit
    Call ShowCurrentMemo
    gblnSearchHit = True
    Else
    'not found,set argument for search next hit
    If gintArguments = 3 Then allArg = gstrArg1 & " + " & gstrArg2 & " + " & gstrArg3
    If gintArguments = 2 Then allArg = gstrArg1 & " + " & gstrArg2
    If gintArguments = 1 Then allArg = gstrArg1
    If frmMain.Combo1.Text = "Important Memos" Then allArg = "Important Memos"
    If gblnSearchHit = True Then
        'if previous was hit
        MsgBox "Search for " & allArg & "  completed.", vbInformation, " Find Memo"
        Else
        'if no previours hit
        MsgBox allArg & "  not found.", vbInformation, " Find Memo"
        End If
    'set pointers to last hit
    gblnSearchHit = False
    gstrPrevSearch = ""
    curPos = lastFound
    SetPosBar curPos
    curStart = oldStart
    curStop = oldStop
    frmMain.Combo1.SetFocus
    End If
gcontinueFind = False
End Sub

Private Function FindMatch(strA As String) As Boolean
Dim ar1 As String
Dim ar2 As String
Dim ar3 As String
Dim blnFound As Boolean
ar1 = UCase(gstrArg1)
ar2 = UCase(gstrArg2)
ar3 = UCase(gstrArg3)
strA = UCase(strA)
Select Case gintArguments
Case 0
    blnFound = True
Case 1
    If InStr(1, strA, ar1) <> 0 Then FindMatch = True
Case 2
    If InStr(1, strA, ar1) <> 0 _
    And InStr(1, strA, ar2) <> 0 Then FindMatch = True
Case 3
    If InStr(1, strA, ar1) <> 0 _
    And InStr(1, strA, ar2) <> 0 _
    And InStr(1, strA, ar3) <> 0 Then FindMatch = True
End Select
End Function

Sub ShowCurrentMemo()
Dim tmp
Dim Header As String
If Len(gstrMemo) < 13 Then
    'no memos = textfield empty
    gstrMemo = ""
    frmMain.txtBox.Text = vbCrLf & "   No memos in file."
    frmMain.txtBox.Enabled = False
    frmMain.txtDate.Text = ""
    frmMain.txtDate.Enabled = False
    frmMain.picPointer.Left = 0
    curPos = 0
    Exit Sub
    End If
If curPos = 0 Then Beep: curPos = curStart: Exit Sub
frmMain.txtBox.Enabled = True
frmMain.txtDate.Enabled = True
'get current memo
If curStop <> 0 Then
    tmp = Mid(gstrMemo, curStart, (curStop - curStart))
    Else
    tmp = Mid(gstrMemo, curStart)
    End If
'get header
Header = Left(tmp, 17)
'get all after header [~XXX~31/12/2000] & 13/10
frmMain.txtBox.Text = Mid(tmp, 20) & vbCrLf
'strip text
beginCut:
If Right(frmMain.txtBox.Text, 4) = vbCrLf & vbCrLf Then frmMain.txtBox.Text = Left(frmMain.txtBox.Text, Len(frmMain.txtBox.Text) - 2): GoTo beginCut
'get date
frmMain.txtDate.Text = Mid(Header, 7, 10)
'get attrib alarm
If Mid(Header, 3, 1) = "A" Then
    gblnAlarm = True
    Else
    gblnAlarm = False
    End If
'get attrib Lock
If Mid(Header, 4, 1) = "L" Then
    gblnLock = True
    Else
    gblnLock = False
    End If
'update toolbar attrib
CheckAttrib
gblnOldLock = gblnLock
gblnOldAlarm = gblnAlarm
gstrOldMemo = frmMain.txtBox.Text
curPos = curStart
lastFound = curStart
oldStart = curStart
oldStop = curStop
SetPosBar curPos
End Sub

Sub StoreCurrentMemo()
Dim tmp As String
'if no text, don't store
If frmMain.txtBox.Text = "" And gstrOldMemo = "" Then Exit Sub
If (gstrOldMemo = frmMain.txtBox.Text And gblnLock = gblnOldLock And gblnAlarm = gblnOldAlarm) _
    Or frmMain.txtBox.Text = vbCrLf & "   No memos in file." Then
    gblnEditNewMemo = False
    Exit Sub
    End If
gblnMemoChanged = True
If gblnEditNewMemo Then
    'store new memo in gstrmemo
    gblnEditNewMemo = False
    gstrOldMemo = frmMain.txtBox.Text
    curStart = Len(gstrMemo) + 1
    gstrMemo = gstrMemo & vbCrLf & ComposeMemo
Else
    If curStop = 0 And Len(gstrMemo) > 17 Then
        'store last memo again
        gstrMemo = Left(gstrMemo, curStart - 1) & ComposeMemo
        Else
        If Len(gstrMemo) > 17 Then
            'store memo inside gstrmemo
            gstrMemo = Left(gstrMemo, curStart - 1) & ComposeMemo & vbCrLf & Mid(gstrMemo, curStop)
            Else
            'store first memo
            gstrMemo = ComposeMemo
        End If
    End If
End If
curPos = curStart
curStop = InStr(curPos + 1, gstrMemo, "[~")
gblnEditNewMemo = False
gstrOldMemo = frmMain.txtBox.Text
Call ShowCurrentMemo
End Sub

Public Function ComposeMemo() As String
'write memo to string
Dim tmpT As String
Dim AttribA As String
Dim AttribL As String
Dim AttribO As String
Dim q As Long
Dim qS As Long
Dim k As Long
tmpT = RTrim(frmMain.txtBox.Text)
beginCut:
If Right(tmpT, 2) = vbCrLf Then tmpT = Left(tmpT, Len(tmpT) - 2): GoTo beginCut
tmpT = RTrim(tmpT)
'get attrib alarm
If gblnAlarm = True Then AttribA = "A" Else AttribA = "X"
'get attrib lock
If gblnLock = True Then AttribL = "L" Else AttribL = "X"
'get attrib other options
AttribO = "X"
'eliminate headers "[~" in text
qS = 1
Do
q = InStr(qS, tmpT, "[~")
If q > 0 Then Mid(tmpT, q, 2) = "[-"
qS = q + 1
If qS > Len(tmpT) Then q = 0
Loop While q > 0
'compose update string
ComposeMemo = "[~" & AttribA & AttribL & AttribO & "~" & Format(Date, "dd/mm/yyyy") & "]" & vbCrLf & tmpT & vbCrLf
End Function

Sub AddNewMemo()
If gblnReadOnly = True Then
    MsgBox "This memo file is Read-Only!", vbExclamation
    Exit Sub
    End If
gblnEditNewMemo = True
gstrOldMemo = ""
gblnLock = False
gblnOldLock = False
Call CheckAttrib
With frmMain
.txtBox.Enabled = True
.txtBox.Text = ""
.txtDate.Enabled = True
.txtDate.Text = Format(Date, "dd/mm/yyyy")
.txtBox.SetFocus
End With
curStart = Len(gstrMemo)
curPos = curStart
curStop = 0
Call CheckButtons
End Sub

Sub DeleteCurrentMemo()
Dim retval As Integer
If gblnReadOnly = True Then
    MsgBox "This memo file is Read-Only!", vbExclamation
    Exit Sub
    End If
If gblnLock = True Then
    MsgBox "This memo is locked and Read-Only!", vbExclamation
    Exit Sub
    End If
If curPos = 0 Then Beep: Exit Sub
retval = MsgBox("Are you sure you want to delete the memo from " & frmMain.txtDate.Text & "?", vbQuestion + vbYesNo, " Delete Memo")
If retval <> vbYes Then Exit Sub
gblnMemoChanged = True
If Len(gstrMemo) < 13 Then gstrMemo = "": Beep: Exit Sub
If curStop <> 0 Then
    gstrMemo = Left(gstrMemo, curStart - 1) + Mid(gstrMemo, curStop)
    Else
    gstrMemo = Left(gstrMemo, curStart - 1)
    End If
If Len(gstrMemo) > 13 Then
    curPos = curStart - 1
    Call GetMemo(Up)
    If curPos = 0 Then curPos = Len(gstrMemo): Call GetMemo(Down)
    End If
Call ShowCurrentMemo
Call CheckButtons
End Sub

Sub SetArguments()
Dim tmp As String
Dim endPos As Integer
tmp = Trim(frmMain.Combo1.Text)
If tmp = "Important Memos" Or tmp = "!" Then gintArguments = 1: gstrArg1 = "[~A": Exit Sub
If tmp = "" Then gintArguments = 0: Exit Sub
endPos = InStr(1, tmp, " ")
gintArguments = 1
If endPos = 0 Then gstrArg1 = Trim(tmp): Exit Sub
gstrArg1 = Trim(Left(tmp, endPos - 1))
If Len(Mid(tmp, endPos)) = 1 Then Exit Sub
tmp = Trim(Mid(tmp, endPos + 1))
endPos = InStr(1, tmp, " ")
gintArguments = 2
If endPos = 0 Then gstrArg2 = Trim(tmp): Exit Sub
gstrArg2 = Trim(Left(tmp, endPos - 1))
If Len(Mid(tmp, endPos)) = 1 Then Exit Sub
gstrArg3 = Trim(Mid(tmp, endPos + 1))
gintArguments = 3
'more than 3 argument = msgbox
If InStr(1, gstrArg3, " ") Then gintArguments = 0: _
    MsgBox "Maximum 3 key words, seperated by a space, are allowed." _
    , vbExclamation, " Find Memo"
End Sub

Public Sub FormResize()
Dim th
Dim ph
With frmMain
th = .Toolbar1.Height
ph = .picPosBar.Height
If .WindowState <> 1 And .ScaleHeight > 1700 And .ScaleWidth > 1700 Then
    If .ScaleHeight - th - 150 > 0 Then
        .txtBox.Height = .ScaleHeight - th - ph - 350
        .txtBox.Top = th
    .picPosBar.Width = .ScaleWidth
    .picPosBar.Top = .ScaleHeight - .Toolbar2.Height - ph
        End If
    .txtBox.Width = .ScaleWidth
    If .ScaleWidth > 1980 Then .Combo1.Width = .ScaleWidth - 1925
    SetPosBar curPos
    End If
End With
End Sub

Function StripFileName(ByVal strNameA As String) As String
Dim pos As Long
Do
pos = InStr(1, strNameA, "\")
strNameA = Right(strNameA, Len(strNameA) - pos)
Loop Until pos = 0
strNameA = Left(strNameA, Len(strNameA) - 4)
StripFileName = strNameA
End Function

Sub SaveWindowPos()
With frmMain
SaveSetting App.EXEName, "CONFIG", "WindowState", .WindowState
If .WindowState <> vbNormal Then Exit Sub
SaveSetting App.EXEName, "CONFIG", "WindowHeight", .Height
SaveSetting App.EXEName, "CONFIG", "WindowWidth", .Width
SaveSetting App.EXEName, "CONFIG", "WindowLeft", .Left
SaveSetting App.EXEName, "CONFIG", "WindowTop", .Top
End With
End Sub

Sub GetWindowPos()
With frmMain
.WindowState = CSng(GetSetting(App.EXEName, "Config", "Windowstate", vbNormal))
If .WindowState <> vbNormal Then Exit Sub
.Height = CSng(GetSetting(App.EXEName, "CONFIG", "WindowHeight", .Height))
.Width = CSng(GetSetting(App.EXEName, "CONFIG", "WindowWidth", .Width))
.Left = CSng(GetSetting(App.EXEName, "CONFIG", "WindowLeft", ((Screen.Width - .Width) / 2)))
.Top = CSng(GetSetting(App.EXEName, "CONFIG", "WindowTop", ((Screen.Height - .Height) / 2)))
End With
End Sub

Public Sub CheckID()
Dim k As Integer
Dim i As Integer
gstrUserId = GetSetting(App.EXEName, "CONFIG", "UserID", "")
gstrUserName = GetSetting(App.EXEName, "CONFIG", "UserName", "")
'if user id found, exit
If gstrUserId <> "" Then Exit Sub
'generate new user id
Randomize
gstrUserId = ""
For k = 1 To 10
    i = Int((90 - 65 + 1) * Rnd + 65)
    gstrUserId = gstrUserId & Chr(i)
Next k
SaveSetting App.EXEName, "CONFIG", "UserID", gstrUserId
End Sub

Public Sub ToggleFind()
If gblnEditNewMemo Then Exit Sub
gstrPrevSearch = ""
If Len(gstrMemo) < 13 Then Beep: Exit Sub
If gblnSearchBusy = False Then
    FindOn
    frmMain.mniFind.Caption = "&Cancel Search"
    Else
    FindOff
    frmMain.mniFind.Caption = "&Find"
    End If
End Sub

Public Sub FindOn()
If gblnEditNewMemo Then Exit Sub
gblnSearchBusy = True
gblnNewSearch = True
With frmMain
    .Combo1.Enabled = True
    .Combo1.BackColor = COL_WHITE
    .Toolbar2.Buttons("find").Image = "stopfind"
    .Toolbar2.Buttons("find").ToolTipText = " Cancel Find [ESC] "
    .Toolbar2.Buttons("first").ToolTipText = " Find First [CTRL+Home] "
    .Toolbar2.Buttons("previous").ToolTipText = " Find Previous [CTRL+Page Down]"
    .Toolbar2.Buttons("next").ToolTipText = " Find Next [CTRL+Page Up] "
    .Toolbar2.Buttons("last").ToolTipText = " Find Last [CTRL+End] "
    .Combo1.SetFocus
End With
End Sub

Public Sub FindOff()
If gblnEditNewMemo Then Exit Sub
gblnSearchBusy = False
With frmMain
    .Combo1.SelLength = 0
    .Combo1.Enabled = False
    .Combo1.BackColor = COL_GRAY
    .Toolbar2.Buttons("find").Image = "find"
    .Toolbar2.Buttons("find").ToolTipText = " FInd F3 "
    .Toolbar2.Buttons("first").ToolTipText = " First [CTRL+Home] "
    .Toolbar2.Buttons("previous").ToolTipText = " Previous [CTRL+Page Down] "
    .Toolbar2.Buttons("next").ToolTipText = " Next [CTRL+Page Up] "
    .Toolbar2.Buttons("last").ToolTipText = " Last [CTRL+End] "
    .txtBox.SetFocus
End With
End Sub

Public Sub CheckAttrib()
With frmMain
If gblnLock = True Then
    .Toolbar1.Buttons("lock").Image = "lock"
    .Toolbar1.Buttons("lock").ToolTipText = " Unlock Memo "
    .mniLock.Caption = "&Unlock Memo"
    .txtBox.Locked = True
    Else
    .Toolbar1.Buttons("lock").Image = "unlock"
    .Toolbar1.Buttons("lock").ToolTipText = " Lock Memo "
    .mniLock.Caption = "&Lock Memo"
    .txtBox.Locked = False
    End If
If gblnAlarm = True Then
    .Toolbar1.Buttons("alarm").Image = "alarm"
    Else
    .Toolbar1.Buttons("alarm").Image = "alarmoff"
    End If
If gblnReadOnly = True Then
    .txtBox.Locked = True
    Else
    .txtBox.Locked = False
    End If
End With
End Sub

Public Sub SetPosBar(posn As Long)
Dim barScale As Single
Dim bLen
bLen = Len(gstrMemo)
With frmMain
If bLen < 13 Then .picPointer.Left = 0: Exit Sub
If posn = 0 Then Exit Sub
barScale = (.picPosBar.Width - .picPointer.Width) / bLen
.picPointer.Left = posn * barScale
.picPosBar.Refresh
End With
End Sub

Public Sub CheckButtons()
With frmMain
    .Toolbar1.Buttons("delete").Enabled = True
    .Toolbar1.Buttons("lock").Enabled = True
    .Toolbar1.Buttons("print").Enabled = True
    .Toolbar1.Buttons("cut").Enabled = True
    .Toolbar1.Buttons("copy").Enabled = True
    .Toolbar1.Buttons("paste").Enabled = True
    .Toolbar1.Buttons("undo").Enabled = True
    .Toolbar1.Buttons("alarm").Enabled = True
    .Toolbar2.Buttons("first").Enabled = True
    .Toolbar2.Buttons("previous").Enabled = True
    .Toolbar2.Buttons("next").Enabled = True
    .Toolbar2.Buttons("last").Enabled = True
    .Toolbar2.Buttons("find").Enabled = True
    .mnuRepair.Enabled = True
    .mnuDelet.Enabled = True
    .mniMemoDel.Enabled = True
    .mniCut.Enabled = True
    .mniFind.Enabled = True
    .mniPrint.Enabled = True
'check for readonly or loc
If gblnReadOnly = False Then
    .Toolbar1.Buttons("new").Enabled = True
    .Toolbar1.Buttons("delete").Enabled = True
    .Toolbar1.Buttons("lock").Enabled = True
    .Toolbar1.Buttons("paste").Enabled = True
    .Toolbar1.Buttons("cut").Enabled = True
    .Toolbar1.Buttons("undo").Enabled = True
    .mniCut.Enabled = True
    .mniAlarm.Enabled = True
    .mniLock.Enabled = True
    .mniCopy.Enabled = True
    .mniPaste.Enabled = True
    .mniMemoAdd.Enabled = True
    .mniMemoDel.Enabled = True
    .mnuDelet.Enabled = True
    .mniUndo.Enabled = True
    .mniSave.Enabled = True
    .mnuRepair.Enabled = True
    Else
    .Toolbar1.Buttons("new").Enabled = False
    .Toolbar1.Buttons("delete").Enabled = False
    .Toolbar1.Buttons("lock").Enabled = False
    .Toolbar1.Buttons("paste").Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .Toolbar1.Buttons("undo").Enabled = False
    .mniCut.Enabled = False
    .mniAlarm.Enabled = False
    .mniLock.Enabled = False
    .mniPaste.Enabled = False
    .mniMemoAdd.Enabled = False
    .mniMemoDel.Enabled = False
    .mnuDelet.Enabled = False
    .mniUndo.Enabled = False
    .mniSave.Enabled = False
    .mnuRepair.Enabled = False
    End If
If gblnLock = True Then
    .Toolbar1.Buttons("delete").Enabled = False
    .Toolbar1.Buttons("alarm").Enabled = False
    .Toolbar1.Buttons("paste").Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .Toolbar1.Buttons("undo").Enabled = False
    .mniCut.Enabled = False
    .mniAlarm.Enabled = False
    .mniPaste.Enabled = False
    .mniMemoDel.Enabled = False
    .mnuDelet.Enabled = False
    .mniUndo.Enabled = False
    .mnuRepair.Enabled = False
    End If
If Len(gstrMemo) < 20 And gblnEditNewMemo = False Then '<<<
    .Toolbar1.Buttons("delete").Enabled = False
    .Toolbar1.Buttons("lock").Enabled = False
    .Toolbar1.Buttons("print").Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .Toolbar1.Buttons("copy").Enabled = False
    .Toolbar1.Buttons("paste").Enabled = False
    .Toolbar1.Buttons("undo").Enabled = False
    .Toolbar1.Buttons("alarm").Enabled = False
    .Toolbar2.Buttons("first").Enabled = False
    .Toolbar2.Buttons("previous").Enabled = False
    .Toolbar2.Buttons("next").Enabled = False
    .Toolbar2.Buttons("last").Enabled = False
    .Toolbar2.Buttons("find").Enabled = False
    .mniAlarm.Enabled = False
    .mniLock.Enabled = False
    .mniPaste.Enabled = False
    .mniMemoDel.Enabled = False
    .mnuRepair.Enabled = False
    .mnuRepair.Enabled = False
    .mnuDelet.Enabled = False
    .mniMemoDel.Enabled = False
    .mniCut.Enabled = False
    .mniFind.Enabled = False
    .mniPrint.Enabled = False
    End If
If Clipboard.GetText = "" Then
    .Toolbar1.Buttons("paste").Enabled = False
    .mniPaste.Enabled = False
    End If
If .txtBox.SelLength = 0 Then
    .Toolbar1.Buttons("copy").Enabled = False
    .mniCopy.Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .mniCut.Enabled = False
    End If
End With
End Sub

Public Sub SetColor()
With frmMain.txtBox
Select Case gintTextCol
Case 1 'Zwart wit
    .BackColor = &HFFFFFF
    .ForeColor = &H0
Case 2 'Lilablauw
    .BackColor = &HFFDDE0
    .ForeColor = &HC00000
Case 3 'Pastelgeel
    .BackColor = &HCEFFFE
    .ForeColor = &HFF0000
Case 4 'Pastelgroen
    .BackColor = &HDFFFDD
    .ForeColor = &HC00000
Case 5 'Woestijnrood
    .BackColor = &H80C0FF
    .ForeColor = &HFF&
Case 6 'Contrast zwart
    .BackColor = &H0
    .ForeColor = &HFF00&
Case 7 'Contrast blauw
    .BackColor = &HFF0000
    .ForeColor = &HFFFFFF
End Select
End With
End Sub

Public Sub FindAlarm()
Dim w1      As Long
Dim w2      As Long
Dim wTxt    As String
'first look for splash
Call FindSplash
'search important memo's
gblnSearchLock = True
frmMain.Combo1.AddItem ("Important Memos")
frmMain.Combo1.Text = "Important Memos"
w1 = InStr(1, gstrMemo, "[~A")
If w1 <> 0 Then
    'more than one?
    w2 = InStr(w1 + 1, gstrMemo, "[~A")
    If w2 = 0 Then
        wTxt = "Please read this important memo first!"
        Else
        wTxt = "Pleas go through these important memos first!"
        End If
    FindOn
    curPos = 0
    Call FindMemo(Up)
    MsgBox wTxt, vbExclamation, " Memo"
    frmMain.Combo1.SetFocus
    'only one = no search
    If w2 = 0 Then frmMain.Combo1.Text = "": FindOff
    End If
End Sub

Public Sub FindSplash()
'search for spash screen
Dim pos1 As Long
Dim pos2 As Long
Dim strSplash As String
pos1 = InStr(1, gstrMemo, "[SPLASH]")
If pos1 = 0 Then Exit Sub
pos2 = InStr(pos1, gstrMemo, "[~")
If pos2 = 0 Then pos2 = Len(gstrMemo)
If pos2 <= pos1 Then Exit Sub
strSplash = Mid(gstrMemo, pos1 + 8, pos2 - pos1 - 8)
frmSplash.Label1.Caption = strSplash
frmSplash.Show (vbModal)
End Sub

