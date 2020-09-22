Attribute VB_Name = "ModFiles"
Option Explicit

Public Function OpenMemoFile(strName As String) As Integer
Dim x As Integer
Dim tmpLOC As String
Dim tmpName As String
Dim retval As Integer
Dim FileNr
Dim fAttrib As Integer
'check write permissions
gblnReadOnly = False
gblnreadLOC = False
'reset LOC requests
frmMain.TimerLOC.Enabled = False
frmMain.mniWriteRequest.Checked = False
'get attribute read only
fAttrib = GetAttr(strName)
'skip loc if file itselve is readonly
If fAttrib And 1 Then gblnReadOnly = True: GoTo skipLOC
'get LOC
tmpLOC = ReadLOC(strName)
'ignore errors on reading ?
'1-if disk is readonly, the file will autom. be readonly
'2-first loc gives also error, so ignore it too
If tmpLOC = "ERROR" Then tmpLOC = ""
'if loc and if not current id, then readonly
If tmpLOC <> "" And Left(tmpLOC, 10) <> gstrUserId Then
    gblnReadOnly = True
    gblnreadLOC = True
    tmpName = Trim(Mid(tmpLOC, 31))
    If tmpName = "" Then tmpName = Left(tmpLOC, 10)
    retval = MsgBox(StripFileName(strName) & " can only be opened as Read-Only." & vbCrLf & vbCrLf & "The memo is in use by " & tmpName & " since " & Mid(tmpLOC, 11, 14) & " Hr" & vbCrLf & vbCrLf & "Dou you want to be notified if you can edit again?", vbYesNo + vbInformation)
    If retval = vbYes Then
        'set timer for LOC request
        frmMain.TimerLOC.Enabled = True
        frmMain.mniWriteRequest.Checked = True
        End If
    End If
If tmpLOC <> "" And Left(tmpLOC, 10) = gstrUserId Then
    'duplicate user !
    End If
skipLOC:
'write loc if free
If gblnReadOnly = False Then
    x = SendLOC(strName, Busy)
    'readonly if the sendloc has failed (no permission or no network)
    If x <> 0 Then gblnReadOnly = True
    End If
'open the file
Screen.MousePointer = 11
Err.Clear
On Error GoTo ErrorHandling
FileNr = FreeFile
'
Open strName For Input As #FileNr
frmMain.txtBox.Text = vbCrLf & "   FILE IS OPENED - " & Format(Int(LOF(FileNr) / 1024), "###,###,##0") & " kB ..."
frmMain.txtBox.Refresh
gstrMemo = Input(LOF(FileNr), FileNr)
Close #FileNr
'
cutCRLF:
If Right(gstrMemo, 2) = vbCrLf Then gstrMemo = Left(gstrMemo, Len(gstrMemo) - 2): GoTo cutCRLF
gstrMemo = gstrMemo & vbCrLf
'
Screen.MousePointer = 0
frmMain.txtBox.Text = ""
frmMain.Caption = " Memo - " & StripFileName(strName)
If gblnReadOnly = True Then frmMain.Caption = frmMain.Caption & " [ Read Only ]"
gstrCurrentFile = strName
curPos = Len(gstrMemo)
gblnMemoChanged = False
Call GetMemo(Down)
Call ShowCurrentMemo
Call CheckButtons
'reset search
Call FindAlarm
OpenMemoFile = 0
Exit Function
'
ErrorHandling:
    frmMain.txtBox.Text = ""
    MsgBox StripFileName(strName) & " cannot be opened." & vbCr & vbCr & "Error: " & Error, vbCritical
    Close #FileNr
    Err.Clear
    Screen.MousePointer = 0
    OpenMemoFile = Err.Number
End Function

Public Function SaveMemoFile(strName As String) As Integer
Dim FileNr
Dim WriteText As String
'prepare save
Call StoreCurrentMemo
If strName = "" Then Exit Function
If gblnReadOnly = True Then Exit Function
If gblnMemoChanged = False Then Exit Function
frmMain.txtBox.Text = vbCrLf & "  FILE IS UPDATED..."
frmMain.txtBox.Refresh
'save memo
On Error GoTo errorHandle
Screen.MousePointer = 11
FileNr = FreeFile
WriteText = gstrMemo
cutCRLF:
If Right(gstrMemo, 2) = vbCrLf Then gstrMemo = Left(gstrMemo, Len(gstrMemo) - 2): GoTo cutCRLF
'
Open strName For Output As #FileNr
Print #FileNr, WriteText
Close #FileNr
'
'set parm after good save
Screen.MousePointer = 0
SaveMemoFile = 0
gblnMemoChanged = False
gstrOldMemo = ""
frmMain.txtBox.Text = ""
Call ShowCurrentMemo
Exit Function
'
errorHandle:
    frmMain.txtBox.Text = ""
    Call ShowCurrentMemo
    MsgBox "Saving " & StripFileName(strName) & " failed." & vbCr & vbCr & "Error: " & Error, vbCritical
    Close #FileNr
    Screen.MousePointer = 0
    SaveMemoFile = Err.Number
End Function

Private Function FilterName(ByVal strFile As String) As String
'returns filename without extension
Dim k   As Integer
Dim pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "." Then pos = k
Next k
If pos = 0 Then
    FilterName = strFile
    Else
    FilterName = Left(strFile, pos - 1)
    End If
End Function

Public Function SendLOC(ByVal fromFile As String, ByVal LOCmsg As LOCstatus) As Integer
Dim FileNr
Dim wText As String
If fromFile = "" Then Exit Function
'choose status
Select Case LOCmsg
Case LOCstatus.Free
    wText = ""
Case LOCstatus.Busy
    wText = gstrUserId & Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm") & ">>>>>>" & gstrUserName
End Select
wText = wText & Space(50 - Len(wText))
'write status
fromFile = FilterName(fromFile) & ".mlc"
Screen.MousePointer = 11
On Error Resume Next
FileNr = FreeFile
Open fromFile For Output As #FileNr
    Print #FileNr, wText
Close #FileNr
Screen.MousePointer = 0
If Err Then
    SendLOC = 1
    If LOCmsg = LOCstatus.Free Then MsgBox "The Write Status from " & StripFileName(gstrCurrentFile) & " could not be send to the station." & vbCrLf & vbCrLf & "Connection to the station could be lost or there's a problem with the network. If the Write Status is not updated, the memo file could stay marked as 'open for writing'. In this case, you need to reopen this memo later, and override the Write Status.", vbExclamation, " Exit current memo file"
    End If
Err.Clear
End Function

Public Function ReadLOC(ByVal fromFile As String) As String
Dim FileNr
If fromFile = "" Then ReadLOC = "": Exit Function
fromFile = FilterName(fromFile) & ".mlc"
On Error Resume Next
FileNr = FreeFile
Screen.MousePointer = 11
Open fromFile For Input As #FileNr
ReadLOC = Input(LOF(FileNr), FileNr)
Close #FileNr
Screen.MousePointer = 0
If Err Then
    ReadLOC = "ERROR"
    Else
    ReadLOC = Trim(ReadLOC)
    End If
Err.Clear
If ReadLOC = "ERROR" Then Exit Function
'strip crlf
If ReadLOC = vbCrLf Then ReadLOC = ""
If Len(ReadLOC) > 2 Then
    If Right(ReadLOC, 2) = vbCrLf Then ReadLOC = Left(ReadLOC, Len(ReadLOC) - 2)
    End If
'fill to buffer lenght 50
If Len(ReadLOC) > 0 Then ReadLOC = ReadLOC & Space(50 - Len(ReadLOC))
End Function
