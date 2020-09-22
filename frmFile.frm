VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Memo File"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Export"
      Height          =   1020
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2205
      Width           =   5655
      Begin VB.CommandButton cmdExport 
         Height          =   550
         Left            =   240
         Picture         =   "frmFile.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   315
         Width           =   550
      End
      Begin VB.Label Label2 
         Caption         =   "Export the currently showed memo"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   15
         Top             =   525
         Width           =   4095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Add "
      Height          =   1020
      Left            =   120
      TabIndex        =   12
      Top             =   1155
      Width           =   5655
      Begin VB.CommandButton cmdAdd 
         Height          =   550
         Left            =   240
         Picture         =   "frmFile.frx":1072
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   550
      End
      Begin VB.Label Label2 
         Caption         =   "Add memos from another memo file"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   13
         Top             =   525
         Width           =   4575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Copy"
      Height          =   1020
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   3255
      Width           =   5655
      Begin VB.CommandButton cmdBackUp 
         Height          =   550
         Left            =   240
         Picture         =   "frmFile.frx":20E4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   550
      End
      Begin VB.Label Label2 
         Caption         =   "Save a copy from this memo file"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   525
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4545
      TabIndex        =   6
      Top             =   5460
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Caption         =   "Create"
      Height          =   1020
      Left            =   105
      TabIndex        =   7
      Top             =   4305
      Width           =   5655
      Begin VB.CommandButton cmdCreateFile 
         Height          =   550
         Left            =   240
         Picture         =   "frmFile.frx":3156
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   315
         Width           =   550
      End
      Begin VB.Label Label2 
         Caption         =   "Create a new memo file and set it as default memo"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   525
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Open"
      Height          =   990
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdOpenFile 
         Height          =   550
         Left            =   240
         Picture         =   "frmFile.frx":41C8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   550
      End
      Begin VB.Label Label1 
         Caption         =   "Open a memo and set it as default memo"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   525
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   420
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      InitDir         =   "c:"
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
On Error Resume Next
StoreCurrentMemo
If gstrCurrentFile <> "" Then 'And Dir(gstrCurrentFile) <> "" Then
    Me.cmdAdd.Enabled = True
    If Len(gstrMemo) > 13 Then
        Me.cmdExport.Enabled = True
        Me.cmdBackUp.Enabled = True
        Else
        Me.cmdExport.Enabled = False
        Me.cmdBackUp.Enabled = False
        End If
    Else
    Me.cmdAdd.Enabled = False
    Me.cmdExport.Enabled = False
    Me.cmdBackUp.Enabled = False
    End If
Me.cmdOpenFile.SetFocus
Err.Clear
End Sub

Private Sub cmdAdd_Click()
Dim FileNr
Dim addName As String
Dim addText As String
On Error Resume Next
With Me.dlgOpen
    .filename = ""
    .DialogTitle = " Add Memo File..."
    .Flags = &H1000 Or &H4
    .DefaultExt = ".mem"
    .InitDir = gstrCurrentFile
    .Filter = "Memo Files (*.mem)|*.mem"
    .FilterIndex = 1
    .ShowOpen
    If Err = 32755 Or .filename = "" Then Exit Sub
    addName = .filename
End With
On Error GoTo errorHandle
'get add file
Screen.MousePointer = 11
FileNr = FreeFile
Open addName For Input As #FileNr
addText = Input(LOF(FileNr), FileNr)
Close #FileNr
'merge memos
beginCut:
If Right(addText, 2) = vbCrLf Then addText = Left(addText, Len(addText) - 2): GoTo beginCut
gstrMemo = gstrMemo & vbCrLf & addText & vbCrLf
Screen.MousePointer = 0
gblnMemoChanged = True
curPos = Len(gstrMemo)
Call FindMemo(Down)
Call ShowCurrentMemo
Call CheckButtons
Me.Hide
Exit Sub
errorHandle:
    Close #FileNr
    Screen.MousePointer = 11
    MsgBox "Failed adding the Memo File.", vbCritical
End Sub

Private Sub cmdBackUp_Click()
Dim x As Integer
Dim strBackUpFile As String
On Error Resume Next
With frmFile.dlgOpen
    .filename = "Copy from " & StripFileName(gstrCurrentFile) & ".mem"
    .DialogTitle = " Save Copy As..."
    .Flags = &H2 Or &H4
    .DefaultExt = ".mem"
    .InitDir = gstrCurrentFile
    .Filter = "Memo Files (*.mem)|*.mem"
    .FilterIndex = 1
    .ShowSave
    If Err = 32755 Or .filename = "" Then Exit Sub
    strBackUpFile = .filename
End With
x = SaveMemoFile(strBackUpFile)
Call ShowCurrentMemo
frmFile.Hide
End Sub

Private Sub cmdExport_Click()
Dim exportFile As String
Dim FileNr
If gblnEditNewMemo = True Then
    Call StoreCurrentMemo
    curPos = Len(gstrMemo)
    Call FindMemo(Down)
    End If
On Error Resume Next
With Me.dlgOpen
    .filename = "Export"
    .DialogTitle = " Export Memo..."
    .Flags = &H2 Or &H4
    .DefaultExt = ".mem"
    .InitDir = gstrCurrentFile
    .Filter = "Memo Files (*.mem)|*.mem"
    .FilterIndex = 1
    .ShowSave
    If Err = 32755 Or .filename = "" Then Exit Sub
End With
On Error GoTo errorHandle
Screen.MousePointer = 11
FileNr = FreeFile
Dim WriteText
WriteText = ComposeMemo
exportFile = Me.dlgOpen.filename
cutCRLF:
If Right(gstrMemo, 2) = vbCrLf Then gstrMemo = Left(gstrMemo, Len(gstrMemo) - 2): GoTo cutCRLF
'write export file
Open exportFile For Output As #FileNr
    Print #FileNr, WriteText
Close #FileNr
'
Screen.MousePointer = 0
Me.Hide
Exit Sub
errorHandle:
    MsgBox "Failed saving Memo." & vbCr & vbCr & "Error: " & Error, 48
    Close #FileNr
    Screen.MousePointer = 0
End Sub

Private Sub cmdOpenFile_Click()
Dim FileNr
Dim x As Integer
Dim tmp As String
Dim retval As Integer
On Error Resume Next
With Me.dlgOpen
    .filename = ""
    .DialogTitle = " Open Memo File..."
    .Flags = &H1000 Or &H4
    .DefaultExt = ".mem"
    .InitDir = gstrCurrentFile
    .Filter = "Memo Files (*.mem)|*.mem"
    .FilterIndex = 1
    .ShowOpen
    If Err = 32755 Or .filename = "" Then Exit Sub
    Me.Hide
    'first save and set current free if owner
    x = SaveMemoFile(gstrCurrentFile)
    If x <> 0 Then
        'warn if not saved
        retval = MsgBox(StripFileName(gstrCurrentFile) & " could not be saved. If you open a new memo file, all changes will be lost." & vbCrLf & vbCrLf & "Do you want to open a new memo file?", vbYesNo + vbExclamation, " Create New Memo File")
        If retval = vbNo Then Exit Sub
        End If
    If gblnReadOnly = False Then x = SendLOC(gstrCurrentFile, Free)
    'load other
    gstrCurrentFile = .filename
    x = OpenMemoFile(gstrCurrentFile)
    curPos = Len(gstrMemo)
    'Direction = memoDn
    Call ShowCurrentMemo
End With
Me.Hide
Call CheckButtons
End Sub

Private Sub cmdCreateFile_Click()
Dim FileNr
Dim x As Integer
Dim tmpName As String
Dim retval As Integer
Dim oldFileName As String
On Error Resume Next
With Me.dlgOpen
    .filename = ""
    .DialogTitle = " Create New Memo File..."
    .Flags = &H2 Or &H4
    .DefaultExt = ".mem"
    .InitDir = gstrCurrentFile
    .Filter = "Memo Files (*.mem)|*.mem"
    .FilterIndex = 1
    .ShowSave
    If Err = 32755 Or .filename = "" Then Exit Sub
    'first save and set current free
    x = SaveMemoFile(gstrCurrentFile)
    If x <> 0 Then
        'warn if not saved
        retval = MsgBox(StripFileName(gstrCurrentFile) & " could not be saved. If you open a new memo file, all changes will be lost." & vbCrLf & vbCrLf & "Do you want to create a new memo file?", vbYesNo + vbExclamation, " Create New Memo File")
        If retval = vbNo Then Exit Sub
        End If
    If gblnReadOnly = False Then x = SendLOC(gstrCurrentFile, Free)
    tmpName = .filename
    oldFileName = gstrCurrentFile
End With
'save new
gstrMemo = ""
gblnMemoChanged = True
gstrCurrentFile = tmpName
x = SaveMemoFile(gstrCurrentFile)
'if error, cancel creating new
'   If x <> 0 Then gstrCurrentFile = oldFileName
frmMain.Caption = " Memo - " & StripFileName(gstrCurrentFile)
frmMain.txtBox.Text = ""
gstrOldMemo = ""
curStop = 0
curStart = 0
curPos = 0
Me.Hide
Call FindMemo(Down)
Call ShowCurrentMemo
Call CheckButtons
End Sub

Private Sub cmdCancel_Click()
frmFile.Hide
End Sub

