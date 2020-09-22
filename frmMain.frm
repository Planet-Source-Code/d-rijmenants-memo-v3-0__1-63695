VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   " Memo"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   5745
   Begin VB.Timer TimerLOC 
      Interval        =   30000
      Left            =   5280
      Top             =   1155
   End
   Begin VB.PictureBox picPosBar 
      BackColor       =   &H00E0E0E0&
      Height          =   150
      Left            =   0
      ScaleHeight     =   90
      ScaleMode       =   0  'User
      ScaleWidth      =   5655
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3375
      Width           =   5655
      Begin VB.PictureBox picPointer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   135
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   158
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   3525
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "first"
            Object.ToolTipText     =   " First [CTRL+Home] "
            ImageKey        =   "first"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previous"
            Object.ToolTipText     =   " Previous [CTRL+Page Down] "
            ImageKey        =   "previous"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "next"
            Object.ToolTipText     =   " Next [CTRL+Page Up] "
            ImageKey        =   "next"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "last"
            Object.ToolTipText     =   " Last [CTRL+End] "
            ImageKey        =   "last"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   " Find [F3] "
            ImageKey        =   "find"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":030A
         Left            =   1875
         List            =   "frmMain.frx":030C
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   3015
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5145
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030E
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0852
            Key             =   "print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D96
            Key             =   "open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11AA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12BE
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1802
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D46
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":228A
            Key             =   "find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":239E
            Key             =   "stopfind"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24B2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29F6
            Key             =   "previous"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B0A
            Key             =   "next"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C1E
            Key             =   "first"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D32
            Key             =   "last"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E46
            Key             =   "unlock"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":338A
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38CE
            Key             =   "alarmoff"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E12
            Key             =   "alarm2"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4356
            Key             =   "alarmoff2"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":489A
            Key             =   "alarm"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DDE
            Key             =   "save"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtBox 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2850
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   5070
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   " Select Memo "
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   " Save Memo "
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   " Add New Memo "
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   " Delete Memo "
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alarm"
            Object.ToolTipText     =   " Important Memo "
            ImageKey        =   "alarmoff"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lock"
            ImageKey        =   "unlock"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   " Print Memo "
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   " Cut "
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   " Cupy "
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   " Paste "
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            Object.ToolTipText     =   " Undo "
            ImageKey        =   "undo"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox txtDate 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   220
         Left            =   4305
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "31/12/2000"
         ToolTipText     =   " Datum laatste wijziging "
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Memo"
      Begin VB.Menu mniFile 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mniSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu ln5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Prop&erties..."
      End
      Begin VB.Menu mnuRepair 
         Caption         =   "&Repair..."
      End
      Begin VB.Menu ln8 
         Caption         =   "-"
      End
      Begin VB.Menu mniPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mniExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mniUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu mniCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mniCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mniPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelet 
         Caption         =   "&Delet"
      End
      Begin VB.Menu ln4 
         Caption         =   "-"
      End
      Begin VB.Menu mniFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu ln6 
         Caption         =   "-"
      End
      Begin VB.Menu mniMemoAdd 
         Caption         =   "&New Memo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mniMemoDel 
         Caption         =   "Delete &Memo"
      End
      Begin VB.Menu ln7 
         Caption         =   "-"
      End
      Begin VB.Menu mniAlarm 
         Caption         =   "&Important Memo"
      End
      Begin VB.Menu mniLock 
         Caption         =   "&Lock Memo"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
         Begin VB.Menu mniFont 
            Caption         =   "&Courier New"
            Index           =   1
         End
         Begin VB.Menu mniFont 
            Caption         =   "&Fixedsys"
            Index           =   2
         End
         Begin VB.Menu mniFont 
            Caption         =   "&Lucida Console"
            Index           =   3
         End
         Begin VB.Menu mniFont 
            Caption         =   "&MS Sans Serif"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSize 
         Caption         =   "&Font Size"
         Begin VB.Menu mniSize 
            Caption         =   "8"
            Index           =   1
         End
         Begin VB.Menu mniSize 
            Caption         =   "9"
            Index           =   2
         End
         Begin VB.Menu mniSize 
            Caption         =   "10"
            Index           =   3
         End
         Begin VB.Menu mniSize 
            Caption         =   "12"
            Index           =   4
         End
         Begin VB.Menu mniSize 
            Caption         =   "14"
            Index           =   5
         End
         Begin VB.Menu mniSize 
            Caption         =   "16"
            Index           =   6
         End
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "&Colors"
         Begin VB.Menu mniCol 
            Caption         =   "Black and &White"
            Index           =   1
         End
         Begin VB.Menu mniCol 
            Caption         =   "&Lila Blue"
            Index           =   2
         End
         Begin VB.Menu mniCol 
            Caption         =   "Pastel &Yellow"
            Index           =   3
         End
         Begin VB.Menu mniCol 
            Caption         =   "Pastel &Green"
            Index           =   4
         End
         Begin VB.Menu mniCol 
            Caption         =   "Desert &Red"
            Index           =   5
         End
         Begin VB.Menu mniCol 
            Caption         =   "Contrast &Black"
            Index           =   6
         End
         Begin VB.Menu mniCol 
            Caption         =   "&Contrast Blue"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "&Extra"
      Begin VB.Menu mniUserName 
         Caption         =   "&User..."
      End
      Begin VB.Menu ln10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShareMemo 
         Caption         =   "Shared Memos"
         Begin VB.Menu mniWriteRequest 
            Caption         =   "&Request Write"
         End
         Begin VB.Menu mniWriteStatusForce 
            Caption         =   "&Override Write Status"
         End
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Info"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
Call FormResize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If gcontinueFind = True Then KeyAscii = 0: Exit Sub
Select Case KeyAscii
Case 14 'ctrl+n nieuw
    KeyAscii = 0
    StoreCurrentMemo
    If gblnSearchBusy Then FindOff
    Call AddNewMemo
Case 15 'ctrl+o open
    If gblnEditNewMemo = True Then
        StoreCurrentMemo
        curPos = Len(gstrMemo)
        Call FindMemo(Down)
        End If
    If gblnSearchBusy = True Then ToggleFind
    frmFile.Show (vbModal)
    KeyAscii = 0
End Select
End Sub

Private Sub mniAlarm_Click()
If gcontinueFind = True Then Exit Sub
If gblnReadOnly = True Then
    MsgBox "This Memo File is Read-Only!", vbExclamation
    Exit Sub
    End If
If gblnLock = True Then
    MsgBox "This Memo is locked and Read-Only!", vbExclamation, " Memo"
    Exit Sub
    End If
If frmMain.txtBox.Enabled = False Then Exit Sub
If gblnAlarm = False Then
    gblnAlarm = True
    Else
    gblnAlarm = False
    End If
CheckAttrib
Call CheckButtons
End Sub

Private Sub mniExit_Click()
Unload Me
End Sub

Private Sub mniFile_Click()
If gcontinueFind = True Then Exit Sub
If gblnEditNewMemo = True Then
    StoreCurrentMemo
    curPos = Len(gstrMemo)
    Call FindMemo(Down)
    End If
If gblnSearchBusy = True Then ToggleFind
frmFile.Show (vbModal)
End Sub

Private Sub mniFind_Click()
If gcontinueFind = True Then Exit Sub
ToggleFind
End Sub

Private Sub mniLock_Click()
If gcontinueFind = True Then Exit Sub
If gblnReadOnly = True Then
    MsgBox "This Memo File is Read-Only!", vbExclamation
    Exit Sub
    End If
If gcontinueFind = True Then Exit Sub
If frmMain.txtBox.Enabled = False Then Exit Sub
If gblnLock = False Then
    gblnLock = True
    Else
    gblnLock = False
    End If
CheckAttrib
Call CheckButtons
End Sub

Private Sub mniMemoAdd_Click()
If gcontinueFind = True Then Exit Sub
StoreCurrentMemo
If gblnSearchBusy Then FindOff
Call AddNewMemo
End Sub

Private Sub mniMemoDel_Click()
If gcontinueFind = True Then Exit Sub
If gblnEditNewMemo = True Then
    gblnEditNewMemo = False
    curPos = Len(gstrMemo)
    Call FindMemo(Down)
    Else
    DeleteCurrentMemo
End If
End Sub

Private Sub mniPrint_Click()
If gcontinueFind = True Then Exit Sub
StoreCurrentMemo
frmPage.Show (vbModal)
End Sub

Private Sub mniSave_Click()
Dim x As Integer
If gcontinueFind = True Then Exit Sub
If gblnEditNewMemo = True Then
    Call StoreCurrentMemo
    Call FindMemo(Up)
    End If
x = SaveMemoFile(gstrCurrentFile)
End Sub

Private Sub mniUserName_Click()
If gcontinueFind = True Then Exit Sub
frmUser.Show (vbModal)
End Sub

Private Sub mniWriteRequest_Click()
If gcontinueFind = True Then Exit Sub
If Me.mniWriteRequest.Checked = False Then
    Me.mniWriteRequest.Checked = True
    Me.TimerLOC.Enabled = True
    Else
    Me.mniWriteRequest.Checked = False
    Me.TimerLOC.Enabled = False
    End If
End Sub

Private Sub mniWriteStatusForce_Click()
If gcontinueFind = True Then Exit Sub
Dim x As Integer
Dim retval As Integer
If gstrCurrentFile = "" Then Exit Sub
    retval = MsgBox("You may only override the Write Status if the status is incorrect," & _
    " for example if a memo file is maked 'open for writing' too long, and you have checked that no there are no other users." & _
    ".  The Write Status can create a conflict if another user opened this file. " & vbOKCancel + vbDefaultButton2 + vbExclamation)
If retval = vbOK Then
    'reopen file
    Me.TimerLOC.Enabled = False
    Me.mniWriteRequest.Checked = False
    x = SendLOC(gstrCurrentFile, Free)
    If x <> 0 Then
        'override faild
        MsgBox "Failed to force Write Status." & vbCrLf & vbCrLf & "Er kan niet geschreven worden naar het station waar het memobestand zich bevind. U heeft geen schrijfbevoegheid of er is een probleem met de netwerkverbinding.", vbCritical
        Exit Sub
        End If
    x = OpenMemoFile(gstrCurrentFile)
    End If
End Sub

Private Sub mnuExtra_Click()
If gcontinueFind = True Then Exit Sub
Dim retval As Integer
If gstrCurrentFile = "" Or gblnReadOnly = False Or gblnreadLOC = False Then
    Me.mnuShareMemo.Enabled = False
    Else
    Me.mnuShareMemo.Enabled = True
    End If
End Sub

Private Sub mnuEdit_Click()
Exit Sub '<<<<<<<<<<<<<<<<
If gcontinueFind = True Then Exit Sub
If gblnLock = False Then
    Me.mniLock.Caption = "&Lock Memo"
    Else
    Me.mniLock.Caption = "&Unlock Memo"
    End If
If gblnSearchBusy = False Then
    Me.mniFind.Caption = "&Find"
    Else
    Me.mniFind.Caption = "&Cancel Search"
    End If
End Sub

Private Sub mniCut_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{x}"
End Sub

Private Sub mniCopy_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{c}"
End Sub

Private Sub mniPaste_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{v}"
End Sub

Private Sub mniUndo_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{z}"
End Sub

Private Sub mnuDelet_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{Del}"
End Sub

Private Sub mnuFile_Click()
If gcontinueFind = True Then Exit Sub
End Sub

Private Sub mnuInfo_Click()
If gcontinueFind = True Then Exit Sub
frmInfo.Show (vbModal)
End Sub

Private Sub mnuRepair_Click()
If gcontinueFind = True Then Exit Sub
Dim x As Integer
Dim retval As Integer
retval = MsgBox("Do you want to cancel all changes to this memo file" & vbCrLf & "and restore the file as before opening it?", vbYesNo + vbExclamation, " Repair Memo")
If retval <> vbYes Then Exit Sub
x = OpenMemoFile(gstrCurrentFile)
curPos = Len(gstrMemo)
ShowCurrentMemo
End Sub

Private Sub mnuProperties_Click()
If gstrCurrentFile = "" Then Exit Sub
MsgBox vbCr & "File name: " & gstrCurrentFile & vbCr & vbCr & _
    "Lengte: " & Format((Int(FileLen(gstrCurrentFile)) / 1024), "###,###,##0") & " kB", vbOKOnly, " Memo Properties"
End Sub

Private Sub mniCol_Click(Index As Integer)
Me.mniCol(gintTextCol).Checked = False
Me.mniCol(Index).Checked = True
gintTextCol = Index
Call SetColor
End Sub

Private Sub mniFont_Click(Index As Integer)
Me.mniFont(gintTextFont).Checked = False
Me.mniFont(Index).Checked = True
gintTextFont = Index
Me.txtBox.Font = Mid(Me.mniFont(Index).Caption, 2)
End Sub

Private Sub mniSize_Click(Index As Integer)
Me.mniSize(gintTextSize).Checked = False
Me.mniSize(Index).Checked = True
gintTextSize = Index
Me.txtBox.FontSize = Val(Me.mniSize(Index).Caption)
End Sub

Private Sub mnuView_Click()
If gcontinueFind = True Then Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Integer
If gcontinueFind = True Then Exit Sub
If (Button.Key <> "new" And Button.Key <> "open") And frmMain.txtBox.Enabled = False Then Exit Sub
Select Case Button.Key
Case "cut"
    SendKeys "^{x}"
Case "copy"
    SendKeys "^{c}"
Case "paste"
    SendKeys "^{v}"
Case "undo"
    SendKeys "^{z}"
Case "print"
    StoreCurrentMemo
    frmPage.Show (vbModal)
Case "open"
    If gblnEditNewMemo = True Then
        StoreCurrentMemo
        curPos = Len(gstrMemo)
        Call FindMemo(Down)
        End If
    If gblnSearchBusy = True Then ToggleFind
    frmFile.Show (vbModal)
Case "save"
    If gblnEditNewMemo = True Then
        Call StoreCurrentMemo
        Call FindMemo(Up)
        End If
    x = SaveMemoFile(gstrCurrentFile)
Case "new"
    If gblnReadOnly = True Then
        MsgBox "This memo file is Read-Only!", vbExclamation
        Exit Sub
        End If
    StoreCurrentMemo
    If gblnSearchBusy Then FindOff
    Call AddNewMemo
Case "delete"
    If gblnReadOnly = True Then
        MsgBox "This memo file is Read-Only!", vbExclamation
        Exit Sub
        End If
    If gblnEditNewMemo = True Then
        gblnEditNewMemo = False
        curPos = Len(gstrMemo)
        Call FindMemo(Down)
    Else
        DeleteCurrentMemo
    End If
Case "lock"
    If gblnReadOnly = True Then
        MsgBox "This memo file is Read-Only!", vbExclamation
        Exit Sub
        End If
    If frmMain.txtBox.Enabled = False Then Exit Sub
    If gblnLock = False Then
        gblnLock = True
        Else
        gblnLock = False
        End If
        CheckAttrib
Case "alarm"
    If gblnReadOnly = True Then
        MsgBox "This memo file is Read-Only!", vbExclamation
        Exit Sub
        End If
    If gblnLock = True Then
        MsgBox "This memo is locked and Read-Only!", vbExclamation, " Memo"
        Exit Sub
        End If
    If frmMain.txtBox.Enabled = False Then Exit Sub
    If gblnAlarm = False Then
        gblnAlarm = True
        Else
        gblnAlarm = False
        End If
        CheckAttrib
End Select
CheckButtons
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
If gcontinueFind = True Then Me.Combo1.SetFocus: Exit Sub
If frmMain.txtBox.Enabled = False Then Exit Sub
StoreCurrentMemo
gblnEditNewMemo = False
Select Case Button.Key
Case "first"
    curPos = 0
    Call FindMemo(Up)
Case "previous"
    Call FindMemo(Down)
Case "next"
    Call FindMemo(Up)
Case "last"
    curPos = Len(gstrMemo)
    Call FindMemo(Down)
Case "find"
    ToggleFind
End Select
CheckButtons
If frmMain.txtBox.Enabled = False Then Exit Sub
If Not gblnSearchBusy Then
    frmMain.txtBox.SetFocus
    Else
    frmMain.Combo1.SetFocus
    End If
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
If gcontinueFind = True Then KeyCode = 0: Exit Sub
Select Case KeyCode
Case 114 'f3
    StoreCurrentMemo
    Call FindOn
    KeyCode = 0
Case 27 'esc
    KeyCode = 0
    FindOff
Case 46 ' del
    If gblnReadOnly = True Then
        MsgBox "This memo file is Read-Only!", vbExclamation
        KeyCode = 0
        Exit Sub
        End If
    If gblnLock = True Then
    MsgBox "This memo is locked and Read-Only!", vbExclamation, " Memo"
    KeyCode = 0
    End If
End Select
If Shift <> 2 Then Exit Sub
Select Case KeyCode
Case 36  'HOME
    StoreCurrentMemo
    curPos = 0
    Call FindMemo(Up)
Case 34  'PDOWN
    StoreCurrentMemo
    Call FindMemo(Down)
Case 33  'PUP
    StoreCurrentMemo
    Call FindMemo(Up)
Case 35  'END
    StoreCurrentMemo
    curPos = Len(gstrMemo)
    Call FindMemo(Down)
End Select
Call CheckButtons
End Sub

Private Sub txtBox_GotFocus()
If gcontinueFind = True And Me.Combo1.Enabled = True Then Me.Combo1.SetFocus
End Sub

Private Sub txtBox_KeyPress(KeyAscii As Integer)
Dim x As Integer
If gcontinueFind = True And KeyAscii = 27 Then
    'exit search
    FindOff
    KeyAscii = 0
    gcontinueFind = False
    Exit Sub
    End If
Select Case KeyAscii
Case 1 ' ctrl+A
    frmMain.txtBox.SelStart = 0
    frmMain.txtBox.SelLength = Len(frmMain.txtBox)
    KeyAscii = 0
Case 3 'ctrl+c copy
    'dummy
Case 4 'ctrl+d delete
    KeyAscii = 0
    If gblnEditNewMemo = True Then
        gblnEditNewMemo = False
        curPos = Len(gstrMemo)
        Call FindMemo(Down)
    Else
        DeleteCurrentMemo
    End If
Case 12 'ctrl+l (lock)
    If gblnLock = False Then
        gblnLock = True
        Else
        gblnLock = False
        End If
    CheckAttrib
    KeyAscii = 0
Case 14 'ctrl+n nieuw
    KeyAscii = 0
    StoreCurrentMemo
    If gblnSearchBusy Then FindOff
    Call AddNewMemo
Case 16 'ctrl+p print
    frmPage.Show vbModal
    KeyAscii = 0
Case 19 'ctrl+s save
    If gblnEditNewMemo = True Then
        Call StoreCurrentMemo
        Call FindMemo(Up)
        End If
    x = SaveMemoFile(gstrCurrentFile)
    KeyAscii = 0
Case 6 'ctrl+f find
    FindOn
    KeyAscii = 0
Case Else
    If gblnReadOnly = True Then
        MsgBox "This memo file is Read-Only!", vbExclamation
    ElseIf gblnLock = True And KeyAscii <> 1 Then
        MsgBox "This memo is locked and Read-Only!", vbExclamation
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim x As Integer
If gcontinueFind = True Then KeyAscii = 0: Exit Sub
Select Case KeyAscii
Case 4 'ctrl+d
    KeyAscii = 0
    If gblnEditNewMemo = True Then
        StoreCurrentMemo
        curPos = Len(gstrMemo)
        Call FindMemo(Down)
        End If
    DeleteCurrentMemo
Case 12 'ctrl+l (lock)
    If gblnLock = False Then
        gblnLock = True
        Else
        gblnLock = False
        End If
    CheckAttrib
    KeyAscii = 0
Case 16 'ctrl+p print
    frmPage.Show vbModal
    KeyAscii = 0
Case 19 'ctrl+s save
    If gblnEditNewMemo = True Then
        Call StoreCurrentMemo
        Call FindMemo(Up)
        End If
    x = SaveMemoFile(gstrCurrentFile)
    KeyAscii = 0
Case 21
    KeyAscii = 0
    Me.Combo1.SetFocus
Case 14 'ctrl+n (nieuw)
    KeyAscii = 0
    StoreCurrentMemo
    If gblnSearchBusy Then FindOff
    Call AddNewMemo
Case 15 'ctrl+o (open)
    KeyAscii = 0
    StoreCurrentMemo
    If gblnSearchBusy = True Then ToggleFind
    frmFile.Show (vbModal)
Case 6
    KeyAscii = 0: ToggleFind
End Select
End Sub

Private Sub Combo1_GotFocus()
Me.Combo1.SelStart = 0
Me.Combo1.SelLength = Len(Me.Combo1.Text)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 And gcontinueFind = True Then
    gcontinueFind = False
    KeyCode = 0
    Exit Sub
    End If
If gcontinueFind = True Then KeyCode = 0: Exit Sub
Select Case KeyCode
Case 27
    FindOff
    KeyCode = 0
Case 13, 114
    Call StoreCurrentMemo
    If gblnNewSearch = True Then curPos = 0
    Call FindMemo(Up)
    Me.Combo1.SelStart = 0
    Me.Combo1.SelLength = Len(Me.Combo1.Text)
    KeyCode = 0
End Select
If Shift <> 2 Then Exit Sub
Select Case KeyCode
Case 36  'HOME
    StoreCurrentMemo
    curPos = 0
    Call FindMemo(Up)
Case 34  'PDOWN
    StoreCurrentMemo
    Call FindMemo(Down)
Case 33  'PUP
    StoreCurrentMemo
    Call FindMemo(Up)
Case 35  'END
    StoreCurrentMemo
    curPos = Len(gstrMemo)
    Call FindMemo(Down)
End Select
End Sub

Private Sub txtBox_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call CheckButtons
End Sub

Private Sub txtDate_GotFocus()
If Me.txtBox.Enabled = True Then Me.txtBox.SetFocus
End Sub

Private Sub picPosBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim newPos As Long
Call FindOff
If Len(gstrMemo) < 20 Then
    If Me.txtBox.Enabled = True Then Me.txtBox.SetFocus
    Exit Sub
    End If
newPos = Int((Len(gstrMemo) / Me.picPosBar.Width) * x)
If newPos > curPos Then
    curPos = newPos
    Call FindMemo(Up)
    Else
    curPos = newPos
    Call FindMemo(Down)
    End If
If Me.txtBox.Enabled = True Then Me.txtBox.SetFocus
End Sub

Private Sub picPosBar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim newPos As Long
newPos = Int((Len(gstrMemo) / Me.picPosBar.Width) * x)
Me.picPosBar.ToolTipText = " " & Format(newPos, "###,###,###,###") & " "
End Sub

Private Sub picPointer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.picPointer.ToolTipText = " >> " & Format(curPos, "###,###,###,###") & " << "
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim x As Integer
Dim retval As Integer
'no exit during search
If gcontinueFind = True Then Cancel = True: Exit Sub
'set loc free and save changes
If gblnReadOnly = False Then
    'save memo if changed
    Call StoreCurrentMemo
    If gblnMemoChanged = True Then x = SaveMemoFile(gstrCurrentFile)
    If x <> 0 Then
        retval = MsgBox("The changes are not saved!" & vbCrLf & vbCrLf & "Do you want to return to the program and try to save later?", vbYesNo + vbExclamation, " Exit Memo")
        If retval = vbYes Then Cancel = True: Exit Sub
        End If
    'send loc
    x = SendLOC(gstrCurrentFile, Free)
    'If x <> 0 Then MsgBox "De schrijfstatus van " & StripFileName(gstrCurrentFile) & " kon niet verzonden worden naar het station." & vbCrLf & vbCrLf & "Mogelijk is de verbinding met het station verbroken waar het bestand is opgeslagen of is er een probleem met de netwerkverbinding. Indien de schrijfstatus niet is bijgewerkt kan het memo-bestand gemarkeerd blijven als 'open voor bewerken'. In dat geval dient uzelf dit bestand later opnieuw te openen of moet een andere gebruiker de schrijfstatus forceren.", vbExclamation, " Afsluiten Memo"
    End If
'save all settings
SaveWindowPos
If gstrCurrentFile <> "" Then
    SaveSetting App.EXEName, "CONFIG", "File", gstrCurrentFile
    End If
SaveSetting App.EXEName, "CONFIG", "TextCol", gintTextCol
SaveSetting App.EXEName, "CONFIG", "TextSize", gintTextSize
SaveSetting App.EXEName, "CONFIG", "TextFont", gintTextFont
Unload frmFile
Unload frmPage
Unload frmInfo
Unload frmSplash
Unload frmUser
End Sub

Private Sub TimerLOC_Timer()
Dim x As Integer
Dim tmp
Dim retval As String
Dim tmpLOC As String
If frmSplash.Visible = True Then frmSplash.Hide
If frmMain.Enabled = False Then Exit Sub
If gcontinueFind = True Then Exit Sub
tmpLOC = ReadLOC(gstrCurrentFile)
If tmpLOC = "ERROR" Then Exit Sub
If tmpLOC <> "" And Left(tmpLOC, 10) <> gstrUserId Then Exit Sub
retval = MsgBox(StripFileName(gstrCurrentFile) & " is no longer in use by others." & vbCrLf & vbCrLf & "Do you want to reopen the memo file for edit?", vbYesNo + vbQuestion + vbSystemModal, "  Memo - Write Request")
If retval = vbNo Then
    Me.TimerLOC.Enabled = False
    Me.mniWriteRequest.Checked = False
    Exit Sub
    End If
'open again
x = OpenMemoFile(gstrCurrentFile)
End Sub
