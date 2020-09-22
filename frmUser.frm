VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Memo User"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   1995
      TabIndex        =   1
      Top             =   2310
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Top             =   2310
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   105
      TabIndex        =   3
      Top             =   0
      Width           =   4530
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   210
         MaxLength       =   20
         TabIndex        =   0
         Top             =   1575
         Width           =   4110
      End
      Begin VB.Label Label2 
         Caption         =   "Programma ID"
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   210
         TabIndex        =   5
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the name of the user. This name is used to identify the user during writing status."
         Height          =   540
         Left            =   210
         TabIndex        =   4
         Top             =   1050
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.lblID.Caption = gstrUserId
Me.txtUser.Text = gstrUserName
Me.txtUser.SetFocus
Me.txtUser.SelStart = 0
Me.txtUser.SelLength = Len(Me.txtUser.Text)
End Sub


Private Sub cmdOK_Click()
gstrUserName = Me.txtUser.Text
If Len(gstrUserName) > 20 Then gstrUserName = Left(gstrUserName, 20)
SaveSetting App.EXEName, "CONFIG", "UserName", gstrUserName
Me.Hide
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Call cmdOK_Click
    End If
End Sub
