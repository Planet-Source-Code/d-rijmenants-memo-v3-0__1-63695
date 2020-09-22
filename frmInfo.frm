VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Memo Info"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   125
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Version 3.00.02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4665
   End
   Begin VB.Label Label1 
      Caption         =   "© RDDS 2003"
      ForeColor       =   &H00000000&
      Height          =   1425
      Index           =   2
      Left            =   105
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MEMO organizer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Me.Label1(2).Caption = "This program is freeware and can be used and distributed under following restrictions: It is forbidden to use this program, copies or parts of it for commercial purposes, sell, lease or make profit of this program by any means." & vbCrLf & vbCrLf & "© Dirk Rijmenants 2003"
End Sub
