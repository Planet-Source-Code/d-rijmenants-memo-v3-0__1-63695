VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   6210
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Press a key or click the mouse to continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3900
      Width           =   5370
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4080
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   5610
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
Me.Hide
End Sub

Private Sub Form_Click()
Me.Hide
End Sub

Private Sub Label1_Click()
Me.Hide
End Sub

Private Sub Label2_Click()
Me.Hide
End Sub
