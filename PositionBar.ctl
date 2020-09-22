VERSION 5.00
Begin VB.UserControl PositionBar 
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   ScaleHeight     =   2760
   ScaleWidth      =   5850
   ToolboxBitmap   =   "PositionBar.ctx":0000
   Begin VB.PictureBox picBar 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.PictureBox picPos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         ScaleHeight     =   225
         ScaleWidth      =   75
         TabIndex        =   1
         Top             =   0
         Width           =   100
      End
   End
End
Attribute VB_Name = "PositionBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private barScale As Single
Private barVal As Long
Private barMax As Long
Private barBack As Long
Private barFront As Long
Private barPoint As Long

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "barmax", barMax, 100
PropBag.WriteProperty "barval", barVal, 0
PropBag.WriteProperty "barback", barBack, &H8000000F
PropBag.WriteProperty "barfront", barFront, &HFF0000
PropBag.WriteProperty "barpoint", barVal, barPoint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
barMax = PropBag.ReadProperty("barmax", 100)
barVal = PropBag.ReadProperty("barval", 0)
barBack = PropBag.ReadProperty("barback", &H8000000F)
barFront = PropBag.ReadProperty("barfront", &HFF0000)
barPoint = PropBag.ReadProperty("barpoint", 100)
End Sub
Private Sub UserControl_Initialize()
'
End Sub

Private Sub UserControl_Resize()
picBar.Top = 10
picBar.Left = 10
picBar.Width = UserControl.Width - 10
picBar.Height = UserControl.Height - 10
picPos.Top = -20
picPos.Height = picBar.Height + 20
End Sub

Public Property Get Value() As Long
Attribute Value.VB_Description = "Sets or return position of the pointer"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
Value = barVal
End Property

Public Property Let Value(lVal As Long)
barVal = Val(lVal)
If barMax < 1 Then Exit Property
barScale = (picBar.Width - picPos.Width - 50) / barMax
picPos.Left = barVal * barScale
End Property

Public Property Get Max() As Long
Max = barMax
End Property

Public Property Let Max(lVal As Long)
Attribute Max.VB_Description = "Sets highest PositionBar value"
Attribute Max.VB_ProcData.VB_Invoke_PropertyPut = ";Misc"
barMax = Val(lVal)
End Property

Public Property Get BackColor() As Variant
Attribute BackColor.VB_Description = "Sets or returns the color of the Positionbar"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_MemberFlags = "200"
BackColor = barBack
End Property

Public Property Let BackColor(ByVal vVal As Variant)
barBack = Val(vVal)
picBar.BackColor = barBack
End Property

Public Property Get ForeColor() As Variant
Attribute ForeColor.VB_Description = "Sets or returns the color of the Pointer"
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
ForeColor = barFront
End Property

Public Property Let ForeColor(ByVal vVal As Variant)
barFront = Val(vVal)
picPos.BackColor = barFront
End Property

Public Property Get PointerWidth() As Variant
Attribute PointerWidth.VB_Description = "Sets or returns the pointer width"
PointerWidth = barPoint
End Property

Public Property Let PointerWidth(ByVal vVal As Variant)
If Val(vVal) < 1 Or Val(vVal) > picBar.Width Then Exit Property
barPoint = Val(vVal)
picPos.Width = barPoint
If barMax < 1 Then Exit Property
barScale = (picBar.Width - picPos.Width - 50) / barMax
picPos.Left = barVal * barScale
End Property
