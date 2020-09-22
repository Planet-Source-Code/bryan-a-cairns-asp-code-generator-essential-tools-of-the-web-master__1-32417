VERSION 5.00
Begin VB.UserControl Pallet 
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ScaleHeight     =   2820
   ScaleWidth      =   3135
   ToolboxBitmap   =   "Pallet.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      MouseIcon       =   "Pallet.ctx":0312
      MousePointer    =   99  'Custom
      Picture         =   "Pallet.ctx":061C
      ScaleHeight     =   2505
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image imgWhite 
      Height          =   480
      Left            =   2640
      Picture         =   "Pallet.ctx":15E40
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgBlack 
      Height          =   480
      Left            =   2760
      Picture         =   "Pallet.ctx":1614A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Pallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Default Property Values:
Const m_def_SelColor = 0
Const m_def_CurColors = 0
Const m_def_WebHex = 0
Dim m_SelColor As Variant
Dim m_CurColors As Variant
Dim m_WebHex As Variant
Dim m_CurHex As Long

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo EH
RaiseEvent MouseMove(Button, Shift, X, Y)
If Y < 720 Then Picture1.MouseIcon = imgBlack.Picture

If Y > 720 Then Picture1.MouseIcon = imgWhite.Picture
m_CurColors = Picture1.Point(X, Y)
m_CurHex = Picture1.Point(X, Y)
GETRGB m_CurHex
Exit Sub
EH:
Exit Sub
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.width = Picture1.width
UserControl.height = Picture1.height
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo EH
RaiseEvent MouseDown(Button, Shift, X, Y)
m_SelColor = Picture1.Point(X, Y)
Exit Sub
EH:
Exit Sub
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo EH
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Exit Sub
EH:
Exit Sub
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    m_SelColor = PropBag.ReadProperty("SelColor", m_def_SelColor)
    m_CurColors = PropBag.ReadProperty("CurColors", m_def_CurColors)
    m_WebHex = PropBag.ReadProperty("WebHex", m_def_WebHex)
End Sub
Private Sub GETRGB(stColor As Long)
On Error GoTo EH
'stColor = m_CurHex
       '     'If r > 255 Then Exit Sub
       '     'If g > 255 Then Exit Sub
       '     'If b > 255 Then Exit Sub
       Dim r, b, g As Long
       
       Dim dts As Variant
       Dim q, w, e As Variant
       Dim qw, we, gq As Variant
       Dim lCol As Long
       lCol = stColor
       r = lCol Mod &H100
       lCol = lCol \ &H100
       g = lCol Mod &H100
       lCol = lCol \ &H100
       b = lCol Mod &H100
       
       '     'Get Red Hex
       q = Hex(r)

              If Len(q) < 2 Then
                     qw = q
                     q = "0" & qw
              End If

       '     'Get Blue Hex
       w = Hex(b)

              If Len(w) < 2 Then
                     we = w
                     w = "0" & we
              End If

       '     'Get Green Hex
       e = Hex(g)

              If Len(e) < 2 Then
                     gq = e
                     e = "0" & gq
              End If

       'GETRGB = "#" & q & e & w
       m_WebHex = q & e & w   '"#" &
Exit Sub
EH:
m_WebHex = q & e & w   '"#" &
Exit Sub
End Sub
Public Property Get SelColor() As Variant
Attribute SelColor.VB_MemberFlags = "400"
 On Error Resume Next
   SelColor = m_SelColor
End Property

Public Property Let SelColor(ByVal New_SelColor As Variant)
 On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 382
    m_SelColor = New_SelColor
    PropertyChanged "SelColor"
End Property

Public Property Get CurColors() As Variant
Attribute CurColors.VB_MemberFlags = "400"
 On Error Resume Next
    CurColors = m_CurColors
End Property

Public Property Let CurColors(ByVal New_CurColors As Variant)
On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 382
    m_CurColors = New_CurColors
    PropertyChanged "CurColors"
End Property

Public Property Get WebHex() As Variant
Attribute WebHex.VB_MemberFlags = "400"
On Error Resume Next
    WebHex = m_WebHex
End Property

Public Property Let WebHex(ByVal New_WebHex As Variant)
On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 382
    m_WebHex = New_WebHex
    PropertyChanged "WebHex"
End Property

Public Property Get About() As Object
On Error Resume Next
   Set About = frmAbout
End Property
Public Property Set About(ByVal SAbout As Object)
On Error Resume Next
    Set frmAbout = SAbout
End Property
Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_Description = "About this OCX"
Attribute ShowAboutBox.VB_UserMemId = -552
On Error Resume Next
   dlgAbout.Show vbModal
    Unload dlgAbout
    Set dlgAbout = Nothing
 End Sub
