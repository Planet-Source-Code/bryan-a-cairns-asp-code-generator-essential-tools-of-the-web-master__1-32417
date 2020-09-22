VERSION 5.00
Object = "{8C99FA6F-84D4-11D2-B300-444553540001}#7.0#0"; "ColPal.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ColPal.Pallet Pallet1 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4471
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Pallet1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = Label2.BackColor

End Sub

Private Sub Pallet1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = Pallet1.CurColors
Label2.BackColor = Pallet1.CurColors
Label4.Caption = Pallet1.WebHex

End Sub
