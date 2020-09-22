VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASP Generator"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2280
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<- Back"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next ->"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   4455
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1800
         TabIndex        =   38
         Top             =   3600
         Width           =   3375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   3960
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Options"
            Object.Width           =   7233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Page Name"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Reset All"
         Height          =   375
         Left            =   5280
         TabIndex        =   39
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Page Filename:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Output Directory:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   6120
         Picture         =   "Form1.frx":380A
         ToolTipText     =   "Select All"
         Top             =   290
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   6360
         Picture         =   "Form1.frx":3954
         ToolTipText     =   "Select None"
         Top             =   290
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Please choose the options you want to use:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generate Pages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   4455
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command6 
         Caption         =   "Click Here to Finish"
         Height          =   375
         Left            =   4560
         TabIndex        =   25
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Overwrite files without prompting."
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3480
         Width           =   4215
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   3840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label14 
         Caption         =   $"Form1.frx":3A9E
         Height          =   735
         Left            =   240
         TabIndex        =   41
         Top             =   1200
         Width           =   6255
      End
      Begin VB.Label Label13 
         Caption         =   $"Form1.frx":3B7F
         Height          =   735
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   4455
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cmbSize 
         Height          =   315
         ItemData        =   "Form1.frx":3C41
         Left            =   3840
         List            =   "Form1.frx":3C5A
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3600
         Width           =   2775
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Underline"
         Height          =   255
         Left            =   5400
         TabIndex        =   34
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italic"
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   4080
         Width           =   975
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   4080
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   1200
         Width           =   3735
      End
      Begin ASPGenerator.Pallet Pallet1 
         Height          =   2535
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4471
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Font Size:"
         Height          =   315
         Left            =   2880
         TabIndex        =   36
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Font Face:"
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label lblFontColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Font Folor:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "These options will allow you to define the font color and size of all information that is generated by this program."
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command4 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   3240
         Width           =   5175
      End
      Begin VB.Label Label16 
         Caption         =   $"Form1.frx":3C73
         Height          =   615
         Left            =   240
         TabIndex        =   43
         Top             =   1320
         Width           =   6255
      End
      Begin VB.Label Label15 
         Caption         =   "To begin, please choose the database you wish to use by clicking the ""Browse"" button."
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   2640
         Width           =   6375
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   240
         Picture         =   "Form1.frx":3D2D
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "ASP Generator 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   615
         Left            =   960
         TabIndex        =   21
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Database:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   3240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Table and Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   4455
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   5415
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Field"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Link Field:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   6360
         Picture         =   "Form1.frx":7537
         ToolTipText     =   "Select None"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   6120
         Picture         =   "Form1.frx":7681
         ToolTipText     =   "Select All"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Please choose the fields you want to use:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Table:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Isection As Integer
Dim bLoadPageName As Boolean
Private Sub Combo1_Click()
LoadColTypes
FieldsToCombo Combo2, Combo1.Text, True
If Combo2.ListCount > 0 Then
    Combo2.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
Isection = Isection + 1
DoValidation
NavSection
End Sub

Private Sub Command2_Click()
Isection = Isection - 1
NavSection
End Sub

Private Sub Command3_Click()
Select Case Isection
Case Is = 0
    ShowHelp 10
Case Is = 1
    ShowHelp 20
Case Is = 2
    ShowHelp 30
Case Is = 3
    ShowHelp 40
Case Is = 4
    ShowHelp 50
End Select
End Sub

Private Sub Command4_Click()
On Error GoTo EH
With dlgCommonDialog
.FileName = ""
.Filter = "Access Database (*.mdb)|*.mdb"
.FilterIndex = 1
.Flags = cdlOFNFileMustExist + cdlOFNExplorer
.DefaultExt = ".mdb"
.CancelError = True
.InitDir = App.Path
.ShowOpen
End With

Text1.Text = dlgCommonDialog.FileName
SaveSetting App.Title, "Database", "Filename", Text1.Text
Exit Sub
EH:
If Err <> cdlCancel Then
ShowError Err.Number, Err.Description, "Selecting File"
End If
Exit Sub
End Sub

Private Sub Command5_Click()
Text2.Text = GetFolderName
SaveSetting App.Title, "Database", "Output", Text2.Text

End Sub

Private Sub Command6_Click()
DoRun
End Sub

Private Sub Command7_Click()
Dim I As Integer
For I = 1 To ListView1(1).ListItems.Count
    ListView1(1).ListItems(I).SubItems(1) = ReturnFilename(ListView1(1).ListItems(I).Text, I, True)
Next I
End Sub

Private Sub Form_Load()
LoadOptions
Text1.Text = GetDatabaseFileName
Text2.Text = GetSetting(App.Title, "Database", "Output", "")
Isection = 0
ProgressBar1.Value = 0
NavSection
FontsToList List1
SelectDefaultFont List1
lblFontColor.BackColor = CLng(GetSetting(App.Title, "Main", "FontColor", &H0))
cmbSize.ListIndex = CInt(GetSetting(App.Title, "Main", "FontSize", 1))
chkBold.Value = CInt(GetSetting(App.Title, "Main", "FontBold", 0))
chkItalic.Value = CInt(GetSetting(App.Title, "Main", "FontItalic", 0))
chkUnderline.Value = CInt(GetSetting(App.Title, "Main", "FontUnderline", 0))
End Sub


Public Sub NavSection()
Dim I As Integer
If Isection < 0 Then Isection = 0
If Isection > Frame1.UBound Then Isection = Frame1.UBound
For I = Frame1.LBound To Frame1.UBound
If I = Isection Then
    Frame1(I).Visible = True
Else
    Frame1(I).Visible = False
End If
Next I
Command1.Enabled = True
Command2.Enabled = True
If Isection = Frame1.UBound Then Command1.Enabled = False
If Isection = Frame1.LBound Then Command2.Enabled = False
End Sub

Private Sub DoValidation()
ProgressBar1.Value = ProgressBar1.Min
Dim I As Integer
Select Case Isection
Case Is = 0
    'do nothing
Case Is = 1 'going from welcome to tables and fields
    LoadDatabaseVals
Case Is = 2
Case Is = 3
If Text2.Text = "" Or Dir(Text2.Text, vbDirectory) = "" Then
    ShowInfo "The output directory you have chosen is not valid!", "User Error"
End If
    For I = 1 To ListView1(1).ListItems.Count
        SaveSetting App.Title, "Options", "File" & I, ListView1(1).ListItems(I).SubItems(1)
    Next I
Case Is = 4
    SaveSetting App.Title, "Main", "FontFace", List1.Text
    SaveSetting App.Title, "Main", "FontColor", lblFontColor.BackColor
    SaveSetting App.Title, "Main", "FontSize", cmbSize.ListIndex
    SaveSetting App.Title, "Main", "FontBold", chkBold.Value
    SaveSetting App.Title, "Main", "FontItalic", chkItalic.Value
    SaveSetting App.Title, "Main", "FontUnderline", chkUnderline.Value
    
End Select
End Sub

Private Sub LoadDatabaseVals()
On Error GoTo EH
Dim bOK As Boolean
SaveSetting App.Title, "Database", "Filename", Text1.Text
Combo1.Clear
Combo2.Clear
ListView1(0).ListItems.Clear
bOK = ConnectToAccess
If bOK = True Then
    TablesToCombo Combo1
    If Combo1.ListCount = 0 Then
        ShowInfo "No Tables found in database!", "File Error"
    Else
        Combo1.ListIndex = 0
    End If
Else
    ShowInfo "Could not connect to database!", "Error"
End If
Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Loading Database Values"
Exit Sub
End Sub

Private Sub DoJobMarking(LST As ListView, bSel As Boolean, bInvert As Boolean)
Dim I As Integer
For I = 1 To LST.ListItems.Count
    If bInvert = True Then
        LST.ListItems(I).Checked = Not LST.ListItems(I).Checked
    Else
        LST.ListItems(I).Checked = bSel
    End If
Next I
End Sub

Private Sub LoadColTypes()
On Error GoTo EH
Dim nCount As Integer
Dim typeString As String
Dim I As Integer
Dim LST As ListItem
 Dim db As Connection
 Dim RS As Recordset
 
ListView1(0).ListItems.Clear

Set db = New Connection
Set RS = New Recordset

Set RS = OpenDB(Combo1.Text, sConnectionString, db)
    For I = 0 To RS.Fields.Count - 1
    Set LST = ListView1(0).ListItems.Add(, , RS.Fields(I).Name)
    LST.SubItems(1) = ConvType(RS.Fields(I).Type)
    LST.SubItems(2) = RS.Fields(I).DefinedSize
    LST.Checked = True
    Next I
Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Loading Field Types"
Exit Sub
End Sub

Private Sub Image1_Click()
DoJobMarking ListView1(0), True, False
End Sub

Private Sub Image2_Click()
DoJobMarking ListView1(0), False, False
End Sub

Private Sub Image3_Click()
DoJobMarking ListView1(1), False, False
End Sub

Private Sub Image4_Click()
DoJobMarking ListView1(1), True, False
End Sub

Private Sub LoadOptions()
Dim I As Integer

With ListView1(1).ListItems
.Add , , "Generate ASP Setup File (setup.asp)."
.Add , , "Generate ADO Constants File (msado.asp)."
.Add , , "Generate Record Navigation Page (recnav.asp)."
.Add , , "Generate Record Display Page (recdisplay.asp)."
.Add , , "Generate Database Variables File (database.asp)."
.Add , , "Allow Add Record (recadd.asp)."
.Add , , "Allow Edit Record (recedit.asp)."
.Add , , "Allow Delete Record (recdelete.asp)."
End With

For I = 1 To ListView1(1).ListItems.Count
    ListView1(1).ListItems(I).SubItems(1) = ReturnFilename(ListView1(1).ListItems(I).Text, I)
Next I
ListView1(1).ListItems(1).Selected = True
Text3.Text = ListView1(1).ListItems(1).SubItems(1)
DoJobMarking ListView1(1), True, False

End Sub

Private Function ReturnFilename(sTXT As String, I As Integer, Optional bReset As Boolean) As String
'Generate ASP Setup File (setup.asp). returns setup.asp
Dim Ipos As Long
Dim Epos As Long
Dim STMP As String
If bReset = False Then
STMP = GetSetting(App.Title, "Options", "File" & I, "")

If STMP <> "" Then
    ReturnFilename = STMP
    Exit Function
End If
End If
Ipos = InStr(1, sTXT, "(")
Epos = InStr(1, sTXT, ")")

If Ipos <> 0 And Epos <> 0 Then
ReturnFilename = Mid(sTXT, Ipos + 1, Epos - Ipos - 1)
Else
    ReturnFilename = ""
End If
End Function


Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
bLoadPageName = True
Text3.Text = Item.SubItems(1)
bLoadPageName = False
End Sub

Private Sub Pallet1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblFontColor.BackColor = Pallet1.CurColors
End Sub

Private Sub Text3_Change()
If bLoadPageName = True Then Exit Sub
If ListView1(1).SelectedItem Is Nothing Then Exit Sub
ListView1(1).SelectedItem.SubItems(1) = Text3.Text
End Sub
