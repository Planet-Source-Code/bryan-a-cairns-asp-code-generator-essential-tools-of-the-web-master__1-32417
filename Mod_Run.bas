Attribute VB_Name = "Mod_Run"
 
Dim sPageVariables As String '# = variables read from the page post
Dim sLoadVars As String '# = page variables loaded from database
Dim sSaveVars As String  '# = page variables saved to current recordset
Dim sDatafile As String '# = database filename
Dim sIncludeDatabase As String '# = include database tag
Dim sIncludeSetup As String '# = include setup file tag
Dim sIncludeMSADO As String '# = include ado file tag
Dim sWriteHead As String '# = write header code
Dim sWriteFoot As String '# = write footer code
Dim sWriteTable As String '# = write table contents code
Dim sReadTable As String '# = readtable# = read table contents code
Dim sTableName As String '# = the current table name
Dim sAdd As String '# = add record
Dim sID As String '# = the record ID tag
Dim sNavTable As String '# = navtable# = naviagtion table

Dim sFontStart As String
Dim sFontStop As String

Public Function isFieldUsed(sTXT As String) As Boolean
On Error GoTo EH
Dim I As Integer
Dim bOK As Boolean

For I = 1 To Form1.ListView1(0).ListItems.Count
    If Form1.ListView1(0).ListItems(I).Text = sTXT Then
        bOK = Form1.ListView1(0).ListItems(I).Checked
    End If
Next I
isFieldUsed = bOK
Exit Function
EH:
    ShowError Err.Number, Err.Description, "Is Field In Use"
Exit Function
End Function

Public Sub DoRun()
'This is where all the files are created
'check for errors
'setup.asp 20
'msado.asp 30
'database.asp 40
'recnav.asp 50
'recdisplay.asp 60
'recadd.asp 70
'recedit.asp 80
'recdelete.asp 90
If CheckProject = False Then Exit Sub
GenerateTags
Form1.ProgressBar1.Value = 10

If CopyLocalToRoot("setup.asp", Form1.ListView1(1).ListItems(1).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(1).SubItems(1)
End If
Form1.ProgressBar1.Value = 20

If CopyLocalToRoot("msado.asp", Form1.ListView1(1).ListItems(2).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(2).SubItems(1)
End If
Form1.ProgressBar1.Value = 30

If CopyLocalToRoot("database.asp", Form1.ListView1(1).ListItems(5).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(5).SubItems(1)
End If
Form1.ProgressBar1.Value = 40

If CopyLocalToRoot("recnav.asp", Form1.ListView1(1).ListItems(3).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(3).SubItems(1)
End If
Form1.ProgressBar1.Value = 50

If CopyLocalToRoot("recdisplay.asp", Form1.ListView1(1).ListItems(4).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(4).SubItems(1)
End If
Form1.ProgressBar1.Value = 60

If CopyLocalToRoot("recadd.asp", Form1.ListView1(1).ListItems(6).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(6).SubItems(1)
End If
Form1.ProgressBar1.Value = 70

If CopyLocalToRoot("recedit.asp", Form1.ListView1(1).ListItems(7).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(7).SubItems(1)
End If
Form1.ProgressBar1.Value = 80

If CopyLocalToRoot("recdelete.asp", Form1.ListView1(1).ListItems(8).SubItems(1)) = True Then
    ReplaceAllPageVars Form1.ListView1(1).ListItems(8).SubItems(1)
End If
Form1.ProgressBar1.Value = 100
ClearVars
End Sub


Public Sub ReplaceAllPageVars(sFile As String)
Dim STMP As String
If CheckFile(Form1.Text2.Text & sFile) = False Then Exit Sub
STMP = OpenTextFile(Form1.Text2.Text & sFile)
    STMP = Replace(STMP, "#pagevariables#", sPageVariables)
    STMP = Replace(STMP, "#loadvars#", sLoadVars)
    STMP = Replace(STMP, "#savevars#", sSaveVars)
    STMP = Replace(STMP, "#datafile#", sDatafile)
    STMP = Replace(STMP, "#includedatabase#", sIncludeDatabase)
    STMP = Replace(STMP, "#includesetup#", sIncludeSetup)
    STMP = Replace(STMP, "#includemsado#", sIncludeMSADO)
    STMP = Replace(STMP, "#writehead#", sWriteHead)
    STMP = Replace(STMP, "#writefoot#", sWriteFoot)
    STMP = Replace(STMP, "#writetable#", sWriteTable)
    STMP = Replace(STMP, "#tablename#", sTableName)
    STMP = Replace(STMP, "#readtable#", sReadTable)
    STMP = Replace(STMP, "#navtable#", sNavTable)
    STMP = Replace(STMP, "#add#", sAdd)
    STMP = Replace(STMP, "#id#", sID)
    
STMP = Replace(STMP, "#pagesetup#", Form1.ListView1(1).ListItems(1).SubItems(1))
STMP = Replace(STMP, "#pagemsado#", Form1.ListView1(1).ListItems(2).SubItems(1))
STMP = Replace(STMP, "#pagerecnav#", Form1.ListView1(1).ListItems(3).SubItems(1))
STMP = Replace(STMP, "#pagerecdisplay#", Form1.ListView1(1).ListItems(4).SubItems(1))
STMP = Replace(STMP, "#pagedatabase#", Form1.ListView1(1).ListItems(5).SubItems(1))
STMP = Replace(STMP, "#pagerecadd#", Form1.ListView1(1).ListItems(6).SubItems(1))
STMP = Replace(STMP, "#pagerecedit#", Form1.ListView1(1).ListItems(7).SubItems(1))
STMP = Replace(STMP, "#pagerecdelete#", Form1.ListView1(1).ListItems(8).SubItems(1))

If Form1.ListView1(1).ListItems(1).Checked = True Then
    STMP = Replace(STMP, "#gonav#", "ShowGoBack")
Else
    STMP = Replace(STMP, "#gonav#", "")
End If

WriteTextFile Form1.Text2.Text & sFile, STMP
End Sub

Public Sub ClearVars()
sPageVariables = ""
sLoadVars = ""
sSaveVars = ""
sWriteTable = ""
sReadTable = ""
sDatafile = ""
sIncludeDatabase = ""
sIncludeSetup = ""
sWriteHead = ""
sWriteFoot = ""
sTableName = ""
sAdd = ""
sID = ""
sNavTable = ""
sIncludeMSADO = ""
sFontStart = ""
sFontStop = ""

End Sub

Public Sub GenerateTags()
 Dim db As Connection
 Dim RS As Recordset
 Dim I As Integer
Set db = New Connection
Set RS = New Recordset
ClearVars

sFontStart = "<FONT FACE=" & Chr(34) & Form1.List1.Text & Chr(34)
sFontStart = sFontStart & " SIZE=" & Chr(34) & Form1.cmbSize.Text & Chr(34)
sFontStart = sFontStart & " COLOR=" & Chr(34) & MakeRGBHex(Form1.lblFontColor.BackColor) & Chr(34) & ">"
sFontStop = "</FONT>"

If Form1.chkBold.Value = 1 Then
sFontStart = sFontStart & "<B>"
sFontStop = "</B>" & sFontStop
End If
If Form1.chkItalic.Value = 1 Then
sFontStart = sFontStart & "<I>"
sFontStop = "</I>" & sFontStop
End If
If Form1.chkUnderline.Value = 1 Then
sFontStart = sFontStart & "<U>"
sFontStop = "</U>" & sFontStop
End If
Set RS = OpenDB(Form1.Combo1.Text, sConnectionString, db)

If Form1.ListView1(1).ListItems(6).Checked = True Then
    sAdd = "<A HREF=" & Chr(34) & "#pagerecadd#" & Chr(34) & ">Add New</A><BR>" & vbCrLf
End If

sID = Form1.Combo2.Text
sTableName = Form1.Combo1.Text

If Form1.ListView1(1).ListItems(6).Checked = True Then
    sAdd = "<A HREF=" & Chr(34) & "#pagerecadd#" & Chr(34) & "><B>Add a new Record</B></A>"
End If

If Form1.ListView1(1).ListItems(2).Checked = True Then
    sIncludeMSADO = "<!-- #include file = " & Chr(34) & "#pagemsado#" & Chr(34) & " -->"
End If

If Form1.ListView1(1).ListItems(5).Checked = True Then
    sIncludeDatabase = "<!-- #include file = " & Chr(34) & "#pagedatabase#" & Chr(34) & " -->"
End If

If Form1.ListView1(1).ListItems(1).Checked = True Then
    sIncludeSetup = "<!-- #include file = " & Chr(34) & "#pagesetup#" & Chr(34) & " -->"
End If

If Form1.ListView1(1).ListItems(1).Checked = True Then
    sWriteHead = "WriteHeader"
    sWriteFoot = "WriteFooter"
End If

sDatafile = Form1.Text1.Text 'ParsePath(Form1.Text1.Text, 2) & ParsePath(Form1.Text1.Text, 3)

'make the page variables
For I = 0 To RS.Fields.Count - 1
If isFieldUsed(RS.Fields(I).Name) = True Then
    sPageVariables = sPageVariables & "s" & RS.Fields(I).Name & " = Trim(Request.Form(" & Chr(34) & RS.Fields(I).Name & Chr(34) & "))" & vbCrLf
End If
Next I

'make the save variables
For I = 0 To RS.Fields.Count - 1
If isFieldUsed(RS.Fields(I).Name) = True Then
    sSaveVars = sSaveVars & "Rs.Fields(" & Chr(34) & RS.Fields(I).Name & Chr(34) & ").value = s" & RS.Fields(I).Name & vbCrLf
End If
Next I

'make the load variables
For I = 0 To RS.Fields.Count - 1
If isFieldUsed(RS.Fields(I).Name) = True Then
    sLoadVars = sLoadVars & "s" & RS.Fields(I).Name & " = Rs.Fields(" & Chr(34) & RS.Fields(I).Name & Chr(34) & ").value" & vbCrLf
End If
Next I

sReadTable = MakeReadTable(RS)
sWriteTable = MakeWriteTable(RS)
sNavTable = MakeNavTable(RS)



CloseDB RS, db
End Sub

Public Function MakeNavTable(RS As Recordset) As String
On Error GoTo EH
Dim STMP As String
Dim I As Integer
Dim bAdd As Boolean
Dim bEdit As Boolean
Dim bDelete As Boolean
Dim bView As Boolean

If RS Is Nothing Then
    MakeNavTable = ""
    Exit Function
End If

bAdd = Form1.ListView1(1).ListItems(6).Checked
bEdit = Form1.ListView1(1).ListItems(7).Checked
bDelete = Form1.ListView1(1).ListItems(8).Checked
bView = Form1.ListView1(1).ListItems(4).Checked
    STMP = STMP & "<TR>" & vbCrLf
        For I = 0 To RS.Fields.Count - 1
        If isFieldUsed(RS.Fields(I).Name) = True Then
            STMP = STMP & "<TD>" & sFontStart & "<% Response.WRite(RS.Fields(" & Chr(34) & RS.Fields(I).Name & Chr(34) & ").value) %>" & sFontStop & "</TD>" & vbCrLf
        End If
        Next I
        
                    'make the view field
                    If bView = True Then
                        STMP = STMP & "<TD>" & sFontStart & "<A HREF=" & Chr(34) & "#pagerecdisplay#?ID=<% Response.Write(rs.fields(" & Chr(34) & Form1.Combo2.Text & Chr(34) & ").value) %>" & Chr(34) & ">View</A>" & sFontStop & "</TD>" & vbCrLf
                    End If
                    
                    'make the edit field
                    If bEdit = True Then
                        STMP = STMP & "<TD>" & sFontStart & "<A HREF=" & Chr(34) & "#pagerecedit#?ID=<% Response.Write(rs.fields(" & Chr(34) & Form1.Combo2.Text & Chr(34) & ").value) %>" & Chr(34) & ">Edit</A>" & sFontStop & "</TD>" & vbCrLf
                    End If
                    
                    'make the delete field
                    If bDelete = True Then
                        STMP = STMP & "<TD>" & sFontStart & "<A HREF=" & Chr(34) & "#pagerecdelete#?ID=<% Response.Write(rs.fields(" & Chr(34) & Form1.Combo2.Text & Chr(34) & ").value) %>" & Chr(34) & ">Delete</A>" & sFontStop & "</TD>" & vbCrLf
                    End If
                    

    STMP = STMP & "</TR>" & vbCrLf

MakeNavTable = STMP
Exit Function
EH:
MakeNavTable = ""
    ShowError Err.Number, Err.Description, "Making Navigation Table"
Exit Function
End Function

Public Function MakeWriteTable(RS As Recordset) As String
On Error GoTo EH
Dim STMP As String
Dim I As Integer


If RS Is Nothing Then
    MakeWriteTable = ""
    Exit Function
End If


'build the col
For I = 0 To RS.Fields.Count - 1
If isFieldUsed(RS.Fields(I).Name) = True Then



Select Case LCase(ConvType(RS.Fields(I).Type))
Case Is = "boolean"
    STMP = STMP & "<TR>" & vbCrLf
    STMP = STMP & "<TD>" & sFontStart & RS.Fields(I).Name & sFontStop & "</TD>" & vbCrLf
    STMP = STMP & "<TD>"
    STMP = STMP & "<SELECT NAME=" & Chr(34) & RS.Fields(I).Name & Chr(34) & " SIZE=" & Chr(34) & "1" & Chr(34) & ">" & vbCrLf
    STMP = STMP & "<OPTION SELECTED><% Response.WRite(" & "s" & RS.Fields(I).Name & ") %>" & vbCrLf
    STMP = STMP & "<OPTION>True" & vbCrLf
    STMP = STMP & "<OPTION>False" & vbCrLf
    STMP = STMP & "</SELECT>" & vbCrLf
    STMP = STMP & "</TD>" & vbCrLf
    STMP = STMP & "</TR>" & vbCrLf
Case Is = "memo"
    STMP = STMP & "<TR>" & vbCrLf
    STMP = STMP & "<TD>" & sFontStart & RS.Fields(I).Name & sFontStop & "</TD>" & vbCrLf
    STMP = STMP & "<TD>"
    STMP = STMP & "<TEXTAREA NAME=" & Chr(34) & RS.Fields(I).Name & Chr(34) & " ROWS=" & Chr(34) & "6" & Chr(34) & " COLS=" & Chr(34) & "40" & Chr(34) & "><% Response.WRite(" & "s" & RS.Fields(I).Name & ") %></TEXTAREA>"
    STMP = STMP & "</TD>" & vbCrLf
    STMP = STMP & "</TR>" & vbCrLf
Case Else
    STMP = STMP & "<TR>" & vbCrLf
    STMP = STMP & "<TD>" & sFontStart & RS.Fields(I).Name & sFontStop & "</TD>" & vbCrLf
    STMP = STMP & "<TD>"
    STMP = STMP & "<INPUT TYPE=" & Chr(34) & "text" & Chr(34) & " NAME=" & Chr(34) & RS.Fields(I).Name & Chr(34) & " SIZE=" & Chr(34) & "30" & Chr(34) & " VALUE=" & Chr(34) & "<% Response.WRite(" & "s" & RS.Fields(I).Name & ") %>" & Chr(34) & ">"
    STMP = STMP & "</TD>" & vbCrLf
    STMP = STMP & "</TR>" & vbCrLf
End Select
End If
Next I
MakeWriteTable = STMP
Exit Function
EH:
MakeWriteTable = ""
    ShowError Err.Number, Err.Description, "Making Write-Only Table"
Exit Function
End Function

Public Function MakeReadTable(RS As Recordset) As String
On Error GoTo EH
Dim STMP As String
Dim I As Integer


If RS Is Nothing Then
    MakeReadTable = ""
    Exit Function
End If


'build the col
For I = 0 To RS.Fields.Count - 1
If isFieldUsed(RS.Fields(I).Name) = True Then

'determine what type of field it is and how to process it
Select Case LCase(ConvType(RS.Fields(I).Type))
Case Is = "memo"
    STMP = STMP & "<TR>" & vbCrLf
    STMP = STMP & "<TD>" & sFontStart & RS.Fields(I).Name & sFontStop & "</TD>" & vbCrLf
    STMP = STMP & "<TD>" & sFontStart & "<PRE><% Response.WRite(" & "s" & RS.Fields(I).Name & ") %></PRE>" & sFontStop & "</TD>" & vbCrLf
    STMP = STMP & "</TR>" & vbCrLf
Case Else
    STMP = STMP & "<TR>" & vbCrLf
    STMP = STMP & "<TD>" & sFontStart & RS.Fields(I).Name & sFontStop & "</TD>" & vbCrLf
    STMP = STMP & "<TD>" & sFontStart & "<% Response.WRite(" & "s" & RS.Fields(I).Name & ") %>" & sFontStop & "</TD>" & vbCrLf
    STMP = STMP & "</TR>" & vbCrLf
End Select
End If
Next I
MakeReadTable = STMP
Exit Function
EH:
MakeReadTable = ""
    ShowError Err.Number, Err.Description, "Making Read-Only Table"
Exit Function
End Function

Public Function AllowFileOverWrite(sSource As String, sDest As String) As Boolean
Dim RET

If Form1.Check1.Value = 1 Then
    AllowFileOverWrite = True
    Exit Function
End If
If CheckFile(sDest) = True Then
    RET = MsgBox("The following file already exists," & vbCrLf & sDest & vbCrLf & "Do you wish to overwrite it?", vbYesNo, "Confirm File Overwrite")
    If RET <> vbYes Then
            AllowFileOverWrite = False
        Exit Function
    End If
End If
AllowFileOverWrite = True
End Function

Public Function CopyLocalToRoot(sFile As String, sDest As String) As Boolean
On Error GoTo EH

If CheckFile(App.Path & "\data\" & sFile) = False Then
    ShowInfo "Could not find " & sFile & ".asp, please re-install!", "File Error"
    CopyLocalToRoot = False
    Exit Function
End If

If AllowFileOverWrite(App.Path & "\data\" & sFile, Form1.Text2.Text & sDest) = False Then
    CopyLocalToRoot = False
    Exit Function
End If
FileCopy App.Path & "\data\" & sFile, Form1.Text2.Text & sDest
CopyLocalToRoot = True
Exit Function
EH:
    ShowError Err.Number, Err.Description, "Creating " & sFile & " File"
Exit Function
End Function


Public Function CheckProject() As Boolean
Dim I As Integer
Dim bFound As Boolean
If CheckFile(Form1.Text1.Text) = False Then
    ShowInfo "Not a valid database", "Configuration Error"
    Form1.Isection = 0
    Form1.NavSection
    CheckProject = False
    Exit Function
End If

If Dir(Form1.Text2.Text, vbDirectory) = "" Then
    ShowInfo "Can not find directory!", "Configuration Error"
    Form1.Isection = 2
    Form1.NavSection
    CheckProject = False
    Exit Function
End If

If Right(Form1.Text2.Text, 1) <> "\" Then Form1.Text2.Text = Form1.Text2.Text & "\"

If Form1.Combo1.ListCount = 0 Then
    ShowInfo "No tables in database", "Configuration Error"
    Form1.Isection = 1
    Form1.NavSection
    CheckProject = False
    Exit Function
End If

If Form1.Combo2.ListCount = 0 Then
    ShowInfo "No fields in table", "Configuration Error"
    Form1.Isection = 1
    Form1.NavSection
    CheckProject = False
    Exit Function
End If

bFound = False
For I = 1 To Form1.ListView1(0).ListItems.Count
If Form1.ListView1(0).ListItems(I).Checked = True Then
    bFound = True
    Exit For
End If
Next I
If bFound = False Then
    ShowInfo "No fields in use", "Configuration Error"
    Form1.Isection = 1
    Form1.NavSection
    CheckProject = False
    Exit Function
End If

bFound = False
For I = 1 To Form1.ListView1(1).ListItems.Count
If Form1.ListView1(1).ListItems(I).Checked = True Then
    bFound = True
    Exit For
End If
Next I
If bFound = False Then
    ShowInfo "Options in use", "Configuration Error"
    Form1.Isection = 2
    Form1.NavSection
    CheckProject = False
    Exit Function
End If
CheckProject = True
End Function
