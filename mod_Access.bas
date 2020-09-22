Attribute VB_Name = "mod_Access"

Global sDataBaseName As String
Global bDataOpened As Boolean
Global sConnectionString As String

Public Enum AccessFieldType
field_Bit = 0
field_BYTE = 1
field_Counter = 2
field_CURRENCY = 3
field_DateTime = 4
field_SINGLE = 5
field_DOUBLE = 6
field_Short = 7
field_LONG = 8
field_LongText = 9
field_LongBinary = 10
field_Text = 11
End Enum
    


Public Function ConnectToAccess() As Boolean
On Error GoTo EH
Dim sFile As String
sFile = GetDatabaseFileName

If CheckFile(sFile) = False Then
    ConnectToAccess = False
    Exit Function
End If

sConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFile & ";"
ConnectToAccess = True
Exit Function
EH:
ConnectToAccess = False
sConnectionString = ""
    ShowError Err.Number, Err.Description, "Connecting to Access"
Exit Function
End Function

Public Function GetDatabaseFileName() As String
    GetDatabaseFileName = GetSetting(App.Title, "Database", "Filename", "")
End Function

Public Function AllowDelete() As Boolean
Dim RET
'returns the inverse to all a delete
RET = MsgBox("Are you sure you wish to delete the current record?", vbYesNo, "Confirm Delete")
If RET = vbYes Then
AllowDelete = False
Else
AllowDelete = True
End If
End Function

Public Sub SortDBGrid(sFRm As Form, grdObject As Object, adoPrimaryRS As Recordset, ColIndex As Integer)
    Dim strColName As String
    Dim STMP As String
    Dim Ipos As Long
    Static bSortAsc As Boolean
    Static strPrevCol As String
On Error GoTo EH
    strColName = grdObject.Columns(ColIndex).DataField
    If strColName = strPrevCol Then


        If bSortAsc Then
            adoPrimaryRS.Sort = strColName & " DESC"
            bSortAsc = False
        Else
            adoPrimaryRS.Sort = strColName
            bSortAsc = True
        End If
    Else
        adoPrimaryRS.Sort = strColName
        bSortAsc = True
    End If
    
    strPrevCol = strColName

    Ipos = InStr(1, LCase(adoPrimaryRS.Source), "order by", vbBinaryCompare)
    If Ipos <> 0 Then
        STMP = Trim(Mid(adoPrimaryRS.Source, 1, Ipos - 1))
    End If
    STMP = STMP & " Order By " & adoPrimaryRS.Sort
    sFRm.CustomSQL = STMP
Exit Sub
EH:
MsgBox Err.Description, vbCritical, "Generating Sort Order"
Exit Sub
End Sub

Public Sub AddField(sConnect As String, sTable As String, sName As String, sType As AccessFieldType)
On Error GoTo EH
Dim db As Connection
Dim sqlStatement As String
Dim sTypeSTR As String

Select Case sType
Case Is = 0
sTypeSTR = "Bit"
Case Is = 1
sTypeSTR = "BYTE"
Case Is = 2
sTypeSTR = "Counter"
Case Is = 3
sTypeSTR = "CURRENCY"
Case Is = 4
sTypeSTR = "DateTime"
Case Is = 5
sTypeSTR = "SINGLE"
Case Is = 6
sTypeSTR = "DOUBLE"
Case Is = 7
sTypeSTR = "Short"
Case Is = 8
sTypeSTR = "LONG"
Case Is = 9
sTypeSTR = "LongText"
Case Is = 10
sTypeSTR = "LongBinary"
Case Is = 11
sTypeSTR = "Text"
End Select
Set db = New Connection
db.Open sConnect
    sqlStatement = "ALTER TABLE " & sTable & " ADD [" & sName & "] " & sTypeSTR
    Call db.Execute(sqlStatement)
db.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Adding Field"
Exit Sub
End Sub

Public Sub DeleteField(sConnect As String, sTable As String, sName As String)
On Error GoTo EH
Dim db As Connection
Dim sqlStatement As String

Set db = New Connection
db.Open sConnect
    sqlStatement = "ALTER TABLE " & sTable & " DROP COLUMN [" & sName & "]"
    Call db.Execute(sqlStatement)
db.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Deleting Field"
Exit Sub
End Sub

Public Sub AddTable(sConnect As String, sName As String)
On Error GoTo EH
Dim db As Connection
Dim sqlStatement As String
Set db = New Connection
db.Open sConnect
    sqlStatement = "CREATE TABLE " & ResolveTable(sName)
    Call db.Execute(sqlStatement)
db.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Adding Table"
Exit Sub
End Sub

Public Sub DeleteTable(sConnect As String, sName As String)
On Error GoTo EH
Dim db As Connection
Dim sqlStatement As String
Set db = New Connection
db.Open sConnect
    sqlStatement = "DROP TABLE " & ResolveTable(sName)
    Call db.Execute(sqlStatement)
db.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Deleting Table"
Exit Sub
End Sub

Public Sub BackupTable(sConnect As String, sName As String, sPrefix As String)
On Error GoTo EH
Dim db As Connection
Dim sqlStatement As String

If DoesTableExist(sConnect, sPrefix & sName) = True Then
'remove the old backup table
    DeleteTable sConnect, sPrefix & sName
End If

Set db = New Connection
db.Open sConnect
    sqlStatement = "SELECT " & ResolveTable(sName) & ".* INTO " & ResolveTable(sPrefix & sName) & " FROM " & ResolveTable(sName)
    Call db.Execute(sqlStatement)
db.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Backup Table"
Exit Sub
End Sub

Public Sub RenameTable(sConnect As String, sFrom As String, sTo As String)
On Error GoTo EH
Dim db As Connection
Dim sqlStatement As String
Set db = New Connection
db.Open sConnect
    sqlStatement = "SELECT " & ResolveTable(sFrom) & ".* INTO " & ResolveTable(sTo) & " FROM " & ResolveTable(sFrom)
    Call db.Execute(sqlStatement)
    sqlStatement = "DROP TABLE " & ResolveTable(sFrom)
    Call db.Execute(sqlStatement)
db.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Rename Table"
Exit Sub
End Sub

Public Function ResolveTable(inputTable As String) As String
    ResolveTable = IIf(InStr(1, inputTable, " ") <> 0 Or IsNumeric(Left(inputTable, 1)), "[" & inputTable & "]", inputTable)
End Function

Public Sub TablesToCombo(CMB As Object)
 Dim db As Connection
 Dim RS As Recordset

Set db = New Connection
Set RS = New Recordset

db.Open sConnectionString
Set RS = db.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
        If Not RS Is Nothing Then
            Do While Not RS.EOF
                If UCase(Left(RS!Table_name, 4)) <> "MSYS" Then
                    If UCase(Left(RS!Table_name, 11)) <> "SWITCHBOARD" Then
                        newtablename = RS!Table_name
                        If newtablename <> "" Then
                        CMB.AddItem newtablename
                        End If
                    End If
                End If
                RS.MoveNext
            Loop
            'CMB.AddItem DEF_CUSTOM_SQL
        End If
CloseDB RS, db
End Sub

Public Function DoesTableExist(sConnection As String, sName As String) As Boolean
On Error GoTo EH
 Dim db As Connection
 Dim RS As Recordset
 Dim bFound As Boolean
Set db = New Connection
Set RS = New Recordset

db.Open sConnection
Set RS = db.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
        If Not RS Is Nothing Then
            Do While Not RS.EOF
                If UCase(Left(RS!Table_name, 4)) <> "MSYS" Then
                    If UCase(Left(RS!Table_name, 11)) <> "SWITCHBOARD" Then
                        newtablename = RS!Table_name
                        If newtablename <> "" Then
                            If LCase(newtablename) = LCase(sName) Then
                                bFound = True
                                Exit Do
                            End If
                        End If
                    End If
                End If
                RS.MoveNext
            Loop
        End If
CloseDB RS, db
DoesTableExist = bFound
Exit Function
EH:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Does Table Exist"
Exit Function
End Function

Public Function DoesFieldExist(RS As Recordset, sName As String) As Boolean
Dim bFound As Boolean
On Error GoTo EH

For I = 0 To RS.Fields.Count - 1
If LCase(RS.Fields(I).Name) = LCase(sName) Then
    bFound = True
    Exit For
End If
Next I
DoesFieldExist = bFound
Exit Function
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Does Field Exist"
Exit Function
End Function

Public Sub FieldsToListview(CMB As Object, sTable As String, Optional bChecked As Boolean)
 Dim db As Connection
 Dim RS As Recordset
 Dim I As Integer
 Dim h As Integer
 Dim bFound As Boolean
 Dim LST As Object
On Error GoTo EH
Set db = New Connection
Set RS = New Recordset

CMB.ListItems.Clear
'add the parts list
Set RS = OpenDB(sTable, sConnectionString, db)
For I = 0 To RS.Fields.Count - 1
Set LST = CMB.ListItems.Add(, , RS.Fields(I).Name)
LST.Checked = bChecked
Next I

CloseDB RS, db
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Fields To Combo"
Exit Sub
End Sub

Public Sub FieldsToCombo(CMB As Object, sTable As String, Optional bNoDuplicates As Boolean)
 Dim db As Connection
 Dim RS As Recordset
 Dim I As Integer
 Dim h As Integer
 Dim bFound As Boolean
On Error GoTo EH
Set db = New Connection
Set RS = New Recordset

CMB.Clear
'add the parts list
Set RS = OpenDB(sTable, sConnectionString, db)
For I = 0 To RS.Fields.Count - 1
bFound = False
If bNoDuplicates = True Then
    For h = 0 To CMB.ListCount - 1
        If LCase(RS.Fields(I).Name) = LCase(CMB.List(h)) Then
            bFound = True
            Exit For
        End If
    Next h
End If
If bFound = False Then
    CMB.AddItem RS.Fields(I).Name
End If
Next I

CloseDB RS, db
If CMB.ListCount > 0 Then
CMB.ListIndex = 0
End If
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Fields To Combo"
Exit Sub
End Sub


Public Sub ClearTable(sConnect As String, sName As String)
On Error GoTo EH
Dim db As Connection
Dim sqlStatement As String


Set db = New Connection
db.Open sConnect
    sqlStatement = "DELETE * FROM " & ResolveTable(sName)
    Call db.Execute(sqlStatement)
db.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Clearing Table"
Exit Sub
End Sub

Public Sub ClearTableEX(sTable As String)
'UnAllocates parts from the current job
 Dim db As Connection
 Dim RS As Recordset
 Dim LS As ListItem
 Dim bFound As Boolean
 Dim I As Long
On Error GoTo EH
Set db = New Connection
Set RS = New Recordset

Set RS = OpenDB(sTable, sConnectionString, db)
If isRecordSetEmpty(RS) = False Then
RS.MoveLast
Do Until RS.BOF = True
    RS.Delete
    RS.UpdateBatch adAffectAllChapters
    RS.MovePrevious
Loop
If GetRecordCount(RS) > 0 Then
RS.Delete
RS.UpdateBatch adAffectAllChapters
End If
End If
CloseDB RS, db
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Clearing Table"
Exit Sub
End Sub

Public Function MakeProgress(cCurrent As Long, cTotal As Long) As Integer
'makes 0 to 100 on file progress
On Error GoTo EH
Dim iProgress As Integer
iProgress = CInt((cCurrent / cTotal) * 100)
MakeProgress = iProgress
Exit Function
EH:
MakeProgress = 0
Exit Function
End Function

Public Sub RecToCombo(CM As Object, sTable As String, sField As String, Optional bNoDuplicates As Boolean)
 Dim db As Connection
 Dim RS As Recordset
 Dim h As Integer
 Dim bFound As Boolean
Set db = New Connection
Set RS = New Recordset
CM.Clear
Set RS = OpenDB(sTable, sConnectionString, db)
    If isRecordSetEmpty(RS) = False Then
    Do Until RS.EOF = True
    DoEvents
    ''''''''
    bFound = False
        If bNoDuplicates = True Then
            For h = 0 To CM.ListCount - 1
                If LCase(RS.Fields(sField).Value) = LCase(CM.List(h)) Then
                    bFound = True
                    Exit For
                End If
            Next h
        End If
If bFound = False Then
If RS.Fields(sField).Value <> vbNull Then
    CM.AddItem RS.Fields(sField).Value
End If
End If
    '''''''''

    RS.MoveNext
    Loop
    End If
CloseDB RS, db
End Sub

Public Function IsMainDBOpen(bShowWarning As Boolean) As Boolean
If bZoneView = False Then
If bShowWarning = True Then
MsgBox "Please open a Zone!", vbInformation, "Can not continue"
End If
IsMainDBOpen = False
Else
IsMainDBOpen = True
End If
End Function

Public Function OpenDB(sTable As String, sConnection As String, Con As ADODB.Connection) As ADODB.Recordset
'Opens a connection to a database
    Dim sConnect As String
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    sSQL = "select * from " & sTable

    Con.Open sConnection
    RS.LockType = adLockBatchOptimistic
    RS.Open sSQL, Con
    Set OpenDB = RS
    'Set RS = Nothing
End Function

Public Sub CloseDB(RS As ADODB.Recordset, Con As ADODB.Connection)
'closes a connection to a database
On Error Resume Next
RS.Close
Con.Close
End Sub


Public Function GetRecordCount(RS As ADODB.Recordset) As Long
'returns the total number of records in the record set
Dim L As Long
If isRecordSetEmpty(RS) = True Then
GetRecordCount = 0
Exit Function
End If
Screen.MousePointer = 11
Do While Not RS.EOF
L = L + 1
DoEvents
RS.MoveNext
Loop
RS.MoveFirst
Screen.MousePointer = 0
GetRecordCount = L
End Function

Public Function isRecordSetEmpty(RS As ADODB.Recordset) As Boolean
'Returns True if recordset is empty
If RS.BOF = True And RS.EOF = True Then
    isRecordSetEmpty = True
Else
    isRecordSetEmpty = False
End If
End Function

Public Function SelectRecords(sConnection As String, Con As ADODB.Connection, sTable As String, Optional sField As String, Optional sWhat As String) As ADODB.Recordset
'Selects a specific record set
Dim RS As ADODB.Recordset
Dim strSQLChange As String

If sWhat <> "" And sWhat <> "*" Then
    strSQLChange = "Select * FROM " & sTable & " WHERE " & sField & " = '" & sWhat & "'"
Else
    strSQLChange = "Select * from " & sTable
End If

    Set RS = New ADODB.Recordset
    RS.LockType = adLockBatchOptimistic
    RS.CursorType = adOpenDynamic
    Con.Open sConnection
    RS.Open strSQLChange, Con, adOpenDynamic, adLockOptimistic
Set SelectRecords = RS
End Function

Public Function SelectLikeRecords(sConnection As String, Con As ADODB.Connection, sTable As String, Optional sField As String, Optional sWhat As String) As ADODB.Recordset
'Selects a specific record set
Dim RS As ADODB.Recordset
Dim strSQLChange As String

If sWhat <> "" And sWhat <> "*" Then
    strSQLChange = "Select * from " & sTable & " Where " & sField & " LIKE '%" & sWhat & "%'"
Else
    strSQLChange = "Select * from " & sTable
End If

    Set RS = New ADODB.Recordset
    RS.LockType = adLockBatchOptimistic
    RS.CursorType = adOpenDynamic
    Con.Open sConnection
    RS.Open strSQLChange, Con, adOpenDynamic, adLockOptimistic
Set SelectLikeRecords = RS
End Function

Public Function SelectContainingRecords(sConnection As String, Con As ADODB.Connection, sTable As String, Optional sField As String, Optional sWhat As String) As ADODB.Recordset
'Selects a specific record set
Dim RS As ADODB.Recordset
Dim strSQLChange As String

If sWhat <> "" And sWhat <> "*" Then
    strSQLChange = "Select * from " & sTable & " Where " & sField & " LIKE '%" & sWhat & "%'"
Else
    strSQLChange = "Select * from " & sTable
End If

    Set RS = New ADODB.Recordset
    RS.LockType = adLockBatchOptimistic
    RS.CursorType = adOpenDynamic
    Con.Open sConnection
    RS.Open strSQLChange, Con, adOpenDynamic, adLockOptimistic
Set SelectContainingRecords = RS
End Function

Public Function SelectRecordsByFeild(sConnection As String, Con As ADODB.Connection, sTable As String, sField As String, bAscending As Boolean) As ADODB.Recordset
'Selects a records by a field value (sorts them)
Dim RS As ADODB.Recordset
Dim strSQLChange As String
Dim sSorttype As String

If bAscending = True Then
    sSorttype = "asc"
Else
    sSorttype = "desc"
End If

    strSQLChange = "Select " & sField & " from " & sTable & " order by " & sField & " " & sSorttype


    Set RS = New ADODB.Recordset
    RS.LockType = adLockBatchOptimistic
    RS.CursorType = adOpenDynamic
    Con.Open sConnection
    RS.Open strSQLChange, Con, adOpenDynamic, adLockOptimistic
Set SelectRecordsByFeild = RS
End Function

Public Sub MassUpdate(sTable As String, sConnection As String, sField As String, sWhat As String)
Dim Con As ADODB.Connection
Dim RS As ADODB.Recordset
Dim lTotal As Long
Dim lCurrent As Long
Dim lDeleted As Long
Screen.MousePointer = 11
Set Con = New ADODB.Connection
Set RS = OpenDB(sTable, sConnection, Con)
If isRecordSetEmpty(RS) = False Then
    Con.Execute "Update " & sTable & " set " & sField & " = '" & sWhat & "'"
End If
CloseDB RS, Con
Screen.MousePointer = 0
End Sub

Public Function ConvType(ByVal TypeVal As Long) As String
  Select Case TypeVal
        Case adBigInt                    ' 20
            ConvType = "Big Integer"
        Case adBinary                    ' 128
            ConvType = "Binary"
        Case adBoolean                   ' 11
            ConvType = "Boolean"
        Case adBSTR                      ' 8 i.e. null terminated string
            ConvType = "Text"
        Case adChar                      ' 129
            ConvType = "Text"
        Case adCurrency                  ' 6
            ConvType = "Currency"
        Case adDate                      ' 7
            ConvType = "Date/Time"
        Case adDBDate                    ' 133
            ConvType = "Date/Time"
        Case adDBTime                    ' 134
            ConvType = "Date/Time"
        Case adDBTimeStamp               ' 135
            ConvType = "Date/Time"
        Case adDecimal                   ' 14
            ConvType = "Float"
        Case adDouble                    ' 5
            ConvType = "Float"
        Case adEmpty                     ' 0
            ConvType = "Empty"
        Case adError                     ' 10
            ConvType = "Error"
        Case adGUID                      ' 72
            ConvType = "GUID"
        Case adIDispatch                 ' 9
            ConvType = "IDispatch"
        Case adInteger                   ' 3
            ConvType = "Integer"
        Case adIUnknown                  ' 13
            ConvType = "Unknown"
        Case adLongVarBinary             ' 205
            ConvType = "Binary"
        Case adLongVarChar               ' 201
            ConvType = "Text"
        Case adLongVarWChar              ' 203
            ConvType = "Memo"
        Case adNumeric                  ' 131
            ConvType = "Long"
        Case adSingle                    ' 4
            ConvType = "Single"
        Case adSmallInt                  ' 2
            ConvType = "Small Integer"
        Case adTinyInt                   ' 16
            ConvType = "Tiny Integer"
        Case adUnsignedBigInt            ' 21
            ConvType = "Big Integer"
        Case adUnsignedInt               ' 19
            ConvType = "Integer"
        Case adUnsignedSmallInt          ' 18
            ConvType = "Small Integer"
        Case adUnsignedTinyInt           ' 17
            ConvType = "Tiny Integer"
        Case adUserDefined               ' 132
            ConvType = "UserDefined"
        Case adVarNumeric                 ' 139
            ConvType = "Long"
        Case adVarBinary                 ' 204
            ConvType = "Binary"
        Case adVarChar                   ' 200
            ConvType = "Text"
        Case adVariant                   ' 12
            ConvType = "Variant"
        Case adVarWChar                  ' 202
            ConvType = "Text"
        Case adWChar                     ' 130
            ConvType = "Text"
        Case Else
            ConvType = "Unknown"
   End Select
End Function

