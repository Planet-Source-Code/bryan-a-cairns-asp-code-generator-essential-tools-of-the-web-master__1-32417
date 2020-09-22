#includesetup#
#includemsado#

<%
DataConnPath = "#datafile#"
Dim rs
Dim connectme
Dim sqlstmt
Dim conn

Sub OpenCon
'Uses a DSNless connection
'get DataConnPath from setup.asp
Set Conn = Server.CreateObject("ADODB.Connection")
connectme = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DataConnPath & ";"
Conn.open(connectme)
End Sub

Sub CloseCon
rs.Close
set rs = Nothing
set conn = Nothing
End Sub


Sub ChooseTable(sTable)
set rs=Server.CreateObject("ADODB.recordset")
sqlstmt = "SELECT * FROM " & stable
rs.open sqlstmt, connectme, adOpenStatic, adLockOptimistic
End Sub

Sub ChooseRecord(sTable, sField, sWhat)
set rs=Server.CreateObject("ADODB.recordset")
sqlstmt = "SELECT * FROM " & stable & " WHERE " & sfield & " = " & "'" & sWhat & "'"
rs.open sqlstmt, connectme, adOpenStatic, adLockOptimistic
End Sub

Sub ChooseTableSort(sTable,sSortBy)
set rs=Server.CreateObject("ADODB.recordset")
sqlstmt = "SELECT * FROM " & stable & " order by " & sSortBy
rs.open sqlstmt, connectme, adOpenStatic, adLockOptimistic
End Sub

Sub ChooseRecordSort(sTable, sField, sWhat,sSortBy)
set rs=Server.CreateObject("ADODB.recordset")
sqlstmt = "SELECT * FROM " & stable & " WHERE " & sfield & " = " & "'" & sWhat & "'" & " order by " & sSortBy
rs.open sqlstmt, connectme, adOpenStatic, adLockOptimistic
End Sub

Sub ChoosePages(sTable,sStart,sSize)
set rs=Server.CreateObject("ADODB.recordset")
sqlstmt = "SELECT * FROM " & stable
rs.CursorType = 3
rs.PageSize = cint(sSize)
rs.open sqlstmt, connectme
rs.AbsolutePage = cINT(sStart)
End Sub

Sub FormatAsMoney(sVal)
stmp = cstr(sval)
if instr(1,stmp,".") = 0 then
stmp = stmp & ".00"
end if
if instr(1,stmp,"$") = 0 then
stmp = "$"  & stmp
end if
Response.write(stmp)
end sub

Function parseoneline (GetTheStr)
'parses normal text and returns HTML formatted'
	Start = 1
	whereis = 1
	NewHTMLStringWithBreaks = ""
	RemaindingHTMLString = GetTheStr		
	RemaindingHTMLString = RemaindingHTMLString & chr(13)
	do until (whereis = 0)
		Whereis = InStr (1, RemaindingHTMLString, Chr(13))
		Chr13Position = Whereis - 1	
		LineOfHTMLBeforeNextChr13 = Mid(RemaindingHTMLString, 1 , Whereis)
		Chr13Position = Whereis + 2
		LineOfHTMLAfterNextChr13 = Mid(RemaindingHTMLString, Chr13Position, Len(RemaindingHTMLString))
		RemaindingHTMLString = LineOfHTMLAfterNextChr13
		NewHTMLStringWithBreaks = NewHTMLStringWithBreaks & LineOfHTMLBeforeNextChr13 & "<br>"
	loop
	Response.Write(NewHTMLStringWithBreaks)
End Function
%>