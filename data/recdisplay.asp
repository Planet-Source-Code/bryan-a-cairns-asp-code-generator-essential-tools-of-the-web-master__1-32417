#includedatabase#
<%
sWhat = trim(Request("ID"))
sPageAction = trim(Request("Action"))
#pagevariables#

#writehead#
    LoadVars
    ShowEditor
#writefoot#

Sub ShowEditor
%>
<B>View a record</B><BR>
<% #gonav# %>
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0">
#readtable#
</TABLE>
<%
End Sub

Sub LoadVars
opencon
ChooseTable "#tablename#"
if rs.bof = true and rs.eof = true then
    Response.Write("Record not found!")
else
    do until rs.eof = true
        if lcase(cstr(rs.fields("#id#").value)) = lcase(cstr(sWhat)) then
            #loadvars#
            exit do
        end if
    rs.movenext
    Loop
end if
closecon
End Sub

%>