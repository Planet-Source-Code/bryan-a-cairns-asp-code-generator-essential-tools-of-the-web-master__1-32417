#includedatabase#
<%
sWhat = trim(Request("ID"))
sPageAction = trim(Request("Action"))

#writehead#

if lcase(sPageAction) = "" then
    ShowWarn
end if
if lcase(sPageAction) = "delete" then
    DeleteRec
    ShowOK
end if
if lcase(sPageAction) = "cancel" then
    ShowCancel
end if

#writefoot#

Sub ShowWarn
%>
<B>Are you sure you wish to delete this record?</B><BR>
<A HREF="#pagerecdelete#?action=delete&id=<% Response.WRite(sWhat) %>">Yes</A>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A HREF="#pagerecdelete#?action=cancel&id=<% Response.WRite(sWhat) %>">No</A>
<%
End Sub

Sub DeleteRec
opencon
ChooseTable "#tablename#"
if rs.bof = true and rs.eof = true then
    Response.Write("Record not found!")
else
    do until rs.eof = true
        if lcase(cstr(rs.fields("#id#").value)) = lcase(cstr(sWhat)) then
            rs.delete
            rs.UpdateBatch adAffectAll
            exit do
        end if
    rs.movenext
    Loop
end if
closecon
End Sub

Sub ShowOK
%>
<B>The record was deleted.</B><BR>
<% #gonav# %>
<%
end Sub

Sub ShowCancel
%>
<B>The record was not deleted.</B><BR>
<% #gonav# %>
<%
end Sub
%>