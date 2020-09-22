#includedatabase#
<%
sWhat = trim(Request("ID"))
sPageAction = trim(Request("Action"))
#pagevariables#

#writehead#

if lcase(sPageAction) = "" then
    LoadVars
    ShowEditor
end if
if lcase(sPageAction) = "save" then
    SaveVars
    ShowOk
end if

#writefoot#

Sub ShowEditor
%>
<B>Edit a record</B><BR>
<% #gonav# %>
<FORM METHOD="POST" ACTION="#pagerecedit#?action=save">
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0">
#writetable#
</TABLE>
<BR><BR>
<INPUT TYPE="submit" VALUE="Save"><INPUT TYPE="reset" Value="Undo"><BR>
</FORM>
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

Sub SaveVars
opencon
ChooseTable "#tablename#"
if rs.bof = true and rs.eof = true then
    Response.Write("Record not found!")
else
    do until rs.eof = true
        if lcase(cstr(rs.fields("#id#").value)) = lcase(cstr(sWhat)) then
            #savevars#
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
<B>Your information has been saved.</B><BR>
<% #gonav# %>
<%
end Sub
%>