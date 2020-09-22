#includedatabase# 
<%
sWhat = trim(Request("ID"))
sPageAction = trim(Request("Action"))
#pagevariables#

#writehead#

if lcase(sPageAction) = "" then
    ShowEditor
end if

if lcase(sPageAction) = "save" then
    SaveVars
    ShowOk
end if

#writefoot#

Sub ShowEditor
%>
<B>Add a new record</B><BR>
<% #gonav# %>
<FORM METHOD="POST" ACTION="#pagerecadd#?action=save">
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0">
#writetable#
</TABLE>
<BR><BR>
<INPUT TYPE="submit" VALUE="Save"><INPUT TYPE="reset" Value="Undo"><BR>
</FORM>
<%
End Sub

Sub SaveVars
opencon
ChooseTable "#tablename#"
rs.addnew
#savevars#
rs.UpdateBatch adAffectAll
closecon
End Sub

Sub ShowOK
%>
<B>Your information has been saved.</B><BR>
<% #gonav# %>
<%
end Sub
%>