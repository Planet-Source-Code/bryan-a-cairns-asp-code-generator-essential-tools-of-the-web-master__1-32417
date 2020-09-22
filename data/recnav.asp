#includedatabase#
<%
#writehead#

Dim currentPage, rowCount, i
currentPage = TRIM(Request("currentPage"))
if currentPage = "" then currentPage = 1
opencon
ChoosePages "#tablename#", currentPage,49
rowCount = 0
DoCount currentPage
%>
#add#
<BR>
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" WIDTH="100%" BGCOLOR="#FFFFFF" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0">
<%
do while not rs.eof
if rowCount = rs.PageSize then exit DO
%>
#navtable#
<% 
rowCount = rowCount + 1
rs.movenext
loop
%>
</TABLE>
<BR>
<%
Response.Write("Total on this page:" & rowCount & "<BR>")
closecon

#writefoot#

Sub DoCount(currentPage) 
h = 0
for i = 1 to rs.PageCount
 Response.Write(" <a href=" & chr(34) & "#pagerecnav#?currentpage=" &  i  & chr(34) &  ">" & i & "</a>")
h = h +1
next
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
end sub
%>