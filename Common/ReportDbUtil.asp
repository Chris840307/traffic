<!--#include file="db.ini" -->
<%
If Not IsObject(Conn) Then
   Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.open connStr
End If

Function ShowPageLink(curpage,allpage,URL,param)
	Dim spg,epg
		spg = (curpage \ PageSize) * PageSize + 1
		epg = spg + (PageSize-1)
	
	if allpage=1 then exit function
	
	if epg > allpage then epg=allpage

	if curpage = 1 then
		Response.Write "<img src='../Image/PREVPAGE.GIF' border=0>"
	else 
		Response.Write "<A href='" & URL & "?page=" & curpage-1 & param & "'><img src='../Image/PREVPAGE.GIF' border=0></A>"
  end if

  Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & curpage & " / " & allpage & "&nbsp;&nbsp;&nbsp;&nbsp;"

	if curpage = allpage then
    Response.Write "<img src='../Image/NextPage.gif' border=0>"
  else 
		Response.Write "<A href='" & URL & "?page=" & curpage+1 & param & "'><img src='../Image/NextPage.gif' border=0></A>"
  end if
End Function
%>