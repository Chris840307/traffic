<!-- #include file="./DciFileOpen.asp"-->
<!--#include virtual="traffic/Common/DCIURL.ini"-->
<HTML>
<HEAD>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
<%
Function getHTTPPage(url)
dim http
set http=Server.createobject("Microsoft.XMLHTTP")
Http.open "GET",url,false
Http.send()
if Http.readystate<>4 then
exit function
end if
bytesToBSTR Http.responseBody,"Big5"
set http=nothing
if err.number<>0 then err.Clear
End function

getHTTPPage(linkURL&trim(request("DCIfile"))) '要抓取的網址

%>
</BODY>
</HTML>
