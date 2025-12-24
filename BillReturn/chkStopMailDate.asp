<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt

EffectDate=gOutDT(request("EffectDate"))
MailDate=""
If Not ifnull(EffectDate) Then
	MailDate=DateAdd("m",3,EffectDate)
End if

%>


