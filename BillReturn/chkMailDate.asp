<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

EffectDate=gOutDT(request("EffectDate"))
MailDate=""
If Not ifnull(EffectDate) Then
	MailDate=DateAdd("m",3,EffectDate)
End if
response.Write "myForm.MailDate[myForm.chkcnt.value-1].value='"&gInitDT(MailDate)&"';"
%>
	myForm.item[myForm.chkcnt.value].focus();
	myForm.item[myForm.chkcnt.value].select();

