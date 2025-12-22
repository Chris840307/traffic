<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
BillStartNumber = trim(Request("BillStartNumber"))
BillEndNumber = trim(Request("BillEndNumber"))

for i=len(BillStartNumber) to 1 step -1
	if not IsNumeric(mid(BillStartNumber,i,1)) then
		Sno=MID(BillStartNumber,1,i)
		Tno=MID(BillStartNumber,i+1,len(BillStartNumber))
		exit for
	end if
next

for i=len(BillEndNumber) to 1 step -1
	if not IsNumeric(mid(BillEndNumber,i,1)) then
		Tno2=MID(BillEndNumber,i+1,len(BillEndNumber))
		exit for
	end if
next

sqlunion="(select BillNo from BillBase where BillNo like '"&Sno&"%' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' and recordStateID=0) union all (select BillNo from PasserBase where BillNo like '"&Sno&"%' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' and recordStateID=0)"

strSQL="select a.BillNo from (select BillNo from getbilldetail where BillStateID <> 461 and BillNo like '"&Sno&"%' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"') a,("&sqlunion&") b where a.BillNo=b.BillNo(+) and b.Billno is null order by a.BillNo"
set rsbill=conn.execute(strSQL)
If Not rsbill.eof Then
	errBillNo=""
	While Not rsbill.eof
		If errBillNo<>"" Then errBillNo=errBillNo&"\n"
		errBillNo=errBillNo&trim(rsbill("BillNo"))&"¡AµL«ØÀÉ¬ö¿ý!!"
		rsbill.movenext
	Wend
	response.write "alert('"&errBillNo&"');"
End if
rsbill.close
%>


updGetBillBase.Submit423.disabled=false;