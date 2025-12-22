<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
GetBillSN = trim(Request("GetBillSN"))
CounterfoiReturn=trim(Request("tag"))

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close


If CounterfoiReturn=1 Then
	strSQL="update GetBillBase set BillReturnDate="&funGetDate(date,1)&",CounterfoiReturn="&CounterfoiReturn&" where GetBillSN="&GetBillSN
	conn.execute(strSQL)

	If sys_City = "台東縣" then
		strSQL="select '(990713_'||(getbillseq.NextVal)||')' getbillseq from Dual"
		set rs=conn.execute(strSQL)
		getbillseq=trim(rs("getbillseq"))
		rs.close

		strSQL="update getbillbase set note=note||'"&getbillseq&"' Where GETBILLSN=" & GetBillSN

		conn.execute(strSQL)
	end if

elseIf CounterfoiReturn=0 Then
	strSQL="update GetBillBase set BillReturnDate=null,CounterfoiReturn="&CounterfoiReturn&" where GetBillSN="&GetBillSN
	conn.execute(strSQL)
end if

If CounterfoiReturn=0 Then
	str_Content="使用中"
	Re_Content="使用完畢"
	Re_Contentvalue=1

elseif CounterfoiReturn=1 Then
	str_Content="使用完畢"
	Re_Content="使用中"
	Re_Contentvalue=0
End if
%>

over_<%=GetBillSN%>.innerHTML='<%=str_Content%>';
btn_<%=GetBillSN%>.innerHTML='<input type=\"button\" name=\"Submit433\" value=\"<%=Re_Content%>\" onclick=\"funBillOver(\'<%=GetBillSN%>\',\'<%=Re_Contentvalue%>\');\">';