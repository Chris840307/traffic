<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	strwhere=""
	strSQL="select TD_RecordDate,TD_RecordCity,TD_OwnerName,TD_ADDRESS,Decode(TD_PROCESS,'0','回報中','1','DCI處理中','已處理完成') TD_ProcessName from TDDT_RAREWORD where TD_CARNO='"&Ucase(trim(Request("Sys_TD_CARNO")))&"'"

	set rschk=conn.execute(strSQL)
	If not rschk.eof Then
		recorddate=split(gArrDT(rschk("TD_RecordDate")),"-")

		TD_RecordDate=recorddate(0)&"年"&recorddate(1)&"月"&recorddate(2)&"日"

		Response.Write "ChkCarNo.innerHTML=""<font color='red'>該車號已在"&TD_RecordDate&"回報<br>目前"&rschk("TD_ProcessName")&"。</font>"";"

		Response.Write "myForm.chk_carno.value=""1"";"
	else
		Response.Write "myForm.chk_carno.value=""0"";"
	End if
	rschk.close
	conn.close
%>
