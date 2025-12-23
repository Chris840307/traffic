<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
if Not ifnull(request("Sys_BatchNumber")) and Instr(1,UCase(request("Sys_BatchNumber")),"W", 1) >0 then
	tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
	for i=0 to Ubound(tmp_BatchNumber)
		if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
		if i=0 then
			Sys_BatchNumber=trim(Sys_BatchNumber)&tmp_BatchNumber(i)
		else
			Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&tmp_BatchNumber(i)
		end if
		if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(Sys_BatchNumber)&"'"
	next
	strSQL="select count(*) as cnt from DCILOG a,BillBase b where a.BillSN=b.SN and a.BillNo=b.BillNo and a.BatchNumber in('"&Sys_BatchNumber&"') and  b.RecordMemberID="&Session("User_ID")
	set rs=conn.execute(strSQL)
	if Cint(rs("cnt"))>0 or trim(Session("Credit_ID"))="A000000000" then
		response.write "myForm.hd_PrintSum.value=""0"";"
		response.write "myForm.PBillSN.value="""";"
		response.write "myForm.DB_Move.value="""";"
		response.write "myForm.DB_Selt.value='BatchSelt';"
		response.write "myForm.DB_Display.value='show';"
		response.write "myForm.submit();"
	else
		response.write "alert(""非原建檔人不可查詢列印!!"");"
	end if
	rs.close
elseif Not ifnull(request("Sys_BillNo1")) then
	if trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
		strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
	elseif trim(request("Sys_BillNo1"))<>"" then
		strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo1"))&"'"
	elseif trim(request("Sys_BillNo2"))<>"" then
		strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo2"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
	end if
	strSQL="select count(*) as cnt from DCILOG a,BillBase b where a.BillSN=b.SN and a.BillNo=b.BillNo"&strwhere&" and  b.RecordMemberID="&Session("User_ID")
	set rs=conn.execute(strSQL)
	if Cint(rs("cnt"))>0 or trim(Session("Credit_ID"))="A000000000" then
		response.write "myForm.hd_PrintSum.value=""0"";"
		response.write "myForm.PBillSN.value="""";"
		response.write "myForm.DB_Move.value="""";"
		response.write "myForm.DB_Selt.value='BatchSelt';"
		response.write "myForm.DB_Display.value='show';"
		response.write "myForm.submit();"
	else
		response.write "alert(""非原建檔人不可查詢列印!!"");"
	end if
	rs.close
elseif Not ifnull(request("Sys_BatchNumber")) then
	response.write "myForm.hd_PrintSum.value=""0"";"
	response.write "myForm.PBillSN.value="""";"
	response.write "myForm.DB_Move.value="""";"
	response.write "myForm.DB_Selt.value='BatchSelt';"
	response.write "myForm.DB_Display.value='show';"
	response.write "myForm.submit();"
end if
%>
