<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<%

Server.ScriptTimeout = 8000
Response.flush

If Not ifnull(request("Sys_SendBillSN")) Then

	sys_billsn=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn=request("hd_BillSN")
else

	sys_billsn=request("BillSN")
End If 

tmp_billsn=split(sys_billsn,",")

sys_billsn=""

For i = 0 to Ubound(tmp_billsn)

	If i >0 then

		If i mod 100 = 0 Then

			sys_billsn=sys_billsn&"@"
		elseif sys_billsn<>"" then

			sys_billsn=sys_billsn&","
		end If 
	end if

	sys_billsn=sys_billsn&tmp_billsn(i)

Next

tmpSQL=""

If Ubound(tmp_billsn) >= 100 Then

	sys_billsn=split(sys_billsn,"@")
	
	For i = 0 to Ubound(sys_billsn)
		
		If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
		
		tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
	Next

else

	tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

End if 

strwhere=""

If request("Sys_checkdate1")<>"" and request("Sys_checkdate2")<>"" Then

	strwhere=" and to_char(checkdate,'YYYY') between '"&(cdbl(request("Sys_checkdate1"))+1911)&"' and '"&(cdbl(request("Sys_checkdate1"))+1911)&"'"
else
	
	strwhere=" and to_char(checkdate,'YYYY')=to_char(sysdate,'YYYY')"
End if 

BasSQL="("&tmpSQL&") tmpPasser"
	
	strSQL="update PasserProperty set postNumber='"&Request("Sys_PostNumber")&"' where exists(select 'Y' from "&BasSQL&" where sn=PasserProperty.billsn) "&strwhere&" and postNumber is null"

	conn.execute(strSQL)
				
conn.close
set conn=nothing

response.write "<script language=""JavaScript"">"
response.write "alert (""儲存完成！！"");"
Response.Write "self.close();"
response.write "</script>"

%>
<TITLE> 裁決批次套印 </TITLE>
</HEAD>
<BODY>
</BODY>
</HTML>