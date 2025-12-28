
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay

if len(year(now)-1911)=2 then
	sYear = "0" & year(now)-1911
else
	sYear = year(now)-1911
end if

fname= sYear & fMnoth & fDay & "°êµ|§½.txt"
Response.Buffer = true
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"
Server.ScriptTimeout = 800
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

BasSQL="("&tmpSQL&") tmpPasser"

	strSQL="select sn,billno from passerbase where exists(select 'Y' from "&BasSQL&" where sn=PasserBase.sn) and not exists(select 'N' from PasserProperty where to_char(checkdate,'YYYY')=to_char(sysdate,'YYYY') and billsn=passerbase.sn)"

	set rs=conn.execute(strSQL)
	While not rs.eof
		
		strSQL="insert into PasserProperty(billsn,billno,checkdate) values("&rs("sn")&",'"&rs("billno")&"',sysdate)"

		conn.execute(strSQL)

		rs.movenext
	Wend
	rs.close
	set rs=nothing

	strQuery="select distinct DriverID,(case when substr(DriverID,1,1) between 'A' and 'Z' and substr(DriverID,2,1) between '1' and '2' then 0 else 3 end) DriverKind,(select to_char(checkdate,'YYYY') from PasserProperty where billsn=passerbase.sn and to_char(checkdate,'YYYY')=to_char(sysdate,'YYYY')) checkdate from PasserBase where exists(select 'Y' from "&BasSQL&" where sn=PasserBase.sn) order by DriverID"

	set rsfound=conn.execute(strQuery)
	If Not rsfound.Bof Then rsfound.MoveFirst 
	cnt=0
	While Not rsfound.Eof
		cnt=cnt+1
		Response.Write right("000000000"&cnt,8)&right("000000000"&cnt&"A",7)
		Response.Write " "
'		tmpstr=(rsfound("checkdate")-1911)&"N"&rsfound("DriverKind")&rsfound("DriverID")
		tmpstr=(rsfound("checkdate")-1912)&rsfound("DriverKind")&rsfound("DriverID")

'		strDate=(rsfound("checkdate")-1911)
		
		rsfound.MoveNext

'		If not rsfound.eof Then
			Response.Write tmpstr
'		else
'			Response.Write replace(tmpstr,strDate&"N",strDate&"L")
'		End if 

		response.write(vbCrLf)
	Wend
	rsfound.close
	set rsfound=nothing
				
conn.close
set conn=nothing

%>