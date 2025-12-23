<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>車籍資料列表</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<style type="text/css">
<!--
.style1 {font-size: 12pt}
.style2 {font-size: 10pt}

.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
-->
</style>
<%
Server.ScriptTimeout = 800
Response.flush
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

	strwhere=trim(request("SQLstr"))

	strCnt="select count(*) as cnt from (select distinct a.*,e.DciReturnCarColor,e.DriverHomeAddress,e.DriverHomeZip,e.Owner,e.OwnerAddress,e.OwnerZip,e.A_Name from BillBase a,DciLog b,BillBaseDciReturn e where a.CarNo=e.CarNo and e.ExchangeTypeID='A' and e.Status='S' and a.Sn=b.BillSn and a.RecordStateID=0 "&strwhere&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close

	tmpSQL=strwhere

	strSQL="select distinct a.*,e.DciReturnCarColor,e.DriverHomeAddress,e.DriverHomeZip,e.Owner,e.OwnerAddress,e.OwnerZip,e.A_Name,e.DciErrorCarData,e.Nwner,e.NwnerAddress,e.NwnerID,e.NwnerZip,e.OWNERNOTIFYADDRESS from BillBase a,DciLog b,BillBaseDciReturn e where a.CarNo=e.CarNo and e.ExchangeTypeID='A' and e.Status='S' and a.Sn=b.BillSn and a.RecordStateID=0 "&strwhere&" order by a.RecordDate"
	set rsfound=conn.execute(strSQL)
'response.write strwhere
%>

</head>
<body>
<form name=myForm method="post">
<%
PageCount=25
mailSN=0
PrintSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
<table width="100%" border="1" cellspacing="0">
	<tr>
		<td colspan="9">
			<center>
			<span class="style1">車籍資料清冊</span><span class="style2">(共 <%=DBsum%> 筆)
			</center>		
<div align="right">
<span class="style1">承辦人：
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
單位主管：
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  </span>
</div>
			<div align="right">印表單位：停車管理處</div>
			<div align="right">印表時間：<%
			response.write year(date)-1911& " - "& month(date)& " - " & day(date)& " - " & hour(now) & " : " & minute(now)
 			%></div>
			</span>
		</td>
	</tr>
	<tr>
		<td align="center" height="40" width="6%"><span class="style4">&nbsp;</span></td>	
		<td width="9%" height="31"><span class="style2">車號</span></td>
		<td width="6%"><span class="style2">違規日期</span></td>
		<td width="5%"><span class="style2">時間</span></td>
		<td width="8%"><span class="style2">廠牌</span></td>
		<td width="6%"><span class="style2">顏色</span></td>
		<td width="12%"><span class="style2">車主姓名</span></td>
		<td width="30%"><span class="style2">車主地址</span></td>
		<td width="18%"><span class="style2">違規地點</span></td>
	</tr>
<%		for i=1 to PageCount
			if rsfound.eof then exit for
			mailSN=mailSN+1 
			PrintSN=PrintSN+1
			   					
%>
	<tr>
		<td height="31"><span class="style2">
		<%	
		response.write PrintSN
		%>
		</span></td>
		<td height="31"><span class="style2">
		<%
		response.write rsfound("CarNo")
		%>
		</span></td>
		<td><span class="style2">
		<%
		response.write year(rsfound("IllegalDate"))-1911&"/"& month(rsfound("IllegalDate"))& "/" &day(rsfound("IllegalDate"))
		%>
		</span></td>
		<td><span class="style2">
		<%
		response.write right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&minute(rsfound("IllegalDate")),2)
		%>
		</span></td>
		<td><span class="style2">
		<%
		if (trim(rsfound("A_Name"))<>"" and not isnull(rsfound("A_Name"))) then
			response.write funcCheckFont(trim(rsfound("A_Name")),17,1)
		end if
		%>
		</span></td>
		<td><span class="style2">
		<%
		if trim(rsfound("DCIReturnCarColor"))<>"" and not isnull(rsfound("DCIReturnCarColor")) then
			ColorLen=cint(Len(rsfound("DCIReturnCarColor")))
			for Clen=1 to ColorLen
				colorID=mid(rsfound("DCIReturnCarColor"),Clen,1)
				strColor="select * from DCIcode where TypeID=4 and ID='"&trim(colorID)&"'"
				set rsColor=conn.execute(strColor)
				if not rsColor.eof then
					response.write trim(rsColor("Content"))
				end if
				rsColor.close
				set rsColor=nothing
			next
		end if
		%>
		</span></td>
		<td><span class="style2">
		<%
		response.write funcCheckFont(trim(rsfound("Owner")),17,1)
		%>
		</span></td>
		<td><span class="style2">
		<%
		if trim(rsfound("OWNERNOTIFYADDRESS"))<>"" and not isnull(rsfound("OWNERNOTIFYADDRESS")) then
			NotifyZip=""
			strNZ="select * from Zip where ZipName like '"&left(trim(rsfound("OWNERNOTIFYADDRESS")),5)&"%'"
			set rsNZ=conn.execute(strNZ)
			if not rsNZ.eof then
				NotifyZip=trim(rsNZ("ZipNo"))
			else
				strNZ2="select * from Zip where ZipName like '"&left(trim(rsfound("OWNERNOTIFYADDRESS")),3)&"%'"
				set rsNZ2=conn.execute(strNZ2)
				if not rsNZ2.eof then
					NotifyZip=trim(rsNZ2("ZipNo"))
				
				end if
				rsNZ2.close
				set rsNZ2=nothing
			end if
			rsNZ.close
			set rsNZ=nothing
			response.write NotifyZip&funcCheckFont(trim(rsfound("OWNERNOTIFYADDRESS")),17,1)
		elseif trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress")) Then
			response.write trim(rsfound("DriverHomeZip"))&funcCheckFont(trim(rsfound("DriverHomeAddress")),17,1)
		else
			response.write trim(rsfound("OwnerZip"))&funcCheckFont(trim(rsfound("OwnerAddress")),17,1)
		end if
		%>
		</span></td>
		<td><span class="style2">
		<%
		response.write trim(rsfound("IllegalAddress"))
		%>
		</span></td>
	</tr>
<%			
		rsfound.MoveNext
		next
%>
</table>

<%
Wend
rsfound.close
set rsfound=nothing
%>
</form>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
window.print();

</script>
<%conn.close%>