<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout = 16800
Response.flush

	if request("RecordDate")<>"" and request("RecordDate1")<>""then
		CaseInDate1=gOutDT(request("RecordDate"))&" 0:0:0"
		CaseInDate2=gOutDT(request("RecordDate1"))&" 23:59:59"

		strwhere=" between TO_DATE('"&CaseInDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&CaseInDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID <> 3552"
	end If 

	strSQL="select RecordDate,UnitOrder,BillTypeUnit,BillUnit,Nvl(sum(Run_Record),0) Run_RecordSum,Nvl(Sum(Run_Del),0) Run_DelSum,Nvl(Sum(RunEqui),0) RunEquiSum,Nvl(Sum(StopEqui),0) StopEquiSum" & _
	",Nvl(sum(Stop_Record),0) Stop_RecordSum,Nvl(Sum(Stop_Suse),0) Stop_SuseSum,Nvl(Sum(Passer_Record),0) Passer_RecordSum" & _	
	" from (" & _
	"select TO_char(RecordDate,'YYYY/MM/DD') RecordDate" & _
	",(select 1 from BillBase a where (a.recordstateid=0 or (a.recordstateid=-1 and billno is not null)) and a.SN=BillBase.SN and '2'=BillBase.BillTypeID) Run_Record" & _
	",(select DeCode(RecordStateID,-1,1,(select Decode(DciReturnStatus,-1,1,null) delCnt from DciReturnStatus where DciActionID='W' and DciReturn=DciLog.DCIReturnStatusID)) from DciLog,BillBase bill where ExchangeTypeID='W' and BillSN=bill.SN and billSN=BillBase.SN and '2'=BillBase.BillTypeID) Run_Del" & _
	",(select DeCode(EquiPmentID,1,1,null) tmpEquip from billbase a where a.sn=billbase.sn and billbase.BillTypeID='2' and (a.recordstateid=0 or (a.recordstateid=-1 and billno is not null))) RunEqui" & _
	",(select DeCode(EquiPmentID,1,1,null) tmpEquip from billbase a where a.sn=billbase.sn and '1'=BillBase.BillTypeID and a.recordstateid=0) StopEqui" & _
	",(select 1 from BillBase a where a.recordstateid=0 and a.SN=BillBase.SN and '1'=BillBase.BillTypeID) Stop_Record" & _
	",(select (select Decode(DciReturnStatus,1,1,null) delCnt from DciReturnStatus where DciActionID='W' and DciReturn=DciLog.DCIReturnStatusID) from DciLog where ExchangeTypeID='W' and billSN=BillBase.SN and '1'=BillBase.BillTypeID and BillBase.RecordStateid=0) Stop_Suse" & _
	",(select 1 from PasserBase a where a.SN=BillBase.SN and '3'=BillBase.BillTypeID and a.recordstateid=0) Passer_Record" & _
	",UnitOrder,BillTypeUnit,BillUnit" & _
	" from (" & _
	"select distinct decode(billno,null,sn,(select max(sn) sn from billbase bs where bs.billno=billbase.billno)) sn" & _
	",decode(billno,null,RecordDate,(select max(RecordDate) RecordDate from billbase bs where bs.billno=billbase.billno)) RecordDate" & _
	",decode(billno,null,RecordStateid,(select max(RecordStateid) RecordStateid from billbase bs where bs.billno=billbase.billno)) RecordStateid,BillTypeID" & _
	",(select (select UnitName from Unitinfo a where UnitID=Unitinfo.UnitTypeID) BillTypeUnit from Unitinfo where UnitID=BillBase.BillUnitID) BillTypeUnit" & _
	",(select (select UnitOrder from Unitinfo a where UnitID=Unitinfo.UnitTypeID) BillTypeUnit from Unitinfo where UnitID=BillBase.BillUnitID) UnitOrder" & _
	",(Select BillBase.BillUnitID||' '||UnitName from Unitinfo b where UnitID=BillBase.BillUnitID) BillUnit" & _
	",decode(billno,null,EquiPmentID,(select max(nvl(EquiPmentID,-1)) EquiPmentID from billbase bs where bs.billno=billbase.billno)) EquiPmentID from BillBase where billtypeid=2 and RecordDate"&strwhere & _
	" Union all " & _
	"select sn,RecordDate,RecordStateid,BillTypeID" & _
	",(select (select UnitName from Unitinfo a where UnitID=Unitinfo.UnitTypeID) BillTypeUnit from Unitinfo where UnitID=BillBase.BillUnitID) BillTypeUnit" & _
	",(select (select UnitOrder from Unitinfo a where UnitID=Unitinfo.UnitTypeID) BillTypeUnit from Unitinfo where UnitID=BillBase.BillUnitID) UnitOrder" & _
	",(Select BillBase.BillUnitID||' '||UnitName from Unitinfo b where UnitID=BillBase.BillUnitID) BillUnit" & _
	",EquiPmentID from BillBase where billtypeid=1 and Exists(select 'Y' from DciLog where billsn=billbase.sn and exchangetypeid='W') and RecordDate"&strwhere & _
	" Union all " & _
	"select sn,RecordDate,RecordStateid,'3' BillTypeID" & _
	",(select (select UnitName from Unitinfo a where UnitID=Unitinfo.UnitTypeID) BillTypeUnit from Unitinfo where UnitID=PasserBase.BillUnitID) BillTypeUnit" & _
	",(select (select UnitOrder from Unitinfo a where UnitID=Unitinfo.UnitTypeID) BillTypeUnit from Unitinfo where UnitID=PasserBase.BillUnitID) BillTypeUnit" & _
	",(Select PasserBase.BillUnitID||' '||UnitName from Unitinfo b where UnitID=PasserBase.BillUnitID) BillUnit" & _
	",'-1' EquiPmentID" & _	
	" from PasserBase where recordstateid=0 and RecordDate"&strwhere & _	
	") BillBase" & _
	") group by RecordDate,UnitOrder,BillTypeUnit,BillUnit order by UnitOrder,BillUnit,RecordDate"

	set rsfound=conn.execute(strSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI 資料交換紀錄</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<%
filecnt=0:tmpUit=""
dim SumCnt(7)
dim totalCnt(7)
while Not rsfound.eof
	if tmpUit<>trim(rsfound("BillTypeUnit")) then
		if not ifnull(tmpUit) then	
			response.write "<tr>"
			response.write "<td align=""right"" colspan=""4"">合計</td>"

			for j=0 to 6
				response.write "<td>"&SumCnt(j)&"</td>"
				SumCnt(j)=0
			next
			response.write "<td>"
			Response.Write "0"
			Response.Write "</td>"

			response.write "</tr>"
			response.write "</table>"
			response.write "</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td>"
			response.write "<br>"
			response.write "前揭違規案件業已建檔傳送資料庫，請查核無誤後，於本表蓋章，擲回交通隊。"
			response.write "<br><br>"
			response.write "舉發單位簽收："
			response.write "</td>"
			response.write "</tr>"
			response.write "</table>"
			response.write "<div class=""PageNext""></div>"
		end if
		tmpUit=trim(rsfound("BillTypeUnit"))
		filecnt=0
		response.write "<table width=""100%"" border=""0"">"
		response.write "<tr>" 
		response.write "<td align=""center"">"
		response.write "<br><strong>苗栗縣警察局交通違規案件委外處理-603表</strong>"
		response.write "<br><br><strong>所屬單位："&tmpUit&"　　　　　　　　　　　　　　　　　　　　　　　收件日期："&request("RecordDate")&"~"&request("RecordDate1")&"</strong>"
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td>"
		response.write "<table width=""100%"" border=""1"" cellpadding=""4"" cellspacing=""0"">"
		response.write "<tr>"
		response.write "<td align=""center"" rowspan=""2"">序號</td>"
		response.write "<td align=""center"" rowspan=""2"">收件日期</td>"
		response.write "<td align=""center"" rowspan=""2"">分局</td>"
		response.write "<td align=""center"" rowspan=""2"">舉發單位</td>"
		response.write "<td align=""center"" colspan=""3"">逕舉</td>"
		response.write "<td align=""center"" colspan=""3"">攔停</td>"
		response.write "<td align=""center"" colspan=""2"">違警</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td>收件數</td>"
		response.write "<td>無法入案</td>"
		response.write "<td>投遞數</td>"
		response.write "<td>收件數</td>"
		response.write "<td>移送數</td>"
		response.write "<td>投遞數</td>"
		response.write "<td>收件數</td>"
		response.write "<td>入案數</td>"
		response.write "</tr>"
	end if
					
		filecnt=filecnt+1
		response.write "<tr>"
		response.write "<td align=""center"">"
		Response.Write filecnt
		Response.Write "&nbsp;</td>"

		response.write "<td>"
		Response.Write gInitDT(trim(rsfound("RecordDate")))
		Response.Write "&nbsp;</td>"

		response.write "<td>"
		Response.Write trim(rsfound("BillTypeUnit"))
		Response.Write "&nbsp;</td>"

		response.write "<td>"
		Response.Write trim(rsfound("BillUnit"))
		Response.Write "&nbsp;</td>"

		response.write "<td>"
		Response.Write trim(rsfound("Run_RecordSum"))
		Response.Write "&nbsp;</td>"
		SumCnt(0)=SumCnt(0)+cdbl(rsfound("Run_RecordSum"))
		totalCnt(0)=totalCnt(0)+cdbl(rsfound("Run_RecordSum"))

		response.write "<td>"
		Response.Write trim(rsfound("Run_DelSum"))
		Response.Write "&nbsp;</td>"
		SumCnt(1)=SumCnt(1)+cdbl(rsfound("Run_DelSum"))
		totalCnt(1)=totalCnt(1)+cdbl(rsfound("Run_DelSum"))

		response.write "<td>"
		Response.Write (cdbl(rsfound("RunEquiSum"))-cdbl(rsfound("Run_DelSum")))
		Response.Write "&nbsp;</td>"
		SumCnt(2)=SumCnt(2)+(cdbl(rsfound("RunEquiSum"))-cdbl(rsfound("Run_DelSum")))
		totalCnt(2)=totalCnt(2)+(cdbl(rsfound("RunEquiSum"))-cdbl(rsfound("Run_DelSum")))

		response.write "<td>"
		Response.Write trim(rsfound("Stop_RecordSum"))
		Response.Write "&nbsp;</td>"
		SumCnt(3)=SumCnt(3)+cdbl(rsfound("Stop_RecordSum"))
		totalCnt(3)=totalCnt(3)+cdbl(rsfound("Stop_RecordSum"))

		response.write "<td>"
		Response.Write trim(rsfound("Stop_SuseSum"))
		Response.Write "&nbsp;</td>"
		SumCnt(4)=SumCnt(4)+cdbl(rsfound("Stop_SuseSum"))
		totalCnt(4)=totalCnt(4)+cdbl(rsfound("Stop_SuseSum"))

		response.write "<td>"
		Response.Write trim(rsfound("StopEquiSum"))
		Response.Write "&nbsp;</td>"
		SumCnt(5)=SumCnt(5)+cdbl(rsfound("StopEquiSum"))
		totalCnt(5)=totalCnt(5)+cdbl(rsfound("StopEquiSum"))

		response.write "<td>"
		Response.Write trim(rsfound("Passer_RecordSum"))
		Response.Write "&nbsp;</td>"
		SumCnt(6)=SumCnt(6)+cdbl(rsfound("Passer_RecordSum"))
		totalCnt(6)=totalCnt(6)+cdbl(rsfound("Passer_RecordSum"))

		response.write "<td>"
		Response.Write "0"
		Response.Write "</td>"

		response.write "</tr>"
		rsfound.movenext
	wend
rsfound.close
response.write "<tr>"
response.write "<td align=""right"" colspan=""4"">合計</td>"

for j=0 to 6
	response.write "<td>"&SumCnt(j)&"</td>"
	SumCnt(j)=0
next
response.write "<td>"
Response.Write "0"
Response.Write "</td>"

response.write "</tr>"
				%>
			</table>
		</td>
	</tr><%
	response.write "<tr>"
	response.write "<td align=""left""><br>總合計："&(totalCnt(2)+totalCnt(4)+totalCnt(6))&"件、逕舉-收件數："&totalCnt(0)&"，逕舉-無法入案："&totalCnt(1)&"，逕舉-投遞數："&totalCnt(2)&"，"
	Response.Write "攔停-收件數："&totalCnt(3)&"，攔停-移送數："&totalCnt(4)&"，攔停-投遞數："&totalCnt(5)&"，違警-收件數："&totalCnt(6)&"，違警-入案數：0</td>"
	response.write "</tr>"%>
	<tr>
		<td>
			<br>
			前揭違規案件業已建檔傳送資料庫，請查核無誤後，於本表蓋章，擲回交通隊。
			<br><br>
			舉發單位簽收：
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>