<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>車籍資料列表</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

if trim(request("kinds"))="CarDataSelect" then
	strwhere=Session("PrintCarDataSQL")&" and a.CarNo like '%"&trim(request("SelCarNo"))&"%'"
	strQry=strQry&",CarNo="&trim(request("SelCarNo"))
else
	strwhere=Session("PrintCarDataSQL")	
	strQry=Session("PrintCarDataSQLCheckItem")
end if
	Session.Contents.Remove("PrintCarDataSQLxls")
	Session("PrintCarDataSQLxls")=strwhere	

	Session.Contents.Remove("PrintCarDataSQLCheckItem")
	Session("PrintCarDataSQLCheckItem")=strQry

	dcitype=trim(request("dcitype"))

	strSQL="select distinct a.SN,a.BillTypeID,a.CarSimpleID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,a.IllegalDate,a.imagefilenameb,a.Note,e.CarNo,e.DCIReturnCarType,e.A_Name,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.Nwner,e.NwnerID,e.NwnerAddress,e.NwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strwhere&" order by a.RecordDate"

	set rsfound=conn.execute(strSQL)
'If  sys_City="台南市" Then
	ConnExecute "舉發單資料維護(車籍資料清冊):查詢事由:"&Trim(request("QryReason"))&"="&strQry ,355
'End If 

	strCnt="select count(*) as cnt from (select distinct a.SN,a.BillTypeID,a.CarSimpleID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,a.IllegalDate,a.imagefilenameb,a.Note,e.CarNo,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strwhere&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
	tmpSQL=strwhere
'response.write strSQL

%>

</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
	<tr>
		<td bgcolor="#1BF5FF">
			<font size="3"><strong>車籍資料清冊</strong></font>
			<img src="space.gif" width="56" height="8">
			每頁
			<select name="sys_MoveCnt" onchange="repage();">
				<option value="0"<%if trim(request("sys_MoveCnt"))="0" then response.write " Selected"%>>10</option>
				<option value="10"<%if trim(request("sys_MoveCnt"))="10" then response.write " Selected"%>>20</option>
				<option value="20"<%if trim(request("sys_MoveCnt"))="20" then response.write " Selected"%>>30</option>
				<option value="30"<%if trim(request("sys_MoveCnt"))="30" then response.write " Selected"%>>40</option>
				<option value="40"<%if trim(request("sys_MoveCnt"))="40" then response.write " Selected"%>>50</option>
				<option value="50"<%if trim(request("sys_MoveCnt"))="50" then response.write " Selected"%>>60</option>
				<option value="60"<%if trim(request("sys_MoveCnt"))="60" then response.write " Selected"%>>70</option>
				<option value="70"<%if trim(request("sys_MoveCnt"))="70" then response.write " Selected"%>>80</option>
				<option value="80"<%if trim(request("sys_MoveCnt"))="80" then response.write " Selected"%>>90</option>
				<option value="90"<%if trim(request("sys_MoveCnt"))="90" then response.write " Selected"%>>100</option>
			</select>
			筆 (共 <%=DBsum%> 筆)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			車號：<input type="text" size="10" name="SelCarNo" value="<%=trim(request("SelCarNo"))%>">
			<input type="button" name="Sel1" value="查詢" onclick="CarDataSelect();">
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="1300" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th></th>
					<th>車號</th>
				<%if sys_City<>"花蓮縣" then %>
					<th>牌類</th>
				<%end if%>
					<th>車別</th>
					<th>廠牌</th>
					<th>顏色</th>
					<th>車主姓名</th>
					<th>車主地址</th>
					<!-- <th>原車主姓名</th>
					<th>原車主地址</th>
					<th>駕駛人戶籍地址</th> -->
					<th>違規地點</th>
				<%if sys_City="雲林縣" or sys_City="南投縣" then %>
					<th>違規日期</th>
				<%end if%>
				<%if sys_City="雲林縣" then %>
					<th>違規法條</th>
				<%end if%>
				<%if sys_City<>"花蓮縣" then %>
					<th>限速、重</th>
					<th>車速、重</th>
				<%end if%>
				<%if trim(Session("SpecUser"))="1" then%>
					<th>業管車</th>
				<%end if%>
					<th>車籍狀態</th>
				<%if sys_City<>"花蓮縣" then %>
					<th>處理狀態</th>
				<%end if%>
				<%if sys_City="花蓮縣" then %>
					<th>違規事實</th>
				<%end if%>
					<th>操作</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center" >
				<%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					ListSN=DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						ListSN=ListSN+1
						if rsfound.eof then exit for
%>						<tr bgcolor="#ffffff" <%lightbarstyle 0 %>>
							<td width="1%"><%=ListSN%></td>
							<td width="6%"><%=rsfound("CarNo")%></td>
						<%if sys_City<>"花蓮縣" then %>
							<td width="4%"><%
							if trim(rsfound("CarSimpleID"))="1" then
								response.write "汽車"
							elseif trim(rsfound("CarSimpleID"))="2" then
								response.write "拖車"
							elseif trim(rsfound("CarSimpleID"))="3" then
								response.write "重機"
							elseif trim(rsfound("CarSimpleID"))="4" then
								response.write "輕機"
							end if								
							%></td>
						<%end if%>
							<td width="5%"><%
							if trim(rsfound("DCIReturnCarType"))<>"" and not isnull(rsfound("DCIReturnCarType")) then
								strCType="select * from DCIcode where TypeID=5 and ID='"&trim(rsfound("DCIReturnCarType"))&"'"
								set rsCType=conn.execute(strCType)
								if not rsCType.eof then
									response.write trim(rsCType("Content"))
								end if
								rsCType.close
								set rsCType=nothing
							end if								
							%></td>
							<td width="6%"><%
							if (trim(rsfound("A_Name"))<>"" and not isnull(rsfound("A_Name"))) then
								response.write trim(rsfound("A_Name"))
							end if
							%></td>
							<td width="4%"><%
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
							%></td>
							<td width="6%"><%=funcCheckFont(rsfound("Owner"),20,1)%></td>
							<td width="14%"><%
							if (trim(rsfound("OwnerAddress"))<>"" and not isnull(rsfound("OwnerAddress"))) then
								response.write trim(rsfound("OwnerZip"))&funcCheckFont(rsfound("OwnerAddress"),20,1)
							end if
							%></td>
							<!-- <td width="8%"> --><%'=rsfound("Nwner")%><!-- </td> -->
							<!-- <td width="14%"> --><%
'							if (trim(rsfound("NwnerAddress"))<>"" and not isnull(rsfound("NwnerAddress"))) then
'								response.write trim(rsfound("NwnerZip"))&trim(rsfound("NwnerAddress"))
'							end if
							%><!-- </td> -->
							<!-- <td width="14%"> --><%
'							if (trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress"))) then
'								response.write trim(rsfound("DriverHomeZip"))&trim(rsfound("DriverHomeAddress"))
'							end if
							%><!-- </td> -->
							
							<td width="10%"><%
							'違規地點
							if trim(rsfound("IllegalAddress"))<>"" and not isnull(rsfound("IllegalAddress")) then
								response.write trim(rsfound("IllegalAddress"))
							else
								response.write "&nbsp;"
							end if
							%></td>
						<%if sys_City="雲林縣" or sys_City="南投縣" then %>
							<td width="7%"><%
							if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
								response.write year(rsfound("IllegalDate"))-1911&" / "&month(rsfound("IllegalDate"))&" / "&day(rsfound("IllegalDate"))&"<br>"&hour(rsfound("IllegalDate"))&" : "&minute(rsfound("IllegalDate"))
							end if
							%></td>
						<%end if%>
						<%if sys_City="雲林縣" then %>
							<td width="7%"><%
							if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
								response.write trim(rsfound("Rule1"))
							end if
							if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
								response.write "<br>"&trim(rsfound("Rule2"))
							end if
							if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
								response.write "<br>"&trim(rsfound("Rule3"))
							end if
							%></td>
						<%end if%>
						<%if sys_City<>"花蓮縣" then %>
							<td width="6%"><%
							'限速
							if trim(rsfound("RuleSpeed"))<>"" and not isnull(rsfound("RuleSpeed")) then
								response.write trim(rsfound("RuleSpeed"))
							else
								response.write "&nbsp;"
							end if
							%></td><td width="6%"><%
							'車速
							if trim(rsfound("IllegalSpeed"))<>"" and not isnull(rsfound("IllegalSpeed")) then
								response.write trim(rsfound("IllegalSpeed"))
							else
								response.write "&nbsp;"
							end if
							%></td>
						<%end if%>
						<%if trim(Session("SpecUser"))="1" then%>
							<td align="center" width="5%"><%
							if sys_City="花蓮縣" then 
								strVip="select * from SpecCar where RecordStateID=0"
								set rsVip=conn.execute(strVip)
								If Not rsVip.Bof Then rsVip.MoveFirst 
								While Not rsVip.Eof
									if instr(trim(rsfound("Owner")),trim(rsVip("CarNo")))>0 then
										response.write "<font color=""red"">＊</font>"
									end if
								rsVip.MoveNext
								Wend
								rsVip.close
								set rsVip=nothing
							else
								strVip="select Count(*) as cnt from SpecCar where CarNo='"&trim(rsfound("CarNo"))&"' and RecordStateID=0"
								set rsVip=conn.execute(strVip)
								if cint(trim(rsVip("cnt"))) > 0 then
									response.write "<font color=""red"">＊</font>"
								end if
								rsVip.close
								set rsVip=nothing
							end if
							%></td>
						<%end if%>
							<td width="6%"><%
							if trim(rsfound("DCIReturnCarStatus"))<>"" and not isnull(rsfound("DCIReturnCarStatus")) then
								strCstatus="select Content from DCIcode where TypeID=10 and ID='"&trim(rsfound("DCIReturnCarStatus"))&"'"
								set rsCS=conn.execute(strCstatus)
								if not rsCS.eof then
									response.write trim(rsCS("COntent"))
								end if 
								rsCS.close
								set rsCS=nothing
							end if
							%></td>
						<%if sys_City<>"花蓮縣" then %>
							<td width="6%"><%
								strStatus="select ExchangeTypeID,DCIReturnStatusID from DCILog where BillSN="&trim(rsfound("SN"))&" order by ExchangeDate Desc"
								set rsStatus=conn.execute(strStatus)
								if not rsStatus.eof then
									strSID="select StatusContent from DCIReturnStatus where DCIactionId='"&trim(rsStatus("ExchangeTypeID"))&"' and DCIreturn='"&trim(rsStatus("DCIReturnStatusID"))&"'"
									set rsSID=conn.execute(strSID)
									if not rsSID.eof then
										response.write trim(rsSID("StatusContent"))
									end if
									rsSID.close
									set rsSID=nothing
								end if
								rsStatus.close
								set rsStatus=nothing
							%></td>
						<%end if%>
						<%if sys_City="花蓮縣" then %>
							<td align="left" width="22%"><%
							if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
								response.write trim(rsfound("Rule1"))
		'						strCarImple=""
		'						if left(trim(rsfound("Rule1")),4)="2110" then
		'							if trim(rsfound("CarSimpleID"))=1 or trim(rsfound("CarSimpleID"))=2 then
		'								strCarImple=" and CarSimpleID in ('5','0')"
		'							elseif trim(rsfound("CarSimpleID"))=3 or trim(rsfound("CarSimpleID"))=4 then
		'								strCarImple=" and CarSimpleID in ('3','0')"
		'							else
		'								strCarImple=""
		'							end if
		'						end if
		'						strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rsfound("Rule1"))&"' and Version='"&trim(rsfound("RuleVer"))&"'"&strCarImple
		'						set rsR1=conn.execute(strR1)
		'						if not rsR1.eof then 
		'							response.write " "&trim(rsR1("IllegalRule"))
		'						end if
		'						rsR1.close
		'						set rsR1=nothing
							end if
							if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
								response.write "<br>"&trim(rsfound("Rule2"))
								strCarImple=""
		'						if left(trim(rsfound("Rule2")),4)="2110" then
		'							if trim(rsfound("CarSimpleID"))=1 or trim(rsfound("CarSimpleID"))=2 then
		'								strCarImple=" and CarSimpleID in ('5','0')"
		'							elseif trim(rsfound("CarSimpleID"))=3 or trim(rsfound("CarSimpleID"))=4 then
		'								strCarImple=" and CarSimpleID in ('3','0')"
		'							else
		'								strCarImple=""
		'							end if
		'						end if
		'
		'						strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rsfound("Rule2"))&"' and Version='"&trim(rsfound("RuleVer"))&"'"&strCarImple
		'						set rsR1=conn.execute(strR1)
		'						if not rsR1.eof then 
		'							response.write " "&trim(rsR1("IllegalRule"))
		'						end if
		'						rsR1.close
		'						set rsR1=nothing
							end if
								if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
								response.write "<br>"&trim(rsfound("Rule3"))
								strCarImple=""
							end if
							if (trim(rsfound("RuleSpeed"))<>"" and not isnull(rsfound("RuleSpeed"))) and (trim(rsfound("IllegalSpeed"))<>"" and not isnull(rsfound("IllegalSpeed"))) then
								response.write "<br>速限"&trim(rsfound("RuleSpeed"))&"公里時速"&trim(rsfound("IllegalSpeed"))&"公里，超速"&trim(rsfound("IllegalSpeed"))-trim(rsfound("RuleSpeed"))&"公里"
							end if
							%></td>
						<%end if%>
							<td align="center" width="5%">
					<%if trim(rsfound("RecordStateID"))=0 then	'未刪除
						'---------刪除-------------
						' 花蓮縣用 imagefilenameb 以及 rule1 是不是56 + note 裡面是不是有催繳的檔案.txt副檔名 來判斷出現刪除按鈕													
						if (checkIsAllowDel(sys_City,trim(rsfound("BillTypeID")))=true) or (trim(rsfound("imagefilenameb"))<>"")  or ( (Instr(rsfound("Rule1"),"56")>0) and (Instr(rsfound("Note"),"txt")>0) and (sys_City="花蓮縣") ) then
							'抓入案日期 及 是否有入案
							CaseInDate=""
							CaseINCnt=0
							strCType="select a.DciCaseInDate from BillBaseDCIReturn a where ((a.BillNo='"&trim(rsfound("BillNo"))&"' and a.CarNo='"&trim(rsfound("CarNo"))&"') or (a.BillNo is null and a.CarNo='"&trim(rsfound("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
							set rsCType=conn.execute(strCType)
							if not rsCType.eof then
								CaseInDate=gOutDT(trim(rsCType("DciCaseInDate")))
								CaseINCnt=1
							end if
							rsCType.close
							set rsCType=nothing

							'計算入案幾天
							CountCaseIN=0
							if CaseInDate<>"" then
								CountCaseIN=dateDiff("d",CaseInDate,now)
							end if
							'response.write CountCaseIN
							'未入案直接刪
							if CaseINCnt=0 then
					%>
								<input type="button" name="b1" value="刪除" onclick="DelBill_NoDCI(<%=trim(rsfound("SN"))%>);">
					<%		else
								'if CountCaseIN<8 then	'超過七天不能刪
					%>
							<input type="button" value="刪除" onclick='window.open("BillBase_Del_DCI.asp?DBillSN=<%=trim(rsfound("SN"))%>","BillDelPage","left=250,top=250,location=0,width=440,height=200,resizable=yes")' <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(234,4)=false then
'							response.write "disabled"
'						end if
						%>>

					<%			'end if
							end if
						end if
					else	'已刪除則列出刪除原因
							strDelRea="select Content from BillDeleteReason a,DciCode b where a.BillSN="&trim(rsfound("SN"))&" and b.TypeID=3 and a.DelReason=b.ID"
							set rsDelRea=conn.execute(strDelRea)
							if not rsDelRea.eof then
								response.write trim(rsDelRea("Content"))
							end if
							rsDelRea.close
							set rsDelRea=nothing
					end if%>
							</td>
						</tr>
<%
						rsfound.movenext
					next
				%>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#1BF5FF" align="center">
			<a href="file:///.."></a>
			<a href="file:///......"></a>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<img src="space.gif" width="13" height="8">
<%
if sys_City<>"花蓮縣" then
	if trim(Session("SpecUser"))="1" then
%>
						<input type="button" name="cancel" value="入案前特殊車輛比對" onClick="funChkVIP();"> 
<%
	end if
else
	if trim(Session("SpecUser"))="1" then
%>
						<input type="button" name="cancel" value="入案前特殊車輛比對" onClick="funChkVIP_HL();"> 
<%
	end if
end if
%>
			<img src="space.gif" width="5" height="8">
<%if sys_City="雲林縣" then%>
			<input type="button" name="btnExecel" value="列印車籍清冊" onclick="funchgList();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(234,1)=false then
'							response.write "disabled"
'						end if
						%>>
<%elseif sys_City="花蓮縣" then%>
			<input type="button" name="btnExecel" value="列印車籍清冊" onclick="funchgList();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(234,1)=false then
'							response.write "disabled"
'						end if
						%>>
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(234,1)=false then
'							response.write "disabled"
'						end if
						%>>
<%else%>
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(234,1)=false then
'							response.write "disabled"
'						end if
						%>>
<%end if%>

			<img src="space.gif" width="5" height="8">
			<input type="button" name="btnExecel" value=" 離 開 " onclick="window.close();">
		</td>
	</tr>

</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="QryReason" value="<%=request("QryReason")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="hidden" name="strSQL" value="<%=tmpSQL%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	function funChkVIP(){
		window.open("DciCarDataChkSpecCar.asp","chk_vip1","width=620,height=440,left=200,top=150,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
	}
function funChkVIP_HL(){
	window.open("DciCarDataChkSpecCar_HL.asp","chk_vip1","width=620,height=440,left=200,top=150,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}
}
function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
function funchgExecel(){
	UrlStr="PrintCarDataList_Excel.asp?QryReason="+myForm.QryReason.value;
	newWin(UrlStr,"PrintWin_xls",790,550,50,10,"yes","yes","yes","no");
}
function funchgList(){
	UrlStr="PrintCarDataList_HL_A3.asp";
	newWin(UrlStr,"DciPrintWin_xls",790,550,50,10,"yes","yes","yes","no","yes");
}
function CarDataSelect(){
	if (myForm.SelCarNo.value==""){
		alert("請輸入車號！");
	}else{
		myForm.DB_Move.value=0;
		myForm.kinds.value="CarDataSelect";
		myForm.submit();	
	}
}
function DelBill_NoDCI(DelSN){
	UrlStr="BillBase_Del.asp?DBillSN="+DelSN;
	newWin(UrlStr,"DelBillBasePage1",290,100,200,110,"no","no","yes","no","no");
}
</script>
<%conn.close%>