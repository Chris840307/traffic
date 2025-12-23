<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>績效獎勵金試算表</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->
<!-- #include file="..\Common\bannernodata.asp" -->
<%
'權限
'AuthorityCheck(234)
RecordDate=split(gInitDT(date),"-")
'組成查詢SQL字串
DB_Selt=request("DB_Selt")
if DB_Selt="Selt" then
		strwhere=""
		if request("IllegalDate")<>"" and request("IllegalDate1")<>""then
			ArgueDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
			ArgueDate2=gOutDT(request("IllegalDate1"))&" 23:59:59"
			strwhere=" and a.IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		end if
		if request("Sys_BillUnitID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillUnitID ='"&request("Sys_BillUnitID")&"'"
			else
				strwhere=" and a.BillUnitID='"&request("Sys_BillUnitID")&"'"
			end if
		end if
		if request("Sys_BillMem")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and (a.BillMemID1='"&request("Sys_BillMem")&"' or a.BillMemID2='"&request("Sys_BillMem")&"' or a.BillMemID3='"&request("Sys_BillMem")&"')"
			else
				strwhere=" and (a.BillMemID1='"&request("Sys_BillMem")&"' or a.BillMemID2='"&request("Sys_BillMem")&"' or a.BillMemID3='"&request("Sys_BillMem")&"')"
			end if
		end if
		if trim(request("RecordStateID"))<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.RecordStateID="&request("RecordStateID")
			else
				strwhere=" and a.RecordStateID="&request("RecordStateID")
			end if
		end if 
		strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,b.LoginID as BillMemID1,c.LoginID as BillMemID2,d.LoginID as BillMemID3,b.CreditID as CreditID1,c.CreditID as CreditID2,d.CreditID as CreditID3,b.Chname as Chname1,c.Chname as Chname2,d.Chname as Chname3,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillBaseTypeID,e.UnitName from BillBaseView a,MemberData b,MemberData c,MemberData d,UnitInfo e where a.BillMemID1=b.MemberID(+) and a.BillMemID2=c.MemberID(+) and a.BillMemID3=d.MemberID(+) and a.BillUnitID=e.UnitID(+)"&strwhere&" order by a.BillMemID1"

		set rsfound=conn.execute(strSQL)

		strCnt="select count(*) as cnt from BillBaseView a,MemberData b,MemberData c,MemberData d,UnitInfo e where a.BillMemID1=b.MemberID(+) and a.BillMemID2=c.MemberID(+) and a.BillMemID3=d.MemberID(+) and a.BillUnitID=e.UnitID(+)"&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close
		tmpSQL=strwhere
end if
%>
<html>

</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">績效獎勵金試算表</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<input type="checkbox" name="IllegalDateCheck" value="1" <%
						if trim(request("IllegalDateCheck"))="1" then
							response.write "checked"
						end if
						%>>
						違規日期
						<input name="IllegalDate" type="text" value="<%=request("IllegalDate")%>" size="8" maxlength="7" class="btn1">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate');">
						~
						<input name="IllegalDate1" type="text" value="<%=request("IllegalDate1")%>" size="8" maxlength="7" class="btn1">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate1');">
						<img src="space.gif" width="8" height="10">
						單位
						<%=SelectUnitOption("Sys_BillUnitID","Sys_BillMem")%>
						<img src="space.gif" width="8" height="10">
						舉發員警
						<%=SelectMemberOption("Sys_BillUnitID","Sys_BillMem")%>
						<img src="space.gif" width="8" height="10">
						舉發單狀態
						<select name="RecordStateID">
							<option value="0" <%if trim(request("RecordStateID"))="0" then response.write "selected"%>>有效</option>
							<!--<option value="-1" <%if trim(request("RecordStateID"))="-1" then response.write "selected"%>>已刪除</option>
							<option value="all" <%if trim(request("RecordStateID"))="all" then response.write "selected"%>>全部</option>-->
						</select>
						<input type="button" name="btnSelt" value="產生績效獎勵金試算表" onclick="funSelt();">
						<input type="button" name="cancel" value="清除" onClick="location='BillLawScore.asp'"> 
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" class="style3">
			舉發單紀錄列表
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
			筆 <font color="#F90000"><strong>(共 <%=DBsum%> 筆)</strong></font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th>單位</th>
					<th>員警臂章號碼</th>
					<th>姓名</th>
					<th>身分證</th>
					<th>法條</th>
					<th>舉發單別</th>
					<th nowrap>舉發日期</th>
					<th>舉發單號</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
				if DB_Selt="Selt" then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						chname="":chRule="":ForFeit="":CreditID="":chnameID=""
						if rsfound("BillMemID1")<>"" then chnameID=rsfound("BillMemID1")
						if rsfound("BillMemID2")<>"" then chnameID=chnameID&","&rsfound("BillMemID2")
						if rsfound("BillMemID3")<>"" then chnameID=chnameID&","&rsfound("BillMemID3")

						if rsfound("BillMemID1")<>"" then Chname=rsfound("Chname1")
						if rsfound("BillMemID2")<>"" then Chname=Chname&","&rsfound("Chname2")
						if rsfound("BillMemID3")<>"" then Chname=Chname&","&rsfound("Chname3")

						if rsfound("CreditID1")<>"" then CreditID=rsfound("CreditID1")
						if rsfound("CreditID2")<>"" then CreditID=CreditID&","&rsfound("CreditID2")
						if rsfound("CreditID3")<>"" then CreditID=CreditID&","&rsfound("CreditID3")

						if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
						if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
						if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")
						if rsfound("Rule4")<>"" then chRule=chRule&"/"&rsfound("Rule4")

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='35'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td>"&rsfound("UnitName")&"</td>"
						response.write "<td>"&chnameID&"</td>"
						response.write "<td>"&Chname&"</td>"
						response.write "<td>"&CreditID&"</td>"
						response.write "<td>"&chRule&"</td>"
						response.write "<td>"
						if trim(rsfound("BillBaseTypeID"))="0" then
							strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
							set rsBTypeVal=conn.execute(strBTypeVal)
							if not rsBTypeVal.eof then response.write rsBTypeVal("Content")
							rsBTypeVal.close
							set rsBTypeVal=nothing
						else
							response.write "攔停"
						end if
						response.write "</td>"
						response.write "<td width='5%'>"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
						response.write "<td width='6%'>"&rsfound("BillNo")&"</td>"
						response.write "</tr>"
						rsfound.movenext
					next
				end if
				%>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFDD77" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
		</td>
	</tr>
	<tr>
		<td>
			<p align="center">&nbsp;</p>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
		<%response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"%>
	function funSelt(){
		var error=0;
		var errorString="";
		if(myForm.IllegalDate.value!=""){
			if(!dateCheck(myForm.IllegalDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：違規日期輸入不正確!!";
			}
		}
		if(myForm.IllegalDate1.value!=""){
			if(!dateCheck(myForm.IllegalDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：違規日期輸入不正確!!";
			}
		}
		if (error>0){
			alert(errorString);
		}else{
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}

	function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
	}
	function repage(){
		myForm.DB_Move.value=0;
		myForm.submit();
	}
	function funchgExecel(){
		myForm.action="BillLawScore_Execel.asp";
		myForm.target="inputWin";
		myForm.submit();
		myForm.action="";
		myForm.target="";
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
</script>
<%
conn.close
set conn=nothing
%>