<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->
<!-- #include file="../Common/Banner.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單車籍查詢</title>
<!--#include virtual="traffic/Common/css.txt"-->
<% Server.ScriptTimeout = 800 %>
<%
'權限
'AuthorityCheck(249)
RecordDate=split(gInitDT(date),"-")
'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
'組成查詢SQL字串
if request("DB_Selt")="Selt" then
		strwhere=""
		if trim(request("RecordDateCheck"))="1" then
			if request("RecordDate")<>"" and request("RecordDate1")<>""then
				RecordDate1=gOutDT(request("RecordDate"))&" 0:0:0"
				RecordDate2=gOutDT(request("RecordDate1"))&" 23:59:59"
				if strwhere<>"" then
					strwhere=strwhere&" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				else
					strwhere=" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				end if
			end if
		end if
		if trim(request("RecordDate_h"))<>"" or trim(request("RecordDate1_h"))<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and to_char(a.RecordDate,'hh') between "&trim(request("RecordDate_h"))&" and "&trim(request("RecordDate1_h"))
			else
				strwhere=" and to_char(a.RecordDate,'hh') between "&trim(request("RecordDate_h"))&" and "&trim(request("RecordDate1_h"))
			end if
		end if
		if request("Sys_BillUnitID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillUnitID ="&request("Sys_BillUnitID")
			else
				strwhere=" and a.BillUnitID="&request("Sys_BillUnitID")
			end if
		end if
		if request("Sys_BillMem")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and (a.BillMem1='"&request("Sys_BillMem")&"' or a.BillMem2='"&request("Sys_BillMem")&"' or a.BillMem3='"&request("Sys_BillMem")&"')"
			else
				strwhere=" and (a.BillMem1='"&request("Sys_BillMem")&"' or a.BillMem2='"&request("Sys_BillMem")&"' or a.BillMem3='"&request("Sys_BillMem")&"')"
			end if
		end if
		if request("Sys_RecordUnit")<>"" and request("Sys_RecordMemberID")="" then
			strwhere=strwhere&" and a.RecordMemberID in (select MemberID from MemberData where UnitID in ('"&trim(request("Sys_RecordUnit"))&"'))"
		end if
		if request("Sys_RecordMemberID")<>"" then
			strwhere=strwhere&" and a.RecordMemberID ="&request("Sys_RecordMemberID")
		end if
		if request("Sys_BillTypeID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillTypeID='"&request("Sys_BillTypeID")&"'"
			else
				strwhere=" and a.BillTypeID='"&request("Sys_BillTypeID")&"'"
			end if
		end if
		if request("Sys_BillNo")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillNo='"&request("Sys_BillNo")&"'"
			else
				strwhere=" and a.BillNo='"&request("Sys_BillNo")&"'"
			end if
		end if
		if request("Sys_CarNo")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.CarNo like '%"&request("Sys_CarNo")&"%'"
			else
				strwhere=" and a.CarNo like '%"&request("Sys_CarNo")&"%'"
			end if
		end if
		if request("Sys_Driver")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.Driver='"&request("Sys_Driver")&"'"
			else
				strwhere=" and a.Driver='"&request("Sys_Driver")&"'"
			end if
		end if
		if request("Sys_DriverID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.DriverID='"&request("Sys_DriverID")&"'"
			else
				strwhere=" and a.DriverID='"&request("Sys_DriverID")&"'"
			end if
		end if
		if request("DCIstatus")<>"" then
			if trim(request("DCIstatus"))="0" then
				if strwhere<>"" then
					strwhere=strwhere&" and a.BillStatus='0'"
				else
					strwhere=" and a.BillStatus='0'"					
				end if
			elseif trim(request("DCIstatus"))="1" then
				if strwhere<>"" then
					strwhere=strwhere&" and a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where BillSN not in (select Billsn from DciLog where exchangeTypeID='A' and (DciReturnStatusID='S' or DciReturnStatusID is null)))"	
				else
					strwhere=" and a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where BillSN not in (select Billsn from DciLog where exchangeTypeID='A' and (DciReturnStatusID='S' or DciReturnStatusID is null)))"					
				end if
			end if
		end if
		if trim(request("BillUseTool"))="0" then
			strwhere=strwhere&" and a.UseTool<>8"
		elseif trim(request("BillUseTool"))="1" then
			strwhere=strwhere&" and a.UseTool=8"
		end if
		if strwhere<>"" then
			strwhere=strwhere&" and a.RecordStateID=0"
		else
			strwhere=" and a.RecordStateID=0"
		end if

		'是否要判斷一打一驗 1:是 0:否
		if Session("DoubleCheck")="1" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.DoubleCheckStatus=1"
			else
				strwhere=" and a.DoubleCheckStatus=1"
			end if
		end if

	strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+)"&strwhere&" order by a.RecordDate"
end if
	'response.write strSQL
'車籍查詢(遇到RecordStateID=-1不做)
if trim(request("kinds"))="CarDataCheck" then
	strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
	set rsSN=conn.execute(strSN)
	if not rsSN.eof then
		theBatchTime=(year(now)-1911)&"A"&trim(rsSN("SN"))
	end if
	rsSN.close
	set rsSN=nothing

	strCCheck="select a.SN,a.IllegalDate,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.BillUnitID,a.BillStatus,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+)"&strwhere
	set rsCCheck=conn.execute(strCCheck)
	If Not rsCCheck.Bof Then
		rsCCheck.MoveFirst 
	else
%>
<script language="JavaScript">
	alert("無可進行車籍查詢之舉發單！");
</script>
<%
	end if
	While Not rsCCheck.Eof
		funcCarDataCheck conn,trim(rsCCheck("SN")),trim(rsCCheck("BillNo")),trim(rsCCheck("BillTypeID")),trim(rsCCheck("CarNo")),trim(rsCCheck("BillUnitID")),trim(rsCCheck("RecordDate")),trim(rsCCheck("RecordMemberID")),theBatchTime
	rsCCheck.MoveNext
	Wend
	If Not rsCCheck.Bof Then
%>
<script language="JavaScript">
	alert("車籍查詢處理完成，批號：<%=theBatchTime%>");
</script>
<%
	end if
	rsCCheck.close
	set rsCCheck=nothing
end if

'做完車籍查詢及入案等動作後再查詢告發單，讓列表取得的資料為最新
if request("DB_Selt")="Selt" then
'response.write strSQL
'response.end
		set rsfound=conn.execute(strSQL)
		strCnt="select count(*) as cnt from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+)"&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close
		tmpSQL=strwhere
		Session.Contents.Remove("BillSQL")
		Session("BillSQL")=strSQL
		Session.Contents.Remove("PrintCarDataSQL")
		Session("PrintCarDataSQL")=strwhere
end if
%>
<style type="text/css">
<!--
.style4 {
	color: #FF0000;
	font-size: 14px
}
.style5 {
	color: #F00000;
	font-size: 16px
}
-->
</style>
</head>
<body>
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">舉發單車籍查詢<%
		if sys_City="雲林縣" then
		%>
		&nbsp;&nbsp;&nbsp;<span class="style5">逕舉手開單不須做車籍查詢，可直接進行『上傳監理站 - 逕舉入案』</span>
		<%
		end if
		%></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<input type="checkbox" name="RecordDateCheck" value="1" <%
						if trim(request("DB_Selt"))="" then
							DateChk="1"
						else
							DateChk=trim(request("RecordDateCheck"))
						end if
						if DateChk="1" then
							response.write "checked"
						end if
						%>>
						建檔日期
						<input name="RecordDate" type="text" value="<%
						if trim(request("DB_Selt"))="" then
							RecordDateTmp=ginitdt(DateAdd("d",-3,now))
						else
							RecordDateTmp=trim(request("RecordDate"))
						end if
						response.write RecordDateTmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate');">
						~
						<input name="RecordDate1" type="text" value="<%
						if trim(request("DB_Selt"))="" then
							RecordDate1Tmp=ginitdt(now)
						else
							RecordDate1Tmp=trim(request("RecordDate1"))
						end if
						response.write RecordDate1Tmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
						<img src="space.gif" width="8" height="10">
						時段
						<input name="RecordDate_h" type="text" value="<%=request("RecordDate_h")%>" size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">時 ~ 
						<input name="RecordDate1_h" type="text" value="<%=request("RecordDate1_h")%>" size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">時
						<img src="space.gif" width="8" height="10">
						<!-- 舉發類別 -->
						<input type="hidden" name="Sys_BillTypeID" value="2">
						DCI作業
						<select name="DCIstatus">
							<option value="0" <%
							if trim(request("DCIstatus"))="0" then response.write "selected"
							%>>車籍查詢</option>
							<option value="1" <%
							if trim(request("DCIstatus"))="1" then response.write "selected"
							%>>車籍查詢失敗</option>
						</select>
						<br>
						建檔單位
						<img src="space.gif" width="8" height="10">
						<%=SelectUnitOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="8" height="10">
						建檔人
						<%=SelectMemberOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="8" height="10">
						舉發單類別
						<select name="BillUseTool">
							<option value="0" <%
							if trim(request("BillUseTool"))="0" then response.write "selected"
							%>>逕舉</option>
							<option value="1" <%
							if trim(request("BillUseTool"))="1" then response.write "selected"
							%>>逕舉手開單</option>
							<option value="2" <%
							if trim(request("BillUseTool"))="2" then response.write "selected"
							%>>全部</option>
						</select>
						<img src="space.gif" width="8" height="10">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" <%
						if CheckPermission(249,1)=false then
							response.write "disabled"
						end if
						%>>
						<input type="button" name="cancel" value="清除" onClick="location='BillBaseQryCar.asp'"> 
						<br><span class="style4">注意：上傳監理站-車籍查詢 並不包含 入案 到監理站</span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#1BF5FF" class="style3">
			舉發單紀錄列表
			<img src="space.gif" width="56" height="8">
			每頁 
			<select name="sys_MoveCnt" onchange="repage1();">
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
					<th width="8%">違規日期</th>
					<th width="8%">舉發員警</th>
					<th width="5%">車號</th>
					<th width="6%">車種</th>
					<th width="10%">法條</th>
					<th width="8%">DCI</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
				if request("DB_Selt")="Selt" then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						chname="":chRule="":ForFeit=""
						if rsfound("BillMem1")<>"" then	chname=rsfound("BillMem1")
						if rsfound("BillMem2")<>"" then	chname=chname&"/"&rsfound("BillMem2")
						if rsfound("BillMem3")<>"" then	chname=chname&"/"&rsfound("BillMem3")
						if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
						if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
						if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")
						if rsfound("Rule4")<>"" then chRule=chRule&"/"&rsfound("Rule4")

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='30'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td width='5%'>"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
						response.write "<td width='8%'>"&chname&"</td>"
'					if trim(rsfound("BillTypeID"))="2" then
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("CarNo")&"</a></td>"
'					else
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("CarNo")&"</a></td>"
'					end if
						response.write "<td width='6%'>"&rsfound("CarNo")&"</td>"
						response.write "<td width='5%'>"
							if trim(rsfound("CarSimpleID"))="1" then
								response.write "汽車"
							elseif trim(rsfound("CarSimpleID"))="2" then
								response.write "拖車"
							elseif trim(rsfound("CarSimpleID"))="3" then
								response.write "重機"
							elseif trim(rsfound("CarSimpleID"))="4" then
								response.write "輕機"
							elseif trim(rsfound("CarSimpleID"))="5" then
								response.write "動力機械"
							elseif trim(rsfound("CarSimpleID"))="6" then
								response.write "臨時車牌"
							end if
						response.write "</td>"
						response.write "<td width='10%'>"&chRule&"</td>"
						response.write "<td width='8%'>"
						if trim(rsfound("BillStatus"))="0" then
							response.write "<font color='#999999'>未處理</font>"
						elseif trim(rsfound("BillStatus"))="1" then
							response.write "<font color='#FF66CC'>車籍查詢</font>"
						elseif trim(rsfound("BillStatus"))="2" then
							response.write "<font color='#009900'>入案</font>"
						elseif trim(rsfound("BillStatus"))="3" then
							response.write "<font color='#0000FF'>退件</font>"
						elseif trim(rsfound("BillStatus"))="4" then
							response.write "<font color='#0000FF'>寄存</font>"
						elseif trim(rsfound("BillStatus"))="5" then
							response.write "<font color='#0000FF'>公示</font>"
						elseif trim(rsfound("BillStatus"))="6" then
							response.write "<font color='#FF0000'>刪除</font>"
						end if
						response.write "</td>"
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
		<td height="35" bgcolor="#1BF5FF" align="center">
			<a href="file:///.."></a>
			<a href="file:///......"></a>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
			<input type="button" name="Submit424" value="進行車籍查詢" onclick="if(confirm('確定要向監理所查詢車籍資料嗎？\n\n注意：上傳監理站-車籍查詢 並不包含 入案 到監理站')){funcCarDataCheck()}">
			
			<span class="style3"><img src="space.gif" width="5" height="8"></span>
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<input type="hidden" name="DelReason" value="">
		</td>
	</tr>
	<tr>
		<td>
			<p align="center">&nbsp;</p>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	<%response.write "UnitMan('Sys_RecordUnit','Sys_RecordMemberID','"&trim(request("Sys_RecordMemberID"))&"');"%>
	function funSelt(){
		var error=0;
		var errorString="";
		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate.value.substr(0,1)=="9" && myForm.RecordDate.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate.value.substr(0,1)=="1" && myForm.RecordDate.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate1.value.substr(0,1)=="9" && myForm.RecordDate1.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate1.value.substr(0,1)=="1" && myForm.RecordDate1.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.RecordDate_h.value!="" || myForm.RecordDate1_h.value!=""){
			if(myForm.RecordDate_h.value=="" || myForm.RecordDate1_h.value==""){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔時段輸入不完整!!";
			}
		}
		if (error>0){
			alert(errorString);
		}else{
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}

	function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
		win.focus();
		return win;
	}
	function repage1(){
		myForm.DB_Move.value=0;
		myForm.submit();
	}
	function funchgExecel(){
		UrlStr="BillBaseQry_Execel.asp?WorkType=1";
		newWin(UrlStr,"inputWin",790,550,50,10,"yes","yes","yes","no");
	}
	//列印車籍清冊
	function funchgCarDataList(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲列印車籍清冊的舉發單！");
		}else{
			UrlStr="PrintCarDataList.asp";
			newWin(UrlStr,"CarListWin",790,575,50,10,"yes","no","yes","no");
		}
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
	//車籍查詢
	function funcCarDataCheck(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲車籍查詢的舉發單！");
		}else{
			myForm.kinds.value="CarDataCheck";
			myForm.submit();
		}
	}
<%if trim(request("DB_Selt"))="" then%>
	funSelt();
<%end if%>
</script>
<%
conn.close
set conn=nothing
%>