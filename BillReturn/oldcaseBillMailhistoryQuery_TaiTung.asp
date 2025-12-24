<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
if request("DB_Selt")="Selt" then
	strwhere=""
	if request("Sys_BatchNo")<>"" then
		If ifnull(strwhere) Then
			strwhere=" where "
		else
			strwhere=strwhere&" and "
		End if
		strwhere=strwhere&"BatchNo='"&request("Sys_BatchNo")&"'"
	end if
	if request("RecordDate1")<>"" and request("RecordDate2")<>""then
		RecordDate1=gOutDT(request("RecordDate1"))&" 0:0:0"
		RecordDate2=gOutDT(request("RecordDate2"))&" 23:59:59"
		If ifnull(strwhere) Then
			strwhere=" where "
		else
			strwhere=strwhere&" and "
		End if
		strwhere=strwhere&"RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
	if request("Sys_BillNo")<>"" then
		If ifnull(strwhere) Then
			strwhere=" where "
		else
			strwhere=strwhere&" and "
		End if
		strwhere=strwhere&"BillNo='"&request("Sys_BillNo")&"'"
	end if
	
	'LEO偷偷加 加入車號查詢 
	if request("Sys_CarNO")<>"" then
		If ifnull(strwhere) Then
			strwhere=" where "
		else
			strwhere=strwhere&" and "
		End if
		strwhere=strwhere&"CarNo='"&request("Sys_CarNO")&"'"
	end if
	
	strSQL="select BatchNo,SninDCIFile,BillNo,CarNo,DeCode(ReaSonID,'1','送達','F','寄存郵局','D','公示送達','') ReaSonName,RecordDate,DeCode(Status,'S','成功','N','找不到資料','n','已結案','p','已做其它送達','B','無此車號/證號','E','日期錯誤','k','已送達不可做未送達註記','h','已開裁決書','未處理') StatusName,FileName from OldCaseBillMailHistory"&strwhere&" order by FileName,SninDCIFile"
	set rsfound=conn.execute(strSQL)
	
	strCnt="select count(*) as cnt from OldCaseBillMailHistory"&strwhere
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
	DB_Display="show"
end if
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE> 舊資料上傳查詢(不上傳)</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><span class="style3">DCI 舊資料上傳查詢紀錄(不上傳)</span></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
				<tr>
					<td>
						批號　
						<input name="Sys_BatchNo" type="text" class="btn1" value="<%=request("Sys_BatchNo")%>" size="10">
						　　　
						註記日期　
						<input name="RecordDate1" type="text" class="btn1" value="<%
							if DB_Display="show" then
								response.write trim(request("RecordDate1"))
							else
								response.write gInitDT(date)
							end if%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
						~
						<input name="RecordDate2" type="text" class="btn1" value="<%
							if DB_Display="show" then
								response.write trim(request("RecordDate2"))
							else
								response.write gInitDT(date)
							end if%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate2');">
						　　　
						舉發單號　
						<input name="Sys_BillNo" type="text" class="btn1" value="<%=request("Sys_BillNo")%>" size="10" maxlength="9">
						
						車號　
						<input name="Sys_CarNO" type="text" class="btn1" value="<%=request("Sys_CarNO")%>" size="10" maxlength="9">
						　
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt('Selt');">
						<input type="button" name="cancel" value="清除" onClick="location='oldcaseBillMailhistoryQuery_TaiTung.asp'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33"><span class="style3">DCI 舊資料上傳查詢紀錄</span>
		每頁<select name="sys_MoveCnt" onchange="repage();">
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
			</select>筆<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄. )</strong></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th>批號</th>
					<th>序號</th>
					<th>註記日期</th>
					<th>單號</th>
					<th>車號</th>
					<th>送達原因</th>
					<th>檔名</th>
					<th>狀態</th>
				</tr><%
					If Not ifnull(request("DB_Selt")) Then
						if Trim(request("DB_Move"))="" then
							DBcnt=0
						else
							DBcnt=request("DB_Move")
						end if
						if Not rsfound.eof then rsfound.move cdbl(DBcnt)
						for i=DBcnt+1 to DBcnt+10
							if rsfound.eof then exit for
							response.write "<tr bgcolor='#FFFFFF' align='center' "
							lightbarstyle 0 
							response.write ">"
							response.write "<td>"&trim(rsfound("BatchNo"))&"</td>"
							response.write "<td>"&trim(rsfound("SninDCIFile"))&"</td>"
							response.write "<td>"&gInitDT(trim(rsfound("RecordDate")))&"</td>"
							response.write "<td>"&trim(rsfound("BillNo"))&"</td>"
							response.write "<td>"&trim(rsfound("CarNo"))&"</td>"
							response.write "<td>"&trim(rsfound("ReaSonName"))&"</td>"
							response.write "<td>"&trim(rsfound("FileName"))&"</td>"
							response.write "<td>"&trim(rsfound("StatusName"))&"</td>"
							response.write "</tr>"
							rsfound.movenext
						next
					end if%>
				</table>
		</td>
	</tr>
	<tr>
		<td height="30" bgcolor="#FFDD77" align="center">
			<a href="file:///.."></a>
			<input type="button" name="MoveFirst" value="第一頁" onclick="funDbMove(0);">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="MoveDown" value="最後一頁" onclick="funDbMove(999);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="Sys_SQL" value="<%=strSQL%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSelt(DBKind){
	var error=0;
	if (error==0){
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=1;
				alert("註記日輸入不正確!!");
			}
		}
	}
	if (error==0){
		if(myForm.RecordDate2.value!=""){
			if(!dateCheck(myForm.RecordDate2.value)){
				error=1;
				alert("上傳日輸入不正確!!");
			}
		}
	}
	if (error==0){
		myForm.DB_Move.value="";
		myForm.DB_Selt.value=DBKind;
		myForm.submit();
	}
}

function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10+eval(myForm.sys_MoveCnt.value))==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))-1)*(10+eval(myForm.sys_MoveCnt.value));
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))*(10+eval(myForm.sys_MoveCnt.value));
		}
		myForm.submit();
	}
}
function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}

function funchgExecel(){
	UrlStr="oldcaseBillMailhistoryQuery_Execel_TaiTung.asp";
	myForm.action=UrlStr;
	myForm.target="inputWin";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
</script>