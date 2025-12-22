<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>罕見字回報系統</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%

strSQL="select * from TDDT_RAREWORD where rownum=1"
set rs=conn.execute(strSQL)

For i=0 to rs.Fields.count-1
	If trim(rs.Fields.item(i).Name)="TD_ADDRESS" Then Exit For
Next
If i>rs.Fields.count-1 Then
	strSQL="Alter Table TDDT_RAREWORD ADD (TD_ADDRESS VarChar2(160))"
	conn.execute(strSQL)
End if
rs.close

if request("DB_Del")="Del" then
	strSQL="Delete from TDDT_RAREWORD where TD_SN="&request("TD_SN")
	conn.execute strSQL
end if

if trim(request("DB_State"))="Selt" then
	strWhere=""
	If (not ifnull(Request("Sys_TD_RecordDate1"))) and (not ifnull(Request("Sys_TD_RecordDate2"))) Then
		TD_RecordDate1=gOutDT(request("Sys_TD_RecordDate1"))&" 0:0:0"
		TD_RecordDate2=gOutDT(request("Sys_TD_RecordDate2"))&" 23:59:59"

		If ifnull(strWhere) Then
			strWhere=" where TD_RecordDate between TO_DATE('"&TD_RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&TD_RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		else
			strWhere=strWhere&" and TD_RecordDate between TO_DATE('"&TD_RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&TD_RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		End if
	End if

	If not ifnull(Request("Sys_TD_CARNO")) Then
		If ifnull(strWhere) Then
			strWhere=" where TD_CARNO='"&trim(Request("Sys_TD_CARNO"))&"'"
		else
			strWhere=strWhere&" and TD_CARNO='"&trim(Request("Sys_TD_CARNO"))&"'"
		End if
	End if

	If not ifnull(Request("Sys_TD_PROCESS")) Then
		If ifnull(strWhere) Then
			strWhere=" where TD_PROCESS='"&trim(Request("Sys_TD_PROCESS"))&"'"
		else
			strWhere=strWhere&" and TD_PROCESS='"&trim(Request("Sys_TD_PROCESS"))&"'"
		End if
	End if
	
	strSQL="select TD_SN,TD_RecordDate,TD_RecordCity,TD_CARNO,TD_ADDRESS,TD_OwnerName,Decode(TD_PROCESS,'0','回報','1','DCI處理中','處理完成') TD_ProcessName,TD_PROCESS from TDDT_RAREWORD"&strWhere&" order by TD_RecordCity,TD_RecordDate DESC"
	set rsfound=conn.execute(strSQL)

	strCnt="select count(*) as cnt from ("&strSQL&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
end if
%>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">罕見字回報系統</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<table width="100%" border="0">
							<tr>
								<td>回報時間</td>
								<td nowrap>
									<input name="Sys_TD_RecordDate1" class="btn1" type="text" value="<%=trim(request("Sys_TD_RecordDate1"))%>" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
									<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_TD_RecordDate1');">
									～
									<input name="Sys_TD_RecordDate2" class="btn1" type="text" value="<%=trim(request("Sys_TD_RecordDate2"))%>" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
									<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_TD_RecordDate2');">
								</td>
								<td nowrap>
									車號
								</td>
								<td>
									<input name="Sys_TD_CARNO" class="btn1" type="text" value="<%=trim(request("Sys_TD_CARNO"))%>" size="10" maxlength="10">
								</td>
								<td nowrap>
									處理進度
								</td>
								<td>
									<select name="Sys_TD_PROCESS" class="btn1"><%
										Response.Write "<option value="""">請選擇</option>"

										Response.Write "<option value=""0"""
										If trim(Request("Sys_TD_PROCESS")) = "0" Then Response.Write " selected"
										Response.Write ">回報</option>"

										Response.Write "<option value=""1"""
										If trim(Request("Sys_TD_PROCESS")) = "1" Then Response.Write " selected"
										Response.Write ">DCI處理中</option>"

										Response.Write "<option value=""2"""
										If trim(Request("Sys_TD_PROCESS")) = "2" Then Response.Write " selected"
										Response.Write ">處理完成</option>"
									%></select>
								</td>
								<td>
									<input type="button" name="btnAdd" value="查詢" onclick="funSelt();">
									<input type="button" name="btnAdd" value="新增" onclick="funAdd();">
									<input type="button" name="cancel" value="清除" onClick="location='RareWord_Query.asp'">
									<input name="btnexit" type="button" value=" 關 閉 " onclick="funExit();">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">罕見字回報紀錄列表
		　　<b>( 查詢 <%=DBsum%> 筆紀錄 )</b></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th>回報時間</th>
					<th>車號</th>
					<th>車主姓名</th>
					<th>車主地址</th>
					<th>處理進度</th>
					<th>操作</th>
				</tr><%
				if trim(request("DB_State"))="Selt" then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move cdbl(DBcnt)
					for i=DBcnt+1 to DBcnt+10
						if rsfound.eof then exit for
						response.write "<tr align='center' bgcolor='#FFFFFF'"
						lightbarstyle 0
						response.write ">"
						Response.Write "<td>"&gInitDT(rsfound("TD_RecordDate"))&"</td>"
						Response.Write "<td>"&rsfound("TD_CARNO")&"</td>"
						Response.Write "<td>"&rsfound("TD_OwnerName")&"</td>"
						Response.Write "<td>"&rsfound("TD_ADDRESS")&"</td>"
						Response.Write "<td>"&rsfound("TD_ProcessName")&"</td>"
						Response.Write "<td>"

						response.write "<input type=""button"" name=""Del"" value=""刪除"" onclick=""funDel('"&rsfound("TD_SN")&"');"""

						'If trim(rsfound("TD_Process"))<>"0" Then Response.Write " disabled"
						
						Response.Write ">"
						Response.Write "</td>"
						Response.Write "</tr>"
						rsfound.movenext
					next
					rsfound.close
				end if%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFDD77" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=cdbl(DBcnt)/10+1&"/"&fix(cdbl(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="匯出上傳檔" onclick="funchgTxt();">	
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_State" value="<%=request("DB_State")%>">
<input type="Hidden" name="DB_Del" value="">
<input type="Hidden" name="TD_SN" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="/traffic/js/date.js"></script>
<script language="javascript">
function funSelt(){ 
	myForm.DB_State.value="Selt";
	myForm.submit();
}
function funAdd(){
	newWin("","inputWin",800,250,50,10,"yes","yes","yes","no");
	myForm.action="RareWord_Insert.asp";
	myForm.target="inputWin";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funchgTxt(){
	UrlStr="RareWord_List.asp";
	myForm.action=UrlStr;
	myForm.target="CHGH";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funDel(SN){
	if(confirm('確定刪除此筆紀錄嗎？')){
		myForm.TD_SN.value=SN;
		myForm.DB_Del.value='Del';
		myForm.submit();
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}
}
function funExit(){
	self.close();
}
</script>
<%
conn.close
set conn=nothing
%>