
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/bannernodata.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Log紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
'AuthorityCheck(225)
DB_Selt=request("DB_Selt")

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

if DB_Selt="Selt" Then
	
	strwhere=""
	strwhereM=""
	if request("Sys_CreditID")<>"" then
		strwhereM=" where CreditID = '"&request("Sys_CreditID")&"'"
	end if
	if request("Sys_ActionChName")<>"" then
		if strwhereM<>"" then
			strwhereM=strwhereM&" and ChName like '%"&request("Sys_ActionChName")&"%'"
		else
			strwhereM=" where ChName like '%"&request("Sys_ActionChName")&"%'"
		end if
	end If
	If strwhereM<>"" Then
		strwhere=strwhere&" and ActionMemberID in (select MemberID from MemberData "&strwhereM&")"
	End If 
	if request("Sys_IP")<>"" then
			strwhere=strwhere&" and ActionIP = '"&request("Sys_IP")&"'"
	end if
	if request("ActionDate")<>"" then
		ArgueDate1=gOutDT(request("ActionDate"))&" 0:0:0"
		ArgueDate2=gOutDT(request("ActionDate2"))&" 23:59:59"
			strwhere=strwhere&" and ActionDate between "&funGetDate(ArgueDate1,1)&" and "&funGetDate(ArgueDate2,1)
	end if
	if request("Sys_TypeID")<>"" then
			strwhere=strwhere&" and TypeID="&request("Sys_TypeID")
	end If
	If Trim(request("KeyWord"))<>"" Then
			strwhere=strwhere&" and ActionContent like '%"&Trim(request("KeyWord"))&"%' "
	End If 
	If Trim(request("ActionUnit"))<>"" Then
		strwhere=strwhere&" and ActionMemberID in (select MemberID from MemberData where UnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(request("ActionUnit"))&"')) "
	End If 

	If Trim(request("chkSmith"))<>"" Then
		strwhere=strwhere&" and not (Lower(ActionContent) like '%billno=%' or Lower(ActionContent) like '%billno =%' or Lower(ActionContent) like '%carno =%' or Lower(ActionContent) like '%carno=%' or Lower(ActionContent) like '%driverid =%' or Lower(ActionContent) like '%driverid=%') "
	End If 

	if trim(strwhere)<>"" then
		strSQL="select * from Log where sn is not null "&strwhere&" order by ActionDate"
		set rsfound=conn.execute(strSQL)
		
		strCnt="select count(*) as cnt from Log where sn is not null "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=cdbl(Dbrs("cnt"))
		Dbrs.close
		tmpSQL=strSQL
	else
		DB_Selt=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
%>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">Log紀錄</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" width="100%" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>身分證號</td>
				    <td nowrap><input name="Sys_CreditID" class="btn1" type="text" value="<%=request("Sys_CreditID")%>" size="10" maxlength="12"></td>

					<td nowrap>姓名</td>
				    <td nowrap><input name="Sys_ActionChName" class="btn1" type="text" value="<%=request("Sys_ActionChName")%>" size="6" maxlength="12"></td>

					<td nowrap>
					IP <input name="Sys_IP" class="btn1" type="text" value="<%=request("Sys_IP")%>" size="15" maxlength="15">
					</td>
					<td nowrap>
						單位 
						<select Name="ActionUnit">
				<%	
					If Trim(Session("UnitLevelID"))="1" Then
						response.write "<option value=''>所有單位</option>"
						strUnit="select * from UnitInfo where UnitID=UnitTypeID and ShowOrder<>-1 order by UnitID"
					Else
						strUnit="select * from UnitInfo where UnitID=(select UnitTypeID from UnitInfo where UnitID='"&Trim(Session("Unit_ID"))&"') and ShowOrder<>-1 order by UnitID"
					End If 
					Set rsUnit=conn.execute(strUnit)
					while Not rsUnit.eof
				%>
							<Option value="<%=Trim(rsUnit("UnitID"))%>" <%
							If Trim(request("ActionUnit"))=Trim(rsUnit("UnitID")) Then
								response.write "selected"
							End If 
							%>><%=Trim(rsUnit("UnitName"))%></option>
				<%
						rsUnit.movenext
					Wend
					rsUnit.close
					Set rsUnit=Nothing 
				%>
						</select>
					</td>
				</tr>
				<tr>
					<td nowrap>紀錄時間</td>
					<td nowrap>
						<input type="text" class="btn1" name="ActionDate" size="5" value="<%=request("ActionDate")%>" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ActionDate');">
						~
						<input type="text" class="btn1" name="ActionDate2" size="5" value="<%=request("ActionDate2")%>" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ActionDate2');">
					</td>
				    <td nowrap>類別</td>
					<td nowrap>
						<select name="Sys_TypeID" class="btn1">
							<option value="">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=12"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								response.write "<option value="""&rs("ID")&""""
								if trim(request("Sys_TypeID"))=trim(rs("ID")) then response.write " selected"
								response.write ">"
								response.write rs("Content")
								response.write "</option>"
								rs.movenext
							wend
							rs.close%>
							<option value="355" <%If Trim(request("Sys_TypeID"))="355" Then response.write "selected"%>>查詢</option>
							<option value="356" <%If Trim(request("Sys_TypeID"))="356" Then response.write "selected"%>>快速查詢</option>
							<option value="357" <%If Trim(request("Sys_TypeID"))="357" Then response.write "selected"%>>登入異常</option>
							<option value="358" <%If Trim(request("Sys_TypeID"))="358" Then response.write "selected"%>>帳號封鎖</option>
							<option value="359" <%If Trim(request("Sys_TypeID"))="359" Then response.write "selected"%>>人員異動</option>
					<%If sys_City="台南市" then%>
							<option value="360" <%If Trim(request("Sys_TypeID"))="360" Then response.write "selected"%>>列印</option>
					<%End If
					If sys_City="高雄市" then%>
							<option value="361" <%If Trim(request("Sys_TypeID"))="361" Then response.write "selected"%>>產生報表</option>
							<option value="362" <%If Trim(request("Sys_TypeID"))="362" Then response.write "selected"%>>產生報表(開始)</option>
							<option value="401" <%If Trim(request("Sys_TypeID"))="401" Then response.write "selected"%>>民眾檢舉匯入</option>
					<%End If
					%>
						</select>
					</td>
					<td nowrap colspan="2">
						關鍵字 <input type="text" class="btn1" name="KeyWord" value="<%=Trim(request("KeyWord"))%>" size="15">&nbsp;&nbsp;&nbsp;&nbsp;
						
						<input class="btn1" type="checkbox" name="chkSmith" onclick="funckIllegalDate()" value="1"<%
						If Not ifnull(request("chkSmith")) Then response.write " checked"%>>S條件&nbsp;&nbsp;

						<input type="button" name="btnSelt" value="查詢" onClick="funSelt();"<%'if Not CheckPermission(225,1) then response.write " disabled"%>>
						&nbsp;&nbsp;
						<input type="button" name="cancel" value="清除" onClick="location='Log.asp'">
					</td>
			    </tr>
			</table>
		</td>
	</tr>
</table>	
<%if DB_Selt="Selt" then%>
<table border="0" width="100%">
	<tr>
		<td colspan="6" bgcolor="#1BF5FF">Log紀錄列表</td>
	</tr>
	<tr bgcolor="#EBFBE3" align="center">
		<td height="34" width="10%">日期</td>
		<td height="34" width="10%">姓名</td>
		<td height="34" width="10%">身份證號</td>
		<td height="34" width="10%">IP</td>
		<td height="34" width="10%">類別</td>
		<td height="34" width="50%">內容</td>
	</tr><%
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
			response.write "<td nowrap width='10%'>"&gArrDT(rsfound("ActionDate"))&" "&Timevalue(rsfound("ActionDate"))&"</td>"
			response.write "<td nowrap>"
			CreditIDTmp=""
			If Trim(rsfound("ActionMemberID"))<>"" And Not IsNull(rsfound("ActionMemberID")) Then
				str="select * from MemberData where Memberid="&Trim(rsfound("ActionMemberID"))
				Set rs=conn.execute(str)
				If Not rs.eof Then
					response.write Trim(rs("ChName"))
					CreditIDTmp=left(Trim(rs("CreditID")),4)&"******"
				End If
				rs.close
				Set rs=Nothing 
			End If 
			response.write "</td>"
			response.write "<td nowrap width='10%'>"&CreditIDTmp&"</td>"
			response.write "<td nowrap width='10%'>"&rsfound("ActionIP")&"</td>"
			response.write "<td width='10%' >"
			'361報表
			If Trim(rsfound("TypeID"))="355" Then
				response.write "查詢"
			elseIf Trim(rsfound("TypeID"))="356" Then
				response.write "快速查詢"
			elseIf Trim(rsfound("TypeID"))="357" Then
				response.write "登入異常"
			elseIf Trim(rsfound("TypeID"))="358" Then
				response.write "帳號封鎖"
			elseIf Trim(rsfound("TypeID"))="359" Then
				response.write "人員異動"
			elseIf Trim(rsfound("TypeID"))="360" Then
				response.write "列印"
			elseIf Trim(rsfound("TypeID"))="361" Then
				response.write "產生報表"
			elseIf Trim(rsfound("TypeID"))="362" Then
				response.write "產生報表(開始)"
			elseIf Trim(rsfound("TypeID"))="401" Then
				response.write "民眾檢舉匯入"
			elseIf Trim(rsfound("TypeID"))<>"" And Not IsNull(rsfound("TypeID")) then
				str="select * from Code where id="&Trim(rsfound("TypeID"))&" and TypeID=12"
				Set rs=conn.execute(str)
				If Not rs.eof Then
					response.write Trim(rs("Content"))
				End If
				rs.close
				Set rs=Nothing 
			End If 
			response.write "</td>"
			response.write "<td width='50%' align='left'>"&Replace(rsfound("ActionContent"),"""","'")&"</td>"
			response.write "</tr>"
			rsfound.movenext
		next%>
	<tr>
		<td bgcolor="#1BF5FF" align="center" colspan="6">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=cdbl(DBcnt)/10+1&"/"&fix(cdbl(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
		</td>
	</tr>
</table>
<%end if%>
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSelt(){
	var error=0;
	if(myForm.ActionDate.value!=""){
		if(!dateCheck(myForm.ActionDate.value)){
			error=1;
			alert("紀錄時間輸入不正確!!");
		}
	}
	if(myForm.ActionDate2.value!=""){
		if(!dateCheck(myForm.ActionDate2.value)){
			error=1;
			alert("紀錄時間輸入不正確!!");
		}
	}
	if(myForm.ActionDate.value!="" || myForm.ActionDate2.value!=""){
		if(myForm.ActionDate.value=="" || myForm.ActionDate2.value==""){
			error=1;
			alert("紀錄時間請輸入開始及截止日期!!");
		}
	}
	if (error==0){
		if (error==0){
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}
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
function funchgExecel(){
	//UrlStr="Log_Execel.asp?SQLstr=<%=tmpSQL%>";
	//newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");

	
	UrlStr="Log_Execel.asp";
	myForm.action=UrlStr;
	myForm.target="log_execel";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}

</script>
<%conn.close%>