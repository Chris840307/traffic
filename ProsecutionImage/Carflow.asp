<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="..\Common\DB.ini"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<!-- #include file="..\Common\Login_Check.asp"-->
<!-- #include file="..\Common\bannernodata.asp" -->
<%
AuthorityCheck(286)
if request("DB_Selt")="Selt" then
	strwhere=""
	if request("Sys_StartPassingDate1")<>"" and request("Sys_StartPassingDate2")<>""then
		ArgueDate1=gOutDT(request("Sys_StartPassingDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_StartPassingDate2"))&" 23:59:59"
		strwhere=" where StartPassingDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	elseif request("Sys_StartPassingDate1")<>"" then
		ArgueDate1=gOutDT(request("Sys_StartPassingDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_StartPassingDate1"))&" 23:59:59"
		strwhere=" where StartPassingDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	elseif request("Sys_StartPassingDate2")<>""then
		ArgueDate1=gOutDT(request("Sys_StartPassingDate2"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_StartPassingDate2"))&" 23:59:59"
		strwhere=" where StartPassingDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	strSQL="select StartPassingDate,EndPassingDate,Passings,Violations,MinSpeed,MaxSpeed from Passing"&strwhere
	set rsfound=conn.execute(strSQL)
	

	strSQL="select Sum(Passings) as PassingSum from Passing"&strwhere
	set Dbrs=conn.execute(strSQL)
	PassingSum=Dbrs("PassingSum")
	Dbrs.close

	strCnt="select count(*) as cnt from Passing"&strwhere
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close

	tmpSQL=strwhere
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>車流查詢</title>
<!-- #include file="..\Common\css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="25">車流查詢</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td>日期　
						<input name="Sys_StartPassingDate1" class="btn1" type="text" value="<%
							if request("DB_Selt")="Selt" then
								response.write request("Sys_StartPassingDate1")
							else
								response.write gInitDT(date)
							end if%>" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StartPassingDate1');">
						　～　日期　
						<input name="Sys_StartPassingDate2" class="btn1" type="text" value="<%
							if request("DB_Selt")="Selt" then
								response.write request("Sys_StartPassingDate2")
							else
								response.write gInitDT(date)
							end if%>" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StartPassingDate2');">
						　<input type="button" name="btnSelt" value="查詢" onClick='funSelt();'>&nbsp;&nbsp;
						<strong>( 查詢期間共 <%=PassingSum%> 輛車輛經過. )</strong>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="25">車流紀錄列表</td>
	</tr>
	<%if request("DB_Selt")="Selt" then%>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="34">開始日期</th>
					<th height="34">結束時間</th>
					<th height="34">通過車輛</th>
					<th height="34">違規車輛</th>
					<th height="34">最低時速</th>
					<th height="34">最高時速</th>
				</tr><%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move Cint(DBcnt)
					for i=DBcnt+1 to DBcnt+10
						if rsfound.eof then exit for
						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						response.write ">"
						response.write "<td>"&rsfound("StartPassingDate")&"</td>"
						response.write "<td>"&rsfound("EndPassingDate")&"</td>"
						response.write "<td align='right'>"&rsfound("Passings")&"</td>"
						response.write "<td align='right'>"&rsfound("Violations")&"</td>"
						response.write "<td align='right'>"&rsfound("MinSpeed")&"</td>"
						response.write "<td align='right'>"&rsfound("MaxSpeed")&"</td>"
						response.write "</tr>"
						rsfound.movenext
					next%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFDD77" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=Cint(DBcnt)/10+1&"/"&fix(Cint(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
		</td>
	</tr>
	<%end if%>
</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSelt(){
	var err=0;
	var selchk=myForm.Sys_StartPassingDate1.value+myForm.Sys_StartPassingDate2.value;
	if(selchk!=''){
		if(myForm.Sys_StartPassingDate1.value!=""){
			if(!dateCheck(myForm.Sys_StartPassingDate1.value)){
				err=1;
				alert("日期輸入不正確!!");
			}
		}
		if (err==0){
			if(myForm.Sys_StartPassingDate2.value!=""){
				if(!dateCheck(myForm.Sys_StartPassingDate2.value)){
					err=1;
					alert("日期輸入不正確!!");
				}
			}
		}
		if (err==0){
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}else{
		alert('必須有查詢條件!!');
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
	if(myForm.DB_Selt.value!=''){
		UrlStr="Carflow.asp_Execel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
	}else{
		alert('必須有查詢條件!!');
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
</script>
<%conn.close%>