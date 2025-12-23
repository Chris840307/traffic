<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>高雄市舊資料查詢</title>
<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/db.ini"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<%

function QuotedStr(Str)
    QuotedStr="'"+Str+"'"
end Function

Function GetCarTypeName(CarTypeID)
	tmp=""
	tmp2=Trim(CarTypeID)
	If tmp2="1" Then 
		tmp="自大客車"
	ElseIf tmp2="4" Then 
		tmp="營大客車"
	ElseIf tmp2="7" Then 
		tmp="營小客車"
	ElseIf tmp2="A" Then 
		tmp="營交通車"
	ElseIf tmp2="G" Then 
		tmp="大型重機"
	ElseIf tmp2="P" Then 
		tmp="併裝車"
	ElseIf tmp2="W" Then 
		tmp="自小客"
	ElseIf tmp2="2" Then 
		tmp="自大貨車"
	ElseIf tmp2="5" Then 
		tmp="營大貨車"
	ElseIf tmp2="8" Then 
		tmp="租賃小客"
	ElseIf tmp2="B" Then 
		tmp="貨櫃曳引"
	ElseIf tmp2="E" Then 
		tmp="外賓小客"
	ElseIf tmp2="H" Then 
		tmp="重機"
	ElseIf tmp2="Q" Then 
		tmp="500cc重機"
	ElseIf tmp2="X" Then 
		tmp="動力機械"
	ElseIf tmp2="3" Then 
		tmp="自小客貨"
	ElseIf tmp2="6" Then 
		tmp="營小貨車"
	ElseIf tmp2="9" Then 
		tmp="遊覽客車"
	ElseIf tmp2="C" Then 
		tmp="自用拖車"
	ElseIf tmp2="F" Then 
		tmp="外賓大客"
	ElseIf tmp2="L" Then 
		tmp="輕機"
	ElseIf tmp2="V" Then 
		tmp="自小貨"
	ElseIf tmp2="Y" Then 
		tmp="租賃小貨"
	End If
	If tmp="" Then tmp=CarTypeID
	GetCarTypeName=tmp

End function

if request("DB_Selt")="Selt" then
    DateSQL=""
    '違規日期 
    if request("IllegalDate")<>"" and request("IllegalDate1")<>"" then
		DateSQL=" and IllegalDate between  " & QuotedStr(trim(request("IllegalDate"))) &" and " & QuotedStr(trim(request("IllegalDate1")))
    end if		

	if request("BillNo")<>"" then DateSQL=DateSQL&" and BillNo = " & QuotedStr(trim(request("BillNo")))

	if request("CarNo")<>"" then DateSQL=DateSQL&" and CarNo = " & QuotedStr(trim(request("CarNo")))

	if request("Driver")<>"" then DateSQL=DateSQL&" and Driver = " & QuotedStr(trim(request("Driver")))

	strSQL="Select Billno,Carno,CarTypeID,Driver,IllegalDate,IllegalPlaceID,IllegalPlace,Rule1 from OldBillBase where 1=1  "  & DateSQL

    set rsfound=conn.execute(strSQL)

	strCnt="Select count(billno) as FieldCount from OldBillBase where 1=1 "  & DateSQL

    set cnt=conn.execute(strCnt)
	DBsum=cnt("FieldCount")
	set cnt=nothing
end if

%>

<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style12 {
	font-size: 20px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr height="25">
					<td bgcolor="#FFCC33">
					    <b>高雄市舊資料查詢</b>
					    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					    <img src="space.gif" width="20" height="2"> <A HREF="..\舊資料查詢系統.doc"><FONT SIZE="3"><b>!!  第一次使用請看.DOC !! </b> </font></A>
					</td>
				</tr>		
				
				<tr>
					<td style="height: 36px">
						違規日期
						<input name="IllegalDate" type="text" value="<%=request("IllegalDate")%>" size="6" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate');">
						~
						<input name="IllegalDate1" type="text" value="<%=request("IllegalDate1")%>" size="6" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate1');">
                        <img src="space.gif" width="20" height="2">
						違規人證號
						<input name="DriverID" type="text" value="<%=request("DriverID")%>" size="9" maxlength="10" class="btn1" onkeyup="value=value.toUpperCase()">

						<img src="space.gif" width="20" height="2">
						車<img src="space.gif" width="20" height="2">號
						<input name="CarNo" type="text" value="<%=request("CarNo")%>" size="9" maxlength="8" class="btn1" onkeyup="value=value.toUpperCase()">					
						<img src="space.gif" width="20" height="2">
						<b>單<img src="space.gif" width="20" height="2">號</b>
						<input name="BillNo" type="text" value="<%=request("BillNo")%>" size="9" maxlength="9" class="btn1" onkeyup="value=value.toUpperCase()">					
					</td>
					<tr><td align="Center">
   						<img src="../Image/space.gif" width="15" height="1"><br>
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();">
						<input type="button" name="cancel" value="清除" onClick="location='OldKaoHsiungCityQuery.asp'"> 
					  </td>
					</td>					
				</tr>
			</table>
		<!--</td>-->
	<!--</tr>-->
	
	<tr height="30">
		<td bgcolor="#FFCC33" class="style3">
			資料紀錄列表
			<img src="space.gif" width="5" height="8">
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
			筆 <font color="#F90000"><strong>(共 <%=DBsum%> 筆 )</strong></font>
			&nbsp; &nbsp; 
			&nbsp;
			
<!--			<select name="sys_OrderType" onchange="repage();">
'				<option value="2" <%if trim(request("sys_OrderType"))="1" then response.write " Selected"%>>違規日期</option>
'				<option value="3" <%if trim(request("sys_OrderType"))="3" then response.write " Selected"%>>綜合資料號</option>
			</select>
			<select name="sys_OrderType2" onchange="repage();">
				<option value="1" <%if trim(request("sys_OrderType2"))="1" then response.write " Selected"%>>由小至大</option>
				<option value="2" <%if trim(request("sys_OrderType2"))="2" then response.write " Selected"%>>由大至小</option>
			</select>
			排列&nbsp; &nbsp;
-->			
		</td>
	</tr>
	
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th nowrap>舉發單號</th>
					<th nowrap>車號</th>
					<th >違規日</th>
					<th >車種</th>
					<th >駕駛人</th>
					<th >違規地點</th>
					<th >法條</th>
					<th >操作</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
               
				if request("DB_Selt")="Selt"  then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						                        
						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td width='10%'>" & trim(rsfound("Billno")) & "</td>"
                        response.write "<td width='10%'>" & trim(rsfound("Carno")) & "</td>"
                        response.write "<td width='8%'>" & trim(rsfound("IllegalDate")) & "</td>"						
						response.write "<td width='15%'>"& GetCarTypeName(trim(rsfound("CarTypeID"))) &"</td>"
						response.write "<td width='12%'>" & trim(rsfound("Driver")) &  "</td>"					
						response.write "<td width='40%' align='left'>" & trim(rsfound("IllegalPlaceID")) & "  " & trim(rsfound("IllegalPlace")) &  "</td>"					
						response.write "<td width='8%'>" & trim(rsfound("Rule1")) &  "</td>"					
						response.write "<td align='left' >"
                %>	
			    <input type="button" name="b1" value="詳細" onclick='window.open("OldKaoHsiungCityDetail.asp?BillNo=<%=trim(rsfound("BillNo"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=520,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
				<%
						response.write "</td>"
						response.write "</tr>"
						rsfound.movenext
					next
				end if
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFFFFF" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
		</td>
	</tr>
<!--</table>-->

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="tmpSQL" value="<%=tempSQL%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

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

		if (myForm.IllegalDate.value=="" && myForm.IllegalDate1.value=="" && myForm.DriverID.value==""  && myForm.CarNo.value=="" && myForm.BillNo.value==""  ) {
				error=error+1;
				errorString=errorString+"\n"+error+"：至少要輸入一項!!";
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
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
	}


	
</script>
<%
conn.close
set conn=nothing
%>