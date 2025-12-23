
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="/traffic/Common/DB.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>結案查詢系統</title>
<!--#include virtual="/traffic/Common/css.txt"-->
<% Server.ScriptTimeout = 800 %>
<%
'防止查詢逾時
Server.ScriptTimeout = 800
Response.flush
'權限
RecordDate=split(gInitDT(date),"-")

function QuotedStr(Str)
    QuotedStr="'"+Str+"'"
end function


'組成查詢SQL字串
if request("DB_Selt")="Selt" then
    strwhere=""
    '查詢車輛
    CarSql  = ""
		if trim(request("CarNo"))<>"" then
				CarSql = " and a.CarNo =" & QuotedStr(trim(request("CarNo")))
		end if
    
    '查詢單號
    BillNoSql  = ""
		if trim(request("Billno"))<>"" then
				BillNoSql = " and a.Billno =" & QuotedStr(trim(request("Billno")))
		end if
    
   strSQL="Select a.* from DCICloseCloseData a,BillBase b where a.Billno=B.billno " &  CarSql   & BillNoSql
     '取筆數
   
   if ((CarSQL <> "") or (BillNoSql <> ""))  then 
    strCnt="Select count(a.PreMonth) as FieldCount from DCICloseCloseData a,BillBase b where a.Billno=B.billno " &  CarSql   & BillNoSql
    set Dbrs=conn.execute(strCnt)

    DBsum=Dbrs("FieldCount")
    Dbrs.close
    set rsfound=conn.execute(strSQL)
    Session.Contents.Remove("BillSQL")
    Session("BillSQL")=strSQL 
   end if
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
		<td bgcolor="#FFCC33">結案查詢系統</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
                        
            單號：
						<input name="Billno" type="text" class="txt1" value="<%=trim(request("Billno"))%>" size="20" maxlength="150">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
                        
						<img src="space.gif" width="8" height="10">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" id="Button1">
						<input type="button" name="cancel" value="清除" onClick="location='CaseCloseQry.asp'"> 
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" class="style3">
			結案查詢列表
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
					<th width="8%">單號</th>
					<th width="8%">車號</th>
					<th width="8%">法條</th>
					<th width="6%">金額</th>
					<th width="10%">違規日期</th>
					<th width="10%">結案日期</th>
					<th width="4%">操作</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
				if ((request("DB_Selt")="Selt") and ((CarSQL <> "") or (BillNoSql <> ""))) then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")                      
						if rsfound.eof then exit for
    					response.write "<tr bgcolor='#FFFFFF' align='center'  height='30'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td width='8%' >" & trim(rsfound("BillNo")) & "</td>"
						response.write "<td width='8%' >" & trim(rsfound("CarNo")) & "</td>"
						'有兩個法條的話都顯示出來
						if (trim(rsfound("Rule2")) <> "0") and (trim(rsfound("Rule2")) <> "") then
						    response.write "<td width='8%' >" & trim(rsfound("Rule1")) & "," & trim(rsfound("Rule2"))  & "</td>"
						else
						    response.write "<td width='8%' >" & trim(rsfound("Rule1")) & "</td>"
						end if
						
                        response.write "<td width='6%' >" & trim(rsfound("ForFeit")) & "</td>"						
						response.write "<td width='10%' >" & trim(rsfound("IllegalDate")) &  "</td>"	
						response.write "<td width='10%' >"  &  trim(rsfound("Close_Date")) & "</td>"
						response.write "<td width='4%' >"
				%>
				<input type="button" name="b1" value="詳細" onclick='window.open("CaseCloseDetail.asp?Billno=<%=trim(rsfound("Billno"))%>&CarNo=<%=trim(rsfound("CarNo"))%>","WebPage2","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")'  style="font-size: 10pt; width: 40px; height:26px;">
				<%
				        response.Write "</td>"
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
		<td bgcolor="#FFDD77" align="center" style="height: 35px">
			<a href="file:///.."></a>
			<a href="file:///......"></a>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
            &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;<span class="style3"><img src="space.gif" width="5" height="8"></span>
			<!--<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();" id="Button2">-->
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

<script type="text/javascript" src="../../js/date.js"></script>
<script language="javascript">

	function funSelt(){
		if ((myForm.CarNo.value == 0) && (myForm.Billno.value == 0))
		{
			alert("請填寫查詢條件，查詢條件最少一項!!");
		}
		else
		{
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
		UrlStr="TD08B02L.asp?WorkType=1";
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
	function dateCheck(Sys_date){
	var error=0;
	var error2=1;
	var date_y=0,date_m=0,date_y=0;
	if(Sys_date.length>5){
		date_y=eval(Sys_date.substr(0,eval(Sys_date.length)-4))+1911;
		date_m=Sys_date.substr(eval(Sys_date.length)-4,2);
		date_d=Sys_date.substr(eval(Sys_date.length)-2,2);
		if (date_m > 0 && date_m < 13){
			if (date_d > 0 && date_d <32){
				error2=0;
				if (date_m == 4 || date_m == 6 || date_m == 9 || date_m == 11){
					if (date_d > 30){
						error=1;
					}
				}
				if(date_m == 2){
					if ((date_y) % 4 == 0){
						if (date_d > 29) error=1;
					}else{
						if (date_d > 28) error=1;
					}
				}
				if(error==0){
					return true;
				}
			}
		}
		if(error2==1){
			return false;
		}
	}else{
		return false;
	}
}

<%'if trim(request("DB_Selt"))="Selt" then%>
//	funSelt();
<%'end if%>
</script>
<%
conn.close
set conn=nothing
%>

