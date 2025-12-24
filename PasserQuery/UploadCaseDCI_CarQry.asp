<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html onkeydown="KeyDown()">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--

.style4 {
	color: #FF0000;
	font-size: 16px
	}

-->
</style>
<title>微電車車查</title>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include file="../Common/Banner.asp"-->
<% Server.ScriptTimeout = 800 %>
<%
'權限
'AuthorityCheck(250)

'抓縣市
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

strwhere=""

if request("RecordDate")<>"" and request("RecordDate1")<>""then
	RecordDate1=gOutDT(request("RecordDate"))&" 0:0:0"
	RecordDate2=gOutDT(request("RecordDate1"))&" 23:59:59"

	strwhere=strwhere&" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"

end If 


if trim(request("RecordDate_h"))<>"" or trim(request("RecordDate1_h"))<>"" then
	strwhere=strwhere&" and to_char(a.RecordDate,'hh') between "&trim(request("RecordDate_h"))&" and "&trim(request("RecordDate1_h"))
end if

if request("BilltypeID")<>"" then
	
	if trim(request("BilltypeID"))="3" then
		strwhere=strwhere&" and a.BillTypeID=2 and a.usetool=8"
	else
		strwhere=strwhere&" and a.BillTypeID="&request("BilltypeID")
	End if 

end If 

if request("Sys_RecordUnit")<>"" then

	strwhere=strwhere&" and a.RecordMemberID in (select MemberID from MemberData where UnitID in ('"&trim(request("Sys_RecordUnit"))&"'))"
end If 

if request("Sys_RecordMemberID")<>"" then
	strwhere=strwhere&" and a.RecordMemberID ="&request("Sys_RecordMemberID")
end If 

if request("Sys_BillNo")<>"" then

	strwhere=strwhere&" and a.BillNo='"&request("Sys_BillNo")&"'"
end If 

if request("Sys_CarNo")<>"" then

	strwhere=strwhere&" and a.CarNo like '%"&request("Sys_CarNo")&"%'"
end If 


if request("sys_BatcuNumber")<>"" then

	strwhere=strwhere&" and a.sn in(select billsn from PASSERDCILOG where batchnumber='"&request("sys_BatcuNumber")&"')"
end If 

'入案(遇到RecordStateID=-1不做)
if trim(request("kinds"))="BillToDCILog" then

	chkcnt=0
	strSQL="select count(1) cnt  " & _
		"from PasserBase a where carno is not null and billtypeid=2 and DCILOGSN is null " & _
		"and recordstateid=0 " & strwhere 
	set rscnt=conn.execute(strSQL)

	chkcnt=cdbl(rscnt("cnt"))
	rscnt.close

	If chkcnt > 0 Then

		strSN="select PASSERDCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"A"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing	

		strSQL="select SN,billno,billtypeid,carno,billunitid " & _
		"from PasserBase a where carno is not null and billtypeid=2 and DCILOGSN is null " & _
		"and recordstateid=0 " & strwhere & _
		" order by a.RecordDate"

		set rs=conn.execute(strSQL)
		While not rs.eof

			DCISN=""
			strDciSN="select passerDCILOG_SEQ.nextval as SN from Dual"
			set rsSN=conn.execute(strDciSN)
			if not rsSN.eof then
				DCISN=trim(rsSN("SN"))
			end if
			rsSN.close
			set rsSN=nothing
			
			strInsCaseIn="insert into PASSERDCILOG(" & _
				"SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" & _
				",RecordMemberID,ExchangeDate,ReturnMarkType,ExchangeTypeID,BatchNumber,DciUnitID)"&_
				"values(" & DCISN & ","&rs("sn")&",'"&rs("billno")&"'"&_
				","&rs("billtypeid")&",'"&rs("carno")&"','"&rs("billunitid")&"'"&_
				",sysdate,"&Session("User_ID")&",sysdate,1,'A','"&theBatchTime&"'" &_
				",(" &_
					"select DciUnitID from UnitInfo ut where UnitID=(" &_
						"select unittypeid from UnitInfo uta where unitid='"&rs("billunitid")&"'" &_
					")" &_
				")" &_
			")" 

			conn.execute strInsCaseIn

			strInsCaseIn="insert into PasserBaseDciReturn(" & _
				"DciLogSN,BillSN,BillNO,CarNo,ExchangeTypeID)"&_
				"values(" & DCISN & ","&rs("sn")&",'"&rs("billno")&"','"&rs("carno")&"','A')" 

			conn.execute strInsCaseIn



			sqlpasserbase="update PasserBase set DCILOGSN=" & DCISN & " where sn="&rs("sn")

			conn.execute sqlpasserbase

			rs.movenext
		Wend
		rs.close

		Response.write "<script>"
		Response.Write "alert(""入案處理完成，批號："&theBatchTime&""");"
		Response.write "</script>"
		
	End if 
	
end if

'組成查詢SQL字串
if request("DB_Selt")="Selt" then		

	strSQL="select SN,a.IllegalDate," & _
	"(select chname from memberdata where memberid=a.billmemid1) billmem1," & _
	"(select chname from memberdata where memberid=a.billmemid2) billmem2," & _
	"(select chname from memberdata where memberid=a.billmemid3) billmem3," & _
	"(select chname from memberdata where memberid=a.billmemid4) billmem4," & _
	"a.BillNo,a.CarNo,a.CarSimpleID,a.BillTypeID,a.Driver,a.Rule1,a.Rule2,a.Rule3,a.BillStatus " & _
	"from PasserBase a where carno is not null and billtypeid=2 and DCILOGSN is null " & strwhere & _
	" order by a.RecordDate"

	set rsfound=conn.execute(strSQL)

	DBsum=0

	strcnt="select count(1) cnt  " & _
		"from PasserBase a where carno is not null and billtypeid=2 and DCILOGSN is null " & _
		"and recordstateid=0 " & strwhere 
	set rscnt=conn.execute(strcnt)

	DBsum=cdbl(rscnt("cnt"))
	rscnt.close
	
end If 
%>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">微電車車查</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						建檔日期
						<input name="RecordDate" type="text" value="<%
							if isempty(request("DB_Selt")) Then
								RecordDateTmp=Year(DateAdd("d",-60,now))-1911&Right("00" & Month(DateAdd("d",-60,now)),2)&Right("00" & Day(DateAdd("d",-60,now)),2)
							else
								RecordDateTmp=trim(request("RecordDate"))
							end if
							response.write RecordDateTmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate');">
						~
						<input name="RecordDate1" type="text" value="<%
							if trim(request("DB_Selt"))="" then
								RecordDate1Tmp=ginitdt(now)
							else
								RecordDate1Tmp=trim(request("RecordDate1"))
							end if
							response.write RecordDate1Tmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
						
						
						<!--時段-->
						<input name="RecordDate_h" type="hidden" value="<%=request("RecordDate_h")%>" size="1" maxlength="2" class="btn1" onKeyup="chknumber(this);"> <!-- 時 ~ -->
						<input name="RecordDate1_h" type="hidden" value="<%=request("RecordDate1_h")%>" size="1" maxlength="2" class="btn1" onKeyup="chknumber(this);"><!-- 時 -->
						 &nbsp;　&nbsp;

						舉發單類別
						<select name="BilltypeID">
							<option value="">請選擇</option>
							<option value="1" <%
							if trim(request("BilltypeID"))="1" then response.write "selected"
							%>>攔停</option>
							<option value="3" <%
							if trim(request("BilltypeID"))="3" then response.write "selected"
							%>>逕舉手開單</option>
							<option value="2" <%
							if trim(request("BilltypeID"))="2" then response.write "selected"
							%>>逕舉</option>
						</select> &nbsp;　&nbsp;

						單號
						<input type="text" name="Sys_BillNo" size="10" maxlength="9" value="<%=trim(request("Sys_BillNo"))%>" onkeyup="value=value.toUpperCase()">
						
						<br>
						建檔單位
						<%=SelectUnitOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						&nbsp;　&nbsp;
						建檔人
						<%=SelectMemberOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						
						&nbsp;　&nbsp;
						<!--
						<font color="red"><strong>車籍查詢批號</strong></font>
						<Select Name="Selt_BatchNumber" onchange="fnBatchNumber();">
							<option value="">請點選</option><%
							strSQL1="select distinct TO_char(ExchangeDate,'YYYY/MM/DD') ExchangeDate,BatchNumber from PasserDCILog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",-5, date)&" 00:00"&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') and ExchangeTypeID='A' order by ExchangeDate DESC"
		
							set rs=conn.execute(strSQL1)
							cut=0
							while Not rs.eof
								ExchangeDate=gInitDT(trim(rs("ExchangeDate")))
								response.write "<option value="""&trim(rs("BatchNumber"))&""">"
								response.write ExchangeDate& " - "&cut&"　"&trim(rs("BatchNumber"))
								response.write "</option>"
								cut=cut+1
								rs.movenext
							wend
							rs.close
						%>
						</select>
						<input type="text" name="sys_BatcuNumber" size="8" value="<%=trim(request("sys_BatcuNumber"))%>" onkeyup="value=value.toUpperCase()">
						&nbsp;　&nbsp;
						-->
						車號
						<input type="text" name="Sys_CarNo" size="8" maxlength="9" value="<%=trim(request("Sys_CarNo"))%>" onkeyup="value=value.toUpperCase()">
	
						
						&nbsp;　&nbsp;
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();">
						<input type="button" name="cancel" value="清除" onClick="location='UploadCaseDCI_CarQry.asp';"> 

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
			&nbsp;&nbsp;
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th width="8%">違規日期</th>
					<th width="8%">舉發員警</th>
					<th width="6%">舉發單號</th>
					<th width="5%">車號</th>
					<th width="6%">車種</th>
					<th width="4%">類別</th>
					<th width="6%">駕駛人</th>
					<th width="10%">法條</th>
					<th width="8%">DCI</th>
				</tr>
				<%
				chkCaseInDelayFlag=0
				CaseInDelayBillNo=""
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
						if rsfound("BillMem1")<>"" then chname=rsfound("BillMem1")
						if rsfound("BillMem2")<>"" then chname=chname&"/"&rsfound("BillMem2")
						if rsfound("BillMem3")<>"" then chname=chname&"/"&rsfound("BillMem3")
						if rsfound("BillMem4")<>"" then chname=chname&"/"&rsfound("BillMem4")
						if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
						if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
						if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='30'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td width='5%'>"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
						response.write "<td width='8%'>"&chname&"</td>"


						response.write "<td width='6%'>"&rsfound("BillNo")&"</td>"

						response.write "<td width='6%'>"&rsfound("CarNo")&"</td>"
						response.write "<td width='5%'>"

							if trim(rsfound("CarSimpleID"))="8" then
								response.write "微電車"
							end If
						
						response.write "</td>"
						response.write "<td width='4%'>"
					strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
					set rsBTypeVal=conn.execute(strBTypeVal)
					if not rsBTypeVal.eof then
						response.write rsBTypeVal("Content")
					end if
					rsBTypeVal.close
					set rsBTypeVal=nothing
						response.write "</td>"

						response.write "<td width='6%'>"&rsfound("Driver")&"</td>"

						response.write "<td width='10%'>"&chRule&"</td>"
						response.write "<td width='8%'>"
						if trim(rsfound("BillStatus"))="0" then
							response.write "<font color='#999999'>建檔</font>"
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
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFDD77" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
			<input type="button" name="Submit4242" value="進行車籍查詢" onclick="BillToDCILog();">

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
<input type="Hidden" name="PKICarchk" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	<%response.write "UnitMan('Sys_RecordUnit','Sys_RecordMemberID','"&trim(request("Sys_RecordMemberID"))&"');"%>

	function funSelt(){
		var error=0;
		var errorString="";

		if(myForm.RecordDate.value==""||myForm.RecordDate1.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：建檔日期須填寫!!";
		}

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
	function repage(){
		myForm.DB_Move.value=0;
		myForm.submit();
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
	//入案
	function BillToDCILog(){
		var Billsum="<%=DBsum%>";

		if (myForm.DB_Selt.value==""){

			alert("請先查詢欲入案的舉發單！");

		}else if (Billsum=="0"){

			alert("查無可入案之舉發單！");

		}else{
			if (myForm.Sys_RecordMemberID.value==""){
				if(confirm('您選擇將所有建檔人的舉發單資料，是否確定要進行車查？')){
					myForm.kinds.value="BillToDCILog";
					myForm.submit();
				}
			}else{
				if(confirm('確定要進行車查嗎？')){
					myForm.kinds.value="BillToDCILog";
					myForm.submit();
				}
			}
		}
	}
	
	function fnBatchNumber(){
		myForm.sys_BatcuNumber.value=myForm.Selt_BatchNumber.value;
	}

	function KeyDown(){ 

		if (event.keyCode==116){	//F5鎖死
			event.keyCode=0;   
			event.returnValue=false;   
		}
	}

</script>
<%
conn.close
set conn=nothing
%>