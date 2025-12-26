<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/BannernoData.asp"-->
	<!--#include File="../Common/css.txt"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body {font-family:新細明體;font-size:10pt}
A:link {text-decoration : none;color=0044ff;line-height:16px;font-size:10pt}
A:visited {text-decoration : none;color=0044ff;line-height:16px;font-size:10pt}
A:hover {text-decoration : underline;color=ff6600;line-height:16px;font-size:10pt}
td {font-family:新細明體;line-height:16px;font-size:10pt}
input {font-family:新細明體;line-height:16px;font-size:10pt}
select {font-family:新細明體;line-height:16px;font-size:10pt}
-->
</style>
<%
	'要到ApConfigure抓法條版本
	strRuleVer="select Value from ApConfigure where ID=3"
	set rsRuleVer=conn.execute(strRuleVer)
	if not rsRuleVer.eof then
		theRuleVer=trim(rsRuleVer("Value"))
	end if
	rsRuleVer.close
	set rsRuleVer=nothing

if trim(request("kinds"))="img_Upload" then
	'Ftp連結位置
	FtpLocation=""
	strftp="select Value from ApConfigure where ID=37"
	set rsftp=conn.execute(strftp)
	if not rsftp.eof then
		FtpLocation=trim(rsftp("Value"))
	end if
	rsftp.close
	set rsftp=nothing
	'影像存放位置
	FileLocation=""
	strfile="select Value from ApConfigure where ID=36"
	set rsFile=conn.execute(strfile)
	if not rsFile.eof then
		FileLocation=trim(rsFile("Value"))
	end if
	rsFile.close
	set rsFile=nothing
	
	if trim(request("UseTool"))="3" then
		FileLocation=FileLocation&"upload\"
		'FtpLocation=FtpLocation&"Type3/"
	elseif trim(request("UseTool"))="4" then
		FileLocation=FileLocation&"Type4\"
		'FtpLocation=FtpLocation&"Type4/"
	elseif trim(request("UseTool"))="5" then
		FileLocation=FileLocation&"Type5\"
		'FtpLocation=FtpLocation&"Type5/"
	end if
	'資料夾名稱
		FileTime=year(now)&Right("00"&month(now),2)&Right("00"&day(now),2)&Right("00"&hour(now),2)&Right("00"&minute(now),2)&Right("00"&second(now),2)
		FileLocation=FileLocation&FileTime
		FtpLocation=FtpLocation&FileTime

'=========資料存入資料庫=========
'上傳張數幾張就要存幾筆
	LoopNum=cint(trim(request("UploadCount")))
	if trim(request("LimitSpeed"))="" then
		LimitSpeedTmp="null"
	else
		LimitSpeedTmp=trim(request("LimitSpeed"))
	end if
	if trim(request("TriggerSpeed"))="" then
		TriggerSpeedTmp="null"
	else
		TriggerSpeedTmp=trim(request("TriggerSpeed"))
	end if
	if trim(request("Line"))="" then
		LineTmp="null"
	else
		LineTmp=trim(request("Line"))
	end if
	if trim(request("IllegalType"))="S" then	'一張
		for qq=1 to LoopNum 
			strPI="insert into ProsecutionImage(FileName,DirectoryName,FixequipType,SiteCode" &_
				",ProsecutionTime,ProsecutionTypeID,Location,District,OperatorA,LimitSpeed,TriggerSpeed" &_
				",RadarID,Direction,Line,ImageFileNameA,ImageFileNameB)" &_
				" values('"&FileTime&Right("0000"&qq,4)&"','"&FileLocation&"\','"&trim(request("UseTool"))&"'" &_
				",'"&trim(request("IllegalAddressID"))&"',sysdate,'"&trim(request("IllegalType"))&"'" &_
				",'"&trim(request("IllegalAddress"))&"','臺灣','"&trim(request("Operator"))&"'" &_
				","&LimitSpeedTmp&","&TriggerSpeedTmp&",'"&trim(request("RadarID"))&"'" &_
				",'"&trim(request("Direction"))&"',"&LineTmp&",'"&FileTime&Right("0000"&qq,4)&".jpg'" &_
				",null)"
			conn.execute strPI

			strPID="Insert into ProsecutionImageDetail(FileName,SN,LawItemID,VerifyResultID)" &_
				" values('"&FileTime&Right("0000"&qq,4)&"',1,'"&trim(request("Rule1"))&"',1)"
			conn.execute strPID
		next
	else	'兩張
		CaseCnt=1
		for qq=1 to LoopNum 
			strPI="insert into ProsecutionImage(FileName,DirectoryName,FixequipType,SiteCode" &_
				",ProsecutionTime,ProsecutionTypeID,Location,District,OperatorA,LimitSpeed,TriggerSpeed" &_
				",RadarID,Direction,Line,ImageFileNameA,ImageFileNameB)" &_
				" values('"&FileTime&Right("0000"&CaseCnt,4)&"','"&FileLocation&"\','"&trim(request("UseTool"))&"'" &_
				",'"&trim(request("IllegalAddressID"))&"',sysdate,'"&trim(request("IllegalType"))&"'" &_
				",'"&trim(request("IllegalAddress"))&"','臺灣','"&trim(request("Operator"))&"'" &_
				","&LimitSpeedTmp&","&TriggerSpeedTmp&",'"&trim(request("RadarID"))&"'" &_
				",'"&trim(request("Direction"))&"',"&LineTmp&",'"&FileTime&Right("0000"&CaseCnt,4)&".jpg'" &_
				",'"&FileTime&Right("0000"&(CaseCnt+1),4)&".jpg')"
			conn.execute strPI

			strPID="Insert into ProsecutionImageDetail(FileName,SN,LawItemID,VerifyResultID)" &_
				" values('"&FileTime&Right("0000"&CaseCnt,4)&"',1,'"&trim(request("Rule1"))&"',1)"
			conn.execute strPID
			CaseCnt=CaseCnt+2
		next
	end if
'=========處理新增資料夾及上傳=========
	'建立資料夾

	set fs=CreateObject("Scripting.FileSystemObject")
	set aa=fs.CreateFolder(FileLocation)

%>
		<script language="JavaScript">
			//alert ("請將欲上傳之影像檔拖曳至FTP視窗中!!");
			window.open("<%=FtpLocation%>","FtpWin","location=0,width=770,height=455,resizable=yes,scrollbars=yes,menubar=yes");
		</script>     
<%
	set fs=nothing
end if
%>
<title>違規數位影像上傳</title>
<script type="text/javascript" src="../js/form.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onContextmenu="return false">
	<form name=myForm method="post">  
		<table width='600' border='0' bgcolor="dddddd" align="center" >
			<tr bgcolor="#FFFFFF">
				<td colspan="2" bgcolor="#FFCC33" height="27" ><strong><span class="font12">違規數位影像上傳</font></strong></td>
			</tr>		
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" width="30%" height="27" align="right"><span class="font12"> 上傳人員</span></td>
				<td width="70%"><span class="font12"><%
				response.write Session("Ch_Name")
				%></span></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">※ 上傳件數</span></td>
				<td>
					<input type="text" name="UploadCount" size="8" value="" onkeyup="value=value.replace(/[^\d]/g,'')"><span class="font12">件</span>   
					<font color="#ff000">(檔名為四碼流水號，闖紅單兩張算一件)
					<br>
					<a href='BatchProcessFile.rar'> <span class="font12"> 下載批次轉檔軟體</span></a></font>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">※ 違規類別</span></td>
				<td>
					<input type="radio" name="IllegalType" value="S" checked>超速(單張)
					<input type="radio" name="IllegalType" value="R">闖紅燈(兩張)

					<input type="hidden" name="UseTool" value="3">
					<!-- <input type="radio" name="UseTool" value="4">固定桿
					<input type="radio" name="UseTool" value="5">雷射測速 -->
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">※ 舉發人</span></td>
				<td>
					<span class="font12">單位：</span>
					<select name="UnitID" onchange="getMemberList()">
						<option value="">==== 請選擇 ==== </option>
<%
				strUnit="select UnitID,UnitName from UnitInfo order by UnitID"
				set rsUnit=conn.execute(strUnit)
				If Not rsUnit.Bof Then rsUnit.MoveFirst 
				While Not rsUnit.Eof
%>
						<option value="<%=trim(rsUnit("UnitID"))%>" <%if trim(rsUnit("UnitID"))=trim(Session("Unit_ID")) then response.write "selected"%>><%=trim(rsUnit("UnitName"))%></option>
<%
				rsUnit.MoveNext
				Wend
				rsUnit.close
				set rsUnit=nothing
%>
					</select>&nbsp;&nbsp;
					<span class="font12">舉發人：</span>
					<select name="Operator">
						<option value="">== 請先選擇單位 ==</option>
<%
				strMember="select MemberID,chName from MemberData where UnitID='"&trim(Session("Unit_ID"))&"' order by MemberID"
				set rsMember=conn.execute(strMember)
				If Not rsMember.Bof Then rsMember.MoveFirst 
				While Not rsMember.Eof
%>
						<option value="<%=trim(rsMember("chName"))%>" <%if trim(rsMember("MemberID"))=trim(Session("User_ID")) then response.write "selected"%>><%=trim(rsMember("chName"))%></option>
<%
				rsMember.MoveNext
				Wend
				rsMember.close
				set rsMember=nothing
%>
					</select>
				</td>
			</tr>
			<!-- <tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">限速</span></td>
				<td>
					<input type="text" name="LimitSpeed" size="8" value="">
				</td>
			</tr> -->
<!-- 			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">觸發時速</span></td>
				<td> -->
					<input type="hidden" name="TriggerSpeed" size="8" value="">
<!-- 				</td>
			</tr> -->
			<!-- <tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">違規車道</span></td>
				<td>
					<input type="text" name="Line" size="8" value="" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
			</tr> -->
			<!-- <tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">違規地點</span></td>
				<td>
					<input type="text" name="IllegalAddressID" size="8" value="" onkeyup="getillStreet()">
					<input type="button" name="b1" value="？" onclick='window.open("../BillKeyIn/Query_Street.asp","WebPage3","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
					<input type="text" name="IllegalAddress" size="35" value="">
				</td>
			</tr> -->
			<!-- <tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">拍照方向</span></td>
				<td>
					<input type="radio" name="Direction" value="車頭" checked><span class="font12">車頭</span>
					<input type="radio" name="Direction" value="車尾"><span class="font12">車尾</span>
				</td>
			</tr> -->
			<!-- <tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">違規類型</span></td>
				<td>
					<input type="radio" name="ProsecutionTypeID" value="S" checked><span class="font12">超速</span>
					<input type="radio" name="ProsecutionTypeID" value="R"><span class="font12">闖紅燈</span>
					<input type="radio" name="ProsecutionTypeID" value="U"><span class="font12">違規左右轉</span>
				</td>
			</tr> -->
			<!-- <tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right" height="32"><span class="font12">違規法條</span></td>
				<td>
					<input type="text" name="Rule1" size="13" value="" onkeyup="getRuleData1()">
					<input type="button" value="？" name="LawSelect" onclick='window.open("Query_LawforUpload.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=550,height=355,resizable=yes,scrollbars=yes")'>
					<div id="Layer1" style="position:absolute ; width:280px; height:24px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" name="ForFeit1" value="<%=request("ForFeit1")%>">
				</td>
			</tr> -->
			<!-- <tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFCC" align="right"><span class="font12">雷達序號</span></td>
				<td>
					<input type="text" name="RadarID" size="13" value="">
				</td>
			</tr> -->
			<tr> 
				<td colspan="2" align='center'>
					<input type="button" name="sub1" value="確 定" onclick="UploadImg()">
					<input type="button" name="cancel" value="關 閉" onclick="window.close()">
					<input type="hidden" name="kinds" value="">
				</td>
			</tr>		
		</table>
		<br>

	</form>
</body>
<%
conn.close
set conn=nothing
%>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var TDLawErrorLog1=0;
function UploadImg(){
	var error=0;
	var errorString="";
	if (myForm.UploadCount.value==""){
		error=error+1;
		errorString=error+"：請輸入上傳張數。";
	}
	if (myForm.UseTool.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入採證工具。";
	}
	if (myForm.Operator.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入舉發人。";
	}
	if (TDLawErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條輸入錯誤。";
	}
	if (error==0){
		myForm.kinds.value="img_Upload";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
function getMemberList(){
	if (myForm.UnitID.value!=""){
		var UnitID=myForm.UnitID.value;
		myForm.Operator.length=0;
		runServerScript("getMemberDataList.asp?UnitID="+UnitID);
	}else{
		myForm.Operator.options[myForm.Operator.length]=new Option("== 請先選擇單位 ==","");
		myForm.Operator.length=1;
	}
}
function setMemberDataList(code,value){
		myForm.Operator.options[myForm.Operator.length]=new Option(value,code);
		myForm.Operator.length=myForm.Operator.length;
}
//違規地點代碼(ajax)
function getillStreet(){
	if (myForm.IllegalAddressID.value.length > 4){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("../BillKeyIn/getIllStreet.asp?illAddrID="+illAddrNum);
	}
}
//違規事實1(ajax)
function getRuleData1(){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("../BillKeyIn/getRuleDetail.asp?RuleOrder=1&RuleID="+Rule1Num+"&RuleVer="+VerNo);
	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=0;
	}
}
</script>
</html>
