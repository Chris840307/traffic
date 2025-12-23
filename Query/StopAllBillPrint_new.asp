<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<!--#include virtual="traffic/Common/DCIURL.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>催繳單 / 各式清冊 列印</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
'檢查是否可進入本系統
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
RecordDate=split(gInitDT(date),"-")

if request("DB_Selt")="BatchSelt" then
	strwhere="":tmp_BatchNumber="":Sys_BatchNumber=""
	if UCase(request("Sys_BatchNumber"))<>"" then
		tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
		for i=0 to Ubound(tmp_BatchNumber)
			if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
			if i=0 then
				Sys_BatchNumber=trim(Sys_BatchNumber)&UCase(tmp_BatchNumber(i))
			else
				Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&UCase(tmp_BatchNumber(i))
			end if
			if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(UCase(Sys_BatchNumber))&"'"
		next
		strwhere=" and b.BatchNumber in ('"&Sys_BatchNumber&"')"
	end if

	if trim(request("Sys_ImageFileNameB1"))<>"" and trim(request("Sys_ImageFileNameB2"))<>"" then
		strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(request("Sys_ImageFileNameB1")))&"' and '"&trim(UCase(request("Sys_ImageFileNameB2")))&"'"
	elseif trim(request("Sys_ImageFileNameB1"))<>"" then
		strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(request("Sys_ImageFileNameB1")))&"' and '"&trim(UCase(request("Sys_ImageFileNameB1")))&"'"
	elseif trim(request("Sys_ImageFileNameB2"))<>"" then
		strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(request("Sys_ImageFileNameB2")))&"' and '"&trim(UCase(request("Sys_ImageFileNameB2")))&"'"
	end if
	if request("Sys_IllegalDate1")<>"" and request("Sys_IllegalDate2")<>""then
		IllegalDate1=gOutDT(request("Sys_IllegalDate1"))&" 0:0:0"
		IllegalDate2=gOutDT(request("Sys_IllegalDate2"))&" 23:59:59"
		strwhere=strwhere&" and a.IllegalDate between TO_DATE('"&IllegalDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&IllegalDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
	DB_Display=request("DB_Display")
end if
if DB_Display="show" then
	If strwhere="" Then DB_Display=""
	if trim(strwhere)<>"" or (trim(request("Sys_UserMarkDate1"))<>"" and trim(request("Sys_UserMarkDate2"))<>"") then
		if trim(strwhere)<>"" then
			strSQL="select distinct a.SN,a.CarNo,a.IllegalDate from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere&" order by a.CarNo,a.IllegalDate"

			set rssn=conn.execute(strSQL)
			BillSN="":tempBillSN=""
			while Not rssn.eof
				If trim(tempBillSN)<>trim(rssn("SN")) Then
					tempBillSN=trim(rssn("SN"))
					if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
					BillSN=BillSN&trim(rssn("SN"))
				end if
				rssn.movenext
			wend
			rssn.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A') b where a.SN=b.BillSN "&strwhere

			set Dbrs=conn.execute(strSQL)
			DBsum=Cint(Dbrs("cnt"))
			Dbrs.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere

			set chksuess=conn.execute(strSQL)
			filsuess=Cint(chksuess("cnt"))
			chksuess.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID in('N','E')) b where a.SN=b.BillSN "&strwhere
			set chksuess=conn.execute(strSQL)
			fildel=Cint(chksuess("cnt"))
			chksuess.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null and RecordStateId = -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A') b where a.SN=b.BillSN "&strwhere
			set Dbrs=conn.execute(strSQL)
			deldata=Cint(Dbrs("cnt"))
			Dbrs.close

			strSQL2=strwhere
		end if
		'單退要用註記日查詢
		if trim(request("Sys_UserMarkDate1"))<>"" and trim(request("Sys_UserMarkDate2"))<>""  then
			UserMarkDate1=gOutDT(request("Sys_UserMarkDate1"))&" 0:0:0"
			UserMarkDate2=gOutDT(request("Sys_UserMarkDate2"))&" 23:59:59"

			strwhere=strwhere&" and c.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"

			strGet="select count(*) as cnt from Billbase a,StopBillMailHistory b" &_
				" where a.Sn=b.BillSn and a.RecordStateID=0 and b.UserMarkResonID in ('A','B','C')" &_
				" and b.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			set rsGet=conn.execute(strGet)
			if not rsGet.eof then
				Getdata=Cint(rsGet("cnt"))
			end if
			rsGet.close
			set rsGet=nothing

			strROpen="select count(*) as cnt from Billbase a,StopBillMailHistory b" &_
				" where a.Sn=b.BillSn and a.RecordStateID=0 and b.UserMarkResonID in ('1','2','3','4','8','K','L','M','O','P','Q')" &_
				" and b.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			set rsROpen=conn.execute(strROpen)
			if not rsROpen.eof then
				Opendata=Cint(rsROpen("cnt"))
			end if
			rsROpen.close
			set rsROpen=nothing

			strRStore="select count(*) as cnt from Billbase a,StopBillMailHistory b" &_
				" where a.Sn=b.BillSn and a.RecordStateID=0 and b.UserMarkResonID in ('5','6','7','T')" &_
				" and b.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			set rsRStore=conn.execute(strRStore)
			if not rsRStore.eof then
				Storedata=Cint(rsRStore("cnt"))
			end if
			rsRStore.close
			set rsRStore=nothing
		end if
	else
		DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
tmpSQL=strwhere
%>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr height="30">
		<td bgcolor="#FFCC33"><span class="style3">催繳單 / 各式清冊 列印</span><img src="space.gif" width="60" height="1"> <strong>請勿升級 Internet Explorer 7 . 避免套印舉發單出現異常</strong></img></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						作業批號
						<Select Name="Selt_BatchNumber" onchange="fnBatchNumber();">
							<option value="">請點選</option><%
							strSQL="select distinct TO_char(ExchangeDate,'YYYY/MM/DD') ExchangeDate,BatchNumber from DCILog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",-5, date)&" 00:00"&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') and ExchangeTypeID='A' and DCIReturnStatusID='S' order by ExchangeDate DESC"
		
							set rs=conn.execute(strSQL)
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
						<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=UCase(request("Sys_BatchNumber"))%>" size="29" onkeyup="funShowBillNo()">
						
						(<strong>多個批號同時處理</strong>，各批號請用,隔開。如：95A361,95A382,95A486）						
						<br>
						催繳單號
						<input name="Sys_ImageFileNameB1" type="text" class="btn1" value="<%=UCase(request("Sys_ImageFileNameB1"))%>" size="16" maxlength="16">
						~
						<input name="Sys_ImageFileNameB2" type="text" class="btn1" value="<%=UCase(request("Sys_ImageFileNameB2"))%>" size="16" maxlength="16"> ( 列印 <strong>單筆</strong> 或 特定範圍 催繳單才需填寫)
						<br>
						註記時間
						<input name="Sys_UserMarkDate1" type="text" class="btn1" value="<%=request("Sys_UserMarkDate1")%>" size="8" maxlength="6">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_UserMarkDate1');">
						~
						<input name="Sys_UserMarkDate2" type="text" class="btn1" value="<%=request("Sys_UserMarkDate2")%>" size="8" maxlength="6">
						<input type="button" name="datestr2" value="..." onclick="OpenWindow('Sys_UserMarkDate2');">
						( 列印 <strong>收受清冊</strong> 或 <strong>單退清冊</strong> 才需填寫)
						<br>
						停車時間
						<input name="Sys_IllegalDate1" type="text" class="btn1" value="<%=request("Sys_IllegalDate1")%>" size="8" maxlength="6">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_IllegalDate1');">
						~
						<input name="Sys_IllegalDate2" type="text" class="btn1" value="<%=request("Sys_IllegalDate2")%>" size="8" maxlength="6">
						<input type="button" name="datestr2" value="..." onclick="OpenWindow('Sys_IllegalDate2');">
							( 產生 <strong>公示檔</strong> 才需填寫)
						<br>
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt('BatchSelt');"<%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						'if CheckPermission(290,1)=false then
						'	response.write " disabled"
						'end if
						%>>
						<input type="button" name="cancel" value="清除" onClick="location='StopAllBillPrint.asp'">

						<img src="space.gif" width="35" height="1"></img><strong>( 查詢 <%=DBsum%> 筆紀錄 , <%=filsuess%>筆成功 ,  <%=fildel%> 筆失敗 , <%=deldata%> 筆刪除  ,  <%=DBsum-filsuess-fildel-deldata%>筆未處理, <%=Getdata%> 筆收受, <%=Storedata%> 筆單退_寄存, <%=Opendata%> 筆單退_公示. )</strong>
											
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<tr>
		<td height="35" bgcolor="#FFDD77" align="left">
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;本批資料繳費期限
				<input name="Sys_DeallineDate" type="text" class="btn1" value="<%
					if ifnull(request("Sys_DeallineDate")) then
						response.write gInitDT(DateAdd("d", 10,date()))
					end if
				%>" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
				<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_DeallineDate');">
				&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="fun_DeallineDate();">
				&nbsp;&nbsp;<font color="red"><B><span id="showBillNoA""></span>&nbsp;&nbsp;<span id="showBillNoB"></span></B></font>
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 催繳單" onclick="funStopBillPrints_HuaLien()">
				<img src="space.gif" width="37" height="1"></img>
				
				<input type="button" name="btnprint" value="列印 補印催繳單" onclick="funStopBillPrints_HuaLien_Mend()">
				<img src="space.gif" width="37" height="1"></img>

				<input type="button" name="btnprint" value="匯出 停管催繳檔 " onclick="funExportTxt_HL()">
				<img src="space.gif" width="37" height="1"></img>

				<input type="button" name="btnprint" value="匯出 停管公示檔 " onclick="funExportOpenGovTxt()">
				<img src="space.gif" width="37" height="1"></img>
				<br>
				
			<hr>
			<!--<span class="style3">
			DCI檔案名稱
			<input name="textfield42324" type="text" value="" size="14" maxlength="13">
			</span>-->
	
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234222" value="車籍資料" onclick="funchgCarDataList_HL()">

			<span class="style3"><img src="space.gif" width="22" height="8"></span>
			<input type="button" name="Submit4234" value="催繳清冊" onclick="funReportSendList_HL()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>

			<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit488423" value="退件清冊_寄存(全部)" onclick="funReturnSendList_Store_All()">
	
			
		<br>

	
		<span class="style3"><img src="space.gif" width="130" height="8"></span>
			<input type="button" name="Submit488423" value="收受清冊" onclick="funGetSendList_HL()">
					
				<span class="style3"><img src="space.gif" width="163" height="8"></span>
			<input type="button" name="Submit4233" value="退件清冊_公示(全部)" onclick="funReturnSendList_Gov_All()">
				
	    <Br>

			
	
		<br>
		<br>
		<!--<HR>
		本批資料發文監理站日期
		<input name="Sys_SendOpenGovDocToStationDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendOpenGovDocToStationDate');">
		&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSendOpenGovDocToStationDate();">
		<br>
		本批資料一次郵寄日期
		&nbsp;&nbsp;&nbsp;&nbsp;<input name="Sys_BillBaseMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BillBaseMailDate');">
		&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSys_MailDate();">
		<Br>
		本批資料二次郵寄日期
		&nbsp;&nbsp;&nbsp;&nbsp;<input name="Sys_StoreAndSendMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StoreAndSendMailDate');">
		&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funStoreAndSendMailDate();">
		<br><br>-->
	</td>
  </tr>
  <tr>
    <td><p align="center">&nbsp;</p>    </td></tr>
<tr>
<td>
<b>催繳單列印設定</b>  <br>
印表機&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;OKI<br>
紙張格式 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;LEGAL 8.5 x 14 <br>
紙張來源 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;進紙夾1  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;(催繳單放最下方進紙夾,背面空白朝上.送達證書區域朝印表機內) <br>
上下左右邊界 : &nbsp;&nbsp; 0.166
</td>
</tr>
</table>

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="BillSN" value="<%=BillSN%>">
<input type="Hidden" name="SQLstr" value="<%=strSQL2%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
funShowBillNo();
function fnBatchNumber(){
	myForm.Sys_BatchNumber.value=myForm.Selt_BatchNumber.value;
	funShowBillNo();
}

function funShowBillNo(){
	if(myForm.Sys_BatchNumber.value.length>=5){
		runServerScript("StopchkShoBillNo.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value);
	}
}

function fun_DeallineDate(){
	if (myForm.DB_Display.value!=""){
		if (myForm.Sys_DeallineDate.value!=''){
			UrlStr="BillBaseDeadLineDate_new.asp";
			myForm.action=UrlStr;
			myForm.target="BillBaseDeadLineDate";
			myForm.submit();
			myForm.action="";
			myForm.target="";
		}
	}
}

function funStopBillPrints_HuaLien_Mend(){
	UrlStr="StopBillPrints_HuaLien_Mend.asp";
	myForm.action=UrlStr;
	myForm.target="StopBillPrints_Mend";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_HuaLien(){
	UrlStr="StopBillPrints_HuaLien_new.asp";
	myForm.action=UrlStr;
	myForm.target="StopBillPrints";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funSelt(DBKind){
	var error=0;
	if(DBKind=='BatchSelt'){
		if(myForm.Sys_BatchNumber.value==''&&myForm.Sys_ImageFileNameB1.value==''&&myForm.Sys_ImageFileNameB2.value==''&&myForm.Sys_UserMarkDate1.value==''&&myForm.Sys_UserMarkDate2.value==''&&myForm.Sys_IllegalDate1.value==''&&myForm.Sys_IllegalDate2.value==''){
			error=1;
			alert("必須有填詢條件!!");
		}
		if(error==0){
			myForm.BillSN.value="";
			myForm.DB_Selt.value=DBKind;
			myForm.DB_Display.value='show';
			myForm.submit();
		}
	}
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}

function funchgCarDataList_HL(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
		alert("請先輸入作業批號或單號查詢欲列印車籍資料清冊的舉發單！");
	}else{
		UrlStr="StopDciPrintCarDataList.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","yes","yes","no");
	}
}

function funReportSendList_HL(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
			alert("請先輸入作業批號或單號查詢欲列印催繳資料清冊的舉發單！");
	}else{
		UrlStr="StopReportSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}

function funExportOpenGovTxt(){
	UrlStr="StopExportOpenGov_txt.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","yes","yes","no");
}

function funExportTxt_HL(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
			alert("請先輸入作業批號或單號查詢欲列印催繳資料清冊的舉發單！");
	}else{
		UrlStr="StopExport_txt.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",10,10,0,0,"no","no","no","no");
	}
}


function funMailList2(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
			alert("請先輸入作業批號或單號查詢欲列印交寄大宗函件的舉發單！");
	}else{
		UrlStr="StopMailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>&MailSendType=S";
		newWin(UrlStr,"MailReportList",300,220,350,200,"no","no","no","no");
	}
}

function funReturnSendList_Store_All(){
	if (myForm.Sys_UserMarkDate1.value=="" || myForm.Sys_UserMarkDate2.value==""){
			alert("請先輸入註記日期查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="StopReturnSendList_Excel_A3_Store_All.asp?Sys_UserMarkDate1="+myForm.Sys_UserMarkDate1.value+"&Sys_UserMarkDate2="+myForm.Sys_UserMarkDate2.value;
		newWin(UrlStr,"inputWin6a",800,700,0,0,"yes","yes","yes","no");
	}
}

function funGetSendList_HL(){
	if (myForm.Sys_UserMarkDate1.value=="" || myForm.Sys_UserMarkDate2.value==""){
			alert("請先輸入註記日期查詢欲列印收受清冊的舉發單！");
	}else{
		UrlStr="StopGetSendList_Excel_A3.asp?Sys_UserMarkDate1="+myForm.Sys_UserMarkDate1.value+"&Sys_UserMarkDate2="+myForm.Sys_UserMarkDate2.value;
		newWin(UrlStr,"inputWin7a",800,700,0,0,"yes","yes","yes","no");
	}
}

function funReturnSendList_Gov_All(){
	if (myForm.Sys_UserMarkDate1.value=="" || myForm.Sys_UserMarkDate2.value==""){
			alert("請先輸入註記日期查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="StopReturnSendList_Excel_A3_Gov_All.asp?Sys_UserMarkDate1="+myForm.Sys_UserMarkDate1.value+"&Sys_UserMarkDate2="+myForm.Sys_UserMarkDate2.value;
		newWin(UrlStr,"inputWin8a",800,700,0,0,"yes","yes","yes","no");
	}
}
</script>
<%conn.close%>