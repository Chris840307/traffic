<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴案件建檔</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
daynow=gInitDT(Date)

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing


if request("DB_Add")="ADD" then
	DocDate=gOutDT(request("DocDate"))
	ArgueDate=gOutDT(request("ArgueDate"))
	ProcessDate=gOutDT(request("ProcessDate"))
	ACTIONDATE=gOutDT(request("ACTIONDATE"))
    if sys_City="澎湖縣" Or sys_City="台南市" Then
			BadCnttmp=0
			WarnCnttmp=0
			If Trim(request("Sys_Punishment"))="劣蹟1次" Then
				BadCnttmp="1"
				WarnCnttmp="0"
			ElseIf Trim(request("Sys_Punishment"))="申誡1次" Then
				BadCnttmp="0"
				WarnCnttmp="1"
			ElseIf Trim(request("Sys_Punishment"))="申誡2次" Then
				BadCnttmp="0"
				WarnCnttmp="2"
			Else
				BadCnttmp="0"
				WarnCnttmp="0"
			End If 
			
			strSQL="Insert into ArgueBase(SN,BillNo,CarNo,Times,ReportContent,DocNo,DocDate,Punishment,ArgueDate,Arguer,Note,ArguerCreditID,ArguerAddr,ArguerResonID,ArguerResonName,ArguerTel,ErrorID,ErrorName,ArguerContent,Cancel,Close,RecordStateID,RecordDate,RecordMemberID,argueway,reportdeparment,reportNo,processdate,processno,DELBILLREASON,VIOLATERULE1,VIOLATERULE2,ACTIONDATE,ACTIONNO,BadCnt,WarnCnt,DelName) values("&funTableSeq("ArgueBase","SN")&",'"&request("Sys_BillNo")&"','"& request("Sys_CarNo") &"',(select count(*)+1 from ArgueBase where BillNo='"&request("Sys_BillNo")&"'),'"&request("Sys_ReportContent")&"','"&request("Sys_DocNo")&"',"&funGetDate(DocDate,0)&",'"&request("Sys_Punishment")&"',"&funGetDate(ArgueDate,0)&",'"&request("Sys_Arguer")&"','"&request("Sys_Note")&"','"&request("Sys_ArguerCreditID")&"','"&request("Sys_ArguerAddr")&"',"&request("Sys_ArguerResonID")&",'"&trim(request("ResonName"))&"','"&request("Sys_ArguerTel")&"',"&request("Sys_ErrorID")&",'"&trim(request("ErrName"))&"','"&request("Sys_ArguerContent")&"','"&request("Sys_Cancel")&"','"&request("Sys_Close")&"',0,"&funGetDate(now,1)&","&Session("User_ID")& ",'"&request("argueway")& "','"&request("reportdeparment")&"','"&request("reportno")&"',"&funGetDate(processdate,0) &",'"&request("processno")&"',"&request("Sys_DelReasonID")&",'"&Trim(request("VIOLATERULE1"))&"','"&Trim(request("VIOLATERULE2"))&"',"&funGetDate(ACTIONDATE,0)&",'"&Trim(request("ACTIONNO"))&"','"&BadCnttmp&"','"&WarnCnttmp&"','"&Trim(request("DelName"))&"')"
			'response.write strSQL
			'response.end
    else
			strSQL="Insert into ArgueBase(SN,BillNo,CarNo,Times,ReportContent,DocNo,DocDate,Punishment,ArgueDate,Arguer,Note,ArguerCreditID,ArguerAddr,ArguerResonID,ArguerResonName,ArguerTel,ErrorID,ErrorName,ArguerContent,Cancel,Close,RecordStateID,RecordDate,RecordMemberID) values("&funTableSeq("ArgueBase","SN")&",'"&request("Sys_BillNo")&"','"& request("Sys_CarNo") &"',(select count(*)+1 from ArgueBase where BillNo='"&request("Sys_BillNo")&"'),'"&request("Sys_ReportContent")&"','"&request("Sys_DocNo")&"',"&funGetDate(DocDate,0)&",'"&request("Sys_Punishment")&"',"&funGetDate(ArgueDate,0)&",'"&request("Sys_Arguer")&"','"&request("Sys_Note")&"','"&request("Sys_ArguerCreditID")&"','"&request("Sys_ArguerAddr")&"',"&request("Sys_ArguerResonID")&",'"&trim(request("ResonName"))&"','"&request("Sys_ArguerTel")&"',"&request("Sys_ErrorID")&",'"&trim(request("ErrName"))&"','"&request("Sys_ArguerContent")&"','"&request("Sys_Cancel")&"','"&request("Sys_Close")&"',0,"&funGetDate(now,1)&","&Session("User_ID")&")"
    end if
	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
%>
<body onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td height="12" bgcolor="#FFCC33">申訴案件</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#dddddd">
				<tr  bgcolor="#ffffff">
					<td width="15%" bgcolor="#FFFF99"><font color="red">* </font>舉發單號</td>
					<td width="35%">
						<input name="Sys_BillNo" class="btn1" type="text" value="" size="12" maxlength="9" onBlur="funInpuMan();">
					</td>
					<td width="15%" bgcolor="#FFFF99">陳情方式</td>
                    <td width="35%">
                        <input name="argueway" class="btn1" type="text" value="" size="25" maxlength="30">
                    </td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">收文號</td>
					<td>
						<input name="Sys_DocNo" class="btn1" type="text" size="12" maxlength="12">
					</td>
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述日期</td>
					<td>
						<input name="ArgueDate" class="btn1" type="text" value="<%=daynow%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ArgueDate');">
					</td>
					
				</tr>
				
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">來文機關</td>
					<td>
						<input name="reportdeparment" class="btn1" type="text" value="" size="25" maxlength="30">
					</td>
					<td bgcolor="#FFFF99">來文文號</td>
					<td>
						<input name="reportno" class="btn1" type="text" value="" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述人姓名</td>
					<td>
						<input name="Sys_Arguer" class="btn1" type="text" value="" size="11" maxlength="11">
					</td>
					<td bgcolor="#FFFF99">陳述人身分證</td>
					<td>
						<input name="Sys_ArguerCreditID" class="btn1" type="text" value="" size="11" maxlength="11">
					</td>
				</tr>
				
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述人地址</td>
					<td>
						<input name="Sys_ArguerAddr" class="btn1" type="text" value="" size="40" maxlength="50">
					</td>
					<td bgcolor="#FFFF99">陳述人電話</td>
					<td>
						<input name="Sys_ArguerTel" class="btn1" type="text" value="" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述原因</td>
					<td colspan="3">
						<select name="Sys_ArguerResonID" class="btn1" >
							<option value="0">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=15 order by ID"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								response.write "<option value="""&rs("ID")&""">"
								response.write rs("Content")
								response.write "</option>"
								rs.movenext
							wend
							rs.close%>
						</select>
						<span id="inputxt"></span>
					</td>
					
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述內容補充摘要</td>
					<td>
						<textarea name="Sys_ArguerContent" cols="40" rows="3" class="btn1"></textarea>
					</td>
					<td bgcolor="#FFFF99">辦理情形</td>
					<td>
						<select name="Sys_ErrorID" class="btn1" onchange="funErrOthen();" >
							<option value="0">無缺失</option><%
							strSQL="select Content,ID from Code where TypeID=16 order by ID"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								response.write "<option value="""&rs("ID")&""">"
								response.write rs("Content")
								response.write "</option>"
								rs.movenext
							wend
							rs.close%>
						</select>
						<span id="inpuerr"></span>
					</td>
				</tr >
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">回覆日期</td>
					<td>
					   <input name="processdate" class="btn1" type="text" value="<%=daynow%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('processdate');">
					</td>
					<td bgcolor="#FFFF99">回覆文號</td>
					<td>
						<input name="processno" class="btn1" type="text" value="" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">違反規定(項)</td>
					<td>
						<select name="VIOLATERULE1">
							<option value="">無違反規定</option>
							<option value="參、一、(一)">參、一、(一)</option>
							<option value="參、一、(二)">參、一、(二)</option>
							<option value="參、二、(一)">參、二、(一)</option>
							<option value="參、二、(二)">參、二、(二)</option>
							<option value="參、三、(一)">參、三、(一)</option>
							<option value="參、三、(一)">參、三、(二)</option>
							<option value="參、四、(一)">參、四、(一)</option>
							<option value="參、四、(二)">參、四、(二)</option>
							<option value="參、四、(三)">參、四、(三)</option>
							<option value="參、四、(四)">參、四、(四)</option>
							<option value="參、四、(五)">參、四、(五)</option>
							<option value="參、四、(六)">參、四、(六)</option>
							<option value="參、四、(七)">參、四、(七)</option>
							<option value="參、四、(八)">參、四、(八)</option>
							<option value="參、四、(九)">參、四、(九)</option>
							<option value="參、五、(一)">參、五、(一)</option>
							<option value="參、五、(二)">參、五、(二)</option>
							<option value="參、五、(三)">參、五、(三)</option>
							<option value="參、五、(四)">參、五、(四)</option>
						</select>
						
					</td>
					<td bgcolor="#FFFF99">違反規定(目)</td>
					<td>
						<select name="VIOLATERULE2">
							<option value="">無違反規定</option>
							<option value="之 1">之 1</option>
							<option value="之 2">之 2</option>
							<option value="之 3">之 3</option>
							<option value="之 4">之 4</option>
							<option value="之 5">之 5</option>
							<option value="之 6">之 6</option>
							<option value="之 7">之 7</option>
							<option value="之 8">之 8</option>
							<option value="之 9">之 9</option>
							<option value="之 10">之 10</option>
						</select>
					</td>
				</tr>	
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99">懲處情形</td>
					<td colspan="3">
						<select name="Sys_Punishment">
							<option value="無">無</option>
							<option value="劣蹟1次">劣蹟1次</option>
							<option value="申誡1次">申誡1次</option>
							<option value="申誡2次">申誡2次</option>
						</select>
					</td>
				</tr>
				
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">處分(通報)日期</td>
					<td>
					   <input name="Actiondate" class="btn1" type="text" value="" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
													<input type="button" name="datestr" value="..." onclick="OpenWindow('Actiondate');">
					</td>
					<td bgcolor="#FFFF99">處分(通報)文號</td>
					<td>
						<input name="Actionno" class="btn1" type="text" value="" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99">是否撤銷</td>
					<td>
						<select name="Sys_Cancel" class="btn1">
							<option value=1>否</option>
							<option value=0>是</option>
						</select>
					</td>
					<td bgcolor="#FFFF99">撤銷舉發單理由</td>
					<td>
						<select name="Sys_DelReasonID" class="btn1" onchange="funDelOthen();" >
							<option value="0">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=21 order by ID"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								response.write "<option value="""&rs("ID")&""">"
								response.write rs("Content")
								response.write "</option>"
								rs.movenext
							wend
							rs.close%>
						</select>
						<span id="inpudel"></span>
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">備註</td>
					<td colspan="3">
						<textarea name="Sys_Note" class="btn1" cols="40"></textarea>
					</td>
				</tr>
				
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99"><font color="red">* </font>案件狀態</td>
					<td>
						<select name="Sys_Close" class="btn1">
							<option value=0>末處理</option>
							<option value=1>結案</option>
							<option value=2>待查中</option>
						</select>
					</td>
					<td bgcolor="#FFFF99">紀錄人員</td>
					<td colspan="3">
						<%=session("Ch_Name")%>
					</td>
				</tr>
                        
                        				
                        
							
						
						
				
		  </table>
		</td>
	</tr>
	<tr bgcolor="#ffffff" align="center">
		<td height="35" bgcolor="#FFDD77">
			<input type="button" name="save" value=" 儲 存 " onclick="funAdd();">
			<input type="button" name="exit" value=" 關 閉 " onclick="funExt();">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Add" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}
function funOthen(){ 
	if(myForm.Sys_ArguerResonID.value=='448'){
		inputxt.innerHTML ="<br>原因：<input type=text name='ResonName' size=30 class='btn1'>";
	}else{
		inputxt.innerHTML ="";
	}
}
function funErrOthen(){ 
	if(myForm.Sys_ErrorID.value=='453'){
		inpuerr.innerHTML ="<br>原因：<input type=text name='ErrName' size=30 class='btn1'>";
	}else{
		inpuerr.innerHTML ="";
	}
}
function funDelOthen(){ 
	if(myForm.Sys_DelReasonID.value=='811'){
		inpudel.innerHTML ="<br>理由：<input type=text name='DelName' size=30 class='btn1'>";
	}else{
		inpudel.innerHTML ="";
	}
}
function funAdd(){
	var err=0;
	if (err==0){
		if(myForm.ArgueDate.value!=""){
			if(!dateCheck(myForm.ArgueDate.value)){
				err=1;
				alert("陳述日輸入不正確!!");
			}
		}
	}
	if (err==0){
		if(myForm.Sys_ArguerResonID.value=="0"){
			err=1;
			alert("請選擇陳述原因!!");
		}
	}
	if (err==0){
		if(myForm.Sys_DelReasonID.value!="0" && myForm.Sys_Cancel.value=="1"){
			err=1;
			alert("如選擇撤銷舉發單原因，是否撤銷必須選擇[是]!!");
		}
	}
	if (err==0){
		if(myForm.Sys_DelReasonID.value=="0" && myForm.Sys_Cancel.value=="0"){
			err=1;
			alert("如是否撤銷選擇[是]，必須輸入撤銷舉發單原因!!");
		}
	}

	if (err==0){
		if(myForm.Sys_BillNo.value==''){
			err=1;
			alert("舉發單號不可空白");
		}else{
			runServerScript("chkAddNew.asp?BillNo="+myForm.Sys_BillNo.value);
		}
	}
}
function funInpuMan(){
	myForm.Sys_BillNo.value=myForm.Sys_BillNo.value.toUpperCase();
	runServerScript("InputMan.asp?BillNo="+myForm.Sys_BillNo.value);
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		opener.myForm.submit();
		self.close();
	}
}
</script>
<%conn.close%>