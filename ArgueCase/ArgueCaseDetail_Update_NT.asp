<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴案件修改</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
daynow=gInitDT(Date)

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing


if request("DB_ADD")="ADD" then
	'DocDate=gOutDT(request("DocDate"))
	ArgueDate=gOutDT(request("ArgueDate"))
	if trim(request("processdate"))<>"" then
		Sprocessdate="TO_DATE('"&gOutDT(request("processdate"))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
	else
		Sprocessdate="null"
	end If
	if trim(request("Actiondate"))<>"" then
		Actiondate="TO_DATE('"&gOutDT(request("Actiondate"))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
	else
		Actiondate="null"
	end if
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
	strSQL="Update ArgueBase set BillNo='"&request("Sys_BillNo")&"',CarNo='"&request("Sys_CarNo")&"'" &_
		",DocNo='"&request("Sys_DocNo")&"'," &_
		"Punishment='"&request("Sys_Punishment")&"',ArgueDate="&funGetDate(ArgueDate,0)&","&_
		"Arguer='"&request("Sys_Arguer")&"',Note='"&request("Sys_Note")&"',ArguerCreditID='"&request("Sys_ArguerCreditID")&"'" &_
		",ArguerAddr='"&request("Sys_ArguerAddr")&"',ArguerResonID="&request("Sys_ArguerResonID")&"," &_
		"ArguerResonName='"&request("ResonName")&"',ArguerTel='"&request("Sys_ArguerTel")&"'" &_
		",ErrorID="&request("Sys_ErrorID")&",ErrorName='"&request("ErrName")&"',ArguerContent='"&request("Sys_ArguerContent")&"'" &_
		",Cancel='"&request("Sys_Cancel")&"',Close='"&request("Sys_Close")&"'" &_
		",argueway='"&trim(request("argueway"))&"',reportdeparment='"&trim(request("reportdeparment"))&"'" &_
		",reportNo='"&trim(request("reportNo"))&"',processno='"&trim(request("processno"))&"'" &_
		",processdate="&Sprocessdate&",DELBILLREASON="&Trim(request("Sys_DelReasonID")) &_
		",VIOLATERULE1='"&Trim(request("VIOLATERULE1"))&"',VIOLATERULE2='"&Trim(request("VIOLATERULE2"))&"'" &_
		",BadCnt='"&BadCnttmp&"',WarnCnt='"&WarnCnttmp&"'" &_
		",Actiondate="&Actiondate &",Actionno='"&Trim(request("Actionno")) &"'"&_
		",DelName='"&Trim(request("DelName"))&"'" &_
		" where SN="&request("SN")
	else
		strSQL="Update ArgueBase set BillNo='"&request("Sys_BillNo")&"',CarNo='"&request("Sys_CarNo")&"',DocNo='"&request("Sys_DocNo")&"',Punishment='"&request("Sys_Punishment")&"',ArgueDate="&funGetDate(ArgueDate,0)&",Arguer='"&request("Sys_Arguer")&"',Note='"&request("Sys_Note")&"',ArguerCreditID='"&request("Sys_ArguerCreditID")&"',ArguerAddr='"&request("Sys_ArguerAddr")&"',ArguerResonID="&request("Sys_ArguerResonID")&",ArguerResonName='"&request("ResonName")&"',ArguerTel='"&request("Sys_ArguerTel")&"',ErrorID="&request("Sys_ErrorID")&",ErrorName='"&request("ErrName")&"',ArguerContent='"&request("Sys_ArguerContent")&"',Cancel='"&request("Sys_Cancel")&"',Close='"&request("Sys_Close")&"' where SN="&request("SN")
	end If
	'response.write strSQL
	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end If

	strSQL="select * from ArgueBase where SN="&request("SN")
	set rs=conn.execute(strSQL)
	DocDate=gInitDT(rs("DocDate"))
	ArgueDate=gInitDT(rs("ArgueDate"))
	sBillNo=rs("BillNo")
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
						<input name="Sys_BillNo" class="btn1" type="text" value="<%
						response.write Trim(rs("BillNo"))
						%>" size="12" maxlength="9" readonly>
					</td>
					<td width="15%" bgcolor="#FFFF99">陳情方式</td>
                    <td width="35%">
                        <input name="argueway" class="btn1" type="text" value="<%
						response.write Trim(rs("argueway"))
						%>" size="25" maxlength="30">
                    </td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">收文號</td>
					<td>
						<input name="Sys_DocNo" class="btn1" type="text" value="<%=rs("DocNo")%>" size="12" maxlength="12">
					</td>
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述日期</td>
					<td>
						<input name="ArgueDate" class="btn1" type="text" value="<%=ArgueDate%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ArgueDate');">
					</td>
					
				</tr>
				
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">來文機關</td>
					<td>
						<input name="reportdeparment" class="btn1" type="text" value="<%=trim(rs("reportdeparment"))%>" size="25" maxlength="30">
					</td>
					<td bgcolor="#FFFF99">來文文號</td>
					<td>
						<input name="reportno" class="btn1" type="text" value="<%=trim(rs("reportNo"))%>" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述人姓名</td>
					<td>
						<input name="Sys_Arguer" class="btn1" type="text" value="<%=rs("Arguer")%>" size="11" maxlength="11">
					</td>
					<td bgcolor="#FFFF99">陳述人身分證</td>
					<td>
						<input name="Sys_ArguerCreditID" class="btn1" type="text" value="<%=rs("ArguerCreditID")%>" size="11" maxlength="11">
					</td>
				</tr>
				
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述人地址</td>
					<td>
						<input name="Sys_ArguerAddr" class="btn1" type="text" value="<%=rs("ArguerAddr")%>" size="40" maxlength="50">
					</td>
					<td bgcolor="#FFFF99">陳述人電話</td>
					<td>
						<input name="Sys_ArguerTel" class="btn1" type="text" value="<%=rs("ArguerTel")%>" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述原因</td>
					<td colspan="3">
						<select name="Sys_ArguerResonID" class="btn1" >
							<option value="0">請選擇</option><%
							strSQLR="select Content,ID from Code where TypeID=15 order by ID"
							set rsR=conn.execute(strSQLR)
							while Not rsR.eof
								response.write "<option value="""&rsR("ID")&""""
								if trim(rs("ArguerResonID"))=trim(rsR("ID")) then response.write " selected"
								response.write ">"
								response.write rsR("Content")
								response.write "</option>"
								rsR.movenext
							wend
							rsR.close%>
						</select>

						<span id="inputxt"></span>
					</td>
					
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述內容補充摘要</td>
					<td>
						<textarea name="Sys_ArguerContent" cols="40" rows="3" class="btn1"><%=rs("ArguerContent")%></textarea>
					</td>
					<td bgcolor="#FFFF99">辦理情形</td>
					<td>
					<%If Trim(rs("ReportContent"))<>"" Then
						response.write Trim(rs("ReportContent"))&"<br>"
					End If %>
						<select name="Sys_ErrorID" class="btn1" onchange="funErrOthen();" >
							<option value="0">無缺失</option><%
							strSQLE="select Content,ID from Code where TypeID=16 order by ID"
							set rsE=conn.execute(strSQLE)
							while Not rsE.eof
								response.write "<option value="""&rsE("ID")&""""
								if trim(rs("ErrorID"))=trim(rsE("ID")) then response.write " selected"
								response.write ">"
								response.write rsE("Content")
								response.write "</option>"
								rsE.movenext
							wend
							rsE.close%>
						</select>

						<span id="inpuerr"></span>
					</td>
				</tr >
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">回覆日期</td>
					<td>
					   <input name="processdate" class="btn1" type="text" value="<%
			if not isnull(rs("processdate")) and trim(rs("processdate"))<>"" then
				response.write year(rs("processdate"))-1911&right("00"&month(rs("processdate")),2)&right("00"&day(rs("processdate")),2)
			end if
					   %>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('processdate');">
					</td>
					<td bgcolor="#FFFF99">回覆文號</td>
					<td>
						<input name="processno" class="btn1" type="text" value="<%=trim(rs("processno"))%>" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">違反規定(項)</td>
					<td>
						<select name="VIOLATERULE1">
							<option value="">無違反規定</option>
							<option value="參、一、(一)" <%If Trim(rs("VIOLATERULE1"))="參、一、(一)" Then response.write "selected"%>>參、一、(一)</option>
							<option value="參、一、(二)" <%If Trim(rs("VIOLATERULE1"))="參、一、(二)" Then response.write "selected"%>>參、一、(二)</option>
							<option value="參、二、(一)" <%If Trim(rs("VIOLATERULE1"))="參、二、(一)" Then response.write "selected"%>>參、二、(一)</option>
							<option value="參、二、(二)" <%If Trim(rs("VIOLATERULE1"))="參、二、(二)" Then response.write "selected"%>>參、二、(二)</option>
							<option value="參、三、(一)" <%If Trim(rs("VIOLATERULE1"))="參、三、(一)" Then response.write "selected"%>>參、三、(一)</option>
							<option value="參、三、(一)" <%If Trim(rs("VIOLATERULE1"))="參、三、(一)" Then response.write "selected"%>>參、三、(二)</option>
							<option value="參、四、(一)" <%If Trim(rs("VIOLATERULE1"))="參、四、(一)" Then response.write "selected"%>>參、四、(一)</option>
							<option value="參、四、(二)" <%If Trim(rs("VIOLATERULE1"))="參、四、(二)" Then response.write "selected"%>>參、四、(二)</option>
							<option value="參、四、(三)" <%If Trim(rs("VIOLATERULE1"))="參、四、(三)" Then response.write "selected"%>>參、四、(三)</option>
							<option value="參、四、(四)" <%If Trim(rs("VIOLATERULE1"))="參、四、(四)" Then response.write "selected"%>>參、四、(四)</option>
							<option value="參、四、(五)" <%If Trim(rs("VIOLATERULE1"))="參、四、(五)" Then response.write "selected"%>>參、四、(五)</option>
							<option value="參、四、(六)" <%If Trim(rs("VIOLATERULE1"))="參、四、(六)" Then response.write "selected"%>>參、四、(六)</option>
							<option value="參、四、(七)" <%If Trim(rs("VIOLATERULE1"))="參、四、(七)" Then response.write "selected"%>>參、四、(七)</option>
							<option value="參、四、(八)" <%If Trim(rs("VIOLATERULE1"))="參、四、(八)" Then response.write "selected"%>>參、四、(八)</option>
							<option value="參、四、(九)" <%If Trim(rs("VIOLATERULE1"))="參、四、(九)" Then response.write "selected"%>>參、四、(九)</option>
							<option value="參、五、(一)" <%If Trim(rs("VIOLATERULE1"))="參、五、(一)" Then response.write "selected"%>>參、五、(一)</option>
							<option value="參、五、(二)" <%If Trim(rs("VIOLATERULE1"))="參、五、(二)" Then response.write "selected"%>>參、五、(二)</option>
							<option value="參、五、(三)" <%If Trim(rs("VIOLATERULE1"))="參、五、(三)" Then response.write "selected"%>>參、五、(三)</option>
							<option value="參、五、(四)" <%If Trim(rs("VIOLATERULE1"))="參、五、(四)" Then response.write "selected"%>>參、五、(四)</option>
						</select>
						
					</td>
					<td bgcolor="#FFFF99">違反規定(目)</td>
					<td>
						<select name="VIOLATERULE2">
							<option value="">無違反規定</option>
							<option value="之 1" <%If Trim(rs("VIOLATERULE2"))="之 1" Then response.write "selected"%>>之 1</option>
							<option value="之 2" <%If Trim(rs("VIOLATERULE2"))="之 2" Then response.write "selected"%>>之 2</option>
							<option value="之 3" <%If Trim(rs("VIOLATERULE2"))="之 3" Then response.write "selected"%>>之 3</option>
							<option value="之 4" <%If Trim(rs("VIOLATERULE2"))="之 4" Then response.write "selected"%>>之 4</option>
							<option value="之 5" <%If Trim(rs("VIOLATERULE2"))="之 5" Then response.write "selected"%>>之 5</option>
							<option value="之 6" <%If Trim(rs("VIOLATERULE2"))="之 6" Then response.write "selected"%>>之 6</option>
							<option value="之 7" <%If Trim(rs("VIOLATERULE2"))="之 7" Then response.write "selected"%>>之 7</option>
							<option value="之 8" <%If Trim(rs("VIOLATERULE2"))="之 8" Then response.write "selected"%>>之 8</option>
							<option value="之 9" <%If Trim(rs("VIOLATERULE2"))="之 9" Then response.write "selected"%>>之 9</option>
							<option value="之 10" <%If Trim(rs("VIOLATERULE2"))="之 10" Then response.write "selected"%>>之 10</option>
						</select>
					</td>
				</tr>	
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99">懲處情形</td>
					<td colspan="3">
						<select name="Sys_Punishment">
							<option value="無" <%If Trim(rs("Punishment"))="無" Then response.write "selected"%>>無</option>
							<option value="劣蹟1次" <%If Trim(rs("Punishment"))="劣蹟1次" Then response.write "selected"%>>劣蹟1次</option>
							<option value="申誡1次" <%If Trim(rs("Punishment"))="申誡1次" Then response.write "selected"%>>申誡1次</option>
							<option value="申誡2次" <%If Trim(rs("Punishment"))="申誡2次" Then response.write "selected"%>>申誡2次</option>
						</select>
					</td>
				</tr>
				
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">處分(通報)日期</td>
					<td>
					   <input name="Actiondate" class="btn1" type="text" value="<%
			if not isnull(rs("Actiondate")) and trim(rs("Actiondate"))<>"" then
				response.write year(rs("Actiondate"))-1911&right("00"&month(rs("Actiondate")),2)&right("00"&day(rs("Actiondate")),2)
			end if
				%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Actiondate');">
					</td>
					<td bgcolor="#FFFF99">處分(通報)文號</td>
					<td>
						<input name="Actionno" class="btn1" type="text" value="<%
							response.write Trim(rs("Actionno"))
								%>" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99">是否撤銷</td>
					<td>
						<select name="Sys_Cancel" class="btn1">
							<option value=0 <%if rs("Cancel")=0 then response.write " selected"%>>是</option>
							<option value=1 <%if rs("Cancel")=1 then response.write " selected"%>>否</option>
						</select>
					</td>
					<td bgcolor="#FFFF99">撤銷舉發單理由</td>
					<td>
						<select name="Sys_DelReasonID" class="btn1" onchange="funDelOthen();" >
							<option value="0">請選擇</option><%
							strSQLD="select Content,ID from Code where TypeID=21 order by ID"
							set rsD=conn.execute(strSQLD)
							while Not rsD.eof
								response.write "<option value="""&rsD("ID")&""""
								if trim(rs("DELBILLREASON"))=trim(rsD("ID")) then response.write " selected"
								response.write ">"
								response.write rsD("Content")
								response.write "</option>"
								rsD.movenext
							wend
							rsD.close%>
						</select>
						<span id="inpudel"></span>
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">備註</td>
					<td colspan="3">
						<textarea name="Sys_Note" class="btn1" cols="40"><%=Trim(rs("Note"))%></textarea>
					</td>
				</tr>
				
				<tr bgcolor="#ffffff">
					
					<td bgcolor="#FFFF99"><font color="red">* </font>案件狀態</td>
					<td>
						<select name="Sys_Close" class="btn1">
							<option value=0<%if rs("Close")=0 then response.write " selected"%>>未處理</option>
							<option value=1<%if rs("Close")=1 then response.write " selected"%>>結案</option>
							<option value=2<%if rs("Close")=2 then response.write " selected"%>>待查中</option>
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
		inpuerr.innerHTML ="<br>原因：<input type=text name='ErrName' value='<%=rs("ErrorName")%>' size=30 class='btn1'>";
	}else{
		inpuerr.innerHTML ="";
	}
}
function funDelOthen(){ 
	if(myForm.Sys_DelReasonID.value=='811'){
		inpudel.innerHTML ="<br>理由：<input type=text name='DelName' value='<%=rs("DelName")%>' size=30 class='btn1'>";
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
funDelOthen();
funErrOthen();
</script>
<%
	rs.close
	Set rs=nothing
%>
<%conn.close%>