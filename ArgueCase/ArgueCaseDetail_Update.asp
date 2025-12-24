<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴案件</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
if request("DB_ADD")="ADD" then
	DocDate=gOutDT(request("DocDate"))
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
		If Trim(request("BadCnt"))<>"" Then
			BadCnttmp=Trim(request("BadCnt"))
		End If 
		If Trim(request("WarnCnt"))<>"" Then
			WarnCnttmp=Trim(request("WarnCnt"))
		End If 
	strSQL="Update ArgueBase set BillNo='"&request("Sys_BillNo")&"',CarNo='"&request("Sys_CarNo")&"'" &_
		",ReportContent='"&request("Sys_ReportContent")&"',DocNo='"&request("Sys_DocNo")&"'," &_
		"DocDate="&funGetDate(DocDate,0)&",Punishment='"&request("Sys_Punishment")&"',ArgueDate="&funGetDate(ArgueDate,0)&","&_
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
		strSQL="Update ArgueBase set BillNo='"&request("Sys_BillNo")&"',CarNo='"&request("Sys_CarNo")&"',ReportContent='"&request("Sys_ReportContent")&"',DocNo='"&request("Sys_DocNo")&"',DocDate="&funGetDate(DocDate,0)&",Punishment='"&request("Sys_Punishment")&"',ArgueDate="&funGetDate(ArgueDate,0)&",Arguer='"&request("Sys_Arguer")&"',Note='"&request("Sys_Note")&"',ArguerCreditID='"&request("Sys_ArguerCreditID")&"',ArguerAddr='"&request("Sys_ArguerAddr")&"',ArguerResonID="&request("Sys_ArguerResonID")&",ArguerResonName='"&request("ResonName")&"',ArguerTel='"&request("Sys_ArguerTel")&"',ErrorID="&request("Sys_ErrorID")&",ErrorName='"&request("ErrName")&"',ArguerContent='"&request("Sys_ArguerContent")&"',Cancel='"&request("Sys_Cancel")&"',Close='"&request("Sys_Close")&"' where SN="&request("SN")
	end if
'response.write strSQL
'response.end
	conn.execute(strSQL)
	%>
	<script language="JavaScript">
		alert ("修改完成!!");
		opener.myForm.submit(); 
		self.close();
	</script><% 
else
	strSQL="select * from ArgueBase where SN="&request("SN")
	set rs=conn.execute(strSQL)
	DocDate=gInitDT(rs("DocDate"))
	ArgueDate=gInitDT(rs("ArgueDate"))
	sBillNo=rs("BillNo")
%>
<body>
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr>
		<td height="27" bgcolor="#FFCC33">申訴案件</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#dddddd">
				<tr  bgcolor="#ffffff">
					<td width="20%" bgcolor="#FFFF99"><font color="red">* </font>舉發單號</td>
					<td width="25%">
						<input name="red_BillNo" class="btn1" type="text" value="<%=sBillNo%>" size="12" maxlength="9" disabled>
						<input type="Hidden" name="Sys_BillNo" value="<%=sBillNo%>">
					</td>
					<td width="13%"  bgcolor="#FFFF99">辦理情形</td>
					<td width="42%">
						<textarea name="Sys_ReportContent" cols="40" class="btn1"><%=rs("ReportContent")%></textarea>
					</td>
				</tr>
				<!--
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">車牌號碼</td>
					<td>																								
						<input name="Sys_CarNo" class="btn1" type="text" value="" size="12" maxlength="12" disabled>
					</td>
				</tr>
				-->
				
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">收文號</td>
					<td>
						<input name="Sys_DocNo" class="btn1" type="text" value="<%=rs("DocNo")%>" size="12" maxlength="30">
					</td>
					<td rowspan="2" bgcolor="#FFFF99">懲處情形</td>
					<td rowspan="2">
						<textarea name="Sys_Punishment" cols="40" class="btn1"><%=rs("Punishment")%></textarea>
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">發文日期</td>
					<td>
						<input name="DocDate" class="btn1" type="text" value="<%=DocDate%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('DocDate');">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述人姓名</td>
					<td>
						<input name="Sys_Arguer" class="btn1" type="text" value="<%=rs("Arguer")%>" size="11" maxlength="11">
					</td>
					<td rowspan="2" bgcolor="#FFFF99">備註</td>
					<td rowspan="2">
						<textarea name="Sys_Note" class="btn1" cols="40"><%=rs("Note")%></textarea>
					</td>
				</tr>
				<tr bgcolor="#ffffff">					
					<td bgcolor="#FFFF99"><font color="red">* </font>陳述日期</td>
					<td>
						<input name="ArgueDate" class="btn1" type="text" value="<%=ArgueDate%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ArgueDate');">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述人身分證</td>
					<td>
						<input name="Sys_ArguerCreditID" class="btn1" type="text" value="<%=rs("ArguerCreditID")%>" size="11" maxlength="11">
					</td>
					<td bgcolor="#FFFF99">陳述原因</td>
					<td>
						<select name="Sys_ArguerResonID" class="btn1" onchange="funOthen();">
							<option  value="0">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=15 order by ID"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								if trim(rs("ArguerResonID"))=trim(rs1("ID")) then response.write " selected"
								response.write ">"
								response.write rs1("Content")
								response.write "</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						<span id="inputxt"></span>
					</td>
				</tr bgcolor="#ffffff">
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述人地址</td>
					<td>
						<input name="Sys_ArguerAddr" class="btn1" type="text" value="<%=rs("ArguerAddr")%>" size="25" maxlength="30">
					</td>
					<td bgcolor="#FFFF99">缺失原因</td>
					<td>
						<select name="Sys_ErrorID" class="btn1" onchange="funErrOthen();">
							<option value="0">無缺失</option><%
							strSQL="select Content,ID from Code where TypeID=16 order by ID"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								if trim(rs("ErrorID"))=trim(rs1("ID")) then response.write " selected"
								response.write ">"
								response.write rs1("Content")
								response.write "</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						<span id="inpuerr"></span>
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述人電話</td>
					<td>
						<input name="Sys_ArguerTel" class="btn1" type="text" value="<%=rs("ArguerTel")%>" size="25" maxlength="30">
					</td>
					<td bgcolor="#FFFF99">是否撤銷</td>
					<td>
						<select name="Sys_Cancel" class="btn1">
							<option value="">請選擇</option>
							<option value=0<%if rs("Cancel")=0 then response.write " selected"%>>是</option>
							<option value=1<%if rs("Cancel")=1 then response.write " selected"%>>否</option>
						</select>
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">陳述內容補充摘要</td>
					<td>
						<input name="Sys_ArguerContent" class="btn1" type="text" value="<%=rs("ArguerContent")%>" size="25" maxlength="30">
					</td>
					<td bgcolor="#FFFF99">案件狀態</td>
					<td>
						<select name="Sys_Close" class="btn1">
							<option value=0<%if rs("Close")=0 then response.write " selected"%>>未處理</option>
							<option value=1<%if rs("Close")=1 then response.write " selected"%>>結案</option>
							<option value=2<%if rs("Close")=2 then response.write " selected"%>>待查中</option>
						</select>
					</td>
				</tr>
		<% if sys_City="澎湖縣" Or sys_City="台南市" then %>
                        <tr bgcolor="#ffffff">
                            <td bgcolor="#FFFF99">陳情方式</td>
                            <td>
                                <input name="argueway" class="btn1" type="text" value="<%=trim(rs("argueway"))%>" size="25" maxlength="30">
                            </td>
                            <td bgcolor="#FFFF99">撤銷舉發單理由</td>
                            <td>
								<select name="Sys_DelReasonID" class="btn1" onchange="funDelOthen();" >
									<option value="0">請選擇</option><%
									strSQL="select Content,ID from Code where TypeID=21 order by ID"
									set rs1=conn.execute(strSQL)
									while Not rs1.eof
										response.write "<option value="""&rs1("ID")&""""
								if trim(rs("DELBILLREASON"))=trim(rs1("ID")) then response.write " selected"
								response.write ">"
										response.write rs1("Content")
										response.write "</option>"
										rs1.movenext
									wend
									rs1.close%>
								</select>
								<span id="inpudel"></span>
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
                                <input name="VIOLATERULE1" class="btn1" type="text" value="<%
							response.write Trim(rs("VIOLATERULE1"))
								%>" size="25" maxlength="30">
                            </td>
                            <td bgcolor="#FFFF99">違反規定(目)</td>
                            <td>
                                <input name="VIOLATERULE2" class="btn1" type="text" value="<%
							response.write Trim(rs("VIOLATERULE2"))
								%>" size="25" maxlength="30">
                            </td>
                        </tr>		
						<tr bgcolor="#ffffff">
                            <td bgcolor="#FFFF99">劣蹟處分</td>
                            <td>
                               <input name="BadCnt" class="btn1" type="text" value="<%
							If isnull(rs("BadCnt")) Or Trim(rs("BadCnt"))="" then
								response.write "0"
							Else
								response.write Trim(rs("BadCnt"))
							End If 
								%>" size="4" maxlength="6" onkeyup="value=value.replace(/[^\d]/g,'')"> 次
	
                            </td>
                            <td bgcolor="#FFFF99">申誡處分</td>
                            <td>
                                <input name="WarnCnt" class="btn1" type="text" value="<%
							If isnull(rs("WarnCnt")) Or Trim(rs("WarnCnt"))="" then
								response.write "0"
							Else
								response.write Trim(rs("WarnCnt"))
							End If 
								%>" size="4" maxlength="6" onkeyup="value=value.replace(/[^\d]/g,'')"> 次
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
                <% end if %>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">紀錄人員</td>
					<td colspan="3">
						<%=Session("Ch_Name")%>
					</td>
				</tr>
				
		  </table>
		</td>
	</tr>
	<tr bgcolor="#ffffff" align="center">
		<td height="35" bgcolor="#FFDD77">
			<input type="button" name="save" value=" 儲 存 " onclick="funAdd();">
			<input type="button" name="exit" value=" 關 閉  " onclick="funExt();">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Add" value="">
<input type="Hidden" name="SN" value="<%=request("SN")%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
funOthen();
funErrOthen();
function funOthen(){ 
	if(myForm.Sys_ArguerResonID.value=='448'){
		inputxt.innerHTML ="<br>原因：<input type=text name='ResonName' value='<%=rs("ArguerResonName")%>' size=30 class='btn1'>";
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
<% if sys_City="澎湖縣" Or sys_City="台南市" then %>
function funDelOthen(){ 
	if(myForm.Sys_DelReasonID.value=='811'){
		inpudel.innerHTML ="<br>理由：<input type=text name='DelName' value='<%=rs("DelName")%>' size=30 class='btn1'>";
	}else{
		inpudel.innerHTML ="";
	}
}
<%end if%>
function funAdd(){
	var err=0;
	if(myForm.DocDate.value!=""){
		if(!dateCheck(myForm.DocDate.value)){
			err=1;
			alert("陳述日輸入不正確!!");
		}
	}
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
		if(myForm.Sys_BillNo.value==''){
			alert("舉發單號不可空白");
		}else{
			runServerScript("chkAddNew.asp?BillNo="+myForm.Sys_BillNo.value);
		}
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		self.close();
	}
}
funDelOthen();
</script>
<%end if
conn.close%>