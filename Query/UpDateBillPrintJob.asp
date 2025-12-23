<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單代印管理</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body><%
	strwhere=""
	If trim(Request("DB_Selt")) = "Upda" Then
		If not ifnull(Request("Sys_tmpBillNo")) Then
			strSQL="Update BillPrintJob set Note='"&trim(request("Sys_Note"))&"' where Batchnumber='"&trim(Request("Sys_tmpBatChnumber"))&"' and BillNo='"&trim(Request("Sys_tmpBillNo"))&"'"
			conn.execute(strSQL)
		else
			strSQL="Update BillPrintJob set Note='"&trim(request("Sys_Note"))&"' where Batchnumber='"&trim(Request("Sys_tmpBatChnumber"))&"' and BillNo is null"
			conn.execute(strSQL)
		End if
		Response.write "<script>"
		Response.Write "alert('已更新完成！');"
		Response.write "</script>"
	end if
	if trim(Request("DB_Selt"))="Selt" then

		if Not ifnull(Request("Sys_ReUnitID")) then
			strwhere=strwhere&" and a.RequestUnitID='"&trim(request("Sys_ReUnitID"))&"'"
		end if

		if Not ifnull(Request("Sys_ReBatChNumber")) then
			strwhere=strwhere&" and a.BatchNumber='"&trim(request("Sys_ReBatChNumber"))&"'"
		end if

		if (Not ifnull(Request("PrintType"))) and Request("PrintType")<>"3" then
			strwhere=strwhere&" and a.PrintStatus="&trim(request("PrintType"))
		end if

		strSQL="select a.*,DeCode(a.PrintStatus,0,'未列印','已列印') PrintTypeName,b.UnitName,c.ChName from BillPrintJob a,UnitInfo b,MemberData c where a.RequestUnitID=b.UnitID and a.RequestMemberID=c.MemberID" & strwhere & " order by RequestUnitID,RequestMemberID,PrintDateTime DESC"

		set rsdata=conn.execute(strSQL)
		If Not rsdata.eof Then DB_Selt="Selt"

		strCnt="select count(*) cnt from BillPrintJob a where batchnumber=batchnumber" & strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close

	else

		strSQL="select a.*,DeCode(a.PrintStatus,0,'未列印','已列印') PrintTypeName,b.UnitName,c.ChName from BillPrintJob a,UnitInfo b,MemberData c where a.RequestUnitID=b.UnitID and a.RequestMemberID=c.MemberID and a.PrintStatus=0 order by PrintDateTime DESC"

		set rsdata=conn.execute(strSQL)
		If Not rsdata.eof Then DB_Selt="Selt"

		strCnt="select count(*) cnt from BillPrintJob where PrintStatus=0"
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close

	End if
%>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="33">代印查詢</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td>代印單位</td>
					<td><%
						strSQL="select UnitID,UnitName from UnitInfo"
						set rsuit=conn.execute(strSQL)
						Response.Write "<select name=""Sys_ReUnitID"" class=""btn1"">"
						Response.Write "<option value="""">全部</option>"
						While Not rsuit.eof
							Response.Write "<option value="""&trim(rsuit("UnitID"))&""""
							If trim(request("Sys_ReUnitID")) = trim(rsuit("UnitID")) Then
								Response.Write " selected"
							End if							
							Response.Write ">"&trim(rsuit("UnitName"))&"</option>"
							rsuit.movenext
						Wend
						Response.Write "</select>"
						rsuit.close
						%>
					</td>
					<td>代印批號</td>
					<td>
						<input name="Sys_ReBatChNumber" class="btn1" type="text" value="<%=request("Sys_ReBatChNumber")%>" size="12" maxlength="12">
					</td>
					<td>列印狀態</td>
					<td>
						<select Name="PrintType">
							<option value="0"<%If trim(Request("PrintType")) = "0" Then Response.Write " selected"%>>未列印</option>
							<option value="3"<%If trim(Request("PrintType")) = "3" Then Response.Write " selected"%>>全部</option>
							<option value="1"<%If trim(Request("PrintType")) = "1" Then Response.Write " selected"%>>已列印</option>
						</select>
					</td>
					<td>
						<input type="submit" name="btnSelt" value="查詢" onClick='funSelt();'>
						<input type="button" name="cancel" value="清除" onClick="location='UpDateBillPrintJob.asp';">
						<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">( 查詢 <%=DBsum%> 筆紀錄 )
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
		</td>
	</tr><%
	if DB_Selt="Selt" then%>
		<tr>
			<td bgcolor="#E0E0E0">
				<table width="100%" height="100%" border="0" cellpadding="1" cellspacing="1">
					<tr bgcolor="#EBFBE3" align="center">
						<th>代印單位</th>
						<th>代印人員</th>
						<th>代印批號</th>
						<th>單號</th>
						<th>件數</th>
						<th>列印狀態</th>
						<th>列印日期</th>
						<th>備註</th>
						<th>操作</th>
					</tr><%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsdata.eof then rsdata.move cdbl(DBcnt)
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsdata.eof then exit for

						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						response.write ">"

						response.write "<td align='left'>"&rsdata("UnitName")&"</td>"
						response.write "<td align='left'>"&rsdata("ChName")&"</td>"
						response.write "<td align='left'>"&rsdata("BatchNumber")&"</td>"
						response.write "<td align='left'>"&trim(rsdata("BillNo"))&"</td>"
						response.write "<td align='left'>"&trim(rsdata("PrintCnt"))&"</td>"
						response.write "<td align='left'>"&rsdata("PrintTypeName")&"</td>"
						response.write "<td align='left'>"&rsdata("PrintDateTime")&"</td>"

						response.write "<td align='left'>"
						Response.Write "<input name=""Sys_Note_"& i &""" class=""btn1"" type=""text"" value="""&rsdata("Note")&""" size=""12"">"
						Response.Write "</td>"

						response.write "<td>"
						response.write "<input type=""button"" name=""Update"" value=""修改"" onclick=""funUpdate("&i&",'"&rsdata("BatchNumber")&"','"&trim(rsdata("BillNo"))&"');"">"

						response.write "<input type=""button"" name=""Stop"" value=""代印"" onclick=""funRePrint('"&rsdata("BatchNumber")&"','"&trim(rsdata("BillNo"))&"');"">"
						response.write "</td>"
						response.write "</tr>"
						rsdata.movenext
					next%>
				</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFDD77" align="center">
				<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
				<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
				<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">		
			</td>
		</tr>
	<%end if%>
</table>
<input type="Hidden" name="DB_Selt" value=<%=DB_Selt%>>
<input type="Hidden" name="Sys_Note" value="">
<input type="Hidden" name="Sys_tmpBillNo" value="">
<input type="Hidden" name="Sys_tmpBatChnumber" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}

function funSelt(){
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}

function funRePrint(BatcNumber,ReBillNo){
	runServerScript("UpdatePrintFile.asp?batchnumber="+BatcNumber+"&ReBillNo="+ReBillNo);
}

function funUpdate(cmt,BatcNumber,ReBillNo){
	myForm.DB_Selt.value='Upda';
	myForm.Sys_tmpBatChnumber.value=BatcNumber;
	myForm.Sys_tmpBillNo.value=ReBillNo;
	myForm.Sys_Note.value=eval("myForm.Sys_Note_"+cmt).value;
	myForm.submit();
}

function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10+eval(myForm.sys_MoveCnt.value))==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))-1)*(10+eval(myForm.sys_MoveCnt.value));
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))*(10+eval(myForm.sys_MoveCnt.value));
		}
		myForm.submit();
	}
}

function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}
</script>