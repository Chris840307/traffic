<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舊資料查詢</title>
<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->

<!--#include virtual="Traffic/Common/OlddbAccessKao.ini"-->

<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<!--#include virtual="Traffic/Common/Login_Check.asp"--> 

<%
Server.ScriptTimeout=12000
if request("DB_Selt")="DelBillno" Then
		strUpdDel="Update FMaster_S set CloseFlag='Y' where FSEQ='"&request("DelBillno")&"'"
		conn1.execute strUpdDel
			Response.write "<script>"
			Response.Write "alert('儲存完成！');"
			response.write "window.location.href='OldBillBaseQueryKao.asp';"
			Response.write "</script>"
End if

'組成查詢SQL字串
		strwhere=" where 1=1 "
if request("DB_Selt")="Selt" then


		if request("IllegalDate")<>"" and request("IllegalDate1")<>"" then
			  if len(request("IllegalDate"))=6 then
				ArgueDate1="0"&request("IllegalDate")
			  else
				ArgueDate1=request("IllegalDate")
			  end if
			  if len(request("IllegalDate"))=6 then
				ArgueDate2="0"&request("IllegalDate1")
			  else
				ArgueDate2=request("IllegalDate1")
			  end if
				strwhere=strwhere&" and IDate between '"&ArgueDate1&"' and '"&ArgueDate2&"'"
		end if
		
		if request("CarNo")<>"" Then strwhere=strwhere&" and CarNo = '"&request("CarNo")&"'"

		if request("BillNo")<>"" Then strwhere=strwhere&" and FSEQ = '"&request("BillNo")&"'"

		if request("UnitID")<>"" Then strwhere=strwhere&" and PBCODE = '"&request("UnitID")&"'"
		if request("DriverName")<>"" Then strwhere=strwhere&" and Iname = '"&request("DriverName")&"'"
		if request("Note")<>"" Then strwhere=strwhere&" and Note = '"&request("Note")&"'"

		if request("DriverID")<>"" then strwhere=strwhere&" and IIDNO='"&request("DriverID")&"'"

		If Trim(request("BillTypeID"))<>"" Then
			If request("BillTypeID")="1" Then
				strwhere=strwhere&" and AccUSeCode='"&request("BillTypeID")&"'"
			ElseIf request("BillTypeID")="2" Then 
				strwhere=strwhere&" and AccUSeCode='"&request("BillTypeID")&"'"
			ElseIf request("BillTypeID")="8" Then 
				strwhere=strwhere&" and AccUSeCode='"&request("BillTypeID")&"'"
			End if
		End if

		strSQL="select AccUSeCode,FSEQ,CarNo,IDate,ITime,CDKIND,IName,IRName,RuleF1,'' AS CloseFlag,'' as ReplyDate,'' as Recvno from FMaster " & strwhere 

		set rsfound1=conn1.execute(strSQL)
		set rsfound2=conn2.execute(strSQL)
		set rsfound3=conn3.execute(strSQL)
		set rsfound4=conn4.execute(strSQL)
		set rsfound5=conn5.execute(strSQL)
		set rsfound6=conn6.execute(strSQL)
		set rsfound7=conn7.execute(strSQL)
		

		
		'set rsfound8=conn8.execute(strSQL)
		
					strSQL="select count(AccUSeCode) as cnt from FMaster " & strwhere
					set cnt1=conn1.execute(strSQL)		
					set cnt2=conn2.execute(strSQL)
					set cnt3=conn3.execute(strSQL)
					set cnt4=conn4.execute(strSQL)
					set cnt5=conn5.execute(strSQL)
					set cnt6=conn6.execute(strSQL)
					set cnt7=conn7.execute(strSQL)


'				set cnt8=conn8.execute(strSQL)

		strSQL="select distinct(FStatus) as FS from FinBack  " 
	   	set rsPolice=conn1.execute(strSQL)

	   if rsPolice.eof then
		   set rsPolice=conn2.execute(strSQL)
		   if rsPolice.eof then
			   set rsPolice=conn3.execute(strSQL)
				if rsPolice.eof then
					set rsPolice=conn4.execute(strSQL)
	 				if rsPolice.eof then
						 set rsPolice=conn5.execute(strSQL)
	 	 				if rsPolice.eof then
							 set rsPolice=conn6.execute(strSQL)
			 				if rsPolice.eof then
								 set rsPolice=conn7.execute(strSQL)
							End if
						End if
					End if
				End if
		   end if
	   end If


'		   	DBsum=CDbl(cnt1("cnt"))+CDbl(cnt2("cnt"))+CDbl(cnt3("cnt"))+CDbl(cnt4("cnt"))+CDbl(cnt5("cnt"))+CDbl(cnt6("cnt"))+CDbl(cnt7("cnt"))+CDbl(cnt8("cnt"))
		   	DBsum=CDbl(cnt1("cnt"))+CDbl(cnt2("cnt"))+CDbl(cnt3("cnt"))+CDbl(cnt4("cnt"))+CDbl(cnt5("cnt"))+CDbl(cnt6("cnt"))+CDbl(cnt7("cnt"))

		set cnt1=Nothing
		set cnt2=Nothing
		set cnt3=Nothing
		set cnt4=nothing		
		set cnt5=Nothing
		set cnt6=nothing		
		set cnt7=nothing

end If

Function GetAccUSeCodeName(AccUSeCode)
	if AccUSeCode="1" then 
	  GetAccUSeCodeName="攔停"
	elseif AccUSeCode="2" then 
	  GetAccUSeCodeName="逕舉"
	elseif AccUSeCode="8" then 
	  GetAccUSeCodeName="行人攤販"
	elseif AccUSeCode="3" then 
	  GetAccUSeCodeName="肇事"
	elseif AccUSeCode="4" then 
	  GetAccUSeCodeName="拖吊"
	elseif AccUSeCode="5" then 
	  GetAccUSeCodeName="戴運砂石土方"
	elseif AccUSeCode="A" then 
	  GetAccUSeCodeName="違規營業"
	elseif AccUSeCode="B" then 
	  GetAccUSeCodeName="違規重標"
	elseif AccUSeCode="N" then 
	  GetAccUSeCodeName="未知"
	Else
	  GetAccUSeCodeName="慢車行人"						
	end if 
End Function

Function GetCarTypeName(CDKIND)
	if CDKIND="1" then
		GetCarTypeName="汽車"
	elseif CDKIND="2" then
		GetCarTypeName="拖車"
	elseif CDKIND="3" then
		GetCarTypeName="重機"
	elseif CDKIND="4" then
		GetCarTypeName="輕機"
	Else
		GetCarTypeName="　"						
	end if
End function

%>
<html>
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
					<td bgcolor="#FFCC33" colspan="5"><b>舊資料查詢</b><img src="space.gif" width="20" height="2"> <A HREF="..\舊資料查詢系統.doc"><FONT SIZE="3"><b>!!  第一次使用請看.DOC !! </b> </font></A>
					</td>
				</tr>			
				<tr>
					<td>

						違規日期
						<input name="IllegalDate" type="text" value="<%=request("IllegalDate")%>" size="5" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate');">
						~
						<input name="IllegalDate1" type="text" value="<%=request("IllegalDate1")%>" size="5" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate1');">

												<img src="space.gif" width="10" height="2">
						證號
						<input name="DriverID" type="text" value="<%=request("DriverID")%>" size="8" maxlength="10" class="btn1" onkeyup="value=value.toUpperCase()">
												<img src="space.gif" width="10" height="2">
						車<img src="space.gif" width="5" height="2">號
						<input name="CarNo" type="text" value="<%=request("CarNo")%>" size="5" maxlength="8" class="btn1" onkeyup="value=value.toUpperCase()">				
						
						<img src="space.gif" width="10" height="2">
						<b>單<img src="space.gif" width="5" height="2">號</b>
						<input name="BillNo" type="text" value="<%=request("BillNo")%>" size="9" maxlength="9" class="btn1" onkeyup="value=value.toUpperCase()">			

								<br>
								舉發單位代碼
								<input name="UnitID" type="text" value="<%=request("UnitID")%>" size="9" maxlength="10" class="btn1"  onkeyup="value=value.toUpperCase()">
								舉發單類型
								 <Select Name="BillTypeID">
								   <option value="" <%if trim(request("BillTypeID"))="" then response.write " Selected"%>>全部</option>
								   <option value="2" <%if trim(request("BillTypeID"))="2" then response.write " Selected"%>>攔停</option>
								   <option value="3" <%if trim(request("BillTypeID"))="3" then response.write " Selected"%>>逕舉</option>
								   <option value="8" <%if trim(request("BillTypeID"))="8" then response.write " Selected"%>>行人攤販</option>
								 </select>

								
								違規人姓名
								<input name="DriverName" type="text" value="<%=request("DriverName")%>" size="9" maxlength="20" class="btn1">
								
								
								備註
								<input name="Note" type="text" value="<%=request("Note")%>" size="30" maxlength="30" class="btn1">
						</td><tr><td align="Center"  colspan="5">
   						<img src="space.gif" width="15" height="1"><br>
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" >
						<input type="button" name="cancel" value="清除" onClick="location='OldBillBaseQuerykao.asp'"> 
						
					  </td>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	
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
					<th nowrap>類別</th>
					<th nowrap>舉發單號</th>
					<th nowrap>車號</th>
					<th >違規日</th>
					<th >車種</th>
					<th >駕駛人</th>
					<th >違規地點</th>
					<th >法條</th>
					<th>結案日期</th>
					<th>收據號碼</th>
					<th >操作</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%

				if request("DB_Selt")="Selt"  then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end If
'1-----------------------------------------------------------------------------------------------------------------------------------------------------------------------

					if Not rsfound1.eof then rsfound1.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound1.eof then exit for
					   response.flush
						chname="":chRule="":ForFeit=""

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"

    					response.write "<td>"
   						response.write GetAccUSeCodeName(rsfound1("AccUSeCode")&"")
						response.write "</td>"

    					response.write "<td>"& rsfound1("FSEQ") & "</td>"
						response.write "<td>"& rsfound1("CarNo")& "</td>"
						response.write "<td>"
   						response.write left(rsfound1("IDate"),3)&"/"& mid(rsfound1("IDate"),4,2)&"/"& Right(rsfound1("IDate"),2)& " " &left(rsfound1("ITime"),2)&":"&right(rsfound1("ITime"),2)
						response.write "</td>"

						response.write "<td>"
   					    response.write GetCarTypeName(rsfound1("CDKIND"))
						response.write "</td>"

						response.write "<td>" & rsfound1("IName") & "</td>"
						response.write "<td align='left'>"& rsfound1("IRName")& "</td>"						
						response.write "<td>"& rsfound1("RuleF1")& "</td>"	

						response.write "<td>&nbsp;</td>"	
						response.write "<td>&nbsp;</td>"	

						response.write "<td align='left' >"


%>
						<input type="button" name="btnModify" value="修改" onclick='window.open("OldBaseNoteModifyKao.asp?BillNo=<%=trim(rsfound1("FSEQ"))%>","OldBaseModify","left=300,top=400,location=0,width=600,height=200,resizable=no,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 40px; height:26px;">
						<input type="button" name="b1" value="詳細" onclick='window.open("OldBaseDetailKao.asp?BillNo=<%=trim(rsfound1("FSEQ"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
						<%

						response.write "</td>"
						response.write "</tr>"
						rsfound1.movenext
					Next
'2-----------------------------------------------------------------------------------------------------------------------------------------------------------------------

					if Not rsfound2.eof then rsfound2.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound2.eof then exit for
					   response.flush					   
						chname="":chRule="":ForFeit=""

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"

    					response.write "<td>"
   						response.write GetAccUSeCodeName(rsfound2("AccUSeCode")&"")
						response.write "</td>"

    					response.write "<td>"& rsfound2("FSEQ") & "</td>"
						response.write "<td>"& rsfound2("CarNo")& "</td>"
						response.write "<td>"
   						response.write left(rsfound2("IDate"),3)&"/"& mid(rsfound2("IDate"),4,2)&"/"& Right(rsfound2("IDate"),2)& " " &left(rsfound2("ITime"),2)&":"&right(rsfound2("ITime"),2)

						response.write "<td>"
   					    response.write GetCarTypeName(rsfound2("CDKIND"))
						response.write "</td>"

						response.write "<td>" & rsfound2("IName") & "</td>"
						response.write "<td align='left'>"& rsfound2("IRName")& "</td>"						
						response.write "<td>"& rsfound2("RuleF1")& "</td>"	

						response.write "<td>&nbsp;</td>"	
						response.write "<td>&nbsp;</td>"	

						response.write "<td align='left' >"


%>
						<input type="button" name="btnModify" value="修改" onclick='window.open("OldBaseNoteModifyKao.asp?BillNo=<%=trim(rsfound2("FSEQ"))%>","OldBaseModify","left=300,top=400,location=0,width=600,height=200,resizable=no,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 40px; height:26px;">

						<input type="button" name="b1" value="詳細" onclick='window.open("OldBaseDetailKao.asp?BillNo=<%=trim(rsfound2("FSEQ"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
						<%
						response.write "</td>"
						response.write "</tr>"
						rsfound2.movenext
					Next

'3-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
					if Not rsfound3.eof then rsfound3.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound3.eof then exit for
					   response.flush					   
						chname="":chRule="":ForFeit=""

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"

    					response.write "<td>"
   						response.write GetAccUSeCodeName(rsfound3("AccUSeCode")&"")
						response.write "</td>"

    					response.write "<td>"& rsfound3("FSEQ") & "</td>"
						response.write "<td>"& rsfound3("CarNo")& "</td>"
						response.write "<td>"
   						response.write left(rsfound3("IDate"),3)&"/"& mid(rsfound3("IDate"),4,2)&"/"& Right(rsfound3("IDate"),2)& " " &left(rsfound3("ITime"),2)&":"&right(rsfound3("ITime"),2)
						response.write "</td>"

						response.write "<td>"
   					    response.write GetCarTypeName(rsfound3("CDKIND"))
						response.write "</td>"

						response.write "<td>" & rsfound3("IName") & "</td>"
						response.write "<td align='left'>"& rsfound3("IRName")& "</td>"						
						response.write "<td>"& rsfound3("RuleF1")& "</td>"	

						response.write "<td>&nbsp;</td>"	
						response.write "<td>&nbsp;</td>"	

						response.write "<td align='left' >"


%>
						<input type="button" name="btnModify" value="修改" onclick='window.open("OldBaseNoteModifyKao.asp?BillNo=<%=trim(rsfound3("FSEQ"))%>","OldBaseModify","left=300,top=400,location=0,width=600,height=200,resizable=no,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 40px; height:26px;">

						<input type="button" name="b1" value="詳細" onclick='window.open("OldBaseDetailKao.asp?BillNo=<%=trim(rsfound3("FSEQ"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
<%
						response.write "</td>"
						response.write "</tr>"
						rsfound3.movenext
					Next

'4-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
					if Not rsfound4.eof then rsfound4.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound4.eof then exit for
					   response.flush					   
						chname="":chRule="":ForFeit=""

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"

    					response.write "<td>"
   						response.write GetAccUSeCodeName(rsfound4("AccUSeCode")&"")
						response.write "</td>"

    					response.write "<td>"& rsfound4("FSEQ") & "</td>"
						response.write "<td>"& rsfound4("CarNo")& "</td>"
						response.write "<td>"
   						response.write left(rsfound4("IDate"),3)&"/"& mid(rsfound4("IDate"),4,2)&"/"& Right(rsfound4("IDate"),2)& " " &left(rsfound4("ITime"),2)&":"&right(rsfound4("ITime"),2)
						response.write "</td>"

						response.write "<td>"
   					    response.write GetCarTypeName(rsfound4("CDKIND"))
						response.write "</td>"

						response.write "<td>" & rsfound4("IName") & "</td>"
						response.write "<td align='left'>"& rsfound4("IRName")& "</td>"						
						response.write "<td>"& rsfound4("RuleF1")& "</td>"	

						response.write "<td>&nbsp;</td>"	
						response.write "<td>&nbsp;</td>"	

						response.write "<td align='left' >"


%>
						<input type="button" name="btnModify" value="修改" onclick='window.open("OldBaseNoteModifyKao.asp?BillNo=<%=trim(rsfound4("FSEQ"))%>","OldBaseModify","left=300,top=400,location=0,width=600,height=200,resizable=no,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 40px; height:26px;">

						<input type="button" name="b1" value="詳細" onclick='window.open("OldBaseDetailKao.asp?BillNo=<%=trim(rsfound4("FSEQ"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
<%
						response.write "</td>"
						response.write "</tr>"
						rsfound4.movenext
					Next

'5-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
					if Not rsfound5.eof then rsfound5.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound5.eof then exit for
					   response.flush					   
						chname="":chRule="":ForFeit=""

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"

    					response.write "<td>"
   						response.write GetAccUSeCodeName(rsfound5("AccUSeCode")&"")
						response.write "</td>"

    					response.write "<td>"& rsfound5("FSEQ") & "</td>"
						response.write "<td>"& rsfound5("CarNo")& "</td>"
						response.write "<td>"
   						response.write left(rsfound5("IDate"),3)&"/"& mid(rsfound5("IDate"),4,2)&"/"& Right(rsfound5("IDate"),2)& " " &left(rsfound5("ITime"),2)&":"&right(rsfound5("ITime"),2)
						response.write "</td>"

						response.write "<td>"
   					    response.write GetCarTypeName(rsfound5("CDKIND"))
						response.write "</td>"

						response.write "<td>" & rsfound5("IName") & "</td>"
						response.write "<td align='left'>"& rsfound5("IRName")& "</td>"						
						response.write "<td>"& rsfound5("RuleF1")& "</td>"	

						response.write "<td>&nbsp;</td>"	
						response.write "<td>&nbsp;</td>"	

						response.write "<td align='left' >"


%>
						<input type="button" name="btnModify" value="修改" onclick='window.open("OldBaseNoteModifyKao.asp?BillNo=<%=trim(rsfound5("FSEQ"))%>","OldBaseModify","left=300,top=400,location=0,width=600,height=200,resizable=no,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 40px; height:26px;">

						<input type="button" name="b1" value="詳細" onclick='window.open("OldBaseDetailKao.asp?BillNo=<%=trim(rsfound5("FSEQ"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
<%
						response.write "</td>"
						response.write "</tr>"
						rsfound5.movenext
					Next

'6-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
					if Not rsfound6.eof then rsfound6.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound6.eof then exit for
					   response.flush					   
						chname="":chRule="":ForFeit=""

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"

    					response.write "<td>"
   						response.write GetAccUSeCodeName(rsfound6("AccUSeCode")&"")
						response.write "</td>"

    					response.write "<td>"& rsfound6("FSEQ") & "</td>"
						response.write "<td>"& rsfound6("CarNo")& "</td>"
						response.write "<td>"
   						response.write left(rsfound6("IDate"),3)&"/"& mid(rsfound6("IDate"),4,2)&"/"& Right(rsfound6("IDate"),2)& " " &left(rsfound6("ITime"),2)&":"&right(rsfound6("ITime"),2)
						response.write "</td>"

						response.write "<td>"
   					    response.write GetCarTypeName(rsfound6("CDKIND"))
						response.write "</td>"

						response.write "<td>" & rsfound6("IName") & "</td>"
						response.write "<td align='left'>"& rsfound6("IRName")& "</td>"						
						response.write "<td>"& rsfound6("RuleF1")& "</td>"	

						response.write "<td>&nbsp;</td>"	
						response.write "<td>&nbsp;</td>"	

						response.write "<td align='left' >"


%>
						<input type="button" name="btnModify" value="修改" onclick='window.open("OldBaseNoteModifyKao.asp?BillNo=<%=trim(rsfound6("FSEQ"))%>","OldBaseModify","left=300,top=400,location=0,width=600,height=200,resizable=no,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 40px; height:26px;">

						<input type="button" name="b1" value="詳細" onclick='window.open("OldBaseDetailKao.asp?BillNo=<%=trim(rsfound6("FSEQ"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
<%
						response.write "</td>"
						response.write "</tr>"
						rsfound6.movenext
					Next

'7-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
					if Not rsfound7.eof then rsfound7.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound7.eof then exit for
					   response.flush					   
						chname="":chRule="":ForFeit=""

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"

    					response.write "<td>"
   						response.write GetAccUSeCodeName(rsfound7("AccUSeCode")&"")
						response.write "</td>"

    					response.write "<td>"& rsfound7("FSEQ") & "</td>"
						response.write "<td>"& rsfound7("CarNo")& "</td>"
						response.write "<td>"
   						response.write left(rsfound7("IDate"),3)&"/"& mid(rsfound7("IDate"),4,2)&"/"& Right(rsfound7("IDate"),2)& " " &left(rsfound7("ITime"),2)&":"&right(rsfound7("ITime"),2)
						response.write "</td>"

						response.write "<td>"
   					    response.write GetCarTypeName(rsfound7("CDKIND"))
						response.write "</td>"

						response.write "<td>" & rsfound7("IName") & "</td>"
						response.write "<td align='left'>"& rsfound7("IRName")& "</td>"						
						response.write "<td>"& rsfound7("RuleF1")& "</td>"	

						response.write "<td>&nbsp;</td>"	
						response.write "<td>&nbsp;</td>"	

						response.write "<td align='left' >"


%>
						<input type="button" name="btnModify" value="修改" onclick='window.open("OOldBaseNoteModifyKao.asp?BillNo=<%=trim(rsfound7("FSEQ"))%>","OldBaseModify","left=300,top=400,location=0,width=600,height=200,resizable=no,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 40px; height:26px;">

						<input type="button" name="b1" value="詳細" onclick='window.open("OldBaseDetailKao.asp?BillNo=<%=trim(rsfound7("FSEQ"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
<%
						response.write "</td>"
						response.write "</tr>"
						rsfound7.movenext
					Next
	End If
	

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
				
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFFFFF" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

				<input type="button" name="bntExcel" value="匯出 Excel" onclick="funExcel();">

		</td>
	</tr>
</table>

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="delBillno" value="">
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

			if (myForm.IllegalDate.value=="" && myForm.IllegalDate1.value=="" && myForm.DriverID.value==""  && myForm.CarNo.value=="" && myForm.BillNo.value=="" && myForm.UnitID.value=="" && myForm.DriverName.value=="" && myForm.Note.value=="") {
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

	function funExcel(){

		urlstr="OldBillBaseListExcel.asp?strwhere=<%=strwhere%>&CloseFlag=<%="N"%>";

        newWin(urlstr,"ListExcel",980,580,0,0,"yes","yes","yes","no");
	}

	function funOldPasserBook(){

		urlstr="OldPasserBook.asp?strwhere=<%=strwhere%>";
        newWin(urlstr,"ListExcel",980,580,0,0,"yes","yes","yes","no");
	}
	
	
//銷案
	function DelBill(Billno){
		myForm.delBillno.value=Billno;
		myForm.DB_Selt.value="DelBillno";
		myForm.submit();
	}

</script>
<%

		conn1.close
		set conn1=nothing
		conn2.close
		set conn2=nothing
		conn3.close
		set conn3=Nothing
		conn4.close
		set conn4=Nothing
		conn5.close
		set conn5=Nothing
		conn6.close
		set conn6=Nothing
		conn7.close
		set conn7=Nothing

%>