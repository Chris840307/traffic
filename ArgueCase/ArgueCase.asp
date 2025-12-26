
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style2 {font-size: 28px; line-height:100%;}
-->
</style>
<title>申訴案件</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/bannernodata.asp"-->
<%
'AuthorityCheck(225)
Server.ScriptTimeout = 60800
Response.flush

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing


if request("DB_state")="Del" then
	sql1 = "Delete From ARGUEDETAIL Where ARGUEBASESN="&request("SN")
    sql2 = "Update ArgueBase set RecordStateID=-1,DelMemberID="&Session("User_ID")&" where SN="&request("SN")
	Conn.BeginTrans
	Conn.Execute(sql1)
	Conn.Execute(sql2)
	if err.number = 0 then
		Conn.CommitTrans
		Response.write "<script>"
		Response.Write "alert('刪除完成！');"
		Response.write "</script>"
	else
		Conn.RollbackTrans
		Response.write "<script>"
		Response.Write "alert('刪除失敗！');"
		Response.write "</script>"
	end if
end if
DB_Selt=trim(request("DB_Selt"))
DB_Display=trim(request("DB_Display"))
if DB_Selt="Selt" then
	strwhere=""
	strCancel=""
	if request("ArgueDate")<>"" and request("ArgueDate1")<>""then
		ArgueDate1=gOutDT(request("ArgueDate"))
		ArgueDate2=gOutDT(request("ArgueDate1"))
		strwhere=" and a.ArgueDate between "&funGetDate(ArgueDate1,0)&" and "&funGetDate(ArgueDate2,0)
	end if
	if request("Sys_DocNo")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.DocNo='"&request("Sys_DocNo")&"'"
		else
			strwhere=" and a.DocNo='"&request("Sys_DocNo")&"'"
		end if
	end if
	if request("Sys_Arguer")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.Arguer like '%"&request("Sys_Arguer")&"%'"
		else
			strwhere=" and a.Arguer like '%"&request("Sys_Arguer")&"%'"
		end if
	end if
	if request("Sys_ArguerCreditID")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.ArguerCreditID='"&request("Sys_ArguerCreditID")&"'"
		else
			strwhere=" and a.ArguerCreditID='"&request("Sys_ArguerCreditID")&"'"
		end if
	end if
	if request("Sys_BillNo")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.BillNo='"&request("Sys_BillNo")&"'"
		else
			strwhere=" and a.BillNo='"&request("Sys_BillNo")&"'"
		end if
	end if
	if request("Sys_Cancel")<>"" then
		if strCancel<>"" then
			strCancel=strCancel&" and a.Cancel='"&request("Sys_Cancel")&"'"
		else
			strCancel=" and a.Cancel='"&request("Sys_Cancel")&"'"
		end if
	end if
	if request("Sys_Close")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.Close='"&request("Sys_Close")&"'"
		else
			strwhere=" and a.Close='"&request("Sys_Close")&"'"
		end if																														
	end if
	if request("Sys_Unit")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.RecordMemberID in (select MemberID from MemberData where UnitID in " &_
				" (select UnitID from UnitInfo where (UnitID='"&request("Sys_Unit")&"') or (UnitTypeID='"&request("Sys_Unit")&"'" &_
				" and ShowOrder=2)))"
		else
			strwhere=" and a.RecordMemberID in (select MemberID from MemberData where UnitID in " &_
				" (select UnitID from UnitInfo where (UnitID='"&request("Sys_Unit")&"') or (UnitTypeID='"&request("Sys_Unit")&"'" &_
				" and ShowOrder=2)))"
		end if																														
	end if
end if
if trim(DB_Selt)="ArgueDateSelt" then
	if trim(request("Sys_ArgueDate"))<>"" then
		strwhere=" and a.ArgueDate <"&funGetDate(DateADD("d",0-Cint(request("Sys_ArgueDate")),Date),0)&" and a.Close='0'"
	else
		strwhere=" and a.Close='0'"
	end if
end if

if DB_Display="show" then
	if trim(strwhere)<>"" then
		if sys_City="澎湖縣" then
			specialsql=" a.argueway,a.reportdeparment,a.reportno,a.processdate,a.processno,a.DELBILLREASON,a.VIOLATERULE1,a.VIOLATERULE2,a.Actiondate,a.Actionno,a.BadCnt,a.WarnCnt,"
		elseif sys_City="台南市" or sys_City="高雄市" then
			specialsql=" a.argueway,a.reportdeparment,a.reportno,a.processdate,a.processno,a.DELBILLREASON,a.VIOLATERULE1,a.VIOLATERULE2,a.Actiondate,a.Actionno,a.BadCnt,a.WarnCnt,a.DelName, "
		else
			specialsql=""
		end if
		strSQL="select " & specialsql & "a.Punishment,a.Note,a.Arguer,a.SN, a.ArgueDate,a.BillNo,a.ArguerResonID,a.ArguerResonName,a.ErrorID,a.ErrorName,a.ReportContent,c.Content as ArguerContent,a.ArguerContent as ArguerContent2,d.Content as ErrorConten,a.Cancel,a.Close,a.DocNo from ArgueBase a,Code c,Code d where a.ArguerResonID=c.ID and a.ErrorID=d.ID and a.RecordStateID=0"&strwhere&strCancel
		set rsfound=conn.execute(strSQL)

					'and a.CarNo=b.CarNo Cancel
		'Response.write strSQL
		strCnt="select count(*) as cnt from ArgueBase a,Code c,Code d where a.ArguerResonID=c.ID and a.ErrorID=d.ID and a.RecordStateID=0 and a.Cancel='0'"&strwhere
		set Dbrs=conn.execute(strCnt)
		DBCancel=Cint(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from ArgueBase a,Code c,Code d where a.ArguerResonID=c.ID and a.ErrorID=d.ID and a.RecordStateID=0"&strwhere&strCancel
		set Dbrs=conn.execute(strCnt)
		DBsum=Cint(Dbrs("cnt"))
		Dbrs.close

		if Cint(DBsum)<>0 then parentCnt=fix(DBCancel/DBsum*100)
		tmpSQL=strSQL
	else
		DB_Selt="":DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
	
end if
'if DB_Selt="" then																																																	'and a.CarNo=b.CarNo
'	strSQL="select a.SN,a.ArgueDate,a.BillNo,b.BillMem1,b.BillMem2,b.BillMem3,c.Content as ArguerContent,b.Rule1,b.Rule2,b.Rule3,b.Rule4,d.Content as ErrorConten,a.Cancel,a.Close,a.DocNo from ArgueBase a,BillBaseView b,Code c,Code d where a.BillNo=b.BillNo and a.ArguerResonID=c.ID and a.ErrorID=d.ID and a.ArgueDate >="&funGetDate(DateADD("d",-3,Date),0)&" and a.Close='0' and a.RecordStateID=0" 'access日期時間沒有單引號
'	set rsfound=conn.execute(strSQL)
'																								'and a.CarNo=b.CarNo 
'	strCnt="select count(*) as cnt from ArgueBase a,BillBaseView b,Code c,Code d where a.BillNo=b.BillNo and a.ArguerResonID=c.ID and a.ErrorID=d.ID and a.ArgueDate >="&funGetDate(DateADD("d",-3,Date),0)&" and a.Close='0' and a.RecordStateID=0"
'
'	set Dbrs=conn.execute(strCnt)
'	DBsum=Cint(Dbrs("cnt"))
'	Dbrs.close
''response.write strCnt
''response.end
'	tmpSQL=strSQL
'end if
%>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF" height="39">申訴案件(<a href="使用說明_申訴系統.doc" target="_blank"><font  class="style2">使用說明下載</font></a>	)</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td width="5%" nowrap>陳述日</td>
					<td width="18%" nowrap>
						&nbsp;&nbsp;<input type="text" class="btn1" name="ArgueDate" size="5" value="<%=request("ArgueDate")%>" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ArgueDate');">
						&nbsp;~&nbsp;
						<input type="text" class="btn1" name="ArgueDate1" size="5" value="<%=request("ArgueDate1")%>" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ArgueDate1');">
						&nbsp;&nbsp;
					</td>
				    <td nowrap><div align="right">收文號</div></td>
				    <td nowrap><input name="Sys_DocNo" class="btn1" value="<%=request("Sys_DocNo")%>" type="text" size="9" maxlength="30"></td>
				    <td nowrap><div align="right">陳述人姓名</div></td>
				    <td nowrap><input name="Sys_Arguer" class="btn1" type="text" value="<%=request("Sys_Arguer")%>" size="6" maxlength="5"></td>
				    <td nowrap><div align="left">身分證號
			          <input name="Sys_ArguerCreditID" class="btn1" type="text" value="<%=request("Sys_ArguerCreditID")%>" size="20" maxlength="20" onkeyup="this.value=this.value.toUpperCase()">
				    </div></td>
			    </tr>
				<tr>
					<td width="5%" nowrap>舉發單號</td>
					<td><input name="Sys_BillNo" class="btn1" type="text" value="<%=request("Sys_BillNo")%>" size="12" maxlength="9" onkeyup="this.value=this.value.toUpperCase()">
  						<input type="button" name="cancel" value="詳細資料" onClick="funBillBaseDetail();">					
					
					</td>
					<td width="4%" nowrap><div align="right">撤銷</div></td>
					<td width="5%" nowrap><select name="Sys_Cancel" class="btn1">
                      <option value="">請選擇...</option>
					  <option value='0'<%if trim(request("Sys_Cancel"))="0" then response.write " selected"%>>是</option>
                      <option value='1'<%if trim(request("Sys_Cancel"))="1" then response.write " selected"%>>否</option>
                      
                    </select></td>
					<td width="3%" nowrap><div align="right">結案否</div></td>
					<td width="7%" nowrap><select name="Sys_Close" class="btn1">
                      <option value="">請選擇...</option>
                      <option value='1'<%if trim(request("Sys_Close"))="1" then response.write " selected"%>>是</option>
                      <option value='0'<%if trim(request("Sys_Close"))="0" then response.write " selected"%>>否</option>
                    </select></td>
					<td nowrap>
					受理單位
						<select name="Sys_Unit" class="btn1">
						<option value="">所有單位</option>
		<%
				strUnit="select * from UnitInfo where ShowOrder in (0,1) order by showorder,Unitid"
				set rsUnit=conn.execute(strUnit)
				If Not rsUnit.Bof Then rsUnit.MoveFirst 
				While Not rsUnit.Eof
		%>
						<option value="<%=trim(rsUnit("UnitID"))%>" <%if trim(request("Sys_Unit"))=trim(rsUnit("UnitID")) then response.write "selected"%>><%=trim(rsUnit("UnitName"))%></option>
		<%
				rsUnit.MoveNext
				Wend
				rsUnit.close
				set rsUnit=nothing
		%>
						</select>
					<input type="button" name="btnSelt" value="查詢" onClick='funSelt();'<%'if Not CheckPermission(225,1) then response.write " disabled"%>>
					  &nbsp;&nbsp;
                      <input type="button" name="btnAdd" value="新增" onClick='funInsert();'<%'if Not CheckPermission(225,2) then response.write " disabled"%>>
                      &nbsp;&nbsp;
                      <input type="button" name="cancel" value="清除" onClick="location='ArgueCase.asp'">
					</td>
			    </tr>
				<tr>
					<td colspan="7">
						<HR>
						申訴超過
						<input name="Sys_ArgueDate" class="btn1" type="text" value="<%=request("Sys_ArgueDate")%>" size="1" maxlength="4" onkeyup="value=value.replace(/[^\d]/g,'')">
						天未結案案件
						<input type="button" name="btnSelt" value="查詢" onclick="funArgueDateSelt();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(225,1)=false then
'							response.write "disabled"
'						end if
						%>>
						<strong>申訴成功比率： <%=Cint(parentCnt)%> % (共<%=Cint(DBsum)%>件，<%=Cint(DBCancel)%>件撤銷 )</strong>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#1BF5FF">申訴案件紀錄列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="34">申訴日期</th>
					<th height="34">舉發單號</th>
					<th height="34">舉發員警</th>
					<th height="34">陳述事由</th>
					<th height="34">適用條款</th>
					<th height="34">造成缺失原因</th>
					<th height="34">是否撤銷</th>
					<th height="34">案件狀態</th>
					<th height="34">操作</th>
				</tr><%
					if DB_Display="show" then
						if Trim(request("DB_Move"))="" then
							DBcnt=0
						else
							DBcnt=request("DB_Move")
						end if
						if Not rsfound.eof then rsfound.move Cint(DBcnt)
						for i=DBcnt+1 to DBcnt+10
							if rsfound.eof then exit For
							chname="":chRule=""
							strB="select * from (select * from BillBaseView where billno='"&Trim(rsfound("BillNo"))&"' order by Recorddate desc) where Rownum<=1"
							Set rsB=conn.execute(strB)
							If Not rsB.eof Then
								
								if rsB("BillMem1")<>"" then	chname=rsB("BillMem1")
								if rsB("BillMem2")<>"" then	chname=chname&"/"&rsB("BillMem2")
								if rsB("BillMem3")<>"" then	chname=chname&"/"&rsB("BillMem3")
								if rsB("Rule1")<>"" then chRule=rsB("Rule1")
								if rsB("Rule2")<>"" then chRule=chRule&"/"&rsB("Rule2")
								if rsB("Rule3")<>"" then chRule=chRule&"/"&rsB("Rule3")
								if rsB("Rule4")<>"" then chRule=chRule&"/"&rsB("Rule4")
							End If
							rsB.close
							Set rsB=Nothing 
							
							if rsfound("Cancel")="0" then
								chkCancel="是"
							else
								chkCancel="否"
							end if
							
							if rsfound("Close")="0" then
								chkClose="未處理"
							elseif rsfound("Close")="1" then
								chkClose="結案"
							elseif rsfound("Close")="2" then
								chkClose="待查中"
							end if
							response.write "<tr bgcolor='#FFFFFF' align='center' "
							lightbarstyle 0 
							response.write ">" 
							response.write "<td>"&gInitDT(rsfound("ArgueDate"))&"</td>"
							response.write "<td>"&rsfound("BillNo")&"</td>"
							response.write "<td>"&chname&"</td>"
							
							If sys_City="台南市" or sys_City="高雄市" Then
								response.write "<td>"&rsfound("ArguerContent")&"</td>"
							Else
								if trim(rsfound("ArguerResonID"))="448" then
									response.write "<td>"&rsfound("ArguerResonName")&"</td>"
								else
									response.write "<td>"&rsfound("ArguerContent")&"</td>"
								end if
							End If 
							
							response.write "<td>"&chRule&"</td>"
							If sys_City="台南市" or sys_City="高雄市" Then
								if trim(rsfound("ErrorID"))="0" then
									response.write "<td>無缺失</td>"
								else
									response.write "<td>"&rsfound("ErrorConten")&"</td>"
								end If
							else
								if trim(rsfound("ErrorID"))="453" then
									response.write "<td>"&rsfound("ErrorName")&"</td>"
								elseif trim(rsfound("ErrorID"))="0" then
									response.write "<td>無缺失</td>"
								else
									response.write "<td>"&rsfound("ErrorConten")&"</td>"
								end If
							End If 

							response.write "<td>"&chkCancel&"</td>"
							response.write "<td>"&chkClose&"</td>"
							response.write "<td>"
							
							response.write "<input type=""button"" name=""Update"" value=""修改"" onclick=""funUpdate('"&rsfound("SN")&"');"""
							'if Not CheckPermission(225,3) then response.writ " disabled"
							response.write ">"
							response.write "<input type=""button"" name=""Del"" value=""刪除"" onclick=""funDel('"&rsfound("SN")&"');"""
							'if Not CheckPermission(225,4) then response.writ " disabled"
							response.write ">"
							response.write "<input type=""button"" name=""addfile"" value=""附件"" onclick=""funDetail('"&rsfound("SN")&"');""></td>"
							response.write "</tr>"
							rsfound.movenext
						next
					end if%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#1BF5FF" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=Cint(DBcnt)/10+1&"/"&fix(Cint(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<% if sys_City="澎湖縣" Or sys_City="台南市" or sys_City="高雄市" then %>
				<input type="button" name="btnExecel" value="民眾申述案件管制一覽表" onclick="funchgExecelNew();">
				<input type="button" name="btnExecel2" value="申訴項目分析表" onclick="funAnaExecel();">
			<% end if %>
			<% if sys_City="澎湖縣" Or sys_City="台南市" or sys_City="高雄市" then %>
				<input type="button" name="btnExece22" value="處理交通違規陳情、陳述統計表"   onclick="funArgExecel();"style="width: 255px; ">
			<% end if %>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="DB_Display" value="<%=trim(DB_Display)%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="SN" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSelt(){
	var error=0;
	if(myForm.ArgueDate.value!=""){
		if(!dateCheck(myForm.ArgueDate.value)){
			error=1;
			alert("陳述日輸入不正確!!");
		}
	}
	if (error==0){
		if(myForm.ArgueDate1.value!=""){
			if(!dateCheck(myForm.ArgueDate1.value)){
				error=1;
				alert("陳述日輸入不正確!!");
			}
		}
		if (error==0){
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.DB_Display.value='show';
			myForm.submit();
		}
	}
}
function funArgueDateSelt(){
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="ArgueDateSelt";
	myForm.DB_Display.value='show';
	myForm.submit();
}
function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}
}
function funchgExecel(){
	UrlStr="ArgueCase_Execel.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funchgExecelNew(){
	UrlStr="ArgueCase_ExecelNew.asp?SQLstr=<%=tmpSQL%>&Sys_Unit="+myForm.Sys_Unit.value;
	newWin(UrlStr,"inputWin1",900,550,50,10,"yes","yes","yes","no");
}
function funInsert(){
<%
	if sys_City="台南市" or sys_City="高雄市" then
%>
	UrlStr="ArgueCaseDetail_Insert_NT.asp";
<%
	else
%>
	UrlStr="ArgueCaseDetail_Insert.asp";
<%
	end if 
%>
	newWin(UrlStr,"inputWin",1000,650,50,10,"yes","no","yes","no");
}
function funDetail(SN){
	UrlStr="ArgueCaseAttach.asp?ArgueBaseSN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funUpdate(SN){
<%
	if sys_City="台南市" or sys_City="高雄市" then
%>
	UrlStr="ArgueCaseDetail_Update_NT.asp?SN="+SN;
<%
	else
%>
	UrlStr="ArgueCaseDetail_Update.asp?SN="+SN;
<%
	end if 
%>
	
	newWin(UrlStr,"inputWin",1000,650,50,10,"yes","no","yes","no");
}
function funAnaExecel(){
	UrlStr="ArgueCaseAnalyze.asp";
	newWin(UrlStr,"inputWin3",900,550,50,10,"yes","yes","yes","no");
}
function funBillBaseDetail(){
	if(myForm.Sys_BillNo.value!=''){
		runServerScript("Bill_SN.asp?BillNo="+myForm.Sys_BillNo.value);
	}
}
function funArgExecel(){
	UrlStr="ArgueCaseReport_set.asp";
	newWin(UrlStr,"inputWin4",500,450,80,50,"yes","no","yes","no");
}
function funDel(SN){
	myForm.SN.value=SN;
	myForm.DB_state.value="Del";
	myForm.submit();
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}

</script>
<%conn.close%>