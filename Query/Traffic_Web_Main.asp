<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
#Layer1 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
	top: 13px;
}
.style1 {font-size: 14px}
.style2 {font-size: 13px; }
#Layer2 {
	position:absolute;
	width:566px;
	height:38px;
	z-index:3;
	top: 15px;
}
#LayerTime {
	position:absolute;
	width:311px;
	height:34px;
	z-index:1;
}
#Layer151 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
}
#D1 {
	BACKGROUND-COLOR: #99FFFF; 
	BORDER-BOTTOM: white 2px outset; 
	BORDER-LEFT: white 2px outset; 
	BORDER-RIGHT: white 2px outset; 
	BORDER-TOP: white 2px outset; 
	LEFT: 0px; POSITION: absolute; 
	TOP: 0px; VISIBILITY: hidden; 
	WIDTH: 150px; 
	layer-background-color: #99FFFF;
	
	z-index:5;
}
-->
</style>
<head>
<!--#include virtual="traffic/Common/css.txt"-->
<title>智慧型交通執法系統</title>
<%
memName=Session("Ch_Name")
GroupID=Session("Group_ID")
UnitNo=Session("Unit_ID")

 	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

'	if sys_City="台中市" then
'		strUn="select * from UnitInfo where UnitID='0469' and UnitName='交通隊直屬第一分隊'"
'		set rsUn=conn.execute(strUn)
'		if rsUn.eof then
'			strUnit1="Update UnitInfo set UnitName='交通隊直屬第一分隊' where UnitID='0469'"
'			conn.execute strUnit1
'			strUnit2="Update UnitInfo set UnitName='交通隊直屬第二分隊' where UnitID='046A'"
'			conn.execute strUnit2
'		end if
'		rsUn.close
'		set rsUn=nothing
'	end if

	if sys_City<>"台中縣" then
		ArgueDate1=DateAdd("d",-10,date) & " 0:0:0"
		ArgueDate2=date & " 23:29:59" 
		strDelErr="select * from Dcilog where ExchangeTypeID='E' and (DciReturnStatusID<>'S')" &_
			" and ExchangeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS')" &_
			" and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID="&Session("User_ID")
		set rsDelErr=conn.execute(strDelErr)
		If Not rsDelErr.Bof Then rsDelErr.MoveFirst 
		While Not rsDelErr.Eof
			if (trim(rsDelErr("DciReturnStatusID"))<>"S" and not isnull(rsDelErr("DciReturnStatusID"))) then
				if trim(rsDelErr("DciReturnStatusID"))="n" then
					strUpd="Update BillBase set RecordStateID=0,BillStatus='2' where RecordStateID<>0 and Sn="&trim(rsDelErr("BillSn"))
					conn.execute strUpd
				else
					strUpd="Update BillBase set RecordStateID=0,BillStatus='2' where RecordStateID<>0 and Sn="&trim(rsDelErr("BillSn"))
					conn.execute strUpd
				end if
			end if
		rsDelErr.MoveNext
		Wend
		rsDelErr.close
		set rsDelErr=nothing
	end if
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>
<body leftmargin="25" topmargin="5" marginwidth="0" marginheight="0" <%
	if sys_City<>"台中市" and sys_City<>"台中縣" then
%>
		onLoad="init()"
<%
	end if
%>>
<div id="D1" style="width:350px">
<%
	if sys_City<>"台中市" and sys_City<>"台中縣" then
%>
  <table border="0" width="350">
    <TBODY> 
	<tr>
		<!-- 處理進度 -->
		<td valign="top"> 
			<table width="100%" border="1">
				<tr bgcolor="#FFFF99">
					<td >處理進度(1~68條)</td>
				</tr>
				<tr bgColor="#FFFFFF">
					<td id="YestodayLayer">

					</td>
				</tr>
				<tr bgColor="#FFFFFF">
					<td id="TodayLayer">
					
					</td>
				</tr>
			</table>
		</td>
		<td align="middle" bgColor="#FF0000" rowSpan="2" width="20" >
			<font color="#FFFFFF"> 
			個<br>人<br>資<br>料
			</font>
		</td>
    </tr>
    <tr>
		<!-- 上傳紀錄 -->
		<td valign="top" id="UpLoadLayer"> 

		</td>
    </tr>
    </TBODY> 
  </table>
<%
	end if
%>
</div>

<font color="#000000"><b>
<%

strUnit="select * from UnitInfo where UnitID='"&UnitNo&"'"
set rsUnit=conn.execute(strUnit)
UnitName=rsUnit("UnitName")
rsUnit.close
set rsUnit=nothing

%>
</b>
</font>
<table width='980' border='0' align="center" >
	<tr>
		<td colspan="5" height="40">
			<div id="Layer1">
			<table width="230" align="left" border="0">
				<tr>
					<td class="style1" nowrap><strong>
						<%=UnitName%>
						<img src="image/space.gif" alt="" width="5" height="2" border="0" align="baseline">
						<%=memName%></strong>
						
							<img src="image/space.gif" alt="" width="5" height="1" border="0" align="baseline">
							<input type="button" name="b10010" value="修改個人資料" onclick="funMember();" style="font-size: 9pt; width: 90px; height:23px;">
				  		<img src="Image/dot.gif" >
				  		<a href="SystemInit.htm" target="_blank" class="style2">** 首次使用系統注意事項 ** </a>			  
				  	
				    	<img src="Image/dot.gif" ><font class="style2">
				   		客服 : (02) 2736-6866
				      <img src="Image/space.gif" width="5" height="1" >
				      <img src="Image/dot.gif" ondblclick="location='UserLogout_Contral.asp'">
				      傳真 : (02) 2736-6777 		    	
				      
				      
				      <!---加上 檢查 flash_recovery_area_usage 是否快滿的顯示--->
				<%if Session("Credit_ID")="A000000000" then%>
				      <img src="Image/dot.gif" >
							v$flash_recovery_area_usage
							<%
								strDB="select PERCENT_SPACE_USED from v$flash_recovery_area_usage  where File_Type='ARCHIVELOG'"
								set rsDB=conn.execute(strDB)
								if not rsDB.eof then
									response.write cint(rsDB("PERCENT_SPACE_USED"))
									if cint(rsDB("PERCENT_SPACE_USED"))>"90" and Session("Credit_ID")="A000000000" then
							%>
									<script language="JavaScript">
										alert("Oracle 暫存區快滿了，請趕快清一清");
									</script>
							<%
									end if
								end if
								rsDB.close
								set rsDB=nothing
			  			%>%			
				<%end if%>
							</font>
				      
					</td>
				</tr>
				<tr>
					<td class="style1" nowrap><strong>
					
					<img src="image/space.gif" alt="" width="5" height="1" border="0" align="baseline">
					<%
'					strGName="select Content from Code where TypeID=10 and ID="&trim(GroupID)
'					set rsGName=conn.execute(strGName)
'					if not rsGName.eof then
'						response.write trim(rsGName("Content"))
'					end if
'					rsGName.close
'					set rsGName=nothing
									
					%></strong>
					</td>
				</tr>
			</table>
			</div>
				<img src="image/space.gif" alt="" width="300" height="2" border="0" align="baseline">
				<!-- <input type="button" name="update1" value="Update" onclick="funcUpdate();"> -->
					<!-- <input type="button" name="b10010" value="登 出" onclick="location='UserLogout_Contral.asp'" > --></span>
		     <!-- <div id="Layer1">
		     <span class="style1">
			  	<table width="181" align="left">
				<tr>
				<td width="93" align="center" class="style2">相片剩餘空間：</td>
				<td width="85" align="right" class="style2"> --><%
'			  set fs=Server.CreateObject("Scripting.FileSystemObject")
'			  FileName=Server.MapPath("storespace.ini")
'			  
'			  if fs.FileExists(FileName) then
'			  	set txtf=fs.OpenTextFile(FileName)
'				if not txtf.atEndOfStream then
'					PicFree=txtf.ReadAll
'					response.write PicFree
'				end if
'				txtf.close
'			  	set txtf=nothing
'			  end if
'			  set fs=nothing
			  %>
			  
			  <!-- &nbsp;MB
			  	</td>
				</tr>
				<tr>
				<td align="center" class="style2">資料剩餘空間：</td>
				<td align="right" class="style2"> --><%
'			  strDB="SELECT a.Tablespace_Name,a.Bytes / 1024 / 1024 TotalMb," &_
'			  	"b.Bytes / 1024 / 1024 UsedMb,c.Bytes / 1024 / 1024 FreeMb," &_
'				"(c.Bytes * 100) / a.Bytes ""% FREE""  FROM Sys.Sm$ts_Avail a, Sys.Sm$ts_Used b,  Sys.Sm$ts_Free c" &_
'				" WHERE a.Tablespace_Name = b.Tablespace_Name  AND a.Tablespace_Name = c.Tablespace_Name" &_
'				" And a.Tablespace_Name='TRAFFIC'"
'				set rsDB=conn.execute(strDB)
'				if not rsDB.eof then
'					response.write rsDB("FreeMb")
'				end if
'				rsDB.close
'				set rsDB=nothing
			  %>
			  
			  <!-- &nbsp;MB
			  	</td>
				</tr>
			  </table>
			  </span> 
		    </div> -->
		    

					<!--
				  <img src="Image/dot.gif" >
				  <a href="Help.rar" class="style2">使用手冊 V3 (0504) </a>
				  -->	
<br><br>
<hr>
	  <td>
 
	</tr>
	<tr>
		<td colspan="5" class="style3">
			<font color="red"><strong>快速查詢</strong></font>
			<img src="image/space.gif" alt="" width="12" height="2" border="0" align="baseline">
			&nbsp;<strong>舉發單號</strong>&nbsp;<input type="text" name="BillNo" size="10" maxlength="9" onkeyup="value=value.toUpperCase()">
			&nbsp;<strong>車號</strong>&nbsp;<input type="text" name="CarNo" size="8" maxlength="9" onkeyup="value=value.toUpperCase()">
			<!-- 花蓮拖吊場限制查詢條件 -->
			<% if Session("Credit_ID")="0000" AND sys_City="花蓮縣" then %>
				&nbsp;<strong></strong>&nbsp;<input type="hidden" name="IllegalName" size="12" maxlength="30">
				&nbsp;<strong></strong>&nbsp;<input type="hidden" name="IllegalID" size="10" maxlength="10" onkeyup="value=value.toUpperCase()">	

                       <% else %>
				&nbsp;<strong>姓名</strong>&nbsp;<input type="text" name="IllegalName" size="12" maxlength="30">
				&nbsp;<strong>身份證號</strong>&nbsp;<input type="text" name="IllegalID" size="10" maxlength="10" onkeyup="value=value.toUpperCase()">	


			 <% end if %>
			<input type="button" value="查詢" onclick='getBillData()'>
			<span class='style2'>可跨分局</span>
			<input type="button" value="事故系統佔用資源" onclick="window.open('tjdbCPU.asp','winother','width=400,height=350,left=0,top=0,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 100px; height: 27px">
		</td>
	</tr>
<%
strFuncGroup="select * from Code where TypeID=19 order by ShowOrder"
set rsFuncGroup=conn.execute(strFuncGroup)
If Not rsFuncGroup.Bof Then rsFuncGroup.MoveFirst 
While Not rsFuncGroup.Eof
%>
	<tr><td width="20%"></td><td width="20%"></td><td width="20%"></td><td width="20%"></td><td width="20%"></td><tr>
	<tr>
		<td colspan="5">
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr>
			    <td bgcolor="#FFCC33">
				<font size="4"><strong><%=trim(rsFuncGroup("Content"))%></strong>
				</font>	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<% if rsFuncGroup("ID")="512" then
				%>
					<font size="3"><strong><div id="LayerTime" ondblclick="NowDownProcess();"></div></strong></font>		
				<%  end if
				
					if rsFuncGroup("ID")="509" then
				%>
					<br>
					<font size="5">
					<strong>
					<img src="Image/dot.gif"></img>
					<a href="send.doc" target="_blank" >民眾簽收 簡要說明.doc</a>
					<img src="Image/dot.gif"></img>
					<a href="opengov.doc" target="_blank" >公示送達 簡要說明.doc</a>
					<img src="Image/dot.gif"></img>
					<a href="storeandsend.doc" target="_blank" >寄存送達 簡要說明.doc</a>
					</strong>
					
					</font>		
				<%  end if				
							
				%>
			    </td>
			</tr>
		</table>
		</td>
	</tr>
<%
	s=1
	strFunc="select a.* from FunctionPageData a,FunctionData b where a.SystemID=b.SystemID and b.GroupID='"&trim(GroupID)&"' and a.SystemGroupID="&trim(rsFuncGroup("ID"))&" and b.Function='1' order by ShowOrder"
	set rsFunc=conn.execute(strFunc)
	If Not rsFunc.Bof Then rsFunc.MoveFirst 
	While Not rsFunc.Eof

	if s=1 then
		response.write "<tr>"
	end if
%>	<td width="190" height="180">
<div id="<%=rsFunc("SystemID")%>" style="position:absolute; width:155px; height:170px; z-index:1 ;">  
		<table id="<%="table"&rsFunc("SystemID")%>" width='100%' border='0' align="center" >
		<tr><td id="td1" align="center">
		<a onclick="OpenSystem('<%=rsFunc("URLLocation")%>','<%=rsFunc("SystemID")%>');" onMouseOver="DivColorChange('<%="table"&rsFunc("SystemID")%>');" onMouseOut="DivColorChange2('<%="table"&rsFunc("SystemID")%>');">
		<%
		if trim(rsFunc("ImageLocation"))="" or isnull(rsFunc("ImageLocation")) then
			picName="tmp.jpg"
		else
			picName=rsFunc("ImageLocation")
		end if
		%>
		<img src="image/<%=picName%>" alt="" width="128" height="128" border="0" align="baseline">
		  <br><font size="4"><%
		  strCode="select * from Code where ID="&trim(rsFunc("SystemID"))
		  set rsCode=conn.execute(strCode)
		  if not rsCode.eof then
			response.write trim(rsCode("Content"))
		  end if
		  rsCode.close
		  set rsCode=nothing
		  %></font></a>
		</td></tr>
		</table>
</div>
	</td>
<%	
	if s=5 then
		response.write "</tr>"
		s=1
	else
		s=s+1
	end if

	rsFunc.MoveNext
	Wend

rsFuncGroup.MoveNext
Wend
rsFuncGroup.close
set rsFuncGroup=nothing
%>
  <%
	rsFunc.close
	set rsFunc=nothing
	strSQL="select TO_Char(sysdate,'YYYY/MM/DD HH24:MI:SS') tmpDate from dual"
	set rstime=conn.execute(strSQL)
	SysDate=rstime("tmpDate")
	rstime.close
conn.close
set conn=nothing
%>
</table>
</body>
<script type="text/javascript" src="./js/date.js"></script>
<script language="JavaScript">
function getBillData(){
	BillNo.value=BillNo.value.toUpperCase();
	if (BillNo.value.length < 9 && BillNo.value!=""){
		alert("舉發單號小於九碼！");
	}else if (CarNo.value=="" && BillNo.value=="" && IllegalName.value=="" && IllegalID.value==""){
		alert("必須填入單號或車號或違規人！");
	}else{
		UrlStr="../traffic/Query/BillBaseData_Detail_Main.asp?BillNo="+BillNo.value+"&CarNo="+CarNo.value+"&IllegalName="+IllegalName.value+"&IllegalID="+IllegalID.value;
		newWin(UrlStr,"winMap",800,550,50,10,"yes","yes","yes","no");
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
}

function OpenSystem(PageUrl,Sn){
	SCheight=screen.availHeight;
	SCWidth=screen.availWidth;
	UrlStr=PageUrl;
	newWin(UrlStr,'',SCWidth,SCheight,0,0,"yes","no","yes","no");
}
function DivColorChange(DivNo){
	eval(DivNo).border="1";
}
function DivColorChange2(DivNo){
	eval(DivNo).border="0";
}
function funMember(){
	UrlStr="UserDataEdit.asp";
	newWin(UrlStr,"winMap",800,450,50,10,"yes","no","yes","no");
}
function change_Time(){
//	var time=new Date();
//	t_Hour=time.getHours();
//	t_Minute=time.getMinutes();
//	t_Second=time.getSeconds();
	//runServerScript("getServerTime.asp");

	//LayerTime.innerHTML="目前時間  "+t_Hour+"："+t_Minute;
	//setTimeout(change_Time,60000);
}
function funOpenWindow(){
	//跳出視窗的網址
	UrlStr="NOTICEMain.asp";
	newWin(UrlStr,"winMap111",600,450,50,10,"yes","no","yes","no");
<%if sys_City="台中市" then%>
	newWin("WeekDeleteCase.asp","winMap181",760,600,50,10,"yes","no","yes","no");
<%end if%>
}
//登入就跳視窗
funOpenWindow();

function menuIn() //隱藏
{
        if(n4) {
                clearTimeout(out_ID)
                if( menu.left > menuH*-1+30+10 ) {  
                        menu.left -= 14
                        in_ID = setTimeout("menuIn()", 1)
                }
                else if( menu.left > menuH*-1+30 ) {
                        menu.left--
                        in_ID = setTimeout("menuIn()", 1)
                }
        }
        else { 
                clearTimeout(out_ID)
                if( menu.pixelLeft > menuH*-1+30+10 ) {
                        menu.pixelLeft -= 14
                        in_ID = setTimeout("menuIn()", 1) 
                }
                else if( menu.pixelLeft > menuH*-1+30 ) {
                        menu.pixelLeft--
                        in_ID = setTimeout("menuIn()", 1)
                }
        }
}
function menuOut() //顯示
{
        if(n4) {
                clearTimeout(in_ID)
                if( menu.left < -10) { 
                        menu.left += 4
                        out_ID = setTimeout("menuOut()", 1)
                }
                else if( menu.left < 0) { 
                        menu.left++
                        out_ID = setTimeout("menuOut()", 1)
                }
                
        }
        else { 
                clearTimeout(in_ID)
                if( menu.pixelLeft < -10) {
                        menu.pixelLeft += 2
                        out_ID = setTimeout("menuOut()", 1)
                }
                else if( menu.pixelLeft < 0 ) {
                        menu.pixelLeft++
                        out_ID = setTimeout("menuOut()", 1)
                }
        }
}
function fireOver() { 
        clearTimeout(F_out)	       
        F_over = setTimeout("menuOut()", 10) 
}
function fireOut() { 
        clearTimeout(F_over)
         F_out = setTimeout("menuIn()", 10)
}
function init() {
        if(n4) {
                menu = document.D1
                menuH = menu.document.width
                menu.left = menu.document.width*-1+30 
                menu.onmouseover = menuOut
                menu.onmouseout = menuIn
				menu.visibility = "visible"
        }
        else if(e4) {
                menu = D1.style
                menuH = D1.offsetWidth
                //D1.style.pixelLeft = D1.offsetWidth*-1+20
                D1.onmouseover = fireOver
                D1.onmouseout = fireOut
				D1.onclick = fireOut
                D1.style.visibility = "visible"
        }
		UpdateLayer();
}
function UpdateLayer(){
	//UpLoadLayer.innerHTML="";
	runServerScript("UpdateMainLayer.asp");
	setTimeout(UpdateLayer,1800000);
	//alert("1");
}
F_over=F_out=in_ID=out_ID=null
n4 = (document.layers)?1:0;
e4 = (document.all)?1:0;
var procesID='<%=Session("Credit_ID")%>'

function DownProcess(nowtime){
	<%if sys_City="花蓮縣" or sys_City="台東縣" or sys_City="雲林縣" then%>
		runServerScript("/traffic/BillReturn/SystemDownloadFile.asp?nowTime="+nowtime);
	<%end if%>
}

function NowDownProcess(){
	<%if sys_City="花蓮縣" or sys_City="台東縣" or sys_City="雲林縣" then%>
		var nowtime="<%=year(SysDate)&"/"&Month(SysDate)&"/"&Day(SysDate)&" "&hour(SysDate)&"："&minute(SysDate)&"："&second(SysDate)%>";
		runServerScript("/traffic/BillReturn/SystemDownloadFile.asp?nowTime="+nowtime);
		alert("已重新處理");
	<%end if%>
}

function SQLDownProcess(nowtime){
	<%if sys_City="台中縣" or sys_City="雲林縣" or sys_City="屏東縣"then%>
//		if(procesID=='A000000000'){
			runServerScript("/traffic/BillReturn/T-SQL.asp?nowTime="+nowtime);
//		}
	<%end if%>
}

function ExportDB(){
	//runServerScript("OracleReturn.asp");
	newWin("/traffic/BillReturn/OracleReturn.aspx","winMap181",760,600,50,10,"yes","no","yes","no");
}

function funcUpdate(){
	newWin("/traffic/Update.asp","winMapUpdate",760,600,50,10,"yes","no","yes","no");
}
</script>
<script language="vbscript"> 
	<%randomize%>
	Dim secondDiff:SecordNow=0:RndSecond=<%=fix(rnd*60)%>:RndMinute=<%=fix(rnd*16)%>
	Sub UpdateTime()
		SecordNow=SecordNow+1
		Sys_nowTime=DateAdd("s", SecordNow,secondDiff)
		LayerTime.innerText = TimeValue(Sys_nowTime)

		'If (minute(Sys_nowTime) mod 20 = 0) and second(Sys_nowTime)=0 Then DownProcess(FormatDateTime(Sys_nowTime))

		'If (hour(Sys_nowTime) mod 1 = 0) and (minute(Sys_nowTime) = 10+RndMinute) and second(Sys_nowTime)=RndSecond Then SQLDownProcess(FormatDateTime(Sys_nowTime,4))

		'If second(Sys_nowTime) mod 10 = 0 Then SQLDownProcess(FormatDateTime(Sys_nowTime,4))
	End Sub

	Sub SetTime(serverDateTime)  		
		secondDiff  = serverDateTime
		oInterval = setInterval("UpdateTime()", 1000)
	End Sub

	SetTime("<%=year(SysDate)&"/"&Month(SysDate)&"/"&Day(SysDate)&" "&hour(SysDate)&"："&minute(SysDate)&"："&second(SysDate)%>")
</script>

</html>
