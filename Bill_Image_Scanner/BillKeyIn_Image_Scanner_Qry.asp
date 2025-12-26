<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<title>掃描備查系統</title>
<style type="text/css">
<!--
.style2 {font-size: 12px}
.style3 {
font-size: 12px ;
color: #FF0000}
.style4 {
font-size: 12px ;
}
.btn2 {font-size: 13px}
.Text1{
font-weight:bold;font-size: 16px;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" Onload="SetFocusOnMe();">

<form name="myForm" method="post">  
<script type="text/javascript" src="../js/date.js"></script>
<% 

  Response.flush
  	'Ftp連結位置
	FtpLocation=""
	strftp="select Value from ApConfigure where ID=53"
	set rsftp=conn.execute(strftp)
	if not rsftp.eof then
		FtpLocation=trim(rsftp("Value")) & Session("User_ID")
	end if
	rsftp.close
	set rsftp=Nothing
	
	strftp="select Value from ApConfigure where ID=31"
	set rsftp=conn.execute(strftp)
	if not rsftp.eof then
		Sys_City=trim(rsftp("Value"))
	end if
	rsftp.close
'Sys_City="雲林縣"
	set rsftp=nothing
  %>
<script language="javascript">

	function NewWindow(Width, Height, URL, WinName){
	    var nWidth = Width;
	    var nHeight = Height;
	    var sURL = URL;
	    var nTop = centerPos(nHeight,0);
	    var nLeft = centerPos(nWidth,1);
	    var sWinSize = "left=" + nLeft.toString(10) + ",top=" + nTop.toString(10) + ",width=" + nWidth.toString(10) + ",height=" + nHeight.toString(10);
	    var sWinStatus = "menubar=0,toolbar=0,scrollbars=1,resizable=1,status=0";
	    var sWinName = WinName;
	    OldObj = window.open(sURL,sWinName,sWinSize + ",left=0,top=0," + sWinStatus);
	}

	//置中 function
	function centerPos(size, type) {
	    switch(type) {
	        case 0:   //Top position
	            return (parseInt(window.screen.height) - size) / 2;
	            break;
	        case 1:   //Left position
	            return (parseInt(window.screen.width) - size) / 2;
	            break;
	        default:
	            alert('centerPos() : Type value error!!');
	    }
	}
	//開啟檢視圖
	function OpenPic(FileName){
		NewWindow(1000, 700, 'ShowMap.asp?PicName=' + FileName.replace(/\+/g, '@2@'), 'MyPic');
    }

function KeyDown(event)
{
<%if Sys_City<>"高雄市" then%>
    if (event.keyCode == 13)
    {
        myForm.btnModify.click();

    }
<%else%>
    if (event.keyCode == 13)
    {
		runServerScript("GetData.asp?BillNo="+myForm.txtBillNo.value);

		myForm.indexScan.value=parseInt(myForm.indexScan.value)+1;

		if (myForm.indexScan.value>="2" && myForm.tmpBillno.value==myForm.txtBillNo.value)
		{
			myForm.btnModify.click();
			myForm.indexScan.value=0;
			myForm.txtBillNo.value="";

		}
		if (parseInt(myForm.indexScan.value)>2)
		{
			myForm.indexScan.value=2;
		}

		myForm.tmpBillno.value=myForm.txtBillNo.value;

		myForm.txtBillNo.select();
    }
<%end if%>
}
function getillStreet(){
	if (event.keyCode==13){
		var illAddrNum=myForm.IllegalAddressIDQry.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
	}
}

function getillStreet2(){
	if (event.keyCode==13){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet2.asp?illAddrID="+illAddrNum);
	}
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no,titlebar=no");
	win.focus();
	return win;
}                   

function ShowPIC(filename,SN)
{
  myForm.PicName.value=filename;
  myForm.SelSN.value=SN;
  myForm.kinds.value="";
  myForm.DB_Selt.value="SP";
  myForm.submit();
}
function Clear()
{
 document.myForm.txtBillNo.value="";
 document.myForm.Sys_ScanUnit.value="";
 document.myForm.Sys_ScanMemberID.value="";
 document.myForm.ScanDate.value="";
 document.myForm.ScanDate1.value="";
 document.myForm.RuleDate.value="";
 document.myForm.RuleDate1.value="";
 document.myForm.CkboxYN.checked=false;
 document.myForm.IllegalAddressQry.value="";
 document.myForm.IllegalAddressIDQry.value="";
}
function toBook()
{
 document.myForm.RuleDateS.value="";
 document.myForm.RuleDateS1.value="";
 document.myForm.IllegalAddress.value="";
 document.myForm.IllegalAddressID.value="";
}

  function SetFocusOnMe()
 {
     document.myForm.txtBillNo.focus();

      <%	
        dim fso
        dim directory 
			if Sys_City="雲林縣" or Sys_City="高雄市" then 
				directory="D:\\F\\Image\\Scan\\" 
'			elseif Sys_City="高雄縣" then 
'			    directory="D:\\Fbackup\\Image\\Scan\\" 
			else
			    'directory="C:\\Image\\Scan\\" 
			    directory="C:\customer\\localuser\\wwwroot\\traffic\Image\\"
			end if

	'strftp="select distinct recordmemberid from billattatchimage"
	'set rsftp=conn.execute(strftp)
	'if rsftp.eof then        
     	set fso=Server.CreateObject("Scripting.FileSystemObject")
	    if (fso.FolderExists(directory&Session("User_ID")))=false then
       		fso.CreateFolder(directory&Session("User_ID"))
     	end if
     	
	    if (fso.FolderExists(directory&Session("User_ID")&"\\YHandle"))=false then
       		fso.CreateFolder(directory&Session("User_ID") &"\\YHandle")
     	end if 

	    if (fso.FolderExists(directory&Session("User_ID")&"\\YHandle\\"&year(date)-1911))=false then
       		fso.CreateFolder(directory&Session("User_ID") &"\\YHandle\\"&year(date)-1911)
     	end if 

	    if (fso.FolderExists(directory&Session("User_ID")&"\\YHandle\\"&year(date)-1911&"\\"&right("0"&month(date),2)))=false then
       		fso.CreateFolder(directory&Session("User_ID") &"\\YHandle\\"&year(date)-1911&"\\"&right("0"&month(date),2))
     	end if 

	    if (fso.FolderExists(directory&Session("User_ID")&"\\YHandle\\"&year(date)-1911&"\\"&right("0"&month(date),2)&"\\"&right("0"&day(date),2)))=false then
       		fso.CreateFolder(directory&Session("User_ID") &"\\YHandle\\"&year(date)-1911&"\\"&right("0"&month(date),2)&"\\"&right("0"&day(date),2))
     	end if 
     	
        set fso=nothing
	'end if 
	'rsftp.close
	'set	rsftp=nothing		

	  %>
 }
  </script>


<script language="JavaScript">
function RunScanner()
{
  myForm.DB_Selt.value="DB_Insert";
  myForm.submit();
}

function UploadData()
{
  window.open("<%=FtpLocation%>","FtpWin","location=0,width=770,height=455,resizable=yes,scrollbars=yes,toolbar=yes");

}
  </script>

<script language="javascript">
function Qry()
{
   myForm.DB_Selt.value="Selt";
   myForm.submit();
}
  </script>
<script language="javascript">
function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
	    myForm.DB_Selt.value="Selt";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
	    myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
	    myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10)==0){
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10)-1)*(10);
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10))*(10);
		}
	    myForm.DB_Selt.value="Selt";
		myForm.submit();
	}
}
  </script>


<!--<table border='1' align="left" cellpadding="0" height="554" width="993">-->
<table border='1' align="left" cellpadding="0" height="100%" width="100%">
		<tr>
		<td  bgcolor="#1BF5FF" width="985" height="24" colspan="2" align="center">
		<p align="center"><font size="4">掃描備查系統</font>&nbsp;&nbsp;&nbsp;<a href="掃描第一次使用時注意事項.doc">掃描第一次使用時注意事項</a>&nbsp;&nbsp;&nbsp;<a href="掃描備查系統.doc">掃描備查系統使用手冊</a>
		<% If Sys_City="高雄縣" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="條碼自動辨識注意事項.doc">條碼自動辨識注意事項</a>
		<%End if%>
		</td>
	</tr>

	    <tr>
		<td height="20" valign="top" width="35%" >
		<table border="1" width="100%">
					<td bgcolor="#EBFBE3" width="100%" align="center">資料新增</td>
					</table>
					
					    <table border="0" width="100%" id="table2">
							<tr>
								<td>類別</td>
								<td colspan="2"><font color="#FF0000">&nbsp;</font>
                         <input type="radio" name="Type" value="Pic"
						<%if request("Type")="Pic" then response.write "Checked" else response.write "" %>>
						<font color="#FF0000">違規相片</font>

						<input type="radio" name="Type" value="Book" onclick="toBook();" 
						<%if request("Type")="Book" then response.write "Checked" else response.write "" 
						  If request("Type")="" Then response.write "Checked"
						%>>
                        <font color="#FF0000">送達證書</font>

						<input type="radio" name="Type" value="BookReturn"  
						<%if request("Type")="BookReturn" then response.write "Checked" else response.write "" %>>
                        <font color="#FF0000">回執聯</font>
						<input type="radio" name="Type" value="BookBillno"  
						<%if request("Type")="BookBillno" then response.write "Checked" else response.write "" %>>
                        <font color="#FF0000">移送聯</font>

						<input type="radio" name="Type" value="BookOther"  
						<%if request("Type")="BookOther" then response.write "Checked" else response.write "" %>>
                        <font color="#FF0000">相關文件</font>

						</td>
							</tr>
							<tr>
								<td>違規日期</td>
								<td colspan="2">&nbsp;<input name="RuleDateS" type="text" class="btn1" value="<%
								response.write trim(request("RuleDateS"))
							%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr4" value="..." onclick="OpenWindow('RuleDateS');">
						~
						<input name="RuleDateS1" type="text" class="btn1" value="<%
								response.write trim(request("RuleDateS1"))
							%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr5" value="..." onclick="OpenWindow('RuleDateS1');"></td>
							</tr>
							<tr>
								<td>違規代碼路段</td>
								<td colspan="2"> <input type="text" size="7" value=""
								<% if trim(request("bIllegalAddressID"))="" then 
								response.write trim(request("IllegalAddressID"))
								else 
								response.write trim(request("bIllegalAddressID"))
								end if%>" 
								name="IllegalAddressID" onKeyUp="getillStreet2();" 
								style=ime-mode:disabled>

								<input type="text" size="8" value="<% 
								if request("bIllegalAddress")="" then 
								  response.write request("IllegalAddress") 
								else
								response.write request("bIllegalAddress") 
								end if
								%>" name="IllegalAddress" style=ime-mode:active ">

					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("../BillKeyIn/Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'> <img src="../Image/BillkeyInButton.jpg" width="25" height="17" onclick='window.open("../BillKeyIn/Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
									</td>
							</tr>
							<tr>
						<!--		<td>掃描方式</td>
								<td align="center" colspan="2">
								<p align="left">
								<input type="radio" name="Print" value="Mutli" <%
                        if request("Print")="Mutli" then 
                          response.write "Checked" 
                        else 
                          response.write "" 
                        end if 
                          if request("Print")="" then 
                            response.write "Checked"
                          end if  
                        %> checked>
					<font color="#000080">多張</font> 
					<input type="radio" name="Print" value="Single" <%if request("Print")="Single" then response.write "Checked" else response.write "" %>  >
					<font color="#000080">單張</font> </td>
							</tr>
							<tr>
								<td align="center">

                 <!--   <input type="button" value="開始掃描" name="btnRunScanner" onclick="RunScanner();"></td>
								<td align="center">-->
					<td></td><td><input type="button" value="上傳" name="UpLoadFile" onclick="UploadData();"></td>
								<td width="96" align="center">
					<input type="button" value="上傳完成開始註記" name="btnStartHandle" onclick="StartHandle();">
						</td>
							</tr>
<tr><td><td colspan="2"><font size="2">掃瞄檔案勿直接放入YHandle目錄內
<%If sys_city="高雄市" Then %>
<br>
<input type="checkbox" value="1" name="cbxAcceptMark"
	<%
	If request("cbxAcceptMark")="1" Then 
	response.write " checked"
	tempcolor="FF9291"	
	Else
	response.write " "
	tempcolor="#FFFFCC"	
	End if
	%>
 onclick="Qry();"><span style='font-size: 16px ;'><B>執行民眾收受註記</B></span>
<%End if%>
</font></td></td></tr>
		</table>
					
					<td width="100%" valign="top" height="20">
					<table border="1" width="100%">
					<td bgcolor="#EBFBE3" align="center">資料查詢</td>
					</table>
					<table border="0" width="100%" id="table1">
						<tr>
							<td width="108">掃描單位</td>
							<td width="140"><%=SelectUnitOption("Sys_ScanUnit","Sys_ScanMemberID")%>
							</td>
							<td width="88">掃描人員</td>
							<td><%=SelectMemberOption("Sys_ScanUnit","Sys_ScanMemberID")%>
							</td>
						</tr>
						<tr>
							<td width="108">掃描日期</td>
							<td width="234" colspan="2">
						<input name="ScanDate" type="text" class="btn1" value="<%
						       If Trim(request("DB_Selt"))<>"" Then 
								response.write trim(request("ScanDate"))
                               Else
								 If trim(request("ScanDate"))<>"" Then 
	                                response.write gInitDT(date)
								 End if
                               End if
							%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')"><input type="button" name="datestr1" value="..." onclick="OpenWindow('ScanDate');">~
							<input name="ScanDate1" type="text" class="btn1" value="<%
						       If Trim(request("DB_Selt"))<>"" Then 
								response.write trim(request("ScanDate1"))
                               Else
								 If trim(request("ScanDate1"))<>"" Then 
	                                response.write gInitDT(date)
								 End if
                               End if
							%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')"><input type="button" name="datestr" value="..." onclick="OpenWindow('ScanDate1');"></td>
							<td>
						<input type="checkbox" name="CkboxYN" value="on" <%
						if request("CkboxYN")<>"" then
						   response.write "Checked"
						end if
                        %>
						>已確認(已輸入舉發單號)&nbsp; </td>
						</tr>
						<tr>
							<td width="108">違規日期</td>
							<td width="234" colspan="2">
						<input name="RuleDate" type="text" class="btn1" value="<%
								response.write trim(request("RuleDate"))
							%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')"><input type="button" name="datestr2" value="..." onclick="OpenWindow('RuleDate');">~<input name="RuleDate1" type="text" class="btn1" value="<%
								response.write trim(request("RuleDate1"))
							%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')"><input type="button" name="datestr3" value="..." onclick="OpenWindow('RuleDate1');"></td>
							<td>

							<input type="radio" name="TypeQry" value="PicQry"
							<%if request("TypeQry")="PicQry" then response.write "Checked" else response.write "" %>>
							<font color="#FF0000">違規相片</font>　

							<input type="radio" name="TypeQry" value="BookQry" <%
                        if request("TypeQry")="BookQry" then response.write "Checked" else response.write "" 
                        if request("TypeQry")="" then response.write "Checked"
                        %> >
					<font color="#FF0000">送達證書</font>
					(<input type="radio" value="A4" name="PicTypeSize" 
					<%if request("PicTypeSize")="A4" then response.write "Checked" else response.write "" %> checked>A4
					<input type="radio" value="B5" name="PicTypeSize" <%if request("PicTypeSize")="B5" then response.write "Checked" else response.write "" %>>B5)

							<input type="radio" name="TypeQry" value="BookReturnQry"
							<%if request("TypeQry")="BookReturnQry" then response.write "Checked" else response.write "" %>>
					<font color="#FF0000">回執聯</font>

							<input type="radio" name="TypeQry" value="BookBillnoReturnQry"
							<%if request("TypeQry")="BookBillnoReturnQry" then response.write "Checked" else response.write "" %>>
					<font color="#FF0000">移送聯</font>

					<input type="radio" name="TypeQry" value="BookOtherQry"
							<%if request("TypeQry")="BookOtherQry" then response.write "Checked" else response.write "" %>>
					<font color="#FF0000">相關文件</font>
					</td>
						</tr>
						<tr>
							<td width="108">違規代碼路段</td>
							<td width="234" colspan="2">
					<input type="text" size="7" name="IllegalAddressIDQry" 	maxlength="10" value="<%=trim(request("IllegalAddressIDQry"))%>" onKeyUp="getillStreet();" >

							<input type="text" size="8"  maxlength="20" name="IllegalAddressQry" value="<%=trim(request("IllegalAddressQry"))
							%>" >

					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage4","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="22" onclick='window.open("Query_Street.asp","WebPage4","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'></td>
					<td>
						舉發單號
						<input type="text" size="15" name="BillnoQry" maxlength="10" value="<%=trim(request("BillnoQry"))%>" onkeyup="javascript:this.value=this.value.toUpperCase();" ></td>
					<tr>
					<td></td>
							<td colspan="5"> 
							<input type="button" value="查詢" name="btnQry" onclick="Qry();">　
							<input type="button" value="清除" name="btnClear" onClick="Clear();"> <font size="2" color="red">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type="button" value="整批刪除" name="btnDelete" onclick="DeleteDatabatch();">(注意:請先「查詢」確認資料內容之後，再按「整批刪除」按鈕)
							<br>請選擇 送達證書格式 ，才能正確的顯示
							</font></td>
					
						</tr>
						
						
						
					</table>
		</td>
	</tr>

		<td width="369" height="430" valign="top">
		<table width='305' border='1' align="left" cellpadding="3" height="31">
			<tr bgcolor="#FFFFCC">
			<td align="center" width="141" height="1"><span class="style2">單位</span></td>
			<td align="center" width="80" height="1"><span class="style2">人員</span></td>
			<td align="center" width="78" height="1"><span class="style2">舉發單號</span></td>
			<td align="center" width="60" height="1"><span class="style2">顯示圖片</span></td>
			<%If Sys_City="高雄市" Then %>
			<td align="center" width="60" height="1"><span class="style2">覆蓋圖片</span></td>
			<%End if%>
			</tr>
	<%	
	Response.flush
'	if request("DB_Selt")="Selt" then
dim bIllegalAddressID,bIllegalAddress,strtmp,strUnitID,strSQL,rs1,ScanDate1,ScanDate2
dim RuleDate1,RuleDate2,strwhere,rs2,DBsum,objFSO,DBcnt,tFileName,SelSnNum,i,tSN,PicName
            strwhere=""

		        ScanDate1=gOutDT(request("ScanDate"))&" 0:0:0"
				ScanDate2=gOutDT(request("ScanDate1"))&" 23:59:59"
				
		        RuleDate1=gOutDT(request("RuleDate"))&" 0:0:0"
				RuleDate2=gOutDT(request("RuleDate1"))&" 23:59:59"

			If request("DB_Selt")="" Then 	
			  ScanDate1=(date)&" 0:0:0"
			  ScanDate2=(date)&" 23:59:59"
    			strwhere=strwhere&" and a.RecordDate between TO_DATE('"&ScanDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ScanDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			End if	

			if request("ScanDate")<>"" and request("ScanDate1")<>"" then 
    			strwhere=strwhere&" and a.RecordDate between TO_DATE('"&ScanDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ScanDate2&"','YYYY/MM/DD/HH24/MI/SS')"
            end if
            
			if request("RuleDate")<>"" and request("RuleDate1")<>"" then 
    			strwhere=strwhere&" and  TO_DATE('"&RuleDate1&"','YYYY/MM/DD/HH24/MI/SS')  between RuleDateS and RuleDateE"
    			strwhere=strwhere&" and  TO_DATE('"&RuleDate2&"','YYYY/MM/DD/HH24/MI/SS')  between RuleDateS and RuleDateE"
            end if       
			
				if request("BillnoQry")<>"" And request("CkboxYN")<>"" Then 
	    			strwhere=strwhere&" and a.Billno='"&request("BillnoQry")&"'"
   		    	end if

            
				if request("Sys_ScanUnit")<>"" then
	    			strwhere=strwhere&" and c.UnitID ='"&request("Sys_ScanUnit")&"'"
   		    	end if
   		    	
   		    	if request("IllegalAddressQry")<>"" then
    		    	  strwhere=strwhere&" and a.IllegalAddress='"&request("IllegalAddressQry")&"'"
   		    	end if
   		    	
   		    	if request("IllegalAddressIDQry")<>"" then
    		    	  strwhere=strwhere&" and a.IllegalAddressID='"&request("IllegalAddressIDQry")&"'"
   		    	end if
   		    	
				if request("Sys_ScanMemberID")<>"" then
	    			strwhere=strwhere&" and a.RecordMemberID='"&request("Sys_ScanMemberID")&"'"
     			end if
     			
    			if request("CkboxYN")="" then
	    			strwhere=strwhere&" and (a.BillNo is null or a.BillNo='')"
     			else
	    			strwhere=strwhere&" and (a.BillNo is not null or a.BillNo<>'')"
     			end if
     			
     			if request("TypeQry")="PicQry" then
	    			strwhere=strwhere&" and a.TypeID=1"
     			ElseIf request("TypeQry")="BookQry" then
	    			strwhere=strwhere&" and a.TypeID=0"
     			ElseIf request("TypeQry")="BookReturnQry" Then
	    			strwhere=strwhere&" and a.TypeID=2"				
			ElseIf request("TypeQry")="BookBillnoReturnQry" Then
	    			strwhere=strwhere&" and a.TypeID=3"				
			ElseIf request("TypeQry")="BookOtherQry" Then
	    			strwhere=strwhere&" and a.TypeID=4"			
     			end if
     		'strwhere=strwhere&" and b.MemberID=" & Session("User_ID")
			strSQL="select count(*) as Total from BillAttatchImage a,MemberData b,UnitInfo c where a.RecordMemberID=b.MemberID and b.unitid=c.unitid and a.RecordStateID=0 " 
			set rs2=conn.execute(strSQL & strwhere)
			' response.write strSQL & strwhere
			DBsum=cdbl(rs2("Total"))
            rs2.close
            
			strSQL="select a.SN,a.BillNo,a.FileName,To_Char(a.RecordDate,'yyyyMMDDhh24miss') as IMGDate,b.ChName,c.unitname from BillAttatchImage a,MemberData b,UnitInfo c where a.RecordMemberID=b.MemberID and b.unitid=c.unitid and a.RecordStateID=0 " 
			'response.write strSQL & strwhere & " order by desc 1"
			set rs2=conn.execute(strSQL & strwhere & " order by 1 ")
			
			Set objFSO = CreateObject("Scripting.FileSystemObject")


					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					
  					if Not rs2.eof then 
  					  if DBcnt<0 then 
  					    rs2.move 0
  					  else
  					    rs2.move DBcnt
  					  end if  
  					end if  
	

                    tFileName=""
                    SelSnNum=""
                    
					for i=DBcnt+1 to DBcnt+10
					  
						if rs2.eof then exit for
     				  if i=DBcnt+1 then 
                        
                        if request("DB_Selt")<>"SP" then 
                            SelSnNum=tSN 
                        else
                            SelSnNum=rs2("SN") 
                            tSN=""
                        end if  

						if i=1 and request("DB_Selt")<>"SP" then
	                      tFileName=rs2("FileName")
                          if tSN="" then tSN=rs2("SN")
						end if
					  end if

						
						'directory=trim("\\"&Request.ServerVariables("SERVER_NAME")&Replace(Replace(rs2("FileName") & "","/","\"),"img","Image"))
						
						'If objFSO.FileExists(directory)=true Then
						'else
						'    strDelete="delete BillAttatchImage where SN=" & rs2("SN")
                        '   	conn.execute strDelete
                        'End If
		               Response.flush 
		%>	
			<tr onMouseOver="this.style.backgroundColor='#CCFFFF'" onMouseOut="this.style.backgroundColor='#FFFFFF'" >
				<td height="21" width="141"><span class="style2">
		        <%
					response.write rs2("unitname")
				%>
				  </span>　</td>
				<td height="21" width="80"><span class="style2">
		        <%
   					response.write rs2("ChName") 
				%>
				  </span>　</td>
				<td height="21" width="84"><span class="style2">
		        <%
					response.write rs2("BillNo") & ""
				%>
				  </span>　</td>				  
				<td height="21" width="60" align="center"><span class="style2">
					<input type="Button" Name="btn" onclick="ShowPIC('<%=rs2("FileName")%>','<%=rs2("SN")%>');" value="顯示">
    			  </span>　</td>
	 			<%If Sys_City="高雄市" Then %>
				<td>
					<input type="Button" Name="btn2" onclick='window.open("UploadScanReFile.asp?SN=<%=rs2("SN")%>","ploadScanReFile","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")' value="覆蓋">
					</td>
				<%End if%>
				  
			</tr>
		<%  
					rs2.movenext
					next
        
    	%>
		  
			<input type="button" name="MoveFirst" value="第一頁" onclick="funDbMove(0);">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(cdbl(DBcnt)/(10)+1)&"/"&fix(cdbl(DBsum)/(10)+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="MoveDown" value="最末頁" onclick="funDbMove(999);">
        <%
		rs2.close
            set objFSO=nothing
			set rs2=nothing
        'end if%>		
<tr>

		</table>

									
		</td>

		<td height="100%" width="100%">
		<!-- 影像大圖 -->
<% if tFileName<>"" or Request("PicName")<>"" then%>
<div style="overflow:hidden">
<img border="0" src="<%if trim(request("PicName"))="" then response.write tFileName else response.write Request("PicName")%>"  id="img" width="300" height="300" alt="點選看原圖" Onclick="OpenPic('<%if trim(request("PicName"))="" then response.write tFileName else response.write Request("PicName")%>')">
</div>

<script language="JavaScript">
//-------設定放大的位置--------------------------------------------------------------------------------------------------------------------------------------------
<%if Sys_City="高雄市" then %>
	var rate=2;
<%else%>
	var rate=5;
<%end if%>
	document.getElementById("img").parentNode.style.width=900;
	document.getElementById("img").parentNode.style.height=500;
if ('<%=request("TypeQry")%>'=="BookQry") 
{
	document.getElementById("img").height = 600*rate;
	document.getElementById("img").width = 600*rate;
<%if Sys_City="高雄市" and request("PicTypeSize")="A4" then %>
	document.getElementById("img").style.marginLeft=-420;
	document.getElementById("img").style.marginTop=0;
<%elseif request("PicTypeSize")="A4" and Sys_City="台中市" then%>
	var rate=2;
	document.getElementById("img").style.marginLeft=-1600;
	document.getElementById("img").style.marginTop=80;
<%elseif request("PicTypeSize")="A4" then%>
	document.getElementById("img").style.marginLeft=-2000;
	document.getElementById("img").style.marginTop=80;
<%else%>
	<%if Sys_City="台中市" then %>
var rate=2;
document.getElementById("img").style.marginLeft=-1700;
document.getElementById("img").style.marginTop=80;
	<%elseif Sys_City="高雄縣" then %>
	document.getElementById("img").style.marginLeft=-1250;
	document.getElementById("img").style.marginTop=-40;
	<%else%>
	document.getElementById("img").style.marginLeft=-1700;
	document.getElementById("img").style.marginTop=80;

	<%end if%>
<%end if%>
}
else
{
<%if Sys_City="台中市" and request("TypeQry")="BookReturnQry" then %>
var rate=5;
	document.getElementById("img").height = 600*rate;
	document.getElementById("img").width = 600*rate;
document.getElementById("img").style.marginLeft=-1200;
document.getElementById("img").style.marginTop=50;
<%else%>
	document.getElementById("img").height = 600;
	document.getElementById("img").width = 600;
<%end if%>



}
//----------------------------------------------------------------------------------------------------------------------------------------------------------------
        </script>
<% end if %>

</td>
	</tr>

	<tr>
		<td height="33" colspan="2" valign="top" width="100%">
     	<div align="left">
     	<table border='0' cellpadding="0" bordercolor="#1BF5FF" cellspacing="0" height="33" style="border-collapse: collapse" width="100%">
			<tr>
				<td bgcolor="<%=tempcolor%>" height="33" width="145" align="center"> 



                <p style="margin-top: 0"> <span class="style3"><font size="5">＊</font></span><font size="5">舉發單號&nbsp; </font></td>
				<td bgcolor="#1BF5FF" width="825">
				&nbsp;<font size="5"><input type="text" size="17" name="txtBillNo" maxlength=9 onkeyup="getkeyupBillno();" onkeydown="KeyDown(event);" onfocus="this.select()"  style="background-color:<%=tempcolor%>"></font>
				<input type="button" value="確定" name="btnModify" onclick="UpdateData('<%if tSN="" then response.write request("SelSN") else response.write tSN%>',document.myForm.txtBillNo.value,'<%=request("TypeQry")%>');">&nbsp;&nbsp;&nbsp;
        <input type="button" value="刪除" name="btnDelete" onclick="DeleteData('<%if tSN="" then response.write request("SelSN") else response.write tSN%>')">
        
        &nbsp;&nbsp;&nbsp;<font size="2">註記前請先點選第一筆紀錄檢查與手中送達證書是否相同，確認上傳順序的正確性。</font>
        </td>
				<input type="Hidden" name="PicName" value="">
				<input type="Hidden" name="kinds" value="">
                <input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
                <input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
                <input type="Hidden" name="SelSN" value="<%=SelSnNum%>">
                <input type="Hidden" name="DB_Selt" value="">
                <input type="Hidden" name="indexScan" value="0">
                <input type="Hidden" name="tmpBillno" value="">
                
							
		</table>
        </div>

<div id="lblData" style="position:absolute ; width:100%; height:140px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>            
</table>


</form>

  

<br>
<!--<OBJECT classid=clsid:5220cb21-c88d-11cf-b347-00aa00a28331><PARAM NAME="LPKPath" VALUE="twaincontrolxtrial.lpk"></object>
<OBJECT id=tcx1 classid=CLSID:C46A8919-8107-11D8-8671-00C1261173F0 codebase=twaincontrolxtrial.cab>
  ACTIVEX 掃描元件讀取失敗，請下載<A href="twaincontrolxtrial.exe">掃描元件</a></object>-->
</body>
     

<%
Response.flush

if request("DB_Selt")="StartHandle" then
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

iSn=0
        dim fp

			if Sys_City="雲林縣" or Sys_City="高雄市" then 
				fp="D:\\F\\Image\\Scan\\" & Session("User_ID") 
'			elseif Sys_City="高雄縣" then 
'			    fp="D:\\Fbackup\\Image\\Scan\\"  & Session("User_ID") 
			else
			    fp="F:\\Image\\Scan\\" & Session("User_ID") 
			end if
        

	  ' fp="D:\\Image\\Scan\\" & Session("User_ID") 

        set fso=Server.CreateObject("Scripting.FileSystemObject")

        set fod=fso.GetFolder(fp)
        set fic=fod.Files
   

        For Each fil In fic
		if fso.GetExtensionName(fil.Name) ="jpg" then
		iSn=iSn+1
'fso.MoveFile "F:\\Image\\Scan\\1\\200705041130522.jpg", "F:\\Image\\Scan\\1\\YHandle\\200705041130522.jpg"

       
                               ' sSQL = "select TO_char(sysdate,'YYYYMMDDHH24MISS') as FileNameSN from Dual"
							    sSQL = "select TO_char(sysdate,'HH24MISS') as FileNameSN from Dual"
                              	set oRST = Conn.execute(sSQL)
                              	
                              	if not oRST.EOF then
                        	       	FileNameSN = oRST("FileNameSN")
                              	end if
                              	oRST.close

 'SN抓最大值
                              	sSQL = "select BillAttatchImage_seq.nextval as SN from Dual"
                              	set oRST = Conn.execute(sSQL)
                              	
                              	if not oRST.EOF then
                        	       	sMaxSN = oRST("SN")
                              	end if
                              	oRST.close
                              	
                              	if request("Type")="Pic" then 
                              	  TypeID="1"
                              	elseif request("Type")="Book" then 
                              	  TypeID="0"
                              	elseif request("Type")="BookReturn" then 
                              	  TypeID="2"
                              	elseif request("Type")="BookBillnoReturn" then 
                              	  TypeID="3"
                              	elseif request("Type")="BookOther" then 
                              	  TypeID="4"
                              	end if
                              	

                                   '/img/scan/memberid/xxxx.jpg
					           	fDir=year(date)-1911&"/"&right("0"&month(date),2)&"/"&right("0"&day(date),2)&"/"
                                FileDirAndName="/img/scan/" & trim(session("User_ID")) & "/YHandle/" & fDir & FileNameSN & iSn & ".jpg"

                                if request("RuleDateS")="" and request("RuleDateE")="" and trim(request("IlgalAddr"))="" and trim(request("IlgalAddrID"))="" then 
                                  strInsert="insert into BillAttatchImage(SN,FileName,BillNo,TypeID,RecordMemberID,RecordDate,RecordStateID)" & _
                    				" values("&sMaxSN&",'"&FileDirAndName&"','','"& TypeID &"','"& trim(session("User_ID")) &"',SYSDATE,0)"
                       			else
                                	strInsert="insert into BillAttatchImage(SN,FileName,BillNo,TypeID,RecordMemberID,RecordDate,RecordStateID,RuleDateS,RuleDateE,IllegalAddress,IllegalAddressID)" & _
                    				" values("&sMaxSN&",'"&FileDirAndName&"','','"& TypeID &"','"& trim(session("User_ID")) &"',SYSDATE,0,to_date('"& gOutDT(trim(request("RuleDateS"))) &" 00:00:00','yyyy/MM/dd HH24:MI:SS') ,to_date('"& gOutDT(trim(request("RuleDateS1")))&" 23:59:59','yyyy/MM/dd HH24:MI:SS'),'"& trim(request("IllegalAddress")) &"','"& trim(request("IllegalAddressID")) &"')"

                       			end if	

                            	conn.execute strInsert



                            	'fso.DeleteFile("F:\\Image\\Scan\\" & Session("User_ID") & "\\YHandle\\" & trim(fil.Name))
           	mDir=year(date)-1911&"\\"&right("0"&month(date),2)&"\\"&right("0"&day(date),2)&"\\"
			if Sys_City="雲林縣" or Sys_City="高雄市" then 

				 fso.MoveFile "D:\\F\\Image\\Scan\\" & Session("User_ID") & "\\" & trim(fil.Name), "D:\\F\\Image\\Scan\\" & Session("User_ID") & "\\YHandle\\" & mDir & FileNameSN & iSn & ".jpg"
'			elseif Sys_City="高雄縣" then 
'                 fso.MoveFile "D:\\Fbackup\\Image\\Scan\\"  & Session("User_ID") & "\\" & trim(fil.Name), "D:\\Fbackup\\Image\\Scan\\"  & Session("User_ID") & "\\YHandle\\" & mDir & FileNameSN & iSn & ".jpg"
			else
                 fso.MoveFile "F:\\Image\\Scan\\" & Session("User_ID") & "\\" & trim(fil.Name), "F:\\Image\\Scan\\" & Session("User_ID") & "\\YHandle\\" & mDir & FileNameSN & iSn & ".jpg"
			end if


'fso.MoveFile "D:\\Image\\Scan\\" & Session("User_ID") & "\\" & trim(fil.Name), "D:\\Image\\Scan\\" & Session("User_ID") & "\\YHandle\\" & FileNameSN & iSn & ".jpg"

         end if 
        Next
		response.write "<script>"
		response.write "alert(""註記完成"")"
		response.write "</script>"
        set fso=nothing
        set fod=nothing
        set fic=nothing
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

end if

if request("DB_Selt")="DB_Insert" then
                       '填單人代碼
                       theRecordMemberID=trim(Session("User_ID"))
                       gCh_Name = session("CH_Name")
                       gUnit_ID = Session("Unit_ID")
                 if request("Type")="Pic" then
                   if (request("RuleDateS")="" or request("RuleDateS1")="") and (request("IllegalAddress")="" or request("IllegalAddressID")="") then 
                     %>
                       <script>
                         alert("請輸入違規日期或違規代碼路段");
                       </script>
                     <%
                     response.End
                   end if
                 end if
                       '新增告發單
	           %>
             	<script language="javascript">
                 var ImageNum;

                  if (tcx1.devicecount > 0)
                  {
                    tcx1.selectdevice();
                    // tcx1.CurrentDevice=1;
                   if (tcx1.connected)
                   {
                     
                     if (document.myForm.Print[0].checked) 
                     {
                       tcx1.MultiImage=true;
                       tcx1.UseADF=true;
                     }
                     else
                     {
                       tcx1.MultiImage=false;
                       tcx1.UseADF=false;
                     }
                     tcx1.KeepImages=true;
                     tcx1.UseInterface=false;
                     tcx1.ShowProgress=true;
                     tcx1.acquire();  
                          for (i = 1; i <=tcx1.ImageCount ; i++) 
                            {
                              ImageNum=getCurDate()+1;
                              //SaveData(ImageNum,'<%=theRecordMemberID%>','<%=gOutDT(request("RuleDateS"))%>','<%=gOutDT(request("RuleDateS1"))%>','<%if request("Type")="Pic" then response.write "1" else response.write "0" %>','<%=request("IllegalAddress")%>','<%=request("IllegalAddressID")%>');         
                              tcx1.SelectedImage = i;
                              tcx1.SaveToFile("C:\\" + ImageNum + ".jpg");
   
                            }
                            alert("共掃描" + tcx1.ImageCount + "張\n\n存檔完畢"); 
                            tcx1.Clear();

                         }   
                       else
                         alert("放棄處理影像");
                       }
                   else
                      alert("無掃描器");
                      
                  function getCurDate(){
                	var d = new Date();
                	var a = ["00","0",""];
                	var yyyy = d.getFullYear();
                	var mm = String(d.getMonth() + 1);
                	mm = a[mm.length] + mm;
                	var dd = String(d.getDay());
                	dd = a[dd.length] + dd;
                	var hh = String(d.getHours());
                	hh = a[hh.length] + hh;
                	var nn = String(d.getMinutes());
                	nn = a[nn.length] + nn;
                	var ss = String(d.getSeconds());
                	ss = a[ss.length] + ss;
                	var ms = String(d.getMilliseconds());
                	return yyyy + mm + dd + hh + nn + ss;
                   }
                   function SaveData(ImageNum,theRecordMemberID,RuleDateS,RuleDateE,TypeID,IlgalAddr,IlgalAddrID){
                     runServerScript("SaveData.asp?ImageNum="+ImageNum+"&theRecordMemberID="+theRecordMemberID+"&RuleDateS="+RuleDateS+"&RuleDateE="+RuleDateE+"&TypeID="+TypeID+"&IlgalAddr="+IlgalAddr+"&IlgalAddrID="+IlgalAddrID);
                   }	
                 </script>     
<%     	
     set oRST = nothing
end if 
%>
<script language="javascript">
		<%'response.write "UnitMan('Sys_ScanUnit','Sys_ScanMemberID','"&request("Sys_ScanMemberID")&"');"%>
		<%response.write "UnitMan('Sys_ScanUnit','Sys_ScanMemberID','"&request("Sys_ScanMemberID")&"');"%>
		


function DeleteData(SN)
{
  if (confirm("是否確定刪除")){
    runServerScript("DeleteData.asp?SN="+SN);
    Qry();
  }  
}


function StartHandle(SN)
{
  if (confirm("是否確定註記")){
    myForm.DB_Selt.value="StartHandle";
    myForm.submit();
  }  
}

function DeleteDatabatch()
{
  if (confirm("是否確定整批刪除")){
    runServerScript("DeleteDataDatabatch.asp?strwhere=<%=strwhere%>");
    Qry();
  }  
}

function UpdateData(SN,BillNo,Type)
{
  var err=0;
  if (myForm.txtBillNo.value=="") 
  {
    err=1;
  }
  else
  {
  <%if Sys_City<>"高雄市" then%>
    if (myForm.txtBillNo.value.length!=9)
    {
      err=1;
    }
  <%end if%>
  //
  }
  
  if (err==0) 
  {
	<%if Sys_City<>"高雄市" then%>
     runServerScript("UpdateData.asp?SN="+SN+"&BillNo="+BillNo+"&Type="+Type);
	<%else%>
	if (myForm.cbxAcceptMark.checked)	{
      runServerScript("UpdateData.asp?SN="+SN+"&BillNo="+BillNo+"&Type="+Type+"&AcceptMark=1");
	}
	else
	{
	   runServerScript("UpdateData.asp?SN="+SN+"&BillNo="+BillNo+"&Type="+Type+"&AcceptMark=0");
	}
	 <%end if%>

    // 
  }  
  else
  {
    alert("不得輸入空值或輸入值需９碼");
  }
    
}

function getkeyupBillno(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
		myForm.txtBillNo.value=myForm.txtBillNo.value.toUpperCase();
	}
}

</script>
<%

conn.close
set conn=nothing
%>
</html>