<% 
Response.Expires = -1
Server.ScriptTimeout = 60000

%>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="freeaspupload.asp" -->
<%
' ****************************************************
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"

  Dim uploadsDirVar
  uploadsDirVar = "d:\Inetpub\wwwroot\Traffic\CaseImport\Data" 
  
' ****************************************************
' Note: this file uploadTester.asp is just an example to demonstrate
' the capabilities of the freeASPUpload.asp class. There are no plans
' to add any new features to uploadTester.asp itself. Feel free to add
' your own code. If you are building a content management system, you
' may also want to consider this script: http://www.webfilebrowser.com/

function chkCarNoFormat(CarNo)

	strHeavy="ABCFGHIJKLMNOPY"	   '重機第一碼
	strSmall="DEQRSTUVWXZ"	'//輕機第一碼
	if InStr(CarNo,"-")>= 0	 then 
		CarNoArray=split(CarNo,"-")
		if len(CarNoArray(0))=2 and len(CarNoArray(1))=2 then 
			chkCarNoFormat=2
		elseif len(CarNoArray(0))=2 and len(CarNoArray(1))=4 or (len(CarNoArray(0))=4 and len(CarNoArray(1))=2) or (len(CarNoArray(0))=2 and  len(CarNoArray(1))=3) or (len(CarNoArray(0))=3 and len(CarNoArray(1))=2) then
			chkCarNoFormat= 1
		elseif (len(CarNoArray(0))=3 and len(CarNoArray(1))=3) then
				 if InStr(CarNoArray(0),strHeavy) = 0 then
					if InStr(CarNoArray(0),"0") = 0 then 
						chkCarNoFormat= 0
					else
						chkCarNoFormat= 3
					end if
				elseif InStr(CarNoArray(0),strSmall) = 0 then
					if InStr(CarNoArray(0),"0") = 0 then 
						chkCarNoFormat= 0
					else
						chkCarNoFormat=4
					end if
				else
					chkCarNoFormat= 3
				end if

			chkCarNoFormat= 0
		end if
	else
		chkCarNoFormat= 0
	end if

end function


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();

</script>
<head>
<script type="text/javascript" src="../js/form.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script type="text/javascript" src="../js/date.js"></script>
<title>催繳資料匯入系統</title>
<script language="javascript">
  function InsertData()
  {
   if (myForm.T1.value!="")
   {
    myForm.action="StopCarCaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_Insert";
    myForm.submit();
   }
   else
   {
    alert("請選擇檔案");
   } 
  }
  function BrowsFile()
  {
    myForm.action="StopCarCaseImport.asp?ImportFileName="+myForm.attach1.value;
    //myForm.T1.value=myForm.attach1.value;
	//myForm.action="StopCarCaseImport.asp";
    myForm.submit();
  }
  function Clear()
  {
    myForm.T1.value="";
    myForm.action="StopCarCaseImport.asp";    
    myForm.submit();
  }
</script>
<body>
<form name="myForm" method="Post" enctype="multipart/form-data">
<%
Dim MemID,memName,fs,FileName,txtf,TempData,UnitID,BillfillerID,strMem,rsMem,MemberName,strVersion,rsVersion,rsMatch
dim Version,strUnit,rsUnit,UnitName,strCheck,rsCheck,ImportDate,ImportMem,strInsertRecord,txtStream,i,Corr,txtline,tempValue

MemID=trim(Session("User_ID"))
memName=Session("Ch_Name")
			' 花蓮  填單日 / 與應到案日期 再入案的時候可以調整 
			
      '讀ini檔
			  set fs=Server.CreateObject("Scripting.FileSystemObject")
			  FileName=Server.MapPath("system.ini")
			  
			  if fs.FileExists(FileName) then
			  	set txtf=fs.OpenTextFile(FileName)
			  	
				while not txtf.atEndOfStream 
					TempData=txtf.readline 
					if InStr(TempData, "UnitID=")>0 then 
					  UnitID=Trim(Replace(TempData,"UnitID=",""))
					end if
					if InStr(TempData, "BillfillerID=")>0 then 
					  BillfillerID=Trim(Replace(TempData,"BillfillerID=",""))
					end if										
				wend
				
			  	set txtf=nothing
			  end if
			  set fs=nothing
			  response.write "<BR>"
		'------------------------------------------------------------------------------------
		'抓出停管業務員警名字
		strMem="select ChName from MemberData where MemberID ='" & BillfillerID & "'"
            	set rsMem=conn.execute(strMem)
             	if not rsMem.eof then
            		MemberName=trim(rsMem("ChName"))
            	end if
            	rsMem.close
            	set rsMem=Nothing

		strVersion="select value from apconfigure where id=3"
            	set rsVersion=conn.execute(strVersion)
             	if not rsVersion.eof then
            		Version=trim(rsVersion("value"))
            	end if
            	rsVersion.close
            	set rsVersion=Nothing

			  	'strCity="select value from apconfigure where id=31"
            	'set rsCity=conn.execute(strCity)
             	'if not rsCity.eof then
            	'	tCity=trim(rsCity("value"))
            	'end if
            	'rsCity.close
            	'set rsCity=Nothing
                tCity="花蓮縣"
            	
            	strUnit="select UnitName from Unitinfo where UnitID ='" & UnitID & "'"
            	set rsUnit=conn.execute(strUnit)
             	if not rsUnit.eof then
            		UnitName =trim(rsUnit("UnitName"))
            	end if
            	rsUnit.close
            	set rsUnit=nothing
	 
	   			 
		response.write "建檔人: " & memName & "&nbsp;&nbsp;&nbsp;"
             ' response.write "停管業務單位: " & UnitName & "&nbsp;&nbsp;&nbsp;"
              response.write "停管業務員警: " & MemberName & "&nbsp;&nbsp;&nbsp;"
			  %>
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>催繳資料匯入系統</td></tr>
    		</table>
    		  <p><p>
  <font size="4">選擇檔案</font><span><font size="4"> </font><input type="text" name="T1" size="53" value="<%=request("ImportFileName")%>" readonly></span>&nbsp;
  <input type="file"  name="attach1" size="1"  onchange="BrowsFile();" style="position: relative;-moz-opacity:0 ;-moz-opacity:0 ;filter:alpha(opacity: 0);opacity: 0;z-index: 2;" /> &nbsp;&nbsp;&nbsp;&nbsp; 
    <input type="button" value="匯入" name="btnInto" onclick="InsertData();">&nbsp;&nbsp;&nbsp;&nbsp; 
  <input type="button" value="清除" name="btnClear" onclick="Clear();">
  <div style="position: absolute;top: 105px;left: 480px;width: 15px;padding: 0;margin: 0;z-index: 1;line-height: 90%;">

		<img src="SelectFile.Jpg" onMouseOver="this.src='SelectFileOn.Jpg'" onMouseOut="this.src='SelectFile.Jpg'">
	</div>
</span></span></span>&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; 
  <br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  注意：點選 匯入 後,系統會開始匯入,並顯示匯入筆數<br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
  注意：每檔案容量限制上傳 600k<br>
  </p>
  <p>
  </p>
  <p>
<%

	if trim(request("DB_Selt"))="DB_Insert" then 
	        '查詢匯入記錄
          set fs=Server.CreateObject("Scripting.FileSystemObject")
					ImportFileName= trim(fs.GetFileName(request("ImportFileName"))) 
      		'檢查是否已匯過  
        	strCheck="select FileName,ImportDate,ImportMem from CaseImport where FileName ='" & ImportFileName & "'"
            	set rsCheck=conn.execute(strCheck)
            	FileName=""
            	ImportDate=""
            	ImportMem=""
             	if not rsCheck.eof then
            		FileName=trim(rsCheck("FileName"))
            		ImportDate=trim(rsCheck("ImportDate"))
                    	ImportMem =trim(rsCheck("ImportMem"))
            	end if
            	rsCheck.close
            	set rsCheck=Nothing
		Set fs=nothing
            	
            	'------------------------------------------------------------------------------------------------------------------
		'判斷是否繼續匯入檔案
                if FileName<>"" then 
                %>
                  <script language="javascript">
	    		    if (confirm("該檔案已匯入過\n\n匯入日期:<%=ImportDate%>\n匯入檔名:<%=FileName%>\n匯入人員:<%=ImportMem%>\n\n  是否繼續匯入"))
	    		    {
                      myForm.action="StopCarCaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
                      myForm.submit();
	    		    }
	    		  </script>
                <%
		else
				%>
                  <script language="javascript">
                      myForm.action="StopCarCaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
                      myForm.submit();
	    		  </script>
                <%
		end if   
	end if   
    	'---------------------------------
    	
    	'讀取每筆資料
    	 Sys_Now=now()
         if trim(request("DB_Selt"))="DB_StillInsert" then 
      	 	set fs=Server.CreateObject("Scripting.FileSystemObject")
                '新增檔案上傳記錄    			  
     		strInsertRecord="Insert into CaseImport (SN,FileName,ImportDate,ImportMem) values(CaseImport_Seq.nextval,'" & trim(fs.GetFileName(request("ImportFileName"))) & "',Sysdate,'" & memName & "')"
                conn.execute(strInsertRecord)
                ImportFileName=trim(fs.GetFileName(request("ImportFileName")))
	    		  
      		FileName=Server.MapPath("Data\" &fs.GetFileName(request("ImportFileName")))
		response.write "資料開始匯入...請稍待下方出現匯入結果<BR>"
		response.flush
		if  tCity="花蓮縣" then
      			  set txtStream = fs.opentextfile(FileName) 
	    		  i = 1
		    	  Corr=0
    			  Err=0

	    		  while not txtStream.atEndOfStream 
		    	    	txtline = txtStream.readline 
      			  	sCarNo=trim(mid(txtline,1,9))
      			  	sDate=trim(mid(txtline,10,7))
      			  	sTime=trim(mid(txtline,18,4))
      			  	sAddress=trim(mid(txtline,31,35))
      			  	sStopCarBillNo=trim(mid(txtline,66,15))
      			  	sCarSimpleID=trim(mid(txtline,29,2))
      			  	sMoney=trim(mid(txtline,23,2))
      			  	'response.write sCarNo & "," & sDate & "," & sTime & "," &  sCarSimpleID & "," & sMoney & "," & sAddress & "," & sStopCarBillNo  
      			  	'response.write "<BR>"
				'判斷是否檔案匯入有空白
      			    	if trim(sCarNo)="" or trim(sDate)="" or trim(sTime)="" or trim(sStopCarBillNo)="" or trim(sMoney)="" then  
      			        	response.write "第" & i  & "行: " & txtline & "<br>"
					response.flush
     			        	Err= Err+1      
     			        	i=i+1
     			     	else 
      			        	'response.write "第" & i & "行: " & txtline & "<br>"
					'新增每筆記錄---------------------------------------------------------------------------------------------------
                        		'違規日期

                         	  	theIllegalDate=funGetDate(gOutDT(sDate) &_
                         	  				" "&left(trim(sTime),2)& ":"&right(trim(sTime),2),1)
					BillFillDate=funGetDate(date(),0)
					BillFillDate2=funGetDate(date()+30,0)
                            		'---------------------------------------------------------------------------				
                            		Sys_Now=DateAdd("s",1,Sys_Now)
                         		strInsert="insert into BillBase(SN,BillTypeID,BillNo,UseTool,Insurance,CarNo,CarSimpleID,IllegalDate" & _
                         				",IllegalAddressID,IllegalAddress,Rule1,ForFeit1,ImagePathName" &_
                         				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
                         				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
                         				",BillFillerMemberID,BillFiller" &_
                         				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
                         				",Note,RuleVer,DriverSex,DOUBLECHECKSTATUS,BILLBASETYPEID,equipmentid)" &_
                         				" values(BillBase_seq.nextval,'2','',0,0" &_
                         				",'"&sCarNo&"',"&sCarSimpleID&_						          
                         				","&theIllegalDate&",null" &_
                         				",'"&sAddress&"','5620001'" &_
                         				","&sMoney&",'"&sStopCarBillNo&"'" &_ 
                         				",null,null,null" &_
                         				",null,null,null" &_
                         				",'"&UnitID&"','"&BillfillerID&"','"&MemberName&"'" &_
                         				",'"&BillfillerID&"','"&MemberName&"'" &_
                         				","&BillFillDate&","&BillFillDate2&",'0',0,"&funGetDate(Sys_Now,1)&",'"&MemID&"'" &_
                         				",'"&sStopCarBillNo&"`"&ImportFileName&"','"&Version&"'" &_
                         				",null" &_
                         				",0,'0','1')"
                         				'response.write strInsert
                         				'response.end
                         				conn.execute strInsert		    
																'smith 因為 停管催繳查車不會入案. 但是要做單退需要有 一般案件入案時寫到billmailhistory
																'      所以在這邊有新增的話就寫入一筆 到 billmailhistory
																'
																'strMail="Insert into StopBillMailHistory(BillSN,BillNo,CarNo,MailDate,MailNumber) "&_
																'				" values("&BillSN&",'"&BillNotemp&"','"&CarNo&"',sysdate,null)"		
																'				conn.execute strMail


    			        	Corr= Corr+1    
     			        	i=i+1  			        
                    		end if
    			  wend 
	    		  set fs=nothing
	              end if
'-------------------------------------------
		     	
	    		response.write "結束資料匯入............................"
	    		response.write "<P>正確筆數：" & Corr & "筆 <br>"
	    		response.write "<font color=""red"">錯誤筆數：" & Err & "筆 </font><br>"
	    		response.write "  總筆數：" & i-1 & "筆 <br>"

	    	End if 
	    		
			%>

  <input type="hidden" name="DB_Selt" value="">
  <input type="hidden" name="ImportFileName" value="<%=request("ImportFileName")%>">
  <p></p>
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>　
			</td></tr>
		</table>


		　</form>
</html>


							
<%		
'上傳檔案到伺服器
function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey

    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)
end function


    SaveFiles()
%> 