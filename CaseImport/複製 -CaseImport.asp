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
<title>舉發單資料匯入系統</title>
<script language="javascript">
  function InsertData()
  {
   if (myForm.T1.value!="")
   {
    myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_Insert";
    myForm.submit();
   }
   else
   {
    alert("請選擇檔案");
   } 
  }
  function BrowsFile()
  {
    myForm.action="CaseImport.asp?ImportFileName="+myForm.attach1.value;
    //myForm.T1.value=myForm.attach1.value;
	//myForm.action="CaseImport.asp";
    myForm.submit();
  }
  function Clear()
  {
    myForm.T1.value="";
    myForm.action="CaseImport.asp";    
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
			' 嘉義   預設填單日就是檔案匯入日期 / 應到案日期為填單日 + 30天
			' 所以兩個縣市預設 填單日就是檔案匯入日 , 應到案日就是填單日 + 30天
			
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

			  	strCity="select value from apconfigure where id=31"
            	set rsCity=conn.execute(strCity)
             	if not rsCity.eof then
            		tCity=trim(rsCity("value"))
            	end if
            	rsCity.close
            	set rsCity=Nothing
                'tCity="台中市"
            	
            	strUnit="select UnitName from Unitinfo where UnitID ='" & UnitID & "'"
            	set rsUnit=conn.execute(strUnit)
             	if not rsUnit.eof then
            		UnitName =trim(rsUnit("UnitName"))
            	end if
            	rsUnit.close
            	set rsUnit=nothing
	   '-------------------smith----台中市 暫時不開放------------------------------------------------
	   'if tCity="台中市" then 
	   '	response.write " 請確認拖吊場工程師修正完成後再行匯入" 
	   '	response.end
	   'end if
	   '
	   			 
			  response.write "建檔人: " & memName & "&nbsp;&nbsp;&nbsp;"
             ' response.write "停管業務單位: " & UnitName & "&nbsp;&nbsp;&nbsp;"
              response.write "停管業務員警: " & MemberName & "&nbsp;&nbsp;&nbsp;"
			  %>
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>舉發單資料匯入系統</td></tr>
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
  注意：每檔案容量限制上傳 200k<br>
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
                      myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
                      myForm.submit();
	    		    }
	    		  </script>
                <%
			    else
				%>
                  <script language="javascript">
                      myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
                      myForm.submit();
	    		  </script>
                <%
				  end if   
			    end if   
    			  '讀取每筆資料
    			  Sys_Now=now()
             if trim(request("DB_Selt"))="DB_StillInsert" then 
      		   	  set fs=Server.CreateObject("Scripting.FileSystemObject")
                  '新增檔案上傳記錄    			  
     			   strInsertRecord="Insert into CaseImport (SN,FileName,ImportDate,ImportMem) values(CaseImport_Seq.nextval,'" & trim(fs.GetFileName(request("ImportFileName"))) & "',Sysdate,'" & memName & "')"
                   conn.execute(strInsertRecord)
                  ImportFileName=trim(fs.GetFileName(request("ImportFileName")))
	    		  
      		   	  FileName=Server.MapPath("Data\" &fs.GetFileName(request("ImportFileName")))
				  response.write "開始資料匯入............................<BR>"
				  response.flush
		if tCity="嘉義市" then
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0

	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
      			  	tempValue=split(trim(txtline),",")
      			  	if UBound(tempValue)=5 then
      			  	  if trim(tempValue(0))="" or trim(tempValue(1))="" or trim(tempValue(2))="" or trim(tempValue(3))="" or trim(tempValue(4))="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
     			      else
      			        'response.write "第" & i & "行: " & txtline & "<br>"
						'新增每筆記錄---------------------------------------------------------------------------------------------------
                        	'違規日期
                           theIllegalDate=funGetDate(gOutDT(tempValue(2)) &" "&tempValue(3),1)
						   BillFillDate=funGetDate(date(),0)
						   	BillFillDate2=funGetDate(date()+30,0)
                            '---------------------------------------------------------------------------
                             '法條金額
                             strLaw="select Level1 from law where itemid='5620001'"
                            	set rsLaw=conn.execute(strLaw)
                             Level1=""
                              if not rsLaw.eof then
                               Level1=trim(rsLaw("Level1"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing
                     
							CarSimpleID=chkCarNoFormat(trim(tempValue(1)))
                          '---------------------------------------------------------------------------				
                            Sys_Now=DateAdd("s",1,Sys_Now)
                         	strInsert="insert into BillBase(SN,BillTypeID,BillNo,UseTool,Insurance,CarNo,CarSimpleID,IllegalDate" & _
                         				",IllegalAddressID,IllegalAddress,Rule1,ForFeit1" &_
                         				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
                         				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
                         				",BillFillerMemberID,BillFiller" &_
                         				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
                         				",Note,RuleVer,DriverSex,DOUBLECHECKSTATUS,BILLBASETYPEID,equipmentid)" &_
                         				" values(BillBase_seq.nextval,'2','',0,0" &_
                         				",'"&tempValue(1)&"',"&CarSimpleID&_						          
                         				","&theIllegalDate&",null" &_
                         				",'"&trim(tempValue(4))&"','5620001'" &_
                         				","&Level1&"" &_
                         				",null,null,null" &_
                         				",null,null,null" &_
                         				",'"&UnitID&"','"&BillfillerID&"','"&MemberName&"'" &_
                         				",'"&BillfillerID&"','"&MemberName&"'" &_
                         				","&BillFillDate&","&BillFillDate2&",'0',0,"&funGetDate(Sys_Now,1)&",'"&MemID&"'" &_
                         				",'"&tempValue(0)&"`"&ImportFileName&"','"&Version&"'" &_
                         				",null" &_
                         				",0,'0','1')"
                         				'response.write strInsert
                         				conn.execute strInsert		    

     			        '--------------------------------------------------------------------------------
    			        Corr= Corr+1    
     			        i=i+1  			        

     			      end if  
     			    else
      			        response.write "第" & i & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1     			        
     			        i=i+1
                    end if
    			  wend 
	    		  set fs=nothing
		elseif tCity="台中市" then
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0

	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
      			  	tempValue=split(trim(txtline),",")
      			  	if UBound(tempValue)=8 then
      			  	  if trim(tempValue(0))="" or trim(tempValue(1))="" or trim(tempValue(2))="" or trim(tempValue(3))="" or trim(tempValue(4))="" or trim(tempValue(5))="" or trim(tempValue(6))="" or trim(tempValue(7))="" or trim(tempValue(8))="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
     			      else
      			        'response.write "第" & i & "行: " & txtline & "<br>"
			'新增每筆記錄---------------------------------------------------------------------------------------------------
                        	'違規日期
                           theIllegalDate=funGetDate(gOutDT(tempValue(2)) &" "&left(trim(tempValue(3)),2)&":"&right(trim(tempValue(3)),2),1)
						   BillFillDate=funGetDate(date(),0)
     				    	BillFillDate2=funGetDate(gOutDT(tempValue(5)),0)
							
                            '---------------------------------------------------------------------------
                             '法條金額

                             strLaw="select Level1 from law where itemid='"&trim(tempValue(7))&"'"
                            	set rsLaw=conn.execute(strLaw)
                             Level1=""
                              if not rsLaw.eof then
                               Level1=trim(rsLaw("Level1"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing

                    
				CarSimpleID=tempValue(6)
				strMem="select ChName,UnitID,MemberID from MemberData where loginid ='" & trim(tempValue(8)) & "'"
                         	set rsMem=conn.execute(strMem)
                        	if not rsMem.eof then
                        		MemberName=trim(rsMem("ChName"))
                        		UnitID=trim(rsMem("UnitID"))
								MemberID=trim(rsMem("MemberID"))
                            else
                                MemberName=""
                        	UnitID=""
				MemberID=""
                            end if
                            rsMem.close
                        	set rsMem=Nothing
                          '---------------------------------------------------------------------------				
                            	Sys_Now=DateAdd("s",1,Sys_Now)
'smith 判斷單號是否已經匯入過 start
strSQL="select billno from BillBase where billno='" & tempValue(0) & "' and recordstateid<>-1 "
set rsMatch=conn.execute(strSQL)	
if rsMatch.eof and MemberID<>"" then 
'smith end
                         	strInsert="insert into BillBase(SN,BillTypeID,BillNo,UseTool,Insurance,CarNo,CarSimpleID,IllegalDate" & _
                         				",IllegalAddressID,IllegalAddress,Rule1,ForFeit1" &_
                         				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
                         				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
                         				",BillFillerMemberID,BillFiller" &_
                         				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
                         				",Note,RuleVer,DriverSex,DOUBLECHECKSTATUS,BILLBASETYPEID,equipmentid,CarAddID)" &_
                         				" values(BillBase_seq.nextval,'2','"&tempValue(0)&"',0,0" &_
                         				",'"&tempValue(1)&"',"&CarSimpleID&_						          
                         				","&theIllegalDate&",null" &_
                         				",'"&trim(tempValue(4))&"','"&trim(tempValue(7))&"'" &_
                         				","&Level1&"" &_
                         				",null,null,null" &_
                         				",null,null,null" &_
                         				",'"&UnitID&"','"&MemberID&"','"&MemberName&"'" &_
                         				",'"&MemberID&"','"&MemberName&"'" &_
                         				","&BillFillDate&","&BillFillDate2&",'0',0,"&funGetDate(Sys_Now,1)&",'"&MemID&"'" &_
                         				",'"&ImportFileName&"','"&Version&"'" &_
                         				",null" &_
                         				",0,'0','1','8')"
                         				'response.write strInsert
                         				conn.execute strInsert		    
				Corr= Corr+1    
     			        i=i+1			        
     			        '--------------------------------------------------------------------------------
'smith start
else
	
	 response.write "第" & i & "行: " & txtline & "  已經匯入過或無人員編號有誤<br>"
	 Err= Err+1	
	i=i+1	
end if
rsMatch.close
set rsMatch=Nothing
'smith end    			     
				     
     			      end if  
     			    else
      			        response.write "第" & i & "行: " & txtline & " 格式有誤 <br>"
						response.flush
     			        Err= Err+1     			        
     			        i=i+1
                    end if
    			  wend 
	    		  set fs=nothing
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				  elseif  tCity="花蓮縣" then
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0

	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
      			  	tempValue0=trim(mid(txtline,1,9))
      			  	tempValue1=trim(mid(txtline,10,7))
      			  	tempValue2=trim(mid(txtline,18,4))
      			  	tempValue3=trim(mid(txtline,31,30))
      			  	tempValue4=trim(mid(txtline,62,20))
      			  	tempValue5=trim(mid(txtline,29,2))

      			  	  if trim(tempValue0)="" or trim(tempValue1)="" or trim(tempValue2)="" or trim(tempValue3)="" or trim(tempValue4)="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
     			     else 
      			        'response.write "第" & i & "行: " & txtline & "<br>"
						'新增每筆記錄---------------------------------------------------------------------------------------------------
                        	'違規日期

                           theIllegalDate=funGetDate(gOutDT(tempValue1) &" "&left(trim(tempValue2),2)&":"&right(trim(tempValue2),2),1)
						   BillFillDate=funGetDate(date(),0)
						   	BillFillDate2=funGetDate(date()+30,0)
                            '---------------------------------------------------------------------------
                             '法條金額
                             strLaw="select Level1 from law where itemid='5620001'"
                            	set rsLaw=conn.execute(strLaw)
                             Level1=""
                              if not rsLaw.eof then
                               Level1=trim(rsLaw("Level1"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing

							CarSimpleID=tempValue5
                          '---------------------------------------------------------------------------				
                            Sys_Now=DateAdd("s",1,Sys_Now)
                         	strInsert="insert into BillBase(SN,BillTypeID,BillNo,UseTool,Insurance,CarNo,CarSimpleID,IllegalDate" & _
                         				",IllegalAddressID,IllegalAddress,Rule1,ForFeit1" &_
                         				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
                         				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
                         				",BillFillerMemberID,BillFiller" &_
                         				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
                         				",Note,RuleVer,DriverSex,DOUBLECHECKSTATUS,BILLBASETYPEID,equipmentid)" &_
                         				" values(BillBase_seq.nextval,'2','',0,0" &_
                         				",'"&tempValue0&"',"&CarSimpleID&_						          
                         				","&theIllegalDate&",null" &_
                         				",'"&tempValue3&"','5620001'" &_
                         				","&Level1&"" &_
                         				",null,null,null" &_
                         				",null,null,null" &_
                         				",'"&UnitID&"','"&BillfillerID&"','"&MemberName&"'" &_
                         				",'"&BillfillerID&"','"&MemberName&"'" &_
                         				","&BillFillDate&","&BillFillDate2&",'0',0,"&funGetDate(Sys_Now,1)&",'"&MemID&"'" &_
                         				",'"&tempValue4&"`"&ImportFileName&"','"&Version&"'" &_
                         				",null" &_
                         				",0,'0','1')"
                         				'response.write strInsert
                         				conn.execute strInsert		    

    			        Corr= Corr+1    
     			        i=i+1  			        
                    end if
    			  wend 
	    		  set fs=nothing
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
              else
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0

	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
                    txtline = Replace(txtline,"'","")
                    'txtline = Replace(txtline,""","")
      			  	tempValue=split(trim(txtline),"@@")
      			  	if UBound(tempValue)=7 then
      			  	  if trim(tempValue(0))="" or trim(tempValue(1))="" or trim(tempValue(2))="" or trim(tempValue(3))="" or trim(tempValue(4))="" or trim(tempValue(5))="" or trim(tempValue(6))="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
     			      else
      			        'response.write "第" & i & "行: " & txtline & "<br>"
						'新增每筆記錄---------------------------------------------------------------------------------------------------
                        	'違規日期
                           theIllegalDate=funGetDate(gOutDT(tempValue(3)) &" "&left(trim(tempValue(4)),2)&":"&right(trim(tempValue(4)),2),1)
                            '---------------------------------------------------------------------------
                             '法條金額
                             strLaw="select Level1 from law where itemid='"&Trim(tempValue(7))&"'"
                            	set rsLaw=conn.execute(strLaw)
                             Level1=""
                              if not rsLaw.eof then
                               Level1=trim(rsLaw("Level1"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing
                          '---------------------------------------------------------------------------				
                         	'應到案日期
                         		theDealLineDate=funGetDate(gOutDT(trim(tempValue(5))),0)
                         	'BillBase
							CarSimpleID=chkCarNoFormat(tempValue(1)) 
                             Sys_Now=DateAdd("s",1,Sys_Now)
                         	strInsert="insert into BillBase(SN,BillTypeID,BillNo,UseTool,Insurance,CarNo,CarSimpleID,IllegalDate" & _
                         				",IllegalAddressID,IllegalAddress,Rule1,ForFeit1" &_
                         				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
                         				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
                         				",BillFillerMemberID,BillFiller" &_
                         				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
                         				",Note,RuleVer,DriverSex,DOUBLECHECKSTATUS,BILLBASETYPEID,equipmentid)" &_
                         				" values(BillBase_seq.nextval,'2','',0,0" &_
                         				",'"&tempValue(1)&"',"&tempValue(6) &_						          
                         				","&theIllegalDate&",null" &_
                         				",'"&tempValue(2)&"','"&tempValue(7)&"'" &_
                         				","&Level1&"" &_
                         				",null,null,null" &_
                         				",null,null,null" &_
                         				",'"&UnitID&"','"&BillfillerID&"','"&MemberName&"'" &_
                         				",'"&BillfillerID&"','"&MemberName&"'" &_
                         				","&theIllegalDate&","&theDealLineDate&",'0',0,"&funGetDate(Sys_Now,1)&",'"&MemID&"'" &_
                         				",'"&tempValue(0)&"`"&ImportFileName&"','"&Version&"'" &_
                         				",null" &_
                         				",0,'0','1')"
                         				'response.write strInsert
                         				conn.execute strInsert		        
     			        '--------------------------------------------------------------------------------
    			        Corr= Corr+1    
     			        i=i+1  			        

     			      end if  
     			    else
      			        response.write "第" & i & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1     			        
     			        i=i+1
                    end if
    			  wend 
	    		  set fs=nothing
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

				  end if
      	
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