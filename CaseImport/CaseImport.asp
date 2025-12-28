<% 
Response.Expires = -1
Server.ScriptTimeout = 60000
%>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="freeaspupload.asp" -->
<%
Function Utf8ToBig5(strUtf8)
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")

    ' 先寫入 UTF8 編碼的字串
    stream.Type = 2 ' Text
    stream.Charset = "big5"
    stream.Open
    stream.WriteText strUtf8

    ' 轉換成 BIG5
    stream.Position = 0
    stream.Charset = "utf-8"
    Utf8ToBig5 = stream.ReadText

    stream.Close
    Set stream = Nothing
End Function
' ****************************************************
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"

  Dim uploadsDirVar
  uploadsDirVar = "C:\Inetpub\wwwroot\Traffic\CaseImport\Data" 
  
' ****************************************************
' Note: this file uploadTester.asp is just an example to demonstrate
' the capabilities of the freeASPUpload.asp class. There are no plans
' to add any new features to uploadTester.asp itself. Feel free to add
' your own code. If you are building a content management system, you
' may also want to consider this script: http://www.webfilebrowser.com/

		strCity="select value from apconfigure where id=31"
            	set rsCity=conn.execute(strCity)
             	if not rsCity.eof then
            		tCity=trim(rsCity("value"))
            	end If
				'tCity="台東縣"
            	rsCity.close
           	set rsCity=Nothing


function chkCarNoFormat(CarNo)

	strHeavy="ABCFGHIJKLMNOPY"	   '重機第一碼
	strSmall="DEQRSTUVWXZ"	'//輕機第一碼
	if InStr(CarNo,"-")>= 0	 then 
		CarNoArray=split(CarNo,"-")
		if len(CarNoArray(0))=2 and len(CarNoArray(1))=2 then 
			chkCarNoFormat=2
		elseif len(CarNoArray(0))=2 and len(CarNoArray(1))=4 or (len(CarNoArray(0))=4 or (len(CarNoArray(0))=3 and len(CarNoArray(1))=4) and len(CarNoArray(1))=2) or (len(CarNoArray(0))=2 and  len(CarNoArray(1))=3) or (len(CarNoArray(0))=3 and len(CarNoArray(1))=2) then
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
   <%if tCity="新竹市" then %>

	var loginid = "";
	var accept = document.getElementsByName("loginid");
　　for(i=0;i<accept.length;i++)
　　{
          var c;
    　　 if(accept[i].checked)
       　　loginid=accept[i].value;   
　　}
		myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_Insert&loginid="+loginid;
		myForm.submit();
	<%else%>
		myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_Insert";
		myForm.submit();
	<%end if%>
   }
   else
   {
    alert("請選擇檔案");
   } 
  }
  function BrowsFile()
  {
   <%if tCity="新竹市" then %>
	var loginid = "";
　　var accept = document.getElementsByName("loginid");
　　for(i=0;i<accept.length;i++)
　　{
          var c;
    　　 if(accept[i].checked)
       　　loginid=accept[i].value;   
　　}

		myForm.action="CaseImport.asp?ImportFileName="+myForm.attach1.value+"&loginid="+loginid;
		myForm.submit();
	<%else%>
		myForm.action="CaseImport.asp?ImportFileName="+myForm.attach1.value;
		myForm.submit();
	<%end if%>
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
				if tCity="高雄縣" Or tCity="新竹市" then
					strMem=strMem & " and UnitID ='" & UnitID &"'"
				end if
				'strMem=strMem & " and RecordStateID=0 and AccountStateID=0 "
	'response.write strMem
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

			  	
                'tCity="高雄縣"
            	
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
	   			 
			  response.write "建檔人員&nbsp;: " & memName & "&nbsp;&nbsp;&nbsp;"
             ' response.write "停管業務單位&nbsp;: " & UnitName & "&nbsp;&nbsp;&nbsp;"

			 if tCity="新竹市" Then 
				  response.write "&nbsp;填單人員&nbsp;:<input type=""radio"" checked value=""J220017575"" name=""loginid"""
				  If request("loginid")="0212" Then 
						response.write " Checked" 
'				  ElseIf request("loginid")="" Then 
'						response.write " Checked" 
				  End if
				  response.write ">曹圃禎&nbsp;&nbsp;"	
				  
				

			 ElseIf tCity="台東縣" Then 
				  response.write "&nbsp;舉發人員1&nbsp;<input type=""text"" name=""BillMemID1"" value='"&session("BillMemID1")&"' size=""8"" onkeyup=""getBillMemID1();""><div id=""Layer1"" style=""position:absolute;z-index:1;"">"&session("BillMemberID1Name")&"</div>"
					response.write "<input type=""hidden"" name=""BillMemberID1"" value="&session("BillMemberID1")&">"
					response.write "<input type=""hidden"" name=""BillMemberID1Name"" value="&session("BillMemberID1Name")&">"

				  response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;舉發人員2&nbsp;<input type=""text"" name=""BillMemID2"" value='"&session("BillMemID2")&"' size=""8"" onkeyup=""getBillMemID2();""><div id=""Layer2"" style=""position:absolute;z-index:1;"">"&session("BillMemberID2Name")&"</div>"
					response.write "<input type=""hidden"" name=""BillMemberID2""  value="&session("BillMemberID2")&">"
					response.write "<input type=""hidden"" name=""BillMemberID2Name"" value="&session("BillMemberID2Name")&">"

				  response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;舉發人員3&nbsp;<input type=""text"" name=""BillMemID3"" value='"&session("BillMemID3")&"' size=""8"" onkeyup=""getBillMemID3();""><div id=""Layer3"" style=""position:absolute;z-index:1;"">"&session("BillMemberID3Name")&"</div>"
					response.write "<input type=""hidden"" name=""BillMemberID3""  value="&session("BillMemberID3")&">"
					response.write "<input type=""hidden"" name=""BillMemberID3Name"" value="&session("BillMemberID3Name")&">"

				  response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;舉發人員4&nbsp;<input type=""text"" name=""BillMemID4"" value='"&session("BillMemID4")&"' size=""8"" onkeyup=""getBillMemID4();""><div id=""Layer4"" style=""position:absolute;z-index:1;"">"&session("BillMemberID4Name")&"</div>"
					response.write "<input type=""hidden"" name=""BillMemberID4""  value="&session("BillMemberID4")&">"
					response.write "<input type=""hidden"" name=""BillMemberID4Name"" value="&session("BillMemberID4Name")&">"

			 End If
			 
			  %>&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
			  <%If tCity="高雄縣" Or tCity="新竹市" Then %> 
				<a target="blank" href="CaseImportFileLog.asp">匯入記錄查詢</a>&nbsp;&nbsp;&nbsp; 
                                <a target="blank" href="停管匯入2.docx">使用說明</a>
								&nbsp;&nbsp;&nbsp; 
                                <a target="blank" href="ExecltoCSV.doc">Excel 轉換 CSV 說明</a>
			  <%End if%>
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#1BF5FF"><font size="4"><strong></strong></font>舉發單資料匯入系統</td></tr>
    		</table>
    		  <p><p>
  <font size="4">選擇檔案</font><span><font size="4"> </font><input type="text" name="T1" size="50" value="<%=request("ImportFileName")%>" readonly></span>&nbsp;
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
  注意：每檔案容量限制上傳 2M &nbsp;&nbsp;&nbsp;<b> * 檔名請以英文、數字命名 。檔案請存放於 C:\ 或D:\</b><br>
  <%If tCity="高雄縣" Or tCity="新竹市"Then %>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   注意 : *  <b>檔案第一行欄位名稱請去除</b>
  <%end if %>
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
				     <%if tCity="新竹市" then %>
					　　var accept = "";
						var loginid = "";
							accept = document.getElementsByName("loginid");
					　　for(i=0;i<accept.length;i++)
					　　{
							  var c;
						　　 if(accept[i].checked)
						   　　loginid=accept[i].value;   
					　　}
loginid = "J220017575";
						if (loginid == "")
						{
							alert("請選擇填單人員1");
						}
						else
						{
						  myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert&loginid="+loginid;
						  myForm.submit();
						}
					  <%else%>
						  myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
						  myForm.submit();
					  <%end if%>
	    		    }
	    		  </script>
                <%
			    else
				%>
                  <script language="javascript">

				     <%if tCity="新竹市" then %>
					　　var accept = "";
						var loginid = "";
							accept = document.getElementsByName("loginid");

					　　for(i=0;i<accept.length;i++)
					　　{
							  var c;
						　　 if(accept[i].checked)
						   　　loginid=accept[i].value;   
					　　}

loginid = "J220017575";
						if (loginid == "")
						{
							alert("請選擇填單人員2");
						}
						else
						{
						  myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert&loginid="+loginid;
						  myForm.submit();
						}
					<%else%>
						 myForm.action="CaseImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
						myForm.submit();
					<%end if%>
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
	    		  
      		   	  FileName=Server.MapPath("Data/" &fs.GetFileName(request("ImportFileName")))






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
							'response.write "1@"&tempValue(2)&"<br>"
                           theIllegalDate=funGetDate(gOutDT(Replace(tempValue(2),"/","")) &" "&tempValue(3),1)

							If DateDiff("d","2015/07/01",gOutDT(Replace(tempValue(2),"/","")))>-1 Then
								Rule="5630001"
							Else
								Rule="5620001"
							End If
							
						   BillFillDate=funGetDate(date(),0)
						   	BillFillDate2=funGetDate(date()+45,0)
                            '---------------------------------------------------------------------------
                             '法條金額
                             strLaw="select Level1 from law where itemid='"&Rule&"'"
                            	set rsLaw=conn.execute(strLaw)
                             Level1=""
                              if not rsLaw.eof then
                               Level1=trim(rsLaw("Level1"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing
                     
							CarSimpleID="1"
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
                         				",'"&trim(tempValue(4))&"','"&Rule&"'" &_
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
							'response.write "2@"&tempValue(2)&"<br>"
                           theIllegalDate=funGetDate(gOutDT(Replace(tempValue(2),"/","")) &" "&left(trim(tempValue(3)),2)&":"&right(trim(tempValue(3)),2),1)
						   BillFillDate=funGetDate(date(),0)
						   'response.write "3@"&tempValue(5)&"<br>"
     				    	BillFillDate2=funGetDate(gOutDT(Replace(tempValue(5),"/","")),0)
							
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
				strMem="select ChName,UnitID,MemberID from MemberData where loginid ='" & trim(tempValue(8)) & "' and RecordStateID=0 and AccountStateID=0 "
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
	    		  set fs=Nothing

				  
		elseif tCity="高雄縣" Or tCity="新竹市" then

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------

      			  set txtStream = fs.opentextfile(FileName,1) 
	    		    i = 1
		    	    Corr=0
    			    Err=0
    If Not txtStream.AtEndOfStream Then
        txtline = txtStream.ReadLine
        fields = Split(txtline, ",")

    End If
	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
      			  	tempValue=split(trim(txtline),",")
If tempValue(0) >= "1" Then

Dim yeara_a, month_a, day_a, hour_a, minute_a
Dim dt_a, result_a, result_b,newDate_a
year_a = CInt(Left(tempValue(5), 3))+1911
month_a = CInt(Mid(tempValue(5), 4, 2))
day_a = CInt(Mid(tempValue(5), 6, 2))

hour_a = CInt(Left(tempValue(6), 2))
minute_a = CInt(Right(tempValue(6), 2))
dt_a = DateSerial(year_a, month_a, day_a) + TimeSerial(hour_a, minute_a, 0)

result_a = Year(dt_a) & "-" & Right("0" & Month(dt_a), 2) & "-" & Right("0" & Day(dt_a), 2) & " " & _
         Right("0" & Hour(dt_a), 2) & ":" & Right("0" & Minute(dt_a), 2) & ":" & Right("0" & Second(dt_a), 2)

result_b = CDate(result_a)
newDate_a = DateAdd("d", 45, result_b)
result_b = Year(newDate_a) & "-" & Right("0" & Month(newDate_a), 2) & "-" & Right("0" & Day(newDate_a), 2) & " " & _
            Right("0" & Hour(newDate_a), 2) & ":" & Right("0" & Minute(newDate_a), 2) & ":" & Right("0" & Second(newDate_a), 2)

      			  	if UBound(tempValue)=11 then
      			  	  if trim(tempValue(0))="" or trim(tempValue(1))="" or trim(tempValue(2))="" or trim(tempValue(3))="" or trim(tempValue(4))="" or trim(tempValue(5))="" or trim(tempValue(6))="" or trim(tempValue(7))="" or trim(tempValue(8))=""  or trim(tempValue(9))=""  or trim(tempValue(10))="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
     			      else
      			        'response.write "第" & i & "行: " & txtline & "<br>"
			'新增每筆記錄---------------------------------------------------------------------------------------------------
                        	'違規日期
							'response.write "4@"&tempValue(3)&"<br>"



                           theIllegalDate=result_a



							Rule=tempValue(11)
				
						   BillFillDate=result_a
     				    	BillFillDate2=result_b
					
                            '---------------------------------------------------------------------------
                             '法條金額

                             strLaw="select Level1 from law where itemid='"&Rule&"'"
                            	set rsLaw=conn.execute(strLaw)
                             Level1=""
                              if not rsLaw.eof then
                               Level1=trim(rsLaw("Level1"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing

                    
							CarSimpleID=1
							if CarSimpleID="" then CarSimpleID="1"
							
							
			RecordMemID=""
       '-----------------------------------------------------------------------------------------------------------------------------------------------------
			if tCity="高雄縣" Then 
				strMem="select ChName,UnitID,MemberID from MemberData where loginid ='K000041' and RecordStateID=0 and AccountStateID=0 "
                         	set rsMem=conn.execute(strMem)
                        	if not rsMem.eof then
                        		MemberName=trim(rsMem("ChName"))
                        		UnitID=trim(rsMem("UnitID"))
								MemberID=trim(rsMem("MemberID"))
								RecordMemID=MemberID
                            else

    ' 取得 GETBILLMEMBERID，根據 billmemid 查詢 memberdata
    Set billmemidRS = Conn.Execute("SELECT MEMBERID,UNITID FROM memberdata WHERE CHNAME = '" & tempValue(9) & "'")
    If billmemidRS.EOF Then
        getbillmemberid = Null
	getbillunitid = Null
    Else
        getbillmemberid = billmemidRS("MEMBERID")
	getbillunitid = billmemidRS("UNITID")
    End If
    billmemidRS.Close
                                MemberName=tempValue(9)
                        		UnitID=getbillunitid
								MemberID=getbillmemberid
								RecordMemID=trim(Session("User_ID"))
                            end if
                            rsMem.close
                        	set rsMem=Nothing

			elseif tCity="新竹市" Then 

				strMem="select ChName,UnitID,MemberID from MemberData where CHNAME ='"&tempValue(9)&"' and RecordStateID=0 and AccountStateID=0 "
                         	set rsMem=conn.execute(strMem)
                        	if not rsMem.eof then
                        		MemberName=trim(rsMem("ChName"))
                        		UnitID=trim(rsMem("UnitID"))
								MemberID=trim(rsMem("MemberID"))
								RecordMemID=trim(Session("User_ID"))
								
                            else
                                MemberName=""
                        		UnitID=""
								MemberID=""
								RecordMemID=""
                            end if
                            rsMem.close
                        	set rsMem=Nothing
							'response.write MemberID
			End If

                          '---------------------------------------------------------------------------				
                            	Sys_Now=DateAdd("s",1,Sys_Now)

'smith 判斷單號是否已經匯入過 start
strSQL="select * from BillBase where carno='" & tempValue(3) & "' and IllegalDate=TO_DATE('"&result_a&"','YYYY-MM-DD HH24:MI:SS') and RULE1= '" & tempValue(11) & "'"
set rsMatch=conn.execute(strSQL)	
'response.write tempValue(0)+"開始寫入資料庫"+Rule+" " + Level1 + " " +UnitID + " "+MemberID + " "+MemberName + " "+RecordMemID+ " "+trim(tempValue(7))+ " "+trim(tempValue(3)) + " "+result_a
						if rsMatch.eof and MemberID<>"" then 

							tmpMemo=""
					

						'smith end
Dim maxSnRS,nbsn
    Set maxSnRS = Conn.Execute("SELECT MAX(SN) AS MAXSN FROM BILLBASE")
	nbsn = maxSnRS("MAXSN")
	nbsn = CLng( nbsn )+1
    'response.write nbsn
    maxSnRS.close
    Set maxSnRS=Nothing

  

													strInsert="insert into BillBase(SN,BillTypeID,BillNo,UseTool,Insurance,CarNo,CarSimpleID,IllegalDate" & _
																",IllegalAddressID,IllegalAddress,Rule1,ForFeit1" &_
																",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
																",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
																",BillFillerMemberID,BillFiller" &_
																",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
																",Note,RuleVer,DriverSex,DOUBLECHECKSTATUS,BILLBASETYPEID,ImageFileNameB,equipmentid)" &_
																" values("&nbsn&",2,null,8,0,'"&trim(tempValue(3))&"',1" &_																						          
																",TO_DATE('"&result_a&"','YYYY-MM-DD HH24:MI:SS'),null,'"&trim(tempValue(7))&"'" &_
																",'"&Rule&"'" &_
																","&Level1&"" &_
																",null,null,null" &_
																",null,null,null" &_
																",'"&UnitID&"','"&MemberID&"','"&MemberName&"'" &_
																",'"&MemberID&"','"&MemberName&"'" &_
																",TO_DATE('"&result_a&"','YYYY-MM-DD HH24:MI:SS'),TO_DATE('"&result_b&"','YYYY-MM-DD HH24:MI:SS'),'0',0,"&funGetDate(Sys_Now,1)&",'"&RecordMemID&"'" &_
																",'"&tmpMemo&"','"&Version&"'" &_
																",null" &_
																",0,'0',null,'1')"
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
                    end if
    			  wend 
	    		  set fs=nothing
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
				  elseif  tCity="花蓮縣" then
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0

	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
      			  	'tempValue0=trim(mid(txtline,1,9))
      			  	'tempValue1=trim(mid(txtline,10,7))
      			  	'tempValue2=trim(mid(txtline,18,4))
      			  	'tempValue3=trim(mid(txtline,31,30))
      			  	'tempValue4=trim(mid(txtline,62,20))
      			  	'tempValue5=trim(mid(txtline,29,2))
					Field=Split(txtline,",")

				tempValue0=Field(0)
				tempValue1=Field(1)
				tempValue2=Field(2)
				tempValue5=Field(3)
				tempValue5=Field(4)
				tempValue3=Field(5)
				tempValue4=Field(6)

      			  	  if trim(tempValue0)="" or trim(tempValue1)="" or trim(tempValue2)="" or trim(tempValue3)="" or trim(tempValue4)="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
     			     else 
      			        'response.write "第" & i & "行: " & txtline & "<br>"
						'新增每筆記錄---------------------------------------------------------------------------------------------------
                        	'違規日期
							'response.write "5@"&tempValue1&"<br>"
                           theIllegalDate=funGetDate(gOutDT(Replace(tempValue1,"/","")) &" "&left(trim(tempValue2),2)&":"&right(trim(tempValue2),2),1)


						If DateDiff("d","2015/07/01",gOutDT(Replace(tempValue1,"/","")))>-1 Then
							Rule="5630001"
						Else
							Rule="5620001"
						End If

						   BillFillDate=funGetDate(date(),0)
						   	BillFillDate2=funGetDate(date()+45,0)
                            '---------------------------------------------------------------------------
                             '法條金額
                             strLaw="select Level1 from law where itemid='"&Rule&"'"
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
                         				",'"&tempValue3&"','"&Rule&"'" &_
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

			'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
			elseif  tCity="台東縣" then

      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0
					dim fos
					folderPath=server.mappath("\traffic\StopCarPicture\"&gInitDT(date))

					set fos=server.CreateObject("Scripting.FileSystemObject")

					If not fos.FolderExists(folderPath) then
						fos.CreateFolder folderPath
					End if

	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
      			  	tempValue0=trim(mid(txtline,1,9))
      			  	tempValue1=trim(mid(txtline,10,7))
      			  	tempValue2=trim(mid(txtline,18,4))
      			  	tempValue3=trim(mid(txtline,31,30))
      			  	tempValue4=trim(mid(txtline,62,20))
      			  	tempValue5=trim(mid(txtline,29,2))
      			  	BillMemberID1=Trim(session("BillMemberID1"))
					MemberName1=Trim(session("BillMemberID1Name"))

      			  	BillMemberID2=Trim(session("BillMemberID2"))
					MemberName2=Trim(session("BillMemberID1Name"))

      			  	BillMemberID3=Trim(session("BillMemberID3"))
					MemberName3=Trim(session("BillMemberID1Name"))

      			  	BillMemberID4=Trim(session("BillMemberID4"))
					MemberName4=Trim(session("BillMemberID1Name"))

					strSQL="select count(1) cnt from billbase where billno is not null and imagefilename like '%\"&tempValue4&".jpg' and recordstateid=0"
					filecnt=0:chk_BillNo=""
					set rscnt=conn.execute(strSQL)
					filecnt=cdbl(rscnt("cnt"))
					rscnt.close

      			  	  if trim(tempValue0)="" or trim(tempValue1)="" or trim(tempValue2)="" or trim(tempValue3)="" or trim(tempValue4)="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
					 ElseIf BillMemberID1="" Then 
						response.flush
     			        Err= Err+1      
     			        i=i+1
						response.write "舉發人員1輸入錯誤"
					elseIf filecnt>0 Then
						response.write "第" & i  & "行: " & tempValue0 & "已舉發過!!<br>"
						response.flush
						Err= Err+1      
     			        i=i+1
     			     else 
      			        'response.write "第" & i & "行: " & txtline & "<br>"
						'新增每筆記錄---------------------------------------------------------------------------------------------------
                        	'違規日期
							'response.write "5@"&tempValue1&"<br>"
                           theIllegalDate=funGetDate(gOutDT(Replace(tempValue1,"/","")) &" "&left(trim(tempValue2),2)&":"&right(trim(tempValue2),2),1)
						   BillFillDate=funGetDate(date(),0)
						   	BillFillDate2=funGetDate(date()+45,0)
							'-----------------------20121118 台東改為用停車日當作違規日期--mark start----------------------
							'strSQL="select DealLineDate from BillBase where CarNo='"&trim(tempValue0)&"' and IllegalDate="&theIllegalDate&" and ImageFileNameB is not null and BillNo is null and Recordstateid=0"
							'
							'set rsstop=conn.execute(strSQL)
							'if Not rsstop.eof then
						    '		theIllegalDate=funGetDate(DateAdd("d",1,rsstop("DealLineDate"))&" 00:00:00",1)
							'end if
							'rsstop.close
'                            '-------------------------------------------------------------mark end--------------
                             '法條金額


						If DateDiff("d","2015/07/01",gOutDT(Replace(tempValue1,"/","")))>-1 Then
							Rule="5630001"
						Else
							Rule="5620001"
						End If

                             strLaw="select Level1 from law where itemid='"&Rule&"'"
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
                         				",MemberStation,BillUnitID,BillMemID1,BillMem1,BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
                         				",BillFillerMemberID,BillFiller" &_
                         				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
                         				",Note,RuleVer,DriverSex,DOUBLECHECKSTATUS,BILLBASETYPEID,equipmentid,imagefilename)" &_
                         				" values(BillBase_seq.nextval,'2','',0,0" &_
                         				",'"&tempValue0&"',"&CarSimpleID&_						          
                         				","&theIllegalDate&",null" &_
                         				",'臺東市"&tempValue3&"','"&Rule&"'" &_
                         				","&Level1&"" &_
                         				",null,null,null" &_
                         				",null,null,null" &_
                         				",'"&UnitID&"','"&BillMemberID1&"','"&MemberName1&"','"&BillMemberID2&"','"&MemberName2&"','"&BillMemberID3&"','"&MemberName3&"','"&BillMemberID4&"','"&MemberName4&"'" &_
                         				",'"&BillMemberID1&"','"&MemberName1&"'" &_
                         				","&BillFillDate&","&BillFillDate2&",'0',0,"&funGetDate(Sys_Now,1)&",'"&MemID&"'" &_
                         				",'"&tempValue4&"`"&ImportFileName&"','"&Version&"'" &_
                         				",null" &_
                         				",0,'0','1','"&gInitDT(date)&"\"&tempValue4&".jpg')"
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
							'response.write "6@"&tempValue(3)&"<br>"
                           theIllegalDate=funGetDate(gOutDT(Replace(tempValue(3),"/","")) &" "&left(trim(tempValue(4)),2)&":"&right(trim(tempValue(4)),2),1)
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
								'response.write "7@"&tempValue(5)&"<br>"
                         		theDealLineDate=funGetDate(gOutDT(Replace(trim(tempValue(5)),"/","")),0)
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
			<tr><td bgcolor="#1BF5FF"><font size="4"><strong></strong></font>　
			</td></tr>
		</table>


		　</form>
</html>
<script>
function getBillMemID1(AccKey){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMemID1.value=myForm.BillMemID1.value.toUpperCase();
	}
	if (myForm.BillMemID1.value.length > 1){
		var BillMemNum=myForm.BillMemID1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
	}else if (myForm.BillMemID1.value.length <= 1 && myForm.BillMemID1.value.length > 0){
		Layer1.innerHTML="";
	}else{
		Layer1.innerHTML="";
	}
}

function getBillMemID2(AccKey){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMemID2.value=myForm.BillMemID2.value.toUpperCase();
	}
	if (myForm.BillMemID2.value.length > 1){
		var BillMemNum=myForm.BillMemID2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
	}else if (myForm.BillMemID2.value.length <= 1 && myForm.BillMemID2.value.length > 0){
		Layer2.innerHTML="";
	}else{
		Layer2.innerHTML="";
	}
}

function getBillMemID3(AccKey){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMemID3.value=myForm.BillMemID3.value.toUpperCase();
	}
	if (myForm.BillMemID3.value.length > 1){
		var BillMemNum=myForm.BillMemID3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
	}else if (myForm.BillMemID3.value.length <= 1 && myForm.BillMemID3.value.length > 0){
		Layer3.innerHTML="";
	}else{
		Layer3.innerHTML="";
	}
}

function getBillMemID4(AccKey){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMemID4.value=myForm.BillMemID4.value.toUpperCase();
	}
	if (myForm.BillMemID4.value.length > 1){
		var BillMemNum=myForm.BillMemID4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
	}else if (myForm.BillMemID4.value.length <= 1 && myForm.BillMemID4.value.length > 0){
		Layer4.innerHTML="";
	}else{
		Layer4.innerHTML="";
	}
}
</script>



							
<%


'上傳檔案到伺服器
function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey

    Set Upload = New FreeASPUpload
SavePath = Server.MapPath("data/")
    Upload.Save SavePath
FileCount = Upload.Files.Count

If FileCount > 0 Then
    For i = 1 To FileCount
        FileName = Upload.Files(i).FileName
        Response.Write "檔案 " & FileName & " 已成功上傳<br>"
    Next
Else
    Response.Write "沒有檔案被上傳"
End If

Set Upload = Nothing
end function


    SaveFiles()
%>