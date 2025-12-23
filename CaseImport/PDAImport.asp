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
<title>PDA舉發單資料匯入系統</title>
<script language="javascript">
  function InsertData()
  {
   if (myForm.T1.value!="")
   {
		myForm.action="PDAImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_Insert";
	    myForm.submit();
   }
   else
   {
		alert("請選擇檔案");
   } 
  }
  function BrowsFile()
  {
	    myForm.action="PDAImport.asp?ImportFileName="+myForm.attach1.value;
	    //myForm.T1.value=myForm.attach1.value;
		//myForm.action="CaseImport.asp";
	    myForm.submit();
  }
  function Clear()
  {
		myForm.T1.value="";
		myForm.action="PDAImport.asp";    
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
UnitID=trim(Session("UnitI_D"))
			

			  response.write "<BR>"
			  '------------------------------------------------------------------------------------

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
                'tCity="高雄縣"
            	
            	strUnit="select UnitName from Unitinfo where UnitID ='" & UnitID & "'"
            	set rsUnit=conn.execute(strUnit)
             	if not rsUnit.eof then
            		UnitName =trim(rsUnit("UnitName"))
            	end if
            	rsUnit.close
            	set rsUnit=nothing

			  response.write "建檔人: " & memName & "&nbsp;&nbsp;&nbsp;"
			  %>
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>PDA舉發單資料匯入系統</td></tr>
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
                      myForm.action="PDAImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
                      myForm.submit();
	    		    }
	    		  </script>
                <%
			    else
				%>
                  <script language="javascript">
                      myForm.action="PDAImport.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
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
if tCity="高雄縣" Or tCity="屏東縣" Or tCity="台東縣" Or tCity="高雄市" then
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0

	    		  while not txtStream.atEndOfStream 
			    	    txtline = txtStream.readline 
	      			  	tempValue=split(trim(txtline),",")
	      			  	if UBound(tempValue)=26 then
	      			  	  if trim(tempValue(0))="" or trim(tempValue(1))="" or trim(tempValue(2))="" or trim(tempValue(3))="" or trim(tempValue(4))="" or trim(tempValue(5))="" or trim(tempValue(6))="" or trim(tempValue(7))="" or trim(tempValue(8))="" then  
			  			        response.write "第" & i  & "行: " & txtline & "<br>"
								response.flush
			 			        Err= Err+1      
			 			        i=i+1
     				      else
      			        'response.write "第" & i & "行: " & txtline & "<br>"
			'新增每筆記錄------------------------------------------------------------------------------------------------------------------------------------------
			                '違規法條
							UseTable=""
							TrafficAccidentType=""
                   ' 1---舉發類型--------------------------------------------------------------------------------------------------------------------------------------
                            '    事故種類 TrafficAccidentType
							                ' 7 掌-欄停  8 掌-行人  9 掌-肇事  0 掌-拖吊
							BillTypeID=trim(tempValue(1))
							'BillTypeID="8"
							if BillTypeID="7" or BillTypeID="0" then 
							  UseTable="BillBase"
							elseif BillTypeID="8" then 
							  UseTable="PasserBase"
							elseif BillTypeID="9" then 
							  UseTable="BillBase"
							  TrafficAccidentType="A3"
							end If
							  BillTypeID="1"							
    				' 2------應到案日期----------------------------------------------------------------------------------------------------------------------------------
							DealLineDate=funGetDate(gOutDT(tempValue(2)),0)
					' 3------到案地點-------------------------------------------------------------------------------------------------------------------------------------


							MemberStateIon=trim(tempValue(3))
						  	strMem="select StationID from Station where dcistationid ='" & MemberStateIon & "'"
				        	set rsMem=conn.execute(strMem)
			             	if not rsMem.eof then
								MemberStateIon=trim(rsMem("StationID"))
			            	end if
			            	rsMem.close
			            	set rsMem=Nothing
					' 4------違規人生日----------------------------------------------------------------------------------------------------------------------------------
							DriverBirthDay=funGetDate(gOutDT(tempValue(4)),0)
					' 5------車種--------------------------------------------------------------------------------------------------------------------------------------------
							CarTypeID=trim(tempValue(5))
					' 6------扣件------------------------XXXXX---------------------------------------------------------------------------------------------------------
							ItemID=trim(tempValue(6))
							If ItemId="000" Then ItemID=""
					' 7------違規人證號-----------------------------------------------------------------------------------------------------------------------------------
							DriverID=trim(tempValue(7))
					' 8------違規人姓名-----------------------------------------------------------------------------------------------------------------------------------
							DriverName=trim(tempValue(8))
					' 9------違規人地址-----------------------------------------------------------------------------------------------------------------------------------
							DriverAddress=trim(tempValue(9))
					' 10------保險證------------------------------------------------------------------------------------------------------------------------------------------
							Insurance=trim(tempValue(10))
					' 11------簽收情形   
					'A簽收/ U拒簽收/ 2拒簽已收/ 3已簽拒收/ 5補開單 
					'0: 正常  1: 拒簽   2: 拒收  3: 拒簽拒收
					'---------------------------------------------------------------------------------------------------------------------------------------------
					      if trim(tempValue(11))="1" then
						    SignType="2"
						  elseif trim(tempValue(11))="2" then
						    SignType="3"
						  elseif trim(tempValue(11))="3" then
						    SignType="U"
						  else
							SignType="A"
						  end if
					' 12------車號---------------------------------------------------------------------------------------------------------------------------------------------
							CarNo=trim(tempValue(12))
					' 13------員警身份證字號-------------------------------------------------------------------------------------------------------------------------
							'PoliceID=trim(tempValue(13))
					' 14------員警姓名--------------------------------------------------------------------------------------------------------------------------------------
							MemberName=trim(tempValue(14))
						  	strMem="select MemberID from MemberData where AccountStateID=0 and RecordStateID=0 and ChName ='" & MemberName & "'"
				        	set rsMem=conn.execute(strMem)
			             	if not rsMem.eof then
								MemberID=trim(rsMem("MemberID"))
			            	end if
			            	rsMem.close
			            	set rsMem=Nothing
					' 15------限制數------------------------------------------------------------------------------------------------------------------------------------------
							RuleSpeed=trim(tempValue(15))
					' 16------實際數-------------------------------------------------------------------------------------------------------------------------------------------
							IllegalSpeed=trim(tempValue(16))
					' 17------違規法條----------------------------------------------------------------------------------------------------------------------------------------
							LawID=trim(tempValue(17))

                             '法條金額
                             strLaw="select Level1 from law where itemid='"&trim(tempValue(17))&"'"
                            	set rsLaw=conn.execute(strLaw)
                             Level1=""
                              if not rsLaw.eof then		
                               Level1=trim(rsLaw("Level1"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing
					' 18------單號-------------------------------------------------------------------------------------------------------------------------------------------------
							BillNo=trim(tempValue(18))
					' 19~22------違規地點-----------XXXXX----------------------------------------------------------------------------------------------------------------
					' 23~24------違規日期、時間等於填單日-------------------------------------------------------------------------------------------------------------
                           theIllegalDate=funGetDate(gOutDT(tempValue(23)) &" "&left(trim(tempValue(24)),2)&":"&right(trim(tempValue(24)),2),1)
						   BillFillDate=funGetDate(gOutDT(tempValue(23)),0)
					' 25------違規地點代碼--------------------------------------------------------------------------------------------------------------------------------------
							IllegalAddressID=trim(tempValue(25))
                             strLaw="select address from street where streetid='"&IllegalAddressID&"'"
                            	set rsLaw=conn.execute(strLaw)
                             IllegalAddress=""
                              if not rsLaw.eof then		
                               IllegalAddress=trim(rsLaw("address"))
                              end if
                          	rsLaw.close
                           	set rsLaw=Nothing
                    ' 26---------------------------------------------------------------------------------------------------------------------------------------------------------------
    						UnitID=trim(tempValue(26))
                    '  -----------------------------------------------------------------------------------------------------------------------------------------------------------------
                            	Sys_Now=DateAdd("s",1,Sys_Now)
'smith 判斷單號是否已經匯入過 start
strSQL="select billno from BillBase where billno='" & BillNo & "' and recordstateid<>-1 "
set rsMatch=conn.execute(strSQL)	
if rsMatch.eof and MemberID<>"" then 
'smith end
  if UseTable="BillBase" then

                         	strInsert="insert into BillBase(SN,BillTypeID,BillNo,UseTool,Insurance,CarNo,CarSimpleID,IllegalDate" & _
                         				",IllegalAddressID,IllegalAddress,Rule1,ForFeit1" &_
                         				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
                         				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
                         				",BillFillerMemberID,BillFiller" &_
                         				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
                         				",Note,RuleVer,DriverSex,"&_
										"DOUBLECHECKSTATUS,BILLBASETYPEID,equipmentid,CarAddID,SignType)" &_
                         				" values("&_
										"BillBase_seq.nextval,'"&BillTypeID&"','"&BillNo&"',0,'"&Insurance&"','"&CarNo&"',"&CarTypeID&","&theIllegalDate&_
										",'"&IllegalAddressID&"','"&IllegalAddress&"','"&LawID&"',"&Level1&"" &_
                         				",'"&DriverID&"',"&DriverBirthDay&",'"&DriverName&"','"&DriverAddress&"',null"&_
										",'"&MemberStateIon&"','"&UnitID&"','"&MemberID&"','"&MemberName&"'" &_
                         				",'"&MemberID&"','"&MemberName&"'" &_
                         				","&BillFillDate&","&DealLineDate&",'0',0,"&funGetDate(Sys_Now,1)&",'"&MemID&"'" &_
                         				",'"&ImportFileName&"','"&Version&"',null" &_
                         				",0,'0','1',null,'"&SignType&"')"
                         				'response.write strInsert
                         				conn.execute strInsert		    
										Corr= Corr+1    
     								    i=i+1			 

  elseif 	UseTable="PasserBase" then
                         	strInsert="insert into PasserBase(SN,  BillTypeID,BillNo,IllegalDate,IllegalAddressID,IllegalAddress" & _
											",Rule1,Forfeit1,DriverID,DriverBirth,Driver,DriverAddress," & _
											"MemberStation,BillUnitID,BillMemID1,BillMem1,BillFillerMemberID," & _
											"BillFiller,BillFillDate,DealLineDate,BillStatus," & _
											"RecordDate,RecordMemberID,Note,RuleVer,DriverSex," & _
											"BillBaseTypeID,DoubleCheckStatus,SignType,RecordStateID" & _
	                         				") values("&_
											"BillBase_seq.nextval,'"&BillTypeID&"','"&BillNo&"',"&theIllegalDate&",'"&IllegalAddressID&"','"&IllegalAddress &_
											"','"&LawID&"',"&Level1&",'"&DriverID&"',"&DriverBirthDay&",'"&DriverName&"','"&DriverAddress&"'"&_
											",'"&MemberStateIon&"','"&UnitID&"','"&MemberID&"','"&MemberName&"','"&MemberID&"'" &_
											",'"&MemberName&"',"&BillFillDate&","&DealLineDate&",'0'"&_
											","&funGetDate(Sys_Now,1)&",'"&MemID&"','"&ImportFileName&"','"&Version&"',null" & _
											",'0',0,'"&SignType&"',0)"
										'response.write strinsert
                         				conn.execute strInsert		    
										Corr= Corr+1    
     								    i=i+1			 

  else
	  	 response.write "第" & i & "行: " & txtline & "  舉發類型有誤<br>"
		 Err= Err+1	
		i=i+1	

  end if
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