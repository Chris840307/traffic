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
<title>拖吊已結舉發單資料匯入系統</title>
<script language="javascript">
  function InsertData()
  {
   if (myForm.T1.value!="")
   {
		myForm.action="CaseImportBillBase3.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_Insert";
		myForm.submit();
   }
   else
   {
    alert("請選擇檔案");
   } 
  }
  function BrowsFile()
  {
  		myForm.action="CaseImportBillBase3.asp?ImportFileName="+myForm.attach1.value;
		myForm.submit();
  }
  function Clear()
  {
    myForm.T1.value="";
    myForm.action="CaseImportBillBase3.asp";    
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
			  
			 
			  response.write "<BR>"
            	
            	strUnit="select UnitName from Unitinfo where UnitID ='" & UnitID & "'"
            	set rsUnit=conn.execute(strUnit)
             	if not rsUnit.eof then
            		UnitName =trim(rsUnit("UnitName"))
            	end if
            	rsUnit.close
            	set rsUnit=nothing

			  %>&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;

		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>拖吊已結舉發單資料匯入系統</td></tr>
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
  注意：每檔案容量限制上傳 2M &nbsp;&nbsp;&nbsp;<b> * 檔名請以英文、數字命名 。檔案請存放於 C:\ 或D:\</b><br>
  <%If tCity="高雄縣" Or tCity="宜蘭縣"Then %>
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
				    	  myForm.action="CaseImportBillBase3.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
						  myForm.submit();
	    		    }
	    		  </script>
                <%
			    else
				%>
                  <script language="javascript">
    					 myForm.action="CaseImportBillBase3.asp?ImportFileName="+myForm.T1.value+"&DB_Selt=DB_StillInsert";
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



'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
      			  set txtStream = fs.opentextfile(FileName) 
	    		    i = 1
		    	    Corr=0
    			    Err=0

	    		  while not txtStream.atEndOfStream 
		    	    txtline = txtStream.readline 
      			  	tempValue=split(trim(txtline),",")

					'0  SN
					SN=tempValue(0)
					'1  BILLTYPEID
					BILLTYPEID="1"
					'2  BILLNO
					BILLNO=tempValue(1)
					'3  CARNO
					CARNO=tempValue(2)
					'4  CARSIMPLEID
					CARSIMPLEID=tempValue(4)
					If CARSIMPLEID="" Then CARSIMPLEID="null"
					'5  CARADDID
					CARADDID="8"
					'6  ILLEGALDATE
					ILLEGALDATE="To_Date('"&tempValue(3)&"','YYYY/MM/DD/HH24/MI/SS')"
					'7  ILLEGALADDRESSID
					ILLEGALADDRESSID=tempValue(16)
					'8  ILLEGALADDRESS
					ILLEGALADDRESS=tempValue(5)
					'9  RULE1
					RULE1=tempValue(6)
					'10  ILLEGALSPEED
					ILLEGALSPEED="null"
					'11  RULESPEED
					RULESPEED="null"
					'12  FORFEIT1
					FORFEIT1=tempValue(7)
					If FORFEIT1="" Then FORFEIT1="null"
					'13  RULE2
					RULE2=""
					'14  FORFEIT2
					FORFEIT2=""
					If FORFEIT2="" Then FORFEIT2="null"
					'15  RULE3
					RULE3=""
					'16  FORFEIT3
					FORFEIT3=""
					If FORFEIT3="" Then FORFEIT3="null"
					'17  RULE4
					RULE4=""
					'18  FORFEIT4
					FORFEIT4=""
					If FORFEIT4="" Then FORFEIT4="null"
					'19  INSURANCE
					INSURANCE=""
					If INSURANCE="" Then INSURANCE="0"
					'20  USETOOL
					USETOOL="null"
					'21  PROJECTID
					PROJECTID=""
					'22  DRIVERID
					DRIVERID=""
					'23  DRIVERBIRTH
					DRIVERBIRTH="null"
					'24  DRIVER
					DRIVER=""
					'25  DRIVERADDRESS
					DRIVERADDRESS=""
					'26  DRIVERZIP
					DRIVERZIP=""
					'27  MEMBERSTATION
					MEMBERSTATION=tempValue(17)
					'28  BILLUNITID
					BILLUNITID=tempValue(11)
							strMem="select MemberID,chname from MemberData where loginid ='" & tempValue(13) & "' and RecordStateID=0 and AccountStateID=0 "
                         	set rsMem=conn.execute(strMem)
                        	if not rsMem.eof then
								MemberID=trim(rsMem("MemberID"))
								chname=trim(rsMem("chname"))
                            else
								MemberID=""
								chname=""
                            end if
                            rsMem.close
					'29  BILLMEMID1
					BILLMEMID1=MemberID
					'30  BILLMEM1
					BILLMEM1=chname
					'31  BILLMEMID2
					BILLMEMID2=""
					'32  BILLMEM2
					BILLMEM2=""
					'33  BILLMEMID3
					BILLMEMID3=""
					'34  BILLMEM3
					BILLMEM3=""
					'35  BILLFILLERMEMBERID
					BILLFILLERMEMBERID=MemberID
					'36  BILLFILLER
					BILLFILLER=chname
					'37  BILLFILLDATE
					BILLFILLDATE="To_Date('"&year(tempValue(3))&"/"&month(tempValue(3))&"/"&day(tempValue(3))&"','YYYY/MM/DD')"

					'38  DEALLINEDATE
					DEALLINEDATE="To_Date('"&tempValue(8)&"','YYYY/MM/DD')"

					NOTE="拖吊已結匯入" & Year(now) & Right("0"&Month(Now),2) & Right("0"&Day(now),2)
					'45  EQUIPMENTID
					EQUIPMENTID="-1"
					'46  RULEVER
					RULEVER="2"
					'47  OWNER
					OWNER=""
					'48  OWNERADDRESS
					OWNERADDRESS=""
					'49  OWNERID
					OWNERID=""
					'50  OWNERZIP
					OWNERZIP=""
					'51  DRIVERSEX
					DRIVERSEX=""
					'52  TRAFFICACCIDENTNO
					TRAFFICACCIDENTNO=""
					'53  TRAFFICACCIDENTTYPE
					TRAFFICACCIDENTTYPE=""
					'54  IMAGEFILENAME
					IMAGEFILENAME=""
					'55  DOUBLECHECKSTATUS
					DOUBLECHECKSTATUS="0"
					'56  IMAGEPATHNAME
					IMAGEPATHNAME="0"
					'57  BILLBASETYPEID
					BILLBASETYPEID="0"
					'58  BILLMEMID4
					BILLMEMID4=""
					'59  BILLMEM4
					BILLMEM4=""
					'60  IMAGEFILENAMEB
					IMAGEFILENAMEB=""
					'61  SIGNTYPE
					SIGNTYPE="A"

      			  	if UBound(tempValue)=17 then
      			  	  if trim(tempValue(2))="" then  
      			        response.write "第" & i  & "行: " & txtline & "<br>"
						response.flush
     			        Err= Err+1      
     			        i=i+1
     			      Else
				      
							'查流水號
							strSN="select BillBase_seq.nextval as SN from Dual"
							set rsSN=conn.execute(strSN)
							if not rsSN.eof then
								theSN=trim(rsSN("SN"))
							end if
							rsSN.close
							set rsSN=nothing

                        	set rsMem=Nothing

                          '---------------------------------------------------------------------------				
							'smith 判斷單號是否已經匯入過 start
							strSQL="select billno from BillBase where billno='" & BILLNO & "' and recordstateid<>-1 "
							set rsMatch=conn.execute(strSQL)	
						if rsMatch.eof and MemberID<>"" then 
							'smith end
							Sys_Now=DateAdd("s",1,Sys_Now)
							strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
									",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
									",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
									",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
									",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
									",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
									",BillFillerMemberID,BillFiller" &_
									",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
									",Note,EquipmentID,RuleVer,DriverSex,TrafficAccidentNo,TrafficAccidentType,SignType)" &_
									" values("&theSN&",'1','"&BillNo&"','"&CarNo&"',"&CarSimpleID&","&CarAddID&","&IllegalDate&"," & _
									"'"&IllegalAddressID&"','"&IllegalAddress&"','"&Rule1&"',"&IllegalSpeed &_
									","&RuleSpeed&","&ForFeit1&",'"&Rule2&"'" &_
									","&ForFeit2&",'"&Rule3&"',"&ForFeit3&",'"&Rule4&"'" &_
									","&ForFeit4&","&Insurance&","&UseTool&",'"&ProjectID&"'" &_
									",'"&DriverPID&"',"& DriverBirth &",'"&Driver&"'" &_
									",'"&DriverAddress&"','"&DriverZip&"','"&MemberStation&"'" &_
									",'"&BillUnitID&"','"&BillMemID1&"','"&BILLMEM1&"'" &_
									",'"&BillMemID2&"','"&BILLMEM2&"'" &_
									",'"&BillMemID3&"','"&BILLMEM3&"'" &_
									",'"&BillMemID4&"','"&BILLMEM4&"'" &_
									","&BillMemID1&",'"&BILLMEM1&"'" &_
									","&BillFillDate&","&DealLineDate&",'9',0,"&funGetDate(Sys_Now,1)&"," & Session("User_ID")  &_
									",'"&Note&"','"&EquipmentID&"','"&RuleVer&"'" &_
									",'"&DriverSex&"','"&TrafficAccidentNo&"','"&TrafficAccidentType&"','"&SignType&"'" &_
									")"
									conn.execute strInsert


								strInsCar="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
									",RecordMemberID,ExchangeDate,ExchangeTypeID,DCIwindowName,BatchNumber)"&_
									"values(DCILOG_SEQ.nextval,"&theSN&",'',2,'"&CarNo&"'" &_
									",'"&BillUnitID&"',"&funGetDate(Sys_Now,1)&","&Session("User_ID")&",sysdate,'A','Z','"&"A"&Right("0"&Year(date)-1911,3)&Right("0"&Month(date),2)&Right("0"&day(date),2)&"'" &_
									")" 
									conn.execute strInsCar
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
			  
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
     	
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