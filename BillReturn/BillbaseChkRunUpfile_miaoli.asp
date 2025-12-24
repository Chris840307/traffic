<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<% 
''將現在的日期和時間轉為檔案名稱
'on Error Resume Next
function makefilename(fname)
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"上午","")
  fname = replace(fname,"下午","")
  makefilename=fname
end Function 

Function ChkNum(strValue)
	if ISNull(strValue) or trim(strValue)="" or IsEmpty(strValue) then
		ChkNum="null"
	else
		ChkNum=strValue
	end if
End Function
%>
<!--#include FILE="../Common/upload_5xsoft.inc"-->
<%
dim upload,file,formName,FileType
set upload=new upload_5xSoft ''建立上傳對象
UpFilePath="./Upaddress/"
FileType=".xls"

strCity="select value from Apconfigure where id=3"
set rsCity=conn.execute(strCity)
sys_RuleVer=trim(rsCity("value"))
rsCity.close

for each formName in upload.file ''列出所有上傳的檔案
	set file=upload.file(formName)  ''生成一個檔案對象
	if Instr(FileType,GetExtendName(file.FileName)) then
		TypeFlag = 1        '檔案為允許的類型
	else
		TypeFlag = 0		'檔案為不允許的類型
		Response.write "<script>"
		Response.Write "alert('不支援您所上傳的檔案類型："&GetExtendName(file.FileName)&"');"
		response.write "self.close();"
		Response.write "</script>"
		exit for
	end if
	if TypeFlag = 1 then

		fname = "tmpRunAccept.xls"
		file.SaveAs Server.mappath(UpFilePath&fname) 

		Set ConnEx = Server.CreateObject("ADODB.Connection")
		Driver = "Provider=Microsoft.Jet.OLEDB.4.0;" &_ 
			"Data Source="&Server.MapPath(UpFilePath&fname) &_ 
			";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
		ConnEx.Open Driver
		
		strCity="select value from Apconfigure where id=31"
		set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
		rsCity.close

		sys_err="":cmt=0

		fileName=split(file.FileName,".")

		CarCode=",1,2,3,4,6,"

		strSQL="select ID,Content from DCICode where typeid=6 order by ID"
		set rscode=conn.execute(strSQL)
		While not rsCode.eof

			FastCode=FastCode&rscode("ID")&","
						
			rsCode.movenext
		Wend

		FastCode=","&FastCode

		rsCode.close

		Sys_now=funGetDate(now,1)

		strSQL="Select * From ["&fileName(0)&"$] order by 序,車號"
		Set Rs=ConnEx.Execute(strSQL)
		While not rs.eof
			Sys_BillNo="":Sys_illegalDate="":Sys_CarNo="":tmp_Rule="":Sys_Rule1="":Sys_Rule2="":Sys_IllegalAddress=""
			Sys_IllegalAddressID="null":Sys_UnitID="":Sys_BillMemID1="":Sys_BillMemID2="":Sys_BillMemID3="":Sys_BillMemID4="":sys_chname=""
			Sys_BillMem=""

			err=0
			If not ifnull(rs("車號")) Then Sys_CarNo=trim(rs("車號"))

			Sys_BillMem=split(trim(rs("舉發人"))&"",",")

			If Ubound(Sys_BillMem) >= 0 Then
				strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMem(0))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

				set rsmen=conn.execute(strSQL)

				If not rsmen.eof Then
					Sys_BillMemID1=trim(rsmen("memberid"))
					Sys_UnitID=trim(rsmen("UnitID"))
					sys_chname=trim(rsmen("UnitName"))&":"&trim(rsmen("chname"))
				end if

				rsmen.close
			End if
			
			If Ubound(Sys_BillMem) >= 1 Then
				strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMem(1))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

				set rsmen=conn.execute(strSQL)

				If not rsmen.eof Then
					Sys_BillMemID2=trim(rsmen("memberid"))
				end if

				rsmen.close
			end If

			If Ubound(Sys_BillMem) >= 2 Then
				strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMem(2))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

				set rsmen=conn.execute(strSQL)

				If not rsmen.eof Then
					Sys_BillMemID3=trim(rsmen("memberid"))
				end if

				rsmen.close
			end If 
			
			If Ubound(Sys_BillMem) >= 3 Then
				strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMem(3))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

				set rsmen=conn.execute(strSQL)

				If not rsmen.eof Then
					Sys_BillMemID4=trim(rsmen("memberid"))
				end if

				rsmen.close
			end if

			If trim(Sys_CarNo)<>"" then			

				if ifnull(rs("違規時間")) Then
					err=1
					sys_err=sys_err&"序："&trim(rs("序"))
					sys_err=sys_err&"違規時間有誤，請確認!!。\n"

				elseif cdbl(left(rs("違規時間"),2))>23 or cdbl(right(rs("違規時間"),2))>59 Then
					err=1
					sys_err=sys_err&"序："&trim(rs("序"))
					sys_err=sys_err&"違規時間有誤，請確認!!。\n"

				end If 

'				if ifnull(rs("車種")) Then
'					err=1
'					sys_err=sys_err&"序："&trim(rs("序"))
'					sys_err=sys_err&"車種有誤，請確認!!。\n"
'
'				end If 
				
				if (not ifnull(rs("車種"))) and (instr(CarCode,rs("車種"))<1) Then
					err=1
					sys_err=sys_err&"序："&trim(rs("序"))
					sys_err=sys_err&"車種有誤，請確認!!。\n"

				end If 

'				if ifnull(rs("違規地點")) Then
'					err=1
'					sys_err=sys_err&"序："&trim(rs("序"))
'					sys_err=sys_err&"違規地點有誤，請確認!!。\n"
'
'				end If

				if ifnull(rs("違反條例")) Then
					err=1
					sys_err=sys_err&"序："&trim(rs("序"))
					sys_err=sys_err&"違反條例有誤，請確認!!。\n"

				end If

'				if ifnull(Sys_UnitID) Then
'					err=1
'					sys_err=sys_err&"序："&trim(rs("序"))
'					sys_err=sys_err&"舉發人員1有誤，請確認!!。\n"
'
'				end If 
'				
'
'				if (Ubound(Sys_BillMem) >= 1) and (ifnull(Sys_BillMemID2)) Then
'					err=1
'					sys_err=sys_err&"序："&trim(rs("序"))
'					sys_err=sys_err&"舉發人員2有誤，請確認!!。\n"
'
'				end If 
'
'				if (Ubound(Sys_BillMem) >= 2) and (ifnull(Sys_BillMemID3)) Then
'					err=1
'					sys_err=sys_err&"序："&trim(rs("序"))
'					sys_err=sys_err&"舉發人員3有誤，請確認!!。\n"
'
'				end If 
'
'				if (Ubound(Sys_BillMem) >= 3) and (ifnull(Sys_BillMemID4)) Then
'					err=1
'					sys_err=sys_err&"序："&trim(rs("序"))
'					sys_err=sys_err&"舉發人員4有誤，請確認!!。\n"
'
'				end If 

			end if


			If trim(Sys_CarNo)<>"" and err=0 Then

				Sys_illegalDate=gOutDT(left(rs("違規時間"),7))&" "&left(right(rs("違規時間"),4),2)&":"&right(right(rs("違規時間"),4),2)

				tmp_Rule=split(rs("違反條例")&"",",")

				if Ubound(tmp_Rule) >= 0 then Sys_Rule1=tmp_Rule(0)
				if Ubound(tmp_Rule) >= 1 then Sys_Rule2=tmp_Rule(1)


				sys_illagalAddr=replace(trim(rs("違規地點"))&"",sys_City,"")
				tmpTown=""
				if instr(sys_illagalAddr,"聯絡")<=0 then
					strSQL="select zipName from Zip where zipName like '"&sys_City&"%'"
					set rszip=conn.execute(strSQL)
					
					while not rszip.eof
						If instr(sys_illagalAddr,replace(rszip("zipName"),sys_City,"")) >0 Then
							tmpTown=replace(rszip("zipName"),sys_City,"")
							sys_illagalAddr=replace(sys_illagalAddr,tmpTown,"")
						end if
						rszip.movenext
					wend
					rszip.close
				end if

				sys_illagalAddr=tmpTown&sys_illagalAddr

				strSQL="select count(1) cmt from BillRunCarAccept where CarNo='"&Sys_CarNo&"' and IllegalDate="&funGetDate(Sys_illegalDate,1)&" and Rule1='"&trim(Sys_Rule1)&"' and recordstateid=0"
		
				set rsnt=conn.execute(strSQL)

				If cdbl(rsnt("cmt"))=0 Then
				
					'strSQL="insert into BillRunCarAccept(CARNO,CARSIMPLEID,BILLUNITID,ILLEGALDATE,ACCEPTDATE,RULE1,RULE2,ILLEGALADDRESS,ILLEGALSPEED,RULESPEED,BILLMEMID1,BILLMEMID2,BILLMEMID3,IMAGEFILE,PICTUREFILE,InformationData,RULEVER,RECORDSTATEID,RECORDMEMBERID1,RECORDDATE) values('"&Sys_CarNo&"',"&ChkNum(rs("車種"))&",'"&Sys_UnitID&"',"&funGetDate(Sys_illegalDate,1)&","&funGetDate(date,0)&",'"&trim(Sys_Rule1)&"','"&trim(Sys_Rule2)&"','"&trim(rs("違規地點"))&"',"&ChkNum(rs("測速"))&","&ChkNum(rs("限速"))&","&ChkNum(Sys_BillMemID1)&","&ChkNum(Sys_BillMemID2)&","&ChkNum(Sys_BillMemID3)&","&ChkNum(rs("圖"))&","&ChkNum(rs("相"))&","&ChkNum(rs("資"))&",'"&trim(sys_RuleVer)&"',0,"&Session("User_ID")&",sysdate)"

					strSQL="insert into BillRunCarAccept(CARNO,BILLUNITID,ILLEGALDATE,ACCEPTDATE,RULE1,RULE2,ILLEGALADDRESS,ILLEGALSPEED,RULESPEED,BILLMEMID1,BILLMEMID2,BILLMEMID3,IMAGEFILE,PICTUREFILE,InformationData,RULEVER,RECORDSTATEID,RECORDMEMBERID1,RECORDDATE,RecordMemberID2,RecordDate2,RecordMemberID3,RecordDate3,COMPANYMEMBERID,COMPANYACCEPTDATE) values('"&Ucase(Sys_CarNo)&"','"&Sys_UnitID&"',"&funGetDate(Sys_illegalDate,1)&","&funGetDate(date,0)&",'"&trim(Sys_Rule1)&"','"&trim(Sys_Rule2)&"','"&replace(sys_illagalAddr,"'","")&"',"&ChkNum(rs("測速"))&","&ChkNum(rs("限速"))&","&ChkNum(Sys_BillMemID1)&","&ChkNum(Sys_BillMemID2)&","&ChkNum(Sys_BillMemID3)&","&ChkNum(rs("圖"))&","&ChkNum(rs("相"))&","&ChkNum(rs("資"))&",'"&trim(sys_RuleVer)&"',0,"&Session("User_ID")&",sysdate,"&Session("User_ID")&",sysdate,"&Session("User_ID")&",sysdate,"&Session("User_ID")&",sysdate)"

					conn.execute(strSQL)
				
				else

					'strSQL="Update BillRunCarAccept set CARNO='"&Sys_CarNo&"',CARSIMPLEID="&ChkNum(rs("車種"))&",BILLUNITID='"&Sys_UnitID&"',ILLEGALDATE='"&funGetDate(Sys_illegalDate,1)&"',ACCEPTDATE="&funGetDate(date,0)&",RULE1='"&trim(Sys_Rule1)&"',RULE2='"&trim(Sys_Rule2)&"',ILLEGALADDRESS='"&trim(rs("違規地點"))&"',ILLEGALSPEED="&ChkNum(rs("測速"))&",RULESPEED="&ChkNum(rs("限速"))&",BILLMEMID1="&ChkNum(Sys_BillMemID1)&",BILLMEMID2="&ChkNum(Sys_BillMemID2)&",BILLMEMID3="&ChkNum(Sys_BillMemID3)&",IMAGEFILE="&ChkNum(rs("圖"))&",PICTUREFILE="&ChkNum(rs("相"))&",InformationData="&ChkNum(rs("資"))&",RECORDMEMBERID1="&Session("User_ID")&",RECORDDATE=sysdate where CarNo='"&Sys_CarNo&"' and IllegalDate="&funGetDate(Sys_illegalDate,1)&" and Rule1='"&trim(Sys_Rule1)&"' and recordstateid=0"

					strSQL="Update BillRunCarAccept set CARNO='"&Ucase(Sys_CarNo)&"',BILLUNITID='"&Sys_UnitID&"',ILLEGALDATE="&funGetDate(Sys_illegalDate,1)&",ACCEPTDATE="&funGetDate(date,0)&",RULE1='"&trim(Sys_Rule1)&"',RULE2='"&trim(Sys_Rule2)&"',ILLEGALADDRESS='"&replace(sys_illagalAddr,"'","")&"',ILLEGALSPEED="&ChkNum(rs("測速"))&",RULESPEED="&ChkNum(rs("限速"))&",BILLMEMID1="&ChkNum(Sys_BillMemID1)&",BILLMEMID2="&ChkNum(Sys_BillMemID2)&",BILLMEMID3="&ChkNum(Sys_BillMemID3)&",IMAGEFILE="&ChkNum(rs("圖"))&",PICTUREFILE="&ChkNum(rs("相"))&",InformationData="&ChkNum(rs("資"))&",RECORDMEMBERID1="&Session("User_ID")&",RECORDDATE=sysdate,RecordMemberID2="&Session("User_ID")&",RecordDate2=sysdate,RecordMemberID3="&Session("User_ID")&",RecordDate3=sysdate,COMPANYMEMBERID="&Session("User_ID")&",COMPANYACCEPTDATE=sysdate where CarNo='"&Sys_CarNo&"' and IllegalDate="&funGetDate(Sys_illegalDate,1)&" and Rule1='"&trim(Sys_Rule1)&"' and recordstateid=0"

					conn.execute(strSQL)

				end if

'				Response.write "<script>"
'				Response.Write "window.opener.myForm.item["&cmt&"].value='"&trim(rs("標示單號碼"))&"';"
'				
'				Response.Write "window.opener.myForm.CarNo["&cmt&"].value='"&trim(rs("車號"))&"';"
'
'				Response.Write "window.opener.myForm.illegalDate["&cmt&"].value='"&trim(rs("違規日期"))&"';"
'
'				Response.Write "window.opener.myForm.Rule1["&cmt&"].value='"&trim(rs("違反條款代碼"))&"';"
'
'				Response.Write "window.opener.myForm.IllegalAddress["&cmt&"].value='"&trim(rs("違規地點"))&"';"
'
'				Response.Write "window.opener.myForm.BillMemName["&cmt&"].value='"&trim(rs("舉發員警代碼"))&"';"
'
'				Response.Write "window.opener.myForm.BillMemID1["&cmt&"].value='"&trim(Sys_BillMemID1)&"';"
'
'				Response.Write "window.opener.myForm.BillUnitID["&cmt&"].value='"&trim(Sys_UnitID)&"';"
'
'				Response.Write "window.opener.BillMemName1["&cmt&"].innerHTML=""<font size=2>"&sys_chname&"</font>"";"
'
'				Response.write "</script>"
'				cmt=cmt+1

			End if
			rs.movenext
		Wend
		rs.close

		strSQL="delete BillRunCarAccept where RecordDate <="&funGetDate(dateadd("d",-20,now),1)

		conn.execute(strSQL)
		
		Response.write "<script>"
		Response.Write "window.opener.myForm.submit();"
		Response.Write "alert('"&file.FilePath&file.FileName&" ("&file.FileSize&") => 匯入成功!');"
		If sys_err<>"" Then Response.Write "alert('以下為錯誤的案件\n"&sys_err&"');"
		response.write "self.close();"
		Response.write "</script>"
	end if
next
set upload=nothing  ''刪除此對象

function GetExtendName(FileName)
dim ExtName
ExtName = LCase(FileName)
ExtName = right(ExtName,3)
ExtName = right(ExtName,3-Instr(ExtName,"."))
GetExtendName = ExtName
end function
%>