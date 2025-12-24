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
end function
%>
<!--#include FILE="../Common/upload_5xsoft.inc"-->
<%
dim upload,file,formName,FileType
set upload=new upload_5xSoft ''建立上傳對象
UpFilePath="./Upaddress/"
FileType=".xls"
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
		fname = "tmpAccept.xls"
		file.SaveAs Server.mappath(UpFilePath&fname) 

		Set ConnEx = Server.CreateObject("ADODB.Connection")
		Driver = "Provider=Microsoft.Jet.OLEDB.4.0;" &_ 
			"Data Source="&Server.MapPath(UpFilePath&fname) &_ 
			";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
		ConnEx.Open Driver
		
		strSQL="Select * From [Sheet1$] order by 違反交通管理通知單字號"
		Set Rs=ConnEx.Execute(strSQL)
		sys_err="":cmt=0
		Sys_now=funGetDate(now,1)
		While not rs.eof
			Sys_BillNo="":Sys_illegalDate="":Sys_CarNo="":Sys_Rule1="":Sys_Driver=""
			Sys_FastenerTypeID="":Sys_UnitID="":Sys_BillMemID1="":sys_chname=""
			If instr(trim(rs("編號")),"：")=0 Then
				If instr(rs("車牌號碼"),"-")>0 Then Sys_CarNo=trim(rs("車牌號碼"))

				strSQL="select count(1) cmt from (select chname,unitid from memberdata where Loginid='"&trim(rs("舉發員警代碼"))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo where UnitTypeid=(select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"')) b where a.unitid=b.unitid"

				set rscnt=conn.execute(strSQL)

				If cdbl(rscnt("cmt"))=1 Then
					
					strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(rs("舉發員警代碼"))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo where UnitTypeid=(select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"')) b where a.unitid=b.unitid"

					set rsmen=conn.execute(strSQL)

					If not rsmen.eof Then
						Sys_BillMemID1=trim(rsmen("memberid"))
						Sys_UnitID=trim(rsmen("UnitID"))
						sys_chname=trim(rsmen("UnitName"))&":"&trim(rsmen("chname"))
					end if

					rsmen.close
					
				End if
				rscnt.close

				If not ifnull(trim(rs("扣件"))) Then
					If instr("1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L",trim(rs("扣件")))>0 Then
						Sys_FastenerTypeID=trim(rs("扣件"))
					else
						Sys_FastenerTypeID="null"
					End if
				End if
				

				If trim(Sys_CarNo)<>"" and trim(sys_BillMemID1)<>"" and Sys_FastenerTypeID<>"null" Then

					strSQL="select count(1) cmt from BillStopCarAccept where BillNo='"&trim(rs("違反交通管理通知單字號"))&"' and recordstateid=0"
			
					set rsnt=conn.execute(strSQL)

					If cdbl(rsnt("cmt"))=0 Then
					
						strSQL="insert into BillStopCarAccept(Billno,Carno,IllegalDate,Rule1,Driver,FastenerTypeID1,BillUnitID,BillMemID1,Acceptdate,RecordStateID,RecordDate) values('"&trim(rs("違反交通管理通知單字號"))&"','"&trim(rs("車牌號碼"))&"',"&funGetDate(gOutDT(rs("違規日期")),1)&",'"&trim(rs("違反條款代碼"))&"','"&trim(rs("違規人姓名"))&"','"&trim(Sys_FastenerTypeID)&"','"&trim(Sys_UnitID)&"','"&trim(Sys_BillMemID1)&"',"&funGetDate(date,0)&",0,"&Sys_now&")"

						conn.execute(strSQL)
					
					else

						strSQL="Update BillStopCarAccept set CarNo='"&trim(rs("車牌號碼"))&"',BillUnitID='"&Sys_UnitID&"',IllegalDate="&funGetDate(gOutDT(rs("違規日期")),1)&",AcceptDate="&funGetDate(date,0)&",Rule1='"&trim(rs("違反條款代碼"))&"',Driver='"&trim(rs("違規人姓名"))&"',FastenerTypeID1='"&Sys_FastenerTypeID&"',BILLMEMID1="&trim(Sys_BillMemID1)&",RECORDDATE="&Sys_now&" where billno='"&trim(rs("違反交通管理通知單字號"))&"' and recordstateid=0"

						conn.execute(strSQL)

					end if

	'				Response.write "<script>"
	'				Response.Write "window.opener.myForm.item["&cmt&"].value='"&trim(rs("違反交通管理通知單字號"))&"';"
	'				
	'				Response.Write "window.opener.myForm.CarNo["&cmt&"].value='"&trim(rs("車牌號碼"))&"';"
	'
	'				Response.Write "window.opener.myForm.illegalDate["&cmt&"].value='"&trim(rs("違規日期"))&"';"
	'
	'				Response.Write "window.opener.myForm.Rule1["&cmt&"].value='"&trim(rs("違反條款代碼"))&"';"
	'
	'				Response.Write "window.opener.myForm.DriverName["&cmt&"].value='"&trim(rs("違規人姓名"))&"';"
	'
	'				Response.Write "window.opener.myForm.Fastener1["&cmt&"].value='"&trim(Sys_FastenerTypeID)&"';"
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
				else
					sys_err=sys_err&"編號："&trim(rs("編號"))

					If ifnull(Sys_CarNo) Then
						sys_err=sys_err&"車號有誤，請確認!!。\n"

					elseif ifnull(sys_BillMemID1) Then
						sys_err=sys_err&"人員代碼重覆或無此人員，請確認!!。\n"

					elseif Sys_FastenerTypeID="null" Then
						sys_err=sys_err&"扣件代碼輸入不正確，請確認!!。\n"

					end if

				End if
			end if
			rs.movenext
		Wend
		rs.close

		
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