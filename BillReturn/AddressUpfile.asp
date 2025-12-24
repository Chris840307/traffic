<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<% 
''將現在的日期和時間轉為檔案名稱
on Error Resume Next
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
FileType=".txt"
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
		errBillNo=""
		fname = "tmpAddress.txt"
		file.SaveAs Server.mappath(UpFilePath&fname)   ''儲存檔案
		Set Fso=CreateObject("Scripting.FileSystemObject")
		Set f=Fso.OpenTextFile( server.mappath(UpFilePath&fname),1,True,0)
		While Not f.AtEndOfStream
			tempstr=f.ReadLine
			tempstr=tempstr
			tempstr=split(tempstr,",")
			If Ubound(tempstr)>0 Then
				If trim(tempstr(1))<>"" Then
					Sys_ZipID="":Sys_ZipName=""
					strSQL="select ZipID,ZipName from Zip where ZipName like '"&left(trim(tempstr(1)),6)&"%'"
					set rszip=conn.execute(strSQL)
					If Not rszip.eof Then
						Sys_ZipID=rszip("ZipID")
						Sys_ZipName=rszip("ZipName")
					else
						rszip.close
						strSQL="select ZipID,ZipName from Zip where ZipName like '%"&left(trim(tempstr(1)),3)&"%'"
						set rszip=conn.execute(strSQL)
						If Not rszip.eof Then
							Sys_ZipID=rszip("ZipID")
							Sys_ZipName=rszip("ZipName")
						end if
					end if
					rszip.close

					strSQL="Select BillTypeID from BillBase where BillNo='"&trim(Ucase(tempstr(0)))&"'"
					set rsbill=conn.execute(strSQL)
					If trim(rsbill("BillTypeID"))="1" Then
						strSQL="Update BillBaseDciReturn set DriverHomeZIP='"&trim(Sys_ZipID)&"',DriverHomeAddress='"&replace(trim(tempstr(1)),trim(Sys_ZipName),"")&"',DriverCounty='"&left(trim(Sys_ZipName),3)&"' where BillNo='"&trim(Ucase(tempstr(0)))&"'"
						conn.execute(strSQL)
					else
						strSQL="Update BillBaseDciReturn set OwnerZip='"&trim(Sys_ZipID)&"',OwnerAddress='"&replace(trim(tempstr(1)),trim(Sys_ZipName),"")&"',OwnerCounty='"&left(trim(Sys_ZipName),3)&"' where BillNo='"&trim(Ucase(tempstr(0)))&"'"
						conn.execute(strSQL)
					end if
					rsbill.close
				else
					If errBillNo<>"" Then errBillNo=errBillNo&"\n"
					errBillNo=errBillNo&trim(tempstr(0))
				end if
			elseif Ubound(tempstr)=0 or Ubound(tempstr)>1 then
				If errBillNo<>"" Then errBillNo=errBillNo&"\n"
				errBillNo=errBillNo&trim(tempstr(0))
			end if
		wend
		Response.write "<script>"
		Response.Write "alert('"&file.FilePath&file.FileName&" ("&file.FileSize&") => 匯入成功!');"
		If errBillNo<>"" Then Response.Write "alert('以下為錯誤的單號\n"&errBillNo&" ');"
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