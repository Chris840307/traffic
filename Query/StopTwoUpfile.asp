<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<% 
''將現在的日期和時間轉為檔案名稱
on Error Resume Next
Server.ScriptTimeout = 60000
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
UpFilePath="./StopExpree/"
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
		errBillNo="":PBillNo1="":PCarNo1="":PBillNo2="":PCarNo2=""
		fname = "tmpAddress.txt"
		file.SaveAs Server.mappath(UpFilePath&fname)   ''儲存檔案
		Set Fso=CreateObject("Scripting.FileSystemObject")
		Set f=Fso.OpenTextFile( server.mappath(UpFilePath&fname),1,True,0)
		fileCount=0
		While Not f.AtEndOfStream
			tempstr=f.ReadLine
			If Not ifnull(tempstr) Then
				fileCount=fileCount+1
				If instr(PBillNo1,trim(right(tempStr,18)))=0 and instr(PBillNo2,trim(right(tempStr,18)))=0 Then
					Sys_Addr_1="":Sys_Addr_2=""
					Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""
					strSQL="select CarNo,Owner,DriverHomeAddress,OwnerZip,OwnerAddress,OwnerNotIfyAddress,DriverHomeZip,DriverHomeAddress from BillbaseDCIReturn where CarNo='"&trim(left(tempStr,8))&"' and ExchangetypeID='A'"
					
					set rsDci=conn.execute(strSQL)
					if Not rsDci.eof then
						Sys_CarNo=trim(rsDci("CarNo"))
						Sys_Owner=trim(rsDci("Owner"))
						If not ifnull(trim(rsDci("OwnerNotIfyAddress"))) Then
							Sys_OwnerAddress=trim(rsDci("OwnerNotIfyAddress"))
							Sys_OwnerZip=getzip(rsDci("OwnerNotIfyAddress"))
						else
							Sys_OwnerAddress=trim(rsDci("OwnerAddress"))
							Sys_OwnerZip=trim(rsDci("OwnerZip"))
						End if

						strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
						set rszip=conn.execute(strSQL)
						if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
						rszip.close
						Sys_Addr_1=Sys_OwnerZip&Sys_OwnerAddress

						Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""

						Sys_OwnerZip=trim(rsDci("DriverHomeZip"))
						Sys_OwnerAddress=trim(rsDci("DriverHomeAddress"))

						strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
						set rszip=conn.execute(strSQL)
						if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
						rszip.close

						Sys_Addr_2=Sys_OwnerZip&Sys_OwnerAddress

						If trim(Sys_Addr_1) <> trim(Sys_Addr_2) and (not ifnull(Sys_Addr_2)) Then
							If Not ifnull(PBillNo1) Then
								PBillNo1=PBillNo1&","
								PCarNo1=PCarNo1&","
							end if
							PBillNo1=PBillNo1&trim(right(tempStr,18))
							PCarNo1=PCarNo1&trim(left(tempStr,8))
						else
							If Not ifnull(PBillNo2) Then
								PBillNo2=PBillNo2&","
								PCarNo2=PCarNo2&","
							end if
							PBillNo2=PBillNo2&trim(right(tempStr,18))
							PCarNo2=PCarNo2&trim(left(tempStr,8))
						End if
					end if
					rsDci.close
				end if
			End if
		wend
		Response.write "<script>"
		Response.Write "alert('"&file.FilePath&file.FileName&" ("&file.FileSize&") => 匯入成功!');"
		Response.Write "opener.myForm.PBillNo.value='"&PBillNo1&"';"
		Response.Write "opener.myForm.PCarNo.value='"&PCarNo1&"';"
		Response.Write "opener.myForm.PBillNo2.value='"&PBillNo2&"';"
		Response.Write "opener.myForm.PCarNo2.value='"&PCarNo2&"';"
		Response.Write "opener.myForm.BillPrintKind.value='2';"
		Response.Write "opener.strCount.innerHTML='( 查詢 "&fileCount&" 筆紀錄 , "&fileCount&"筆成功)';"
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