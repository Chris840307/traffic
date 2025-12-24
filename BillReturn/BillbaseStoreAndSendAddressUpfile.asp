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
		fname = "tmpStoreAndSendAddress.xls"
		file.SaveAs Server.mappath(UpFilePath&fname) 

		Set ConnEx = Server.CreateObject("ADODB.Connection")
		Driver = "Provider=Microsoft.Jet.OLEDB.4.0;" &_ 
			"Data Source="&Server.MapPath(UpFilePath&fname) &_ 
			";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"

		ConnEx.Open Driver
		
		strSQL="Select * From [Sheet1$]"
		Set Rs=ConnEx.Execute(strSQL)
		sys_err="":cmt=0:filecmt=1

		Response.Write "<script language=""JavaScript"">"

		While not rs.eof
			filecmt=filecmt+1
			If (not ifnull(rs("單號"))) and (not ifnull(rs("戶籍地址"))) Then
				Response.Write "window.opener.myForm.item["&cmt&"].value='"&trim(rs("單號"))&"';"& vbcrlf
				Response.Write "window.opener.myForm.CarNo["&cmt&"].value='"&trim(rs("車號"))&"';"& vbcrlf
				Response.Write "window.opener.myForm.OwnerAddress["&cmt&"].value='"&trim(rs("戶籍地址"))&"';"& vbcrlf

				
				strSQL = "select ZipID,ZipName from Zip where ZipName like '"&replace(left(trim(rs("戶籍地址")),5),"臺","台")&"%'"

				set rscnt=conn.execute(strSQL)
				if Not rscnt.eof then
					Response.Write "window.opener.myForm.OwnerZip["&cmt&"].value='"&trim(rscnt("ZipID"))&"';"& vbcrlf
					Response.Write "window.opener.myForm.Sys_ZipName["&cmt&"].value='"&trim(rscnt("ZipName"))&"';"& vbcrlf

				else
					rscnt.close
		
					strSQL = "select ZipID,ZipName from Zip where ZipName like '"&replace(left(trim(rs("戶籍地址")),3),"臺","台")&"%'"
					set rscnt=conn.execute(strSQL)
					if Not rscnt.eof then
						Response.Write "window.opener.myForm.OwnerZip["&cmt&"].value='"&trim(rscnt("ZipID"))&"';"& vbcrlf
						Response.Write "window.opener.myForm.Sys_ZipName["&cmt&"].value='"&trim(rscnt("ZipName"))&"';"& vbcrlf
					else
						sys_err=sys_err&"資料第"&filecmt&"行，找不到郵遞區號，請確認!!\n"
					end if

					rscnt.close
				end if

				If cmt>=119 and ((abs(cmt-119) mod 30)=0) then Response.Write "window.opener.insertRow(window.opener.document.all.fmyTable);"& vbcrlf

				cmt=cmt+1
			elseIf (not ifnull(rs("單號"))) and (ifnull(rs("戶籍地址"))) Then
				sys_err=sys_err&"資料第"&filecmt&"行，無戶籍地址，請確認!!\n"
			End if
			rs.movenext
		Wend
		rs.close

		Response.Write "alert('匯入成功!');"
		If sys_err<>"" Then Response.Write "alert('以下為錯誤的案件\n"&sys_err&"');"
		response.write "self.close();"& vbcrlf
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