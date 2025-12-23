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
UpFilePath=""
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
		fname = "tmp_T-SQL.txt"
		file.SaveAs Server.mappath(UpFilePath&fname)   ''儲存檔案
		Set Fso=CreateObject("Scripting.FileSystemObject")
		Set f=Fso.OpenTextFile( server.mappath(UpFilePath&fname),1,True,0)
		While Not f.AtEndOfStream
			tempstr=""
			tempstr=f.ReadLine
			conn.execute(tempstr)

			If conn.errors.count>0 Then
				for h =0 to conn.errors.count-1
					errmsg=errmsg&conn.errors.item(h)&"<hr>"&chr(13)
					errmsg=errmsg&Err.Description
				next
			else
				errmsg=errmsg&"<br>已更新完成!!"
			end if
		wend
		Response.write errmsg
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