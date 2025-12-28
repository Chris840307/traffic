<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<% 
''將現在的日期和時間轉為檔案名稱
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
<!--#include FILE="../Common/wang_upload_5xsoft.inc"-->
<%
dim upload,file,formName,FileType
set upload=new upload_5xSoft ''建立上傳對象
upload.ProgressID=Request("progressID") '一定是這行在先
upload.GetUpFile
UpFilePath="./Picture/"

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

changDir=0
If sys_City = "台中市" Then
	changDir=1
End if 

FileType=".jpg@.jpeg@.tif@.pdf"

'fileStr=split("@.jpg@.jpeg@.tif@.pdf","@")
cnt_file=0

errmsg=""

for each formName in upload.objFile ''列出所有上傳的檔案
	TypeFlag=0
	ExName=""

	cnt_file=cnt_file+1

	set file=upload.objFile(formName)  ''生成一個檔案對象
	
	If file.FileName <>"" Then

		if Instr(FileType,GetExtendName(file.FileName)) then
			TypeFlag = 1  '檔案為允許的類型
			ExName=GetExtendName(file.FileName)

		else
			TypeFlag = 0		'檔案為不允許的類型
			'Response.write "<script>"
			'Response.Write "alert('不支援您所上傳的檔案類型："&GetExtendName(file.FileName)&"');"
			'response.write "self.close();"
			'Response.write "</script>"
			'exit for

			errmsg=errmsg&"不支援您所上傳的檔案類型："&GetExtendName(file.FileName)&"\n"
		end if
	End if 
	
	if TypeFlag = 1 then
		fname = "PassersEndArrived_"&trim(Session("Map_SN"))&"."&ExName

		If changDir = 1 Then
				
			file.SaveAs Server.mappath("./tmpCreditor/"&fname)   ''儲存檔案
		else
			
			file.SaveAs Server.mappath(UpFilePath&fname)  '儲存檔案
		End if 

		
		strSQL="Update PassersEndArrived set Imagedirectoryname='"&UpFilePath&"',Imagefilename='"&fname&"' where SN="&trim(Session("Map_SN"))
		conn.execute(strSQL)

		If changDir = 1 Then

			set fso=Server.CreateObject("Scripting.FileSystemObject")

			fso.CopyFile Server.mappath("./tmpCreditor/"&fname), Server.mappath(UpFilePath&fname), True
			fso.DeleteFile Server.mappath("./tmpCreditor/"&fname)
			set fso=nothing
		end If 

		Response.write "<script>"
		Response.Write "alert('"&file.FilePath&file.FileName&" ("&file.FileSize&") => 上傳檔案成功!');"
		response.write "opener.myForm.submit();"
		response.write "self.close();"
		Response.write "</script>"
		'fileStr = fileStr & "<img src='pic/addon.gif'><a href='"& UpFilePath&file.FileName&"' target='_blank'>查看上傳的檔案﹕<font color='red'>" & file.FileName &"</font> ("& file.FileSize &" kb)</a><br>"
		'FileNameStr = UpFilePath&fname
	end if
next

If errmsg <>"" Then
	Response.write "<script>"
	Response.Write "alert("""& errmsg &""");"
	response.write "opener.myForm.submit();"
	response.write "self.close();"
	Response.write "</script>"
End if 

set upload=nothing  ''刪除此對象

function GetExtendName(FileName)
dim ExtName
ExtName = LCase(FileName)
ExtName = right(ExtName,7)
ExtName = right(ExtName,7-Instr(ExtName,"."))
GetExtendName = ExtName
end function
%>