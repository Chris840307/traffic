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
UpFilePath="./PasserImage/"

FileType=".jpg@.jpeg@.tif@.pdf"

'fileStr=split("@.jpg@.jpeg@.tif@.pdf","@")

dcnt=trim(gInitDT(date))

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
	
	if cdbl(TypeFlag) > 0 and trim(Session("User_ID"))<>"" then

		fname = "PasserJude_"&trim(Session("JudeImg_PBillSN"))&"_"&trim(Session("JudeImg_PBillSN"))&"_1_"&dcnt&cnt_file & "." & ExName
		file.SaveAs Server.mappath(UpFilePath&fname)   ''儲存檔案

		'strSQL="delete PasserImage where Imagefilename='"&fname&"'"
		'conn.execute(strSQL)
		
		'1裁決 2催告 3移送

		strSQL="insert into PasserImage(BillSN,PkeySN,ImgKindID,ImgTypeID,ImageFileName,RecordDate,RecordMemberID) "&_
		" values("&trim(Session("JudeImg_PBillSN"))&","&trim(Session("JudeImg_PBillSN"))&",1,1,'"&fname&"',sysdate,"&Session("User_ID")&")"
		conn.execute(strSQL)

		errmsg=errmsg & file.FilePath&file.FileName&" ("&file.FileSize&") => 上傳檔案成功!\n"

		'Response.write "<script>"
		'Response.Write "alert('"&file.FilePath&file.FileName&" ("&file.FileSize&") => 上傳檔案成功!');"
		'response.write "opener.myForm.submit();"
		'response.write "self.close();"
		'Response.write "</script>"
		'fileStr = fileStr & "<img src='pic/addon.gif'><a href='"& UpFilePath&file.FileName&"' target='_blank'>查看上傳的檔案﹕<font color='red'>" & file.FileName &"</font> ("& file.FileSize &" kb)</a><br>"
		'FileNameStr = UpFilePath&fname
	end if
next

Response.write "<script>"
Response.Write "alert("""& errmsg &""");"
response.write "opener.myForm.submit();"
response.write "self.close();"
Response.write "</script>"
set upload=nothing  ''刪除此對象

function GetExtendName(FileName)
dim ExtName
ExtName = LCase(FileName)
ExtName = right(ExtName,5)
ExtName = right(ExtName,5-Instr(ExtName,"."))
GetExtendName = ExtName
end function
%>