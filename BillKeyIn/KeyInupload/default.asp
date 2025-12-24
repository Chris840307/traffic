<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/BannernoData.asp"-->
<!--#include virtual="traffic/Common/css.txt"-->
<%
DB_Selt	= trim(request("DB_Selt"))

	TodayFolder = year(now)-1911 & right("00"&month(now),2) & right("00"&day(now),2)
	
	'影像存放位置
	FileLocation=""
	strfile="select Value from ApConfigure where ID=110"
	set rsFile=conn.execute(strfile)
	if not rsFile.eof then
		FileLocation=trim(rsFile("Value"))
	end if
	rsFile.close
	set rsFile=nothing
	
	FileLocation=FileLocation&"Type8\"
	dim fso1 
	set fso1=Server.CreateObject("Scripting.FileSystemObject")
	if (fso1.FolderExists(FileLocation&Session("Credit_ID")))=false then
		fso1.CreateFolder FileLocation&Session("Credit_ID")
	end if

application("sessionID")=Session.SessionID
application("uploadpath")="Userfiles\" &Session("Credit_ID")
dim i
dim fs
set fs=Server.CreateObject("Scripting.FileSystemObject")
'資料夾要開權限

tmpDir   = "d:\Inetpub\wwwroot\Traffic\BillKeyIn\KeyInupload\Userfiles\" &Session("Credit_ID")
tmpDirTo = FileLocation&Session("Credit_ID")&"\"&TodayFolder

	if (fso1.FolderExists(tmpDir))=false then
		fso1.CreateFolder tmpDir
	end if
	
	if (fso1.FolderExists(tmpDirTo))=false then
		fso1.CreateFolder tmpDirTo
	end if
if DB_Selt = "ADD" then	
	ff = fs.GetFolder(tmpDir).Files.Count

	if ff > 0 then 
i = 0
	 Set objFolder = fs.GetFolder(tmpDir)
	 Set objFiles = objFolder.Files
	 For Each objFile in objFiles
i = i+1
		if fs.FileExists(tmpDirTo & "\\" & objFile.Name) Then
			fs.DeleteFile tmpDirTo & "\\" & objFile.Name, True
		end if
		tmpfileName=year(now) & right("0"&month(now),2) & right("0"&day(now),2) & right("0"&hour(now),2) & right("0"&minute(now),2) & right("0"&Second(now),2) & "_" & Right("000" & i,3) & ".jpg"

		fs.MoveFile tmpDir & "\\" & objFile.Name , tmpDirTo & "\\" & tmpfileName
	 Next
	end if  
end if
set fs=nothing
set fso1=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Big5" />
<title>多檔上傳</title>
<link href="uploadify214/css/default.css" rel="stylesheet" type="text/css" />
<link href="uploadify214/css/uploadify.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="uploadify214/jquery.min.js"></script>
<script type="text/javascript" src="uploadify214/swfobject.js"></script>
<script type="text/javascript" src="uploadify214/jquery.uploadify.v2.1.4.min.js"></script>
<script type="text/javascript">
function Send_document()
{
	$('#uploadify').uploadifyUpload();
}
</script>

<script type="text/javascript">
var sizelimit = '51200'; //or '5120'
$(document).ready(function() {
	$("#uploadify").uploadify({
		'uploader'       : 'uploadify214/uploadify.swf',
		'script'         : 'uploader214.asp?sId=<%=session.sessionID%>',
		'cancelImg'      : 'uploadify214/cancel.png',		
		'fileDesc'		 : 'JPG (*.jpg)',
		'fileExt'		 : '*.jpg;',
		'folder'         : '<%=application("uploadpath")%>',
		'multi'          : true,
		'sizeLimit': 1024*5000, // 100 KB
		'buttonText': 'Select Files',
		'buttonImg': 'uploadify214/SelectFile.jpeg',
		'auto':true,		
		'checkExisting':true,				
        'simUploadLimit': 1,		
		onError: function (a, b, c, d) {
         if (d.status == 404)
            alert('Could not find upload script. Use a path relative to: '+'<?= getcwd() ?>');
         else if (d.type === "HTTP")
            alert('HTTP發生錯誤 '+d.type+": "+d.status);
         else if (d.type ==="File Size")
            alert(c.name+' '+'檔案大小不得超過 5MB');
         else
            alert('發生錯誤 '+d.type+": "+d.text);
},
		onComplete		 : function(event, queueID, fileObj, response, data) {
     							var path = escape(fileObj.filePath);
							//	$('#filesUploaded').append('<div class=\'uploadifyQueueItem\'><a href='+path+' target=\'_blank\'>'+fileObj.name+'</a></div>');
							}
	});
});

function funUploadData()
{
	formIDdoc.DB_Selt.value="ADD";
	formIDdoc.submit();
}

</script>
</head>
<body>
<form id="formIDdoc" name="formIDdoc" class="form" method="post" action="default.asp">
<font size="4">上傳檔名請勿有中文、符號等</font><p>
<font size="4">請勿重覆上傳檔案</font><br>
<p><input class="text-input" name="uploadify" id="uploadify" type="file" size="20" /></p>
<div id="filesUploaded"></div>
<p id="sending" name="sending">
<input type="button" class="btn2" name="b1" value="上傳完畢" onclick="funUploadData();">
&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" name="b1" value="清除" onclick="javascript:window.location.reload()">
</p>
<input type="hidden" name="DB_Selt" value="">
</form>
</body>
</html>
