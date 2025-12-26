<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<title>違規影像</title>
<%

%>

<style type="text/css">
<!--
.style1 {
	font-size: 28px;
	font-weight: bold;
	color: #0000FF;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
	  <span class="style1">違規影像上按滑鼠左鍵可將該圖片放大</span>
<%
 '違規影像資料
			strImage="select b.FileName,b.Sn,a.PROSECUTIONTIME,a.FIXEQUIPTYPE,a.FileName,a.DIRECTORYNAME" &_
			",a.IMAGEFILENAMEA,a.IMAGEFILENAMEB" &_
			" from ProsecutionImage a,ProsecutionImageDetail b where b.BillSn="&trim(request("ImgSN")) &_
			" and a.FileName=b.FileName"
			set rsImage=conn.execute(strImage)
			if not rsImage.eof then
				sSmallPicFileName=""
				sSmallPicFileName2=""
				sImgWebPath = toImageDir(rsImage("PROSECUTIONTIME"))
				if trim(rsImage("FIXEQUIPTYPE"))="3" then	
					bPicWebPath=""
					RealFileName=right(rsImage("FileName"),4)
					WebPicPathTmp=left(rsImage("FileName"),14)
					bPicWebPath=sImgWebPath & "upload/" & WebPicPathTmp & "/" & RealFileName & ".jpg"
					sSmallPicFileName=bPicWebPath
				else
					'sSmallPicFileName = sImgWebPath & rsImage("DIRECTORYNAME") &  "/" & lcase(rsImage("IMAGEFILENAMEA"))
					sSmallPicFileName = rsImage("DIRECTORYNAME") &  "/" & lcase(rsImage("IMAGEFILENAMEA"))
				end if
				response.write "<br><FONT SIZE='5'>"&rsImage("PROSECUTIONTIME")&"</FONT><br>"
%>
			<img src="<%=sSmallPicFileName%>" name="imgA1" width="800" onclick="OpenImg('imgA1')">
<%				if trim(rsImage("IMAGEFILENAMEB"))<>"" then
				'sSmallPicFileName2 = sImgWebPath & rsImage("DIRECTORYNAME") &  "/" & lcase(rsImage("IMAGEFILENAMEB"))
				sSmallPicFileName2 = rsImage("DIRECTORYNAME") &  "/" & lcase(rsImage("IMAGEFILENAMEB"))
%>
			<img src="<%=sSmallPicFileName2%>" name="imgB1" width="800" onclick="OpenImg('imgA1')">
<%				end if
			else
				imgno=0
				strImage2="select b.FileName,b.Sn,a.PROSECUTIONTIME,a.FIXEQUIPTYPE,a.FileName,a.DIRECTORYNAME" &_
					",a.IMAGEFILENAMEA,a.IMAGEFILENAMEB" &_
					" from ProsecutionImage a,ProsecutionImageDetail b" &_
					" where a.FileName=b.FileName" &_
					" and ProsecutionTime between TO_DATE('"&year(trim(request("illdate")))&"/"&Month(trim(request("illdate")))&"/"&day(trim(request("illdate")))&" "&Hour(trim(request("illdate")))&":"&Minute(trim(request("illdate")))&":00','YYYY/MM/DD/HH24/MI/SS') " &_
					" and TO_DATE('"&year(trim(request("illdate")))&"/"&Month(trim(request("illdate")))&"/"&day(trim(request("illdate")))&" "&Hour(trim(request("illdate")))&":"&Minute(trim(request("illdate")))&":59','YYYY/MM/DD/HH24/MI/SS') order by PROSECUTIONTIME"
				'Location='"&trim(rs1("IllegalAddress"))&"'
				'response.write strImage2
				set rsImage2=conn.execute(strImage2)
				if not rsImage2.eof then
					While Not rsImage2.Eof
						imgno=imgno+1
						sSmallPicFileName=""
						sSmallPicFileName2=""
						sImgWebPath = toImageDir(rsImage2("PROSECUTIONTIME"))
						if trim(rsImage2("FIXEQUIPTYPE"))="3" then	
							bPicWebPath=""
							RealFileName=right(rsImage2("FileName"),4)
							WebPicPathTmp=left(rsImage2("FileName"),14)
							bPicWebPath=sImgWebPath & "upload/" & WebPicPathTmp & "/" & RealFileName & ".jpg"
							sSmallPicFileName=bPicWebPath
						else
							'sSmallPicFileName = sImgWebPath & rsImage2("DIRECTORYNAME") &  "/" & lcase(rsImage2("IMAGEFILENAMEA"))
							sSmallPicFileName = rsImage2("DIRECTORYNAME") &  "/" & lcase(rsImage2("IMAGEFILENAMEA"))
						end if
						response.write "<br><FONT SIZE='5'>"&rsImage2("PROSECUTIONTIME")&"</FONT><br>"
%>
			<img src="<%=sSmallPicFileName%>" name="imgA<%=imgno%>" width="800" onclick="OpenImg('imgA<%=imgno%>')">
<%						if trim(rsImage2("IMAGEFILENAMEB"))<>"" then
							sSmallPicFileName2 = sImgWebPath & rsImage2("DIRECTORYNAME") &  "/" & lcase(rsImage2("IMAGEFILENAMEB"))
%>
			<img src="<%=sSmallPicFileName2%>" name="imgB<%=imgno%>" width="800" onclick="OpenImg('imgB<%=imgno%>')">
<%	
						end if
					rsImage2.MoveNext
					Wend
				else
					response.write "&nbsp;"
				end if
				rsImage2.close
				set rsImage2=nothing
			end if
			rsImage.close
			set rsImage=nothing
%>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function OpenImg(ImgNo){
	
	if (eval("myForm."+ImgNo).width=="800"){
		eval("myForm."+ImgNo).width="1200";
	}else{
		eval("myForm."+ImgNo).width="800";
	}
}
</script>
</html>
