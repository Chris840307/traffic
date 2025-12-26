<!--#include File="../Common/DbUtil.asp"-->
<!--#include File="../Common/Allfunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
'AuthorityCheck(270)
'***********************************************************************************
'主資料處理
'***********************************************************************************
Dim i
Dim j	'機算需要補多少空白
Dim oRST
Dim sTypeID	'ProsecutionTypeID
Dim sFileName, sSmallPicFileName, sPicFileName, sVideoFileName, sVideoPath, sVerifyResult, sSN
Dim sSQL, sOption, sProsecutionTime, sLocation, sRadarID, sLimitSpeed, sOverSpeed, sLine, sRedLightTime, sIntervalTime
Dim sImgWebPath

'@1@ 換回 +
sFileName = replace(Request.QueryString("FileName"),"@2@", "+")
sSN = Request.QueryString("SN")

'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
	
if Conn.state and sFileName <> "" then
	j = 1
	sCarNo = ""
	
	sSQL = "select A.FileName, case when FIXEQUIPTYPE = 1 then 'Type1' when FIXEQUIPTYPE = 2 then 'Type2' when FIXEQUIPTYPE = 5 then 'Type5' when FIXEQUIPTYPE = 3 then 'Type3' else 'Type10' end FIXEQUIPTYPE, A.DirectoryName, A.ProsecutionTime, A.Location, A.ProsecutionTypeID, A.RadarID, " + _
			"A.LimitSpeed, A.OverSpeed,A.PROSECUTIONTIME, A.Line, A.RedLightTime, A.IntervalTime, A.IMAGEFILENAMEA, A.IMAGEFILENAMEB, A.VIDEOFILENAME, B.VerifyResultID, B.CarNo, B.RealCarNo, B.LPRResultID, B.MemberID " + _
			"from ProsecutionImage A, ProsecutionImageDetail B " + _
			"where A.FileName = B.FileName and A.FileName = '" + sFileName + "' and B.SN = " + sSN

	set oRST = Conn.execute(sSQL)
	if not oRST.EOF then
		do while not oRST.EOF
			'SmallPicFileName
			sImgWebPath = toImageDir(oRST("PROSECUTIONTIME"))
			'***********員警上傳違規影像Kevin*************
			if trim(oRST("FIXEQUIPTYPE").value)="Type3" then	
				bPicWebPath=""
				RealFileName=right(oRST("FileName").value,4)
				WebPicPathTmp=left(oRST("FileName").value,14)
				bPicWebPath=sImgWebPath & "upload/" & WebPicPathTmp & "/" & RealFileName & ".jpg"
				sSmallPicFileName=bPicWebPath
			else
				'sSmallPicFileName = replace(sImgWebPath & replace(oRST("DIRECTORYNAME"),"\","/") &  "/" & lcase(oRST("IMAGEFILENAMEA").value),"//","/")
				sSmallPicFileName = replace( replace(oRST("DIRECTORYNAME"),"\","/") &  "/" & lcase(oRST("IMAGEFILENAMEA").value),"//","/")
			end if
			'********************************************
			'VideoFileName
			sVideoFileName = oRST("VIDEOFILENAME").value
			sVideoPath = sImgWebPath & oRST("FIXEQUIPTYPE").value & "/" & oRST("FileName").value & "/" & oRST("VIDEOFILENAME").value
		
			'PicFileName
'			if trim(oRST("FIXEQUIPTYPE").value)="Type1" or trim(oRST("FIXEQUIPTYPE").value)="Type2" or trim(oRST("FIXEQUIPTYPE").value)="Type5" then
				sPicFileName = oRST("IMAGEFILENAMEA").value
'			else
'				sPicFileName = right(oRST("FileName").value,4)&".jpg"
'			end if
						
			'車牌號碼
			sCarNo = oRST("CarNo").value & "<br>"

			'實際車牌號碼
			sRealCarNo = oRST("RealCarNo").value & "<br>"

			'辨識結果
			sLprResult = ""
			if trim(oRST("LPRResultID").value)="0" then
				sLprResult = "成功"
			end if

			'審核狀態
			if oRST("VerifyResultID").value = "1" or isnull(oRST("VerifyResultID").value) then
				sVerifyResultID = "未判別"
			else
				if oRST("VerifyResultID").value = "0" then
					sVerifyResultID = "有效"
				else
					sVerifyResultID = "無效"
				end if
			end if
			
			'審核人員
			sMem = ""
			if trim(oRST("MemberID").value)<>"" and not isnull(oRST("MemberID").value) then
				strMem="select chName from MemberData where MemberId="&trim(oRST("MemberID").value)
				set rsMem=conn.execute(strMem)
				if not rsMem.eof then
					sMem = trim(rsMem("chName"))
				end if
				rsMem.close
				set rsMem=nothing
			end if
				
			'按紐功能
			sMenuButton = ""
		
			'違規類型
			sTypeID = oRST("ProsecutionTypeID").value
			select case(sTypeID)
				case "R":	sTypeID = "紅燈"
							sMenuButton = "<input type=""button"" value=""檢視動態影像圖""><input type=""button"" value=""檢視詳細資料"">"
						'***********員警上傳違規影像Kevin*************
						if trim(oRST("FIXEQUIPTYPE").value)="Type3" then	
							sPicWebPath=""
							sRealFileName=right(replace(oRST("ImageFileNameB"),".jpg",""),4)
							sWebPicPathTmp=left(replace(oRST("ImageFileNameB"),".jpg",""),14)
							sPicWebPath=sImgWebPath & "upload/" & sWebPicPathTmp & "/" & sRealFileName & ".jpg"

							sSmallPicFileNameB=sPicWebPath
						else
							'sSmallPicFileNameB=sImgWebPath & oRST("DIRECTORYNAME").value & "/" & lcase(oRST("IMAGEFILENAMEB").value)
							sSmallPicFileNameB= oRST("DIRECTORYNAME").value & "/" & lcase(oRST("IMAGEFILENAMEB").value)
						end if
						'*********************************************
							sPicFileNameB = oRST("IMAGEFILENAMEB").value
				case "S":	sTypeID = "超速"
							sMenuButton = "<input type=""button"" onClick=""OpenPic('" & sImgWebPath & oRST("FileName").value & "/" & oRST("IMAGEFILENAMEA").value & "')"" value=""檢視縮圖""><input type=""button"" onClick=""OpenDetail('" & oRST("FileName").value & "')"" value=""檢視詳細資料"">"
				case "SR":	sTypeID = "超速、闖紅燈"
							sMenuButton = "<input type=""button"" onClick=""OpenPic('" & sImgWebPath & oRST("FileName").value & "/" & replace(oRST("IMAGEFILENAMEA").value, ".jpg", "_small.jpg") & "')"" value=""檢視縮圖""><input type=""button"" onClick=""OpenPic('" & sImgWebPath & oRST("DirectoryName").value & oRST("FileName").value & "')"" value=""檢視原圖""><input type=""button"" onClick=""OpenDetail('" & oRST("FileName").value & "')"" value=""檢視詳細資料"">"
				case "L":	sTypeID = "越線"
				case "T":	sTypeID = "緊跟著前車駕駛"
				case "U":	sTypeID = "左轉"
				defaule:	sTypeID = "未知類型"
			end select		
			
			'ProsecutionTime
		if trim(oRST("ProsecutionTime").value)<>"" and not isnull(oRST("ProsecutionTime").value) then
			sProsecutionTime = gInitDT(oRST("ProsecutionTime").value) & " " & TimeValue(oRST("ProsecutionTime").value)
		else
			sProsecutionTime = ""
		end if
			'Location
			sLocation = oRST("Location").value
			
			'RadarID
			sRadarID = oRST("RadarID").value
			
			'LimitSpeed
			sLimitSpeed = oRST("LimitSpeed").value
			
			'OverSpeed
			sOverSpeed = oRST("OverSpeed").value
			
			'Line
			sLine = oRST("Line").value
			
			'RedLightTime
			sRedLightTime = oRST("RedLightTime").value
			
			'IntervalTime
			sIntervalTime = oRST("IntervalTime").value
			
			j = j + 1
			oRST.movenext
		loop
	end if
	oRST.close
	set oRST = nothing
end if

'補空白
for i = j to 10
	sTR = sTR + "<tr bgcolor=""#FFFFFF"">" + _
				"	<td height=""25""></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"</tr>"
next
'***********************************************************************************

if isObject(Conn) then
    if Conn.state then    	  
        Conn.close
    end if
    Set Conn = Nothing
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title>數位固定桿處理系統</title>
	<!--#include File="../Common/css.txt"-->
</head>

<body topmargin="0" leftmargin="0" onfocus="ClosePic();">
<table width="100%" border="0">
  <tr>
    <td height="22" bgcolor="#1BF5FF"><span class="font12">數位固定桿處理系統-詳細資料</span></td>
  </tr>
  <tr>
    <td height="100%" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1">
      <tr bgcolor="#FFFFFF">
        <td width="17%" nowrap height="23" bgcolor="#1BF5FF"><span class="font12">舉發時間</span></td>
        <td width="20%" height="23"><span class="font12"><%=sProsecutionTime%></span></td>
        <td width="63%" rowspan="16"><img src="<%=sSmallPicFileName%>" width="700" height="500"></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">舉發地點</span></td>
        <td height="23"><span class="font12"><%=sLocation%></span></td>
        </tr>
<!--  <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">雷達認證</span></td>
        <td height="23"><span class="font12"><%=sRadarID%></span></td>
        </tr>-->
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">舉發類型</span></td>
        <td height="23"><span class="font12"><%=sTypeID%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><%If sys_City<>"南投縣" then%><span class="font12">限定車速</span><%End if%>&nbsp;</td>
        <td height="23"><%If sys_City<>"南投縣" then%><span class="font12"><%=sLimitSpeed%></span><%End if%>&nbsp;</td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><%If sys_City<>"南投縣" then%><span class="font12">違規車速</span><%End if%>&nbsp;</td>
        <td height="23"><%If sys_City<>"南投縣" then%><span class="font12"><%=sOverSpeed%></span><%End if%>&nbsp;</td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">違規車道</span></td>
        <td height="23"><span class="font12"><%=sLine%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">紅燈亮起時間</span></td>
        <td height="23"><span class="font12"><%=sRedLightTime%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">違規舉發相片拍攝間隔時間</span></td>
        <td height="23"><span class="font12"><%=sIntervalTime%></span></td>
        </tr>
 <!--<tr bgcolor="#FFFFFF">
        <td height="23" nowarp bgcolor="#1BF5FF"><span class="font12">辨識車牌號碼</span></td>
        <td height="23"><span class="font12"><%=sCarNo%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">實際車牌號碼</span></td>
        <td height="23"><span class="font12"><%=sRealCarNo%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">辨識結果</span></td>
        <td height="23"><span class="font12"><%=sLprResult%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="23" bgcolor="#1BF5FF"><span class="font12">相片有效性</span></td>
        <td height="23"><span class="font12"><%=sVerifyResultID%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="30" bgcolor="#1BF5FF"><span class="font12">審核人員</span></td>
        <td height="30"><span class="font12"><%=sMem%></span></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td height="29" bgcolor="#1BF5FF"><span class="font12">違規動態影像</span></td>
        <td height="29"><a href="#" onClick="OpenPic('<%=sVideoPath%>')" class="font12"><%=sVideoFileName%></a></td>
      </tr>-->
      <tr bgcolor="#FFFFFF">
        <td height="49" bgcolor="#1BF5FF"><span class="font12">違規影像檔名稱</span></td>
        <td height="49">
			<%
			if not isnull(sVideoFileName) and sVideoFileName<>"" then
			%>
			<a href="#" onClick="OpenPic('<%=sVideoPath%>')" class="font12"><%=sVideoFileName%></a><br>
			<%
			end if
			%>
			<a href="#" onClick="OpenPic('<%=replace(sSmallPicFileName,"\","/")%>')" class="font12"><%=sPicFileName%></a><%
			If trim(sTypeID) = "紅燈" then
				response.write "<BR><a href=""#"" onClick=""OpenPic('"&replace(sSmallPicFileNameB,"\","/")&"')"" class=""font12"">"&sPicFileNameB&"</a>"
			end if%>
		</td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#1BF5FF"><p align="center" class="style1"><a href="sssss">
      </a>
        <input type="button" onclick="javascript:window.close();" value="關 閉">
    </p>    </td>
  </tr>
</table>
<script language="javascript">
	var OldObj = '';
	
	//置中 function
	function centerPos(size, type) {
	    switch(type) {
	        case 0:   //Top position
	            return (parseInt(window.screen.height) - size) / 2;
	            break;
	        case 1:   //Left position
	            return (parseInt(window.screen.width) - size) / 2;
	            break;
	        default:
	            alert('centerPos() : Type value error!!');
	    }
	}
	
	function OpenWindow(Width, Height, URL, WinName){
	    var nWidth = Width;
	    var nHeight = Height;
	    var sURL = URL;
	    var nTop = centerPos(nHeight,0);
	    var nLeft = centerPos(nWidth,1);
	    var sWinSize = "left=" + nLeft.toString(10) + ",top=" + nTop.toString(10) + ",width=" + nWidth.toString(10) + ",height=" + nHeight.toString(10);
	    var sWinStatus = "menubar=0,toolbar=0,scrollbars=1,resizable=1,status=0";
	    var sWinName = WinName;
	    OldObj = window.open(sURL,sWinName,sWinSize + "," + sWinStatus);
	}

	//開啟檢視圖
	function OpenPic(FileName){
		OpenWindow(800, 600, FileName, 'MyPic');
	}

	//關閉檢是圖
	function ClosePic(){
		try{
			if(OldObj!=''){
				OldObj.close();
				OldObj = '';
			}
		}
		catch(e){
		}
	}
</script>
</body>
</html>
