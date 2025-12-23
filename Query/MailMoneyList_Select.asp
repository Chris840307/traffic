<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title><%
if trim(request("MailSendType"))="S" then
	response.write "列印大宗函件"
else
	response.write "列印郵費單"
end if
%></title>
<%
tmpSQL=request("SQLstr")

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return select_street();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td>選擇郵資</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="1" >掛號 28 元
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="2">郵簡 26 元
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="3" <%response.write "checked"%>>自行輸入&nbsp;
					<input type="text" name="MailMoneyValue" value="<%
						MailMoney=0
					if sys_City="台中市" then
						MailMoney="28"
					Elseif sys_City="台南市" then
						MailMoney="39"
					Else
						strMailMoney="select Value from ApConfigure where ID=28"
						set rsMailMoney=conn.execute(strMailMoney) 
						if not rsMailMoney.eof then
							MailMoney=cint(rsMailMoney("Value"))
						end if
						rsMailMoney.close
						set rsMailMoney=nothing
					End if
						response.write MailMoney
					%>" size="5" onkeyup="value=value.replace(/[^\d]/g,'')"> 元
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="4">不計算郵資
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center">
					<%if sys_City="宜蘭縣" then %>
					<input type="button" value="列  印" onclick="funMailListCity();">
					
					<%elseif sys_City="台中市" or sys_City="高雄縣" or sys_City="嘉義市" then %>
					<input type="button" value="列印(A4)" onclick="funMailListCity();">
					<%elseIf sys_City="南投縣" then%>
					<input type="button" value="列印(回執聯式舉發單專用)&nbsp; &nbsp; " onclick="funMailListCity();">
					<%else%>
					<input type="button" value="列  印" onclick="funMailListCity();">
					<%end if%>
					

				<%if trim(request("MailSendType"))="S" then%>
					<%if sys_City="台中市" or sys_City="高雄縣" or sys_City="高雄市" then %>
					<input type="button" value="清冊套印" onclick="funMailListCity_b();">
					<%elseif sys_City="嘉義市" then %>
					<input type="button" value="列印(連續報表紙)" onclick="funMailListCity_b();">
					<%end if%>
					<%if sys_City="高雄縣" then %>
					<input type="button" value="監理站" onclick="funMailListStation();">
					<%end if%>
					<%if sys_City="南投縣" then %>
					<br>
					<input type="button" value="列印(送達証書式舉發單專用)" onclick="funMailListCity_NT();">
					<%end if%>
					<%If sys_City="台中市" then%>
					<br>
					郵遞區號<input type="text" name="StartZip" value="" size="6" maxlength="4"> ~ 
					<input type="text" name="EndZip" value="" size="6" maxlength="4">
					<br>
					<input type="button" value="清冊套印(郵遞區號排序)" onclick="funMailListCity_TC();">
					<%End If %>
					<%if sys_City="台東縣" then %>
					<input type="button" value="肇事案件" onclick="funMailListCity_TD();">
					<%end if%>
				<%else%>
					<%if sys_City="宜蘭縣" then %>
					<input type="button" value="列  印(件數x郵費)" onclick="funMailListCity_2();" style=" width: 140px;">
					<%end if%>
				<%end if%>
				
				<%if trim(request("MailSendType"))="S" then%>
					<%if sys_City="南投縣" Or sys_City="保二總隊四大隊二中隊" then %>
					<br>
					<input type="button" value="監理站" onclick="funMailListStation();">
					<%end if%>
				<%end if%>

				<%if sys_City="宜蘭縣" then %>
				逕舉案件第一次郵寄列印大宗清冊前，請先列印舉發單
				<%End If %>
				</td>
			</tr>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	winopen.focus();
	return win;
}
function funMailListCity_b(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			MailMoneyTypeID="1";
			MailMoneyValueTmp2="28";
		}else if (myForm.MailMoneyType[1].checked==true){
			MailMoneyTypeID="2";
			MailMoneyValueTmp2="26";
		}else if (myForm.MailMoneyType[2].checked==true){
			MailMoneyTypeID="3";
			MailMoneyValueTmp2=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			MailMoneyTypeID="4";
			MailMoneyValueTmp2="";
		}
	
	<%if trim(request("MailSendType"))="S" then%>
		<%if sys_City="台中市" or sys_City="高雄市" Or sys_City=ApconfigureCityName then %>	
			UrlStr="MailReportList_TaiChung.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
			newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		<%elseif sys_City="嘉義市" then %>	
			UrlStr="MailReportList_ChiayiCity.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
			newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		<%elseif sys_City="高雄縣" then %>	
			UrlStr="MailReportList_Ka.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
			newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		<%end if%>
	<%end if%>
	window.close();
	}


}
<%If sys_City="台中市" then%>
function funMailListCity_TC(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			MailMoneyTypeID="1";
			MailMoneyValueTmp2="28";
		}else if (myForm.MailMoneyType[1].checked==true){
			MailMoneyTypeID="2";
			MailMoneyValueTmp2="26";
		}else if (myForm.MailMoneyType[2].checked==true){
			MailMoneyTypeID="3";
			MailMoneyValueTmp2=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			MailMoneyTypeID="4";
			MailMoneyValueTmp2="";
		}
	
	<%if trim(request("MailSendType"))="S" then%>
			UrlStr="MailReportList_TaiChung_Order.asp?MailMoneyType="+MailMoneyTypeID+"&StartZip="+myForm.StartZip.value+"&EndZip="+myForm.EndZip.value+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
			newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");

	<%end if%>
	window.close();
	}
}
<%end if %>
function funMailListStation(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			MailMoneyTypeID="1";
			MailMoneyValueTmp2="28";
		}else if (myForm.MailMoneyType[1].checked==true){
			MailMoneyTypeID="2";
			MailMoneyValueTmp2="26";
		}else if (myForm.MailMoneyType[2].checked==true){
			MailMoneyTypeID="3";
			MailMoneyValueTmp2=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			MailMoneyTypeID="4";
			MailMoneyValueTmp2="";
		}
        UrlStr="MailReportList_Station.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		window.close();
	}

}

function funMailListCity(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			MailMoneyTypeID="1";
			MailMoneyValueTmp2="28";
		}else if (myForm.MailMoneyType[1].checked==true){
			MailMoneyTypeID="2";
			MailMoneyValueTmp2="26";
		}else if (myForm.MailMoneyType[2].checked==true){
			MailMoneyTypeID="3";
			MailMoneyValueTmp2=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			MailMoneyTypeID="4";
			MailMoneyValueTmp2="";
		}
	
	<%if trim(request("MailSendType"))="S" then%>
		<%if sys_City="彰化縣" then '彰化只印一份用複寫%>
			UrlStr="MailReportList_CH.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
			newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		<%elseif sys_City="高雄縣" then '高雄縣要套印%>
			UrlStr="MailReportList.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
			newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		<%else%>
			UrlStr="MailReportList.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
			newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		<%end if%>
	<%elseif trim(request("MailSendType"))="SM" then%>
		UrlStr="MailReportList_ML.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList_ML_3",1000,700,0,0,"yes","yes","yes","no");
	<%elseif trim(request("MailSendType"))="SM_A" then%>
		UrlStr="MailReportList_ML.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SendUnit=1&SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList_ML_3",1000,700,0,0,"yes","yes","yes","no");
	<%elseif trim(request("MailSendType"))="SM_B" then%>
		UrlStr="MailReportList_ML.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SendUnit=2&SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList_ML_3",1000,700,0,0,"yes","yes","yes","no");
	<%else%>
	  	<%if sys_City="台東縣" then '台東縣格式與其他縣市不同%>
    		UrlStr="MailMoneyList_TaiTung_Excel.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp+"&SQLstr=<%=tmpSQL%>";
	    	newWin(UrlStr,"MailMoneyList_2",1000,700,0,0,"yes","yes","yes","no");
        <%else%>
            UrlStr="MailMoneyList_Excel.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp+"&SQLstr=<%=tmpSQL%>";
	    	newWin(UrlStr,"MailMoneyList_2",1000,700,0,0,"yes","yes","yes","no");
        <%end if%> 

	<%end if%>
	window.close();
	}
}

function funMailListCity_2(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			MailMoneyTypeID="1";
			MailMoneyValueTmp2="28";
		}else if (myForm.MailMoneyType[1].checked==true){
			MailMoneyTypeID="2";
			MailMoneyValueTmp2="26";
		}else if (myForm.MailMoneyType[2].checked==true){
			MailMoneyTypeID="3";
			MailMoneyValueTmp2=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			MailMoneyTypeID="4";
			MailMoneyValueTmp2="";
		}
	
        UrlStr="MailMoneyList_YiLan_Excel.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp+"&SQLstr=<%=tmpSQL%>";
	    	newWin(UrlStr,"MailMoneyList_2",1000,700,0,0,"yes","yes","yes","no");

	window.close();
	}
}

function funMailListCity_NT(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			MailMoneyTypeID="1";
			MailMoneyValueTmp2="28";
		}else if (myForm.MailMoneyType[1].checked==true){
			MailMoneyTypeID="2";
			MailMoneyValueTmp2="26";
		}else if (myForm.MailMoneyType[2].checked==true){
			MailMoneyTypeID="3";
			MailMoneyValueTmp2=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			MailMoneyTypeID="4";
			MailMoneyValueTmp2="";
		}
	
		UrlStr="MailReportList_NT.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList_3",1000,700,0,0,"yes","yes","yes","no");
		window.close();
	}
}

function funMailListCity_TD(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			MailMoneyTypeID="1";
			MailMoneyValueTmp2="28";
		}else if (myForm.MailMoneyType[1].checked==true){
			MailMoneyTypeID="2";
			MailMoneyValueTmp2="26";
		}else if (myForm.MailMoneyType[2].checked==true){
			MailMoneyTypeID="3";
			MailMoneyValueTmp2=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			MailMoneyTypeID="4";
			MailMoneyValueTmp2="";
		}
	
	
		UrlStr="MailReportList.asp?MailMoneyType="+MailMoneyTypeID+"&MailMoneyValue="+MailMoneyValueTmp2+"&SendUnit=2&CitySpecFlag=TD01&SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList_ML_3",1000,700,0,0,"yes","yes","yes","no");

		window.close();
	}

}
</script>
</html>
